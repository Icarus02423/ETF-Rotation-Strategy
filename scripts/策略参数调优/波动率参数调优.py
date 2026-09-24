#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
波动率过滤参数调优。

每次为一个更新频率、过滤状态、聚类阈值、趋势窗口和错峰天数组合，
遍历波动率收益窗口与波动率保留比例。程序复用正式交易回测的交易函数，
只输出两个趋势公式、两种成交方式对应的四张累计收益率热力图。
组合过滤状态使用固定的BIAS参数，因此结果是条件最优，不是两类参数的联合最优。
"""

from __future__ import annotations

import csv
import importlib.util
import math
import statistics
import sys
from bisect import bisect_right
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from types import ModuleType
from typing import Mapping, Sequence

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Rectangle
from matplotlib.ticker import PercentFormatter


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from 实验配置 import (
    CORRELATION_THRESHOLDS,
    FILTER_PROFILE_SETTINGS,
    STAGGER_DAYS,
    TREND_WINDOWS,
    UPDATE_FREQUENCIES,
)
from 输出路径 import factor_dir, index_price_dir, tuning_result_dir

BACKTEST_FILE = PROJECT_ROOT / "scripts" / "ETF趋势策略回测" / "交易策略回测.py"

# ============================= 策略基准参数 =============================
UPDATE_FREQUENCY = "daily"
FILTER_PROFILE = "volatility_only"
CLUSTER_CORRELATION_THRESHOLD = 0.9
TREND_WINDOW = 20
BENCHMARK_CODE = "000510.CSI"
SCORE_METHODS_TO_RUN = ("return_r2", "return_vol")
TOP_PERCENT = 0.10
ACCOUNT_REBALANCE_INTERVAL = 1
ACCOUNT_COUNT = ACCOUNT_REBALANCE_INTERVAL
VOL_FILTER_ENABLED = True
VOL_FILTER_MODE = "high"
VOL_RETURN_DAYS = 18
VOL_KEEP_TOP_RATIO = 0.55
BIAS_MODE = "none"
BIAS_WINDOW = 36
BIAS_LOWER = -0.05
BIAS_UPPER = +0.09
MAX_PRICE_STALENESS_CALENDAR_DAYS = 7
ACCOUNT_VARIANT_DIR = f"staggered_{ACCOUNT_REBALANCE_INTERVAL}d"
TRANSACTION_COST_RATE = 0.001
ANNUAL_TRADING_DAYS = 252
ANNUAL_RISK_FREE_RATE = 0.015
INITIAL_NAV = 1.0
CAPACITY_DAILY_AMOUNT_RATIO = 0.10
CAPACITY_DESCENDING_QUANTILE = 0.95
# =======================================================================

# ============================== 调优参数 ==============================
# 波动率使用日收益率的样本标准差，因此窗口至少需要2个日收益。
# 100%保留全部指数，可作为“不使用波动率过滤”的对照组。
VOL_WINDOWS = tuple(range(2, 61))
VOL_KEEP_RATIOS = tuple(value / 100.0 for value in range(5, 101, 5))
# ======================================================================

SCORE_METHODS = SCORE_METHODS_TO_RUN
TRADE_MODES = ("close", "next_day_vwap")
ALLOWED_FILTER_PROFILES = ("volatility_only", "bias_and_volatility")

OUTPUT_FILES = {
    ("return_r2", "close"): "return_r2_close_total_return_heatmap.png",
    ("return_r2", "next_day_vwap"): "return_r2_next_day_vwap_total_return_heatmap.png",
    ("return_vol", "close"): "return_vol_close_total_return_heatmap.png",
    ("return_vol", "next_day_vwap"): "return_vol_next_day_vwap_total_return_heatmap.png",
}


@dataclass(frozen=True)
class VolatilityTuningSettings:
    update_frequency: str
    filter_profile: str
    correlation_threshold: float
    trend_window: int
    stagger_days: int
    factor_directory: Path
    index_price_directory: Path
    output_directory: Path
    bias_mode: str
    vol_filter_enabled: bool


@dataclass(frozen=True)
class FilterMetrics:
    volatilities: tuple[float, ...]
    bias: float | None


@dataclass(frozen=True)
class MethodContext:
    trading_dates: tuple[date, ...]
    benchmark_end_date: date


def validate_tuning_dimensions(
    update_frequency: str,
    filter_profile: str,
    correlation_threshold: float,
    trend_window: int,
    stagger_days: int,
) -> None:
    """在计算任何输入输出路径前，先验证公共实验维度。"""

    if update_frequency not in UPDATE_FREQUENCIES:
        raise ValueError(
            f"update_frequency只能是：{', '.join(UPDATE_FREQUENCIES)}"
        )
    if filter_profile not in ALLOWED_FILTER_PROFILES:
        raise ValueError(
            "波动率参数调优的filter_profile只能是："
            + "、".join(ALLOWED_FILTER_PROFILES)
        )
    if not any(
        math.isclose(correlation_threshold, allowed, abs_tol=1e-12)
        for allowed in CORRELATION_THRESHOLDS
    ):
        raise ValueError(
            f"correlation_threshold只能是：{CORRELATION_THRESHOLDS}"
        )
    if trend_window not in TREND_WINDOWS:
        raise ValueError(f"trend_window只能是：{TREND_WINDOWS}")
    if stagger_days not in STAGGER_DAYS:
        raise ValueError(f"stagger_days只能是：{STAGGER_DAYS}")


def build_tuning_settings(
    update_frequency: str,
    filter_profile: str,
    correlation_threshold: float,
    trend_window: int,
    stagger_days: int,
) -> VolatilityTuningSettings:
    validate_tuning_dimensions(
        update_frequency,
        filter_profile,
        correlation_threshold,
        trend_window,
        stagger_days,
    )
    profile_settings = FILTER_PROFILE_SETTINGS[filter_profile]
    bias_mode = profile_settings["bias_mode"]
    vol_filter_enabled = profile_settings["vol_filter_enabled"]
    if bias_mode not in {"none", "upper"}:
        raise ValueError("波动率参数调优只支持关闭BIAS或启用BIAS上限")
    if vol_filter_enabled is not True:
        raise ValueError("波动率参数调优要求过滤状态启用波动率过滤")
    return VolatilityTuningSettings(
        update_frequency=update_frequency,
        filter_profile=filter_profile,
        correlation_threshold=correlation_threshold,
        trend_window=trend_window,
        stagger_days=stagger_days,
        factor_directory=factor_dir(
            update_frequency,
            correlation_threshold,
            trend_window,
        ),
        index_price_directory=index_price_dir(
            update_frequency,
            correlation_threshold,
        ),
        output_directory=tuning_result_dir(
            update_frequency,
            filter_profile,
            "volatility_parameter_grid",
            correlation_threshold,
            trend_window,
            stagger_days,
        ),
        bias_mode=bias_mode,
        vol_filter_enabled=vol_filter_enabled,
    )


def configure_backtest_module(
    backtest: ModuleType,
    settings: VolatilityTuningSettings,
) -> None:
    """显式把本次调优设置注入隔离加载的正式回测模块。"""

    parameter_names = (
        "BENCHMARK_CODE",
        "SCORE_METHODS_TO_RUN",
        "TOP_PERCENT",
        "VOL_FILTER_MODE",
        "VOL_RETURN_DAYS",
        "VOL_KEEP_TOP_RATIO",
        "BIAS_WINDOW",
        "BIAS_LOWER",
        "BIAS_UPPER",
        "MAX_PRICE_STALENESS_CALENDAR_DAYS",
        "ACCOUNT_VARIANT_DIR",
        "TRANSACTION_COST_RATE",
        "ANNUAL_TRADING_DAYS",
        "ANNUAL_RISK_FREE_RATE",
        "INITIAL_NAV",
        "CAPACITY_DAILY_AMOUNT_RATIO",
        "CAPACITY_DESCENDING_QUANTILE",
    )
    for name in parameter_names:
        setattr(backtest, name, globals()[name])

    backtest.UPDATE_FREQUENCY = settings.update_frequency
    backtest.FILTER_PROFILE = settings.filter_profile
    backtest.CLUSTER_CORRELATION_THRESHOLD = settings.correlation_threshold
    backtest.TREND_WINDOW = settings.trend_window
    backtest.ACCOUNT_REBALANCE_INTERVAL = settings.stagger_days
    backtest.ACCOUNT_COUNT = settings.stagger_days
    backtest.ACCOUNT_VARIANT_DIR = f"staggered_{settings.stagger_days}d"
    backtest.BIAS_MODE = settings.bias_mode
    backtest.VOL_FILTER_ENABLED = settings.vol_filter_enabled
    backtest.FACTOR_DIR = settings.factor_directory
    backtest.INDEX_RAW_DATA_DIR = settings.index_price_directory
    backtest.ETF_DATA_FILE = PROJECT_ROOT / "outputs" / "etf_data" / "etf_data.csv"
    backtest.BENCHMARK_DIR = PROJECT_ROOT / "outputs" / "benchmark_data"
    # 调优一次计算两个趋势因子，且自己管理输出；不使用正式回测的
    # TREND_FACTOR、DEFAULT_EXPERIMENT_CASE和BACKTEST_DIR。


def load_backtest_module(settings: VolatilityTuningSettings) -> ModuleType:
    if not BACKTEST_FILE.exists():
        raise FileNotFoundError(f"找不到正式交易回测脚本：{BACKTEST_FILE}")
    module_name = "etf_trend_backtest_for_volatility_tuning"
    spec = importlib.util.spec_from_file_location(module_name, BACKTEST_FILE)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载正式交易回测脚本：{BACKTEST_FILE}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    configure_backtest_module(module, settings)
    return module


def validate_grid(
    backtest: ModuleType,
    settings: VolatilityTuningSettings,
) -> None:
    backtest.validate_parameters()
    if (
        backtest.VOL_FILTER_ENABLED is not settings.vol_filter_enabled
        or not backtest.VOL_FILTER_ENABLED
    ):
        raise ValueError("波动率参数调优要求正式回测启用波动率过滤")
    if backtest.BIAS_MODE != settings.bias_mode:
        raise ValueError("正式回测的BIAS过滤模式与本次调优不一致")
    if backtest.FACTOR_DIR != settings.factor_directory:
        raise ValueError("正式回测的因子输入目录与本次调优不一致")
    if backtest.INDEX_RAW_DATA_DIR != settings.index_price_directory:
        raise ValueError("正式回测的指数行情目录与本次调优不一致")
    if (
        backtest.ACCOUNT_COUNT != settings.stagger_days
        or backtest.ACCOUNT_REBALANCE_INTERVAL != settings.stagger_days
    ):
        raise ValueError("正式回测的错峰账户设置与本次调优不一致")
    if tuple(backtest.SCORE_METHODS_TO_RUN) != SCORE_METHODS:
        raise ValueError(
            "正式回测需同时启用return_r2和return_vol，"
            f"当前为{backtest.SCORE_METHODS_TO_RUN}"
        )
    if not VOL_WINDOWS or any(
        not isinstance(window, int) or window < 2 for window in VOL_WINDOWS
    ):
        raise ValueError("VOL_WINDOWS必须由不小于2的整数构成")
    if tuple(sorted(set(VOL_WINDOWS))) != VOL_WINDOWS:
        raise ValueError("VOL_WINDOWS必须严格递增且不能重复")
    if not VOL_KEEP_RATIOS or any(
        not math.isfinite(ratio) or ratio <= 0 or ratio > 1
        for ratio in VOL_KEEP_RATIOS
    ):
        raise ValueError("VOL_KEEP_RATIOS必须位于(0, 1]区间")
    if tuple(sorted(set(VOL_KEEP_RATIOS))) != VOL_KEEP_RATIOS:
        raise ValueError("VOL_KEEP_RATIOS必须严格递增且不能重复")


def load_factor_universes(
    backtest: ModuleType,
) -> dict[str, dict[date, tuple[object, ...]]]:
    """一次读取因子文件，同时保留两个趋势公式的完整日度指数池。"""

    files = sorted(
        path
        for path in backtest.FACTOR_DIR.glob("*.csv")
        if not path.name.startswith(".") and path.stem.isdigit()
    )
    if not files:
        raise FileNotFoundError(f"趋势因子目录没有年度CSV：{backtest.FACTOR_DIR}")

    required_columns = {
        "日期",
        "对标指数代码",
        "对标指数",
        "窗口收益率",
    } | {backtest.SCORE_COLUMNS[method] for method in SCORE_METHODS}
    members: dict[str, dict[date, dict[str, object]]] = {
        method: defaultdict(dict) for method in SCORE_METHODS
    }

    for path in files:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            missing = required_columns - set(reader.fieldnames or [])
            if missing:
                raise ValueError(f"{path.name}缺少列：{sorted(missing)}")
            for row_number, row in enumerate(reader, start=2):
                window_return = backtest.finite_float(row.get("窗口收益率"))
                index_code = backtest.clean_text(row.get("对标指数代码"))
                if window_return is None or not index_code:
                    continue
                signal_date = backtest.parse_date(row.get("日期"), "日期")
                index_name = backtest.clean_text(row.get("对标指数"))
                for method in SCORE_METHODS:
                    factor = backtest.finite_float(
                        row.get(backtest.SCORE_COLUMNS[method])
                    )
                    if factor is None:
                        continue
                    if index_code in members[method][signal_date]:
                        raise ValueError(
                            f"{path.name}第{row_number}行出现重复指数："
                            f"{signal_date} {index_code}"
                        )
                    members[method][signal_date][index_code] = backtest.FactorMember(
                        signal_date=signal_date,
                        index_code=index_code,
                        index_name=index_name,
                        trend_factor=factor,
                        window_return=window_return,
                    )

    result: dict[str, dict[date, tuple[object, ...]]] = {}
    for method in SCORE_METHODS:
        result[method] = {
            signal_date: tuple(
                sorted(
                    members_by_code.values(),
                    key=lambda member: (-member.trend_factor, member.index_code),
                )
            )
            for signal_date, members_by_code in sorted(members[method].items())
            if members_by_code
        }
        if not result[method]:
            raise ValueError(f"{method}没有有效趋势因子")
    return result


def build_universal_selections(
    backtest: ModuleType,
    factor_universes: Mapping[str, Mapping[date, Sequence[object]]],
) -> dict[date, object]:
    """构造全指数日度池，只用于一次性确定每个指数对应的代表ETF。"""

    all_dates = sorted(
        {
            signal_date
            for method_universe in factor_universes.values()
            for signal_date in method_universe
        }
    )
    selections: dict[date, object] = {}
    for signal_date in all_dates:
        by_code: dict[str, object] = {}
        for method in SCORE_METHODS:
            for member in factor_universes[method].get(signal_date, ()):
                by_code.setdefault(member.index_code, member)
        ordered_members = tuple(by_code[code] for code in sorted(by_code))
        selections[signal_date] = backtest.DailyFactorSelection(
            signal_date=signal_date,
            planned_index_count=len(ordered_members),
            members=ordered_members,
            filtered_index_count=0,
        )
    return selections


def build_candidate_lookup(
    universal_targets: Mapping[date, object],
) -> dict[date, dict[str, object]]:
    return {
        signal_date: {
            member.index_code: member for member in target.members
        }
        for signal_date, target in universal_targets.items()
    }


def prepare_method_contexts(
    backtest: ModuleType,
    factor_universes: Mapping[str, Mapping[date, Sequence[object]]],
    full_trading_calendar: Sequence[date],
) -> dict[str, MethodContext]:
    benchmark = backtest.load_benchmark_data()
    benchmark_end_date = max(benchmark.closes)
    contexts: dict[str, MethodContext] = {}
    calendar_set = set(full_trading_calendar)

    for method in SCORE_METHODS:
        signal_dates = sorted(factor_universes[method])
        first_signal_date = signal_dates[0]
        effective_last_date = min(signal_dates[-1], benchmark_end_date)
        if effective_last_date < first_signal_date:
            raise ValueError(
                f"{method}的基准截止日早于首个趋势信号日："
                f"{benchmark_end_date} < {first_signal_date}"
            )
        missing_dates = [
            signal_date
            for signal_date in signal_dates
            if first_signal_date <= signal_date <= effective_last_date
            and signal_date not in calendar_set
        ]
        if missing_dates:
            raise ValueError(
                f"{method}趋势信号日期不在ETF交易日历中："
                + ",".join(value.isoformat() for value in missing_dates[:10])
            )
        trading_dates = tuple(
            current_date
            for current_date in full_trading_calendar
            if first_signal_date <= current_date <= effective_last_date
        )
        if not trading_dates:
            raise ValueError(f"{method}没有可用ETF交易日期")
        contexts[method] = MethodContext(
            trading_dates=trading_dates,
            benchmark_end_date=benchmark_end_date,
        )
    return contexts


def precompute_filter_metrics(
    backtest: ModuleType,
    factor_universes: Mapping[str, Mapping[date, Sequence[object]]],
    price_history: Mapping[str, tuple[Sequence[date], Sequence[float]]],
    windows: Sequence[int],
) -> dict[tuple[date, str], FilterMetrics]:
    """一次计算全部指数日期的所有波动率窗口和固定BIAS。"""

    events_by_code: dict[str, set[date]] = defaultdict(set)
    for method_universe in factor_universes.values():
        for signal_date, members in method_universe.items():
            for member in members:
                events_by_code[member.index_code].add(signal_date)

    max_window = max(windows)
    needs_bias = backtest.BIAS_MODE != "none"
    required_prices = max(
        max_window + 1,
        backtest.BIAS_WINDOW if needs_bias else 0,
    )
    metrics: dict[tuple[date, str], FilterMetrics] = {}

    for index_position, index_code in enumerate(sorted(events_by_code), start=1):
        price_dates, prices = price_history.get(index_code, ((), ()))
        if not price_dates:
            raise ValueError(f"{index_code}没有过滤所需的指数历史价格")
        daily_returns = [
            later / earlier - 1.0
            for earlier, later in zip(prices, prices[1:])
        ]
        return_prefix_sums = [0.0]
        return_square_prefix_sums = [0.0]
        for value in daily_returns:
            return_prefix_sums.append(return_prefix_sums[-1] + value)
            return_square_prefix_sums.append(
                return_square_prefix_sums[-1] + value * value
            )

        for signal_date in sorted(events_by_code[index_code]):
            end = bisect_right(price_dates, signal_date)
            if end < required_prices:
                raise ValueError(
                    f"{signal_date} {index_code}过滤历史不足"
                    f"{required_prices}个收盘价"
                )
            if (
                signal_date - price_dates[end - 1]
            ).days > backtest.MAX_PRICE_STALENESS_CALENDAR_DAYS:
                raise ValueError(f"{signal_date} {index_code}过滤历史价格过期")

            return_end = end - 1
            volatilities = []
            for window in windows:
                return_start = return_end - window
                return_sum = (
                    return_prefix_sums[return_end]
                    - return_prefix_sums[return_start]
                )
                return_square_sum = (
                    return_square_prefix_sums[return_end]
                    - return_square_prefix_sums[return_start]
                )
                variance_numerator = (
                    return_square_sum - return_sum * return_sum / window
                )
                # 浮点相减可能产生极小负数；真实样本方差不小于0。
                variance = max(variance_numerator / (window - 1), 0.0)
                volatility = math.sqrt(variance)
                if not math.isfinite(volatility):
                    raise ValueError(
                        f"{signal_date} {index_code}过滤波动率无效"
                    )
                volatilities.append(volatility)

            bias: float | None = None
            if needs_bias:
                bias_prices = prices[end - backtest.BIAS_WINDOW : end]
                bias = prices[end - 1] / statistics.mean(bias_prices) - 1.0
                if not math.isfinite(bias):
                    raise ValueError(f"{signal_date} {index_code} BIAS无效")
            metrics[(signal_date, index_code)] = FilterMetrics(
                volatilities=tuple(volatilities),
                bias=bias,
            )

        if index_position % 100 == 0 or index_position == len(events_by_code):
            print(
                f"  已预计算过滤指标：{index_position}/{len(events_by_code)}个指数",
                flush=True,
            )
    return metrics


def build_volatility_orders(
    backtest: ModuleType,
    factor_universes: Mapping[str, Mapping[date, Sequence[object]]],
    metrics: Mapping[tuple[date, str], FilterMetrics],
    window_position: int,
) -> dict[str, dict[date, tuple[str, ...]]]:
    """按当前波动率窗口排序全池；BIAS不参与波动率排名。"""

    orders: dict[str, dict[date, tuple[str, ...]]] = {
        method: {} for method in SCORE_METHODS
    }
    for method in SCORE_METHODS:
        for signal_date, members in factor_universes[method].items():
            orders[method][signal_date] = tuple(sorted(
                (member.index_code for member in members),
                key=lambda code: (
                    -metrics[(signal_date, code)].volatilities[window_position]
                    if backtest.VOL_FILTER_MODE == "high"
                    else metrics[(signal_date, code)].volatilities[window_position],
                    code,
                ),
            ))
    return orders


def select_members(
    backtest: ModuleType,
    members: Sequence[object],
    volatility_order: Sequence[str],
    metrics: Mapping[tuple[date, str], FilterMetrics],
    keep_ratio: float,
) -> tuple[tuple[tuple[object, int], ...], int]:
    if not members:
        return (), 0
    signal_date = members[0].signal_date
    keep_count = math.ceil(len(members) * keep_ratio)
    kept_volatility_codes = frozenset(volatility_order[:keep_count])
    eligible = [
        member
        for member in members
        if member.index_code in kept_volatility_codes
        and (
            backtest.BIAS_MODE == "none"
            or backtest.passes_bias_filter(
                metrics[(signal_date, member.index_code)].bias
            )
        )
    ]
    planned_count = math.ceil(len(eligible) * backtest.TOP_PERCENT)
    return (
        tuple(
            (member, factor_rank)
            for factor_rank, member in enumerate(
                eligible[:planned_count],
                start=1,
            )
        ),
        len(eligible),
    )


def build_parameter_targets(
    backtest: ModuleType,
    method: str,
    factor_universe: Mapping[date, Sequence[object]],
    candidate_lookup: Mapping[date, Mapping[str, object]],
    volatility_orders: Mapping[date, Sequence[str]],
    metrics: Mapping[tuple[date, str], FilterMetrics],
    keep_ratio: float,
    trading_dates: Sequence[date],
) -> dict[date, object]:
    first_date = trading_dates[0]
    last_date = trading_dates[-1]
    targets: dict[date, object] = {}

    for signal_date, members in factor_universe.items():
        if signal_date < first_date or signal_date > last_date:
            continue
        selected, eligible_count = select_members(
            backtest,
            members,
            volatility_orders[signal_date],
            metrics,
            keep_ratio,
        )
        planned_count = math.ceil(eligible_count * backtest.TOP_PERCENT)

        mapped = [
            (member, factor_rank, candidate_lookup.get(signal_date, {}).get(member.index_code))
            for member, factor_rank in selected
        ]
        mapped = [item for item in mapped if item[2] is not None]
        target_weight = 1.0 / len(mapped) if mapped else 0.0
        target_members = tuple(
            backtest.TargetMember(
                signal_date=signal_date,
                index_code=member.index_code,
                index_name=member.index_name or candidate.index_name,
                etf_code=candidate.etf_code,
                etf_name=candidate.etf_name,
                target_weight=target_weight,
                trend_factor=member.trend_factor,
                factor_rank=factor_rank,
                selection_volume=candidate.selection_volume,
                selection_amount=candidate.selection_amount,
                selection_scale=candidate.selection_scale,
            )
            for member, factor_rank, candidate in mapped
        )
        targets[signal_date] = backtest.DailyTarget(
            signal_date=signal_date,
            planned_index_count=planned_count,
            selected_index_count=len(selected),
            members=target_members,
            filtered_index_count=len(members) - eligible_count,
        )
    if not targets:
        raise ValueError(f"{method}在当前回测区间没有目标持仓")
    return targets


def run_total_return(
    backtest: ModuleType,
    mode: str,
    trading_dates: Sequence[date],
    targets: Mapping[date, object],
    prices: Mapping[date, Mapping[str, object]],
) -> float:
    """复用正式调仓函数，只保留计算累计收益率所需的账户状态。"""

    if mode not in TRADE_MODES:
        raise ValueError(f"未知回测模式：{mode}")
    accounts = [
        backtest.AccountState(
            account_id=account_index + 1,
            positions={},
            cash=backtest.INITIAL_NAV / backtest.ACCOUNT_COUNT,
        )
        for account_index in range(backtest.ACCOUNT_COUNT)
    ]
    for account_index, account in enumerate(accounts):
        first_position = account_index if mode == "close" else account_index + 1
        account.next_rebalance_date = (
            trading_dates[first_position]
            if first_position < len(trading_dates)
            else None
        )

    last_closes: dict[str, float] = {}
    nav = backtest.INITIAL_NAV

    for date_position, current_date in enumerate(trading_dates):
        scheduled_accounts = [
            account
            for account in accounts
            if account.next_rebalance_date == current_date
        ]
        if len(scheduled_accounts) > 1:
            raise RuntimeError(f"{current_date}存在多个计划调仓账户")
        rebalance_account = scheduled_accounts[0] if scheduled_accounts else None
        execution_target = None
        if rebalance_account is not None:
            signal_date = (
                current_date
                if mode == "close"
                else trading_dates[date_position - 1]
            )
            execution_target = targets.get(signal_date)

        relevant_codes = {
            code
            for account in accounts
            for code in account.positions
        }
        if execution_target is not None:
            relevant_codes.update(
                member.etf_code for member in execution_target.members
            )

        daily_prices = prices.get(current_date, {})
        close_prices: dict[str, float] = {}
        vwap_prices: dict[str, float] = {}
        for code in relevant_codes:
            point = daily_prices.get(code)
            if point is None:
                continue
            if point.close is not None:
                close_prices[code] = point.close
            if point.vwap is not None:
                vwap_prices[code] = point.vwap
        last_closes.update(close_prices)
        execution_prices = close_prices if mode == "close" else vwap_prices

        if rebalance_account is not None:
            if execution_target is None:
                target_weights: dict[str, float] = {}
                target_metadata: dict[str, object] = {}
            else:
                target_weights, target_metadata = backtest.aggregate_target_weights(
                    execution_target,
                    execution_prices,
                    close_prices,
                )
            rebalance_result = backtest.rebalance_portfolio(
                rebalance_account.positions,
                rebalance_account.cash,
                target_weights,
                target_metadata,
                execution_prices,
            )
            rebalance_account.positions = rebalance_result.positions
            rebalance_account.cash = rebalance_result.cash
            next_position = date_position + backtest.ACCOUNT_REBALANCE_INTERVAL
            rebalance_account.next_rebalance_date = (
                trading_dates[next_position]
                if next_position < len(trading_dates)
                else None
            )

        nav = 0.0
        for account in accounts:
            nav += account.cash
            for code, position in account.positions.items():
                close = close_prices.get(code, last_closes.get(code))
                if close is None:
                    raise RuntimeError(f"持仓ETF缺少可用收盘价：{code}")
                nav += position.shares * close

    return nav / backtest.INITIAL_NAV - 1.0


def run_grid(
    backtest: ModuleType,
    windows: Sequence[int],
    keep_ratios: Sequence[float],
) -> tuple[
    dict[tuple[str, str], np.ndarray],
    dict[str, object],
]:
    print("读取两个趋势公式的因子数据……", flush=True)
    factor_universes = load_factor_universes(backtest)
    universal_selections = build_universal_selections(
        backtest,
        factor_universes,
    )

    print("一次性扫描ETF总表，建立每日指数到代表ETF的映射……", flush=True)
    universal_targets, full_trading_calendar = backtest.build_daily_targets(
        universal_selections
    )
    candidate_lookup = build_candidate_lookup(universal_targets)
    contexts = prepare_method_contexts(
        backtest,
        factor_universes,
        full_trading_calendar,
    )

    all_index_codes = {
        member.index_code
        for method_universe in factor_universes.values()
        for members in method_universe.values()
        for member in members
    }
    print("一次性读取指数历史价格并预计算全部过滤指标……", flush=True)
    price_history = backtest.load_filter_price_history(
        all_index_codes,
        backtest.INDEX_RAW_DATA_DIR,
    )
    filter_metrics = precompute_filter_metrics(
        backtest,
        factor_universes,
        price_history,
        windows,
    )
    del price_history

    required_dates = {
        current_date
        for context in contexts.values()
        for current_date in context.trading_dates
    }
    selected_codes = {
        member.etf_code
        for signal_date, target in universal_targets.items()
        if signal_date in required_dates
        for member in target.members
    }
    print(
        f"一次性读取{len(selected_codes)}只可能入选ETF的成交价格……",
        flush=True,
    )
    prices = backtest.load_selected_prices(selected_codes, required_dates)

    matrices = {
        (method, mode): np.full(
            (len(keep_ratios), len(windows)),
            np.nan,
            dtype=float,
        )
        for method in SCORE_METHODS
        for mode in TRADE_MODES
    }
    total_pairs = len(windows) * len(keep_ratios)
    completed_pairs = 0

    for window_position, window in enumerate(windows):
        volatility_orders = build_volatility_orders(
            backtest,
            factor_universes,
            filter_metrics,
            window_position,
        )
        for ratio_position, keep_ratio in enumerate(keep_ratios):
            for method in SCORE_METHODS:
                context = contexts[method]
                targets = build_parameter_targets(
                    backtest,
                    method,
                    factor_universes[method],
                    candidate_lookup,
                    volatility_orders[method],
                    filter_metrics,
                    keep_ratio,
                    context.trading_dates,
                )
                for mode in TRADE_MODES:
                    matrices[(method, mode)][ratio_position, window_position] = (
                        run_total_return(
                            backtest,
                            mode,
                            context.trading_dates,
                            targets,
                            prices,
                        )
                    )
            completed_pairs += 1
            if completed_pairs % 25 == 0 or completed_pairs == total_pairs:
                print(
                    f"  参数网格进度：{completed_pairs}/{total_pairs} "
                    f"({completed_pairs / total_pairs:.1%})",
                    flush=True,
                )

    metadata = {
        "factor_universes": factor_universes,
        "candidate_lookup": candidate_lookup,
        "contexts": contexts,
        "filter_metrics": filter_metrics,
        "prices": prices,
    }
    return matrices, metadata


def save_heatmaps(
    backtest: ModuleType,
    matrices: Mapping[tuple[str, str], np.ndarray],
    windows: Sequence[int],
    keep_ratios: Sequence[float],
    settings: VolatilityTuningSettings,
) -> list[Path]:
    values = np.concatenate([matrix.ravel() for matrix in matrices.values()])
    finite_values = values[np.isfinite(values)]
    if finite_values.size == 0:
        raise ValueError("参数网格没有可绘制的累计收益率")
    color_min = float(finite_values.min())
    color_max = float(finite_values.max())
    if math.isclose(color_min, color_max, rel_tol=0.0, abs_tol=1e-15):
        color_min -= 1e-6
        color_max += 1e-6

    plt.rcParams["font.sans-serif"] = [
        "Arial Unicode MS",
        "PingFang SC",
        "Heiti SC",
        "SimHei",
        "DejaVu Sans",
    ]
    plt.rcParams["axes.unicode_minus"] = False
    settings.output_directory.mkdir(parents=True, exist_ok=True)

    method_labels = {
        "return_r2": "收益率×R平方",
        "return_vol": "收益率÷波动率",
    }
    mode_labels = {
        "close": "收盘价成交",
        "next_day_vwap": "次日VWAP成交",
    }
    window_ticks = [
        position
        for position, window in enumerate(windows)
        if window == windows[0] or window % 5 == 0 or window == windows[-1]
    ]
    ratio_ticks = list(range(len(keep_ratios)))
    saved: list[Path] = []
    current_parameter = (
        backtest.VOL_RETURN_DAYS,
        backtest.VOL_KEEP_TOP_RATIO,
    )
    bias_description = (
        f"固定BIAS={backtest.BIAS_WINDOW}日/"
        f"{backtest.BIAS_MODE}/上限{backtest.BIAS_UPPER:.0%}"
        if settings.bias_mode != "none"
        else "BIAS过滤关闭"
    )

    for method in SCORE_METHODS:
        for mode in TRADE_MODES:
            matrix = matrices[(method, mode)]
            figure, axis = plt.subplots(figsize=(20, 11))
            image = axis.imshow(
                matrix,
                origin="lower",
                aspect="auto",
                interpolation="nearest",
                cmap="RdYlGn",
                vmin=color_min,
                vmax=color_max,
            )
            axis.set_title(
                "波动率参数网格："
                f"{method_labels[method]} / {mode_labels[mode]}",
                fontsize=17,
                fontweight="bold",
                pad=18,
            )
            axis.set_xlabel("VOL_RETURN_DAYS（日收益个数）", fontsize=13)
            axis.set_ylabel("VOL_KEEP_TOP_RATIO（保留比例）", fontsize=13)
            axis.set_xticks(window_ticks)
            axis.set_xticklabels([str(windows[index]) for index in window_ticks])
            axis.set_yticks(ratio_ticks)
            axis.set_yticklabels(
                [f"{keep_ratios[index]:.0%}" for index in ratio_ticks]
            )
            axis.set_xticks(
                np.arange(-0.5, len(windows), 1),
                minor=True,
            )
            axis.set_yticks(
                np.arange(-0.5, len(keep_ratios), 1),
                minor=True,
            )
            axis.grid(
                which="minor",
                color="white",
                linewidth=0.22,
                alpha=0.50,
            )
            axis.tick_params(which="minor", bottom=False, left=False)

            if (
                current_parameter[0] in windows
                and current_parameter[1] in keep_ratios
            ):
                current_x = windows.index(current_parameter[0])
                current_y = keep_ratios.index(current_parameter[1])
                axis.add_patch(
                    Rectangle(
                        (current_x - 0.5, current_y - 0.5),
                        1.0,
                        1.0,
                        fill=False,
                        edgecolor="#1f3a93",
                        linewidth=2.2,
                    )
                )

            colorbar = figure.colorbar(image, ax=axis, pad=0.015)
            colorbar.set_label("累计总收益率（扣费后）", fontsize=12)
            colorbar.ax.yaxis.set_major_formatter(PercentFormatter(xmax=1.0))
            axis.text(
                0.0,
                -0.095,
                "固定参数："
                f"更新频率{settings.update_frequency}；"
                f"过滤状态{settings.filter_profile}；"
                f"聚类阈值{backtest.CLUSTER_CORRELATION_THRESHOLD:g}；"
                f"趋势窗口{backtest.TREND_WINDOW}日；"
                f"波动方向{backtest.VOL_FILTER_MODE}；"
                f"{bias_description}；"
                f"过滤后Top{backtest.TOP_PERCENT:.0%}；"
                f"{backtest.ACCOUNT_COUNT}账户错峰；"
                f"单边成本{backtest.TRANSACTION_COST_RATE:.1%}；"
                f"蓝框=当前波动率({current_parameter[0]}日, "
                f"{current_parameter[1]:.0%})",
                transform=axis.transAxes,
                fontsize=10,
                color="#444444",
            )
            figure.tight_layout()

            output_path = (
                settings.output_directory / OUTPUT_FILES[(method, mode)]
            )
            temporary_path = output_path.with_name(
                f".{output_path.stem}.tmp{output_path.suffix}"
            )
            figure.savefig(temporary_path, dpi=220, bbox_inches="tight")
            plt.close(figure)
            temporary_path.replace(output_path)
            saved.append(output_path)
    return saved


def main(
    update_frequency: str = UPDATE_FREQUENCY,
    filter_profile: str = FILTER_PROFILE,
    correlation_threshold: float = CLUSTER_CORRELATION_THRESHOLD,
    trend_window: int = TREND_WINDOW,
    stagger_days: int = ACCOUNT_REBALANCE_INTERVAL,
) -> None:
    settings = build_tuning_settings(
        update_frequency,
        filter_profile,
        correlation_threshold,
        trend_window,
        stagger_days,
    )
    backtest = load_backtest_module(settings)
    validate_grid(backtest, settings)
    bias_summary = (
        f"固定BIAS={backtest.BIAS_WINDOW}日/"
        f"{backtest.BIAS_MODE}/上限{backtest.BIAS_UPPER:.0%}"
        if settings.bias_mode != "none"
        else "BIAS过滤关闭"
    )
    print(
        f"实验组合：{settings.update_frequency} / "
        f"{settings.filter_profile} / "
        f"阈值{settings.correlation_threshold:g} / "
        f"趋势窗口{settings.trend_window}日 / "
        f"错峰{settings.stagger_days}日。\n"
        f"波动率参数网格：窗口{VOL_WINDOWS[0]}至{VOL_WINDOWS[-1]}个日收益，"
        f"保留比例{VOL_KEEP_RATIOS[0]:.0%}至{VOL_KEEP_RATIOS[-1]:.0%}；"
        f"共{len(VOL_WINDOWS) * len(VOL_KEEP_RATIOS)}组参数，"
        f"{bias_summary}；"
        "每组回测两个公式和两种成交方式。",
        flush=True,
    )
    matrices, metadata = run_grid(
        backtest,
        VOL_WINDOWS,
        VOL_KEEP_RATIOS,
    )
    del metadata
    output_paths = save_heatmaps(
        backtest,
        matrices,
        VOL_WINDOWS,
        VOL_KEEP_RATIOS,
        settings,
    )
    print("完成，仅生成以下四张参数热力图：", flush=True)
    for path in output_paths:
        print(f"  {path}", flush=True)


if __name__ == "__main__":
    main()
