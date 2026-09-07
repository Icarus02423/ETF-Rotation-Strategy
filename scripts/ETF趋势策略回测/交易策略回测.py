#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETF趋势轮动策略回测。

策略：
1. 每日分别按两种趋势得分取全池前10%，再检查正收益、RSI、BIAS，不向后补选；
2. ETF映射沿用同指数信号日成交量最大者，成交额、规模和代码用于并列排序；
3. 单账户按买入批次独立记账，买入日为第0日，最短持有期按交易日计算；
4. 每日收盘逐批次检查固定跌幅止损及浮盈回撤退出；风控优先，未到期批次不因排名变化卖出；
5. 到期且未入选的批次正常卖出；未卖出的批次数量、买入价和计时保持不变；
6. 卖出释放资金与现金在本轮过滤后的ETF间等额投入，每次追加买入另建批次；
   同轮风控退出ETF禁止买回，无合格标的时留现金，组合不每日重新等权；
7. 主口径为信号日收盘决策、次日VWAP成交；close仅保留为同收盘理想化对照；
8. 买卖各收0.1%费用；无成交价格或成交额时不虚构成交，未成交卖单后续重试。
   收盘价与VWAP沿用下载数据的统一前复权口径，不在回测中额外调整。

输入：
- outputs/etf_trend_strategy/threshold_<阈值>/factors/window_<窗口>/YYYY.csv
- outputs/etf_trend_strategy/threshold_<阈值>/index_prices/*.csv
- outputs/etf_data/etf_data.csv
- outputs/benchmark_data/*.csv（按BENCHMARK_CODE选择）

输出：
- <公式>/daily_rotation/top_<比例>pct__r_gt_<门槛>/
  min_hold_<H>d__fixed_stop_<S>pct__profit_trigger_<P>pct__peak_drawdown_<D>pct/
  <RSI与BIAS条件>/<成交方式>/
- 文件前缀包含最短持有期、固定止损、浮盈启动和最高点回撤参数；每种成交方式独立输出
  年度指标、总指标、合并持仓、时序、批次明细五个Excel及原有五张图。
- batch_details.xlsx包含账户每日状态、批次持仓、批次交易；批次编号贯穿买入和卖出。
"""

from __future__ import annotations

import csv
import math
import statistics
from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path
from typing import Any, Mapping, Sequence

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path(__file__).resolve().parents[2]

# ============================= 回测参数 =============================
CLUSTER_CORRELATION_THRESHOLD = 0.9
TREND_WINDOW = 15
# 评价基准代码；需与“下载基准数据.py”中的BENCHMARK_CODE保持一致。
BENCHMARK_CODE = "000510.CSI"
# 默认一次运行两种得分。若以后只想跑其中一种，可只保留对应英文键。
SCORE_METHODS_TO_RUN = ("return_r2", "return_vol")
TOP_PERCENT = 0.10
# H和S相互独立，RSI周期也不再与交易频率或最短持有期绑定。
MIN_HOLD_DAYS = 10
STOP_LOSS_PCT = 0.05
PROFIT_TRAILING_TRIGGER_PCT = 0.10
PROFIT_TRAILING_DRAWDOWN_PCT = 0.05
ACCOUNT_COUNT = 1
RSI_PERIOD = 7
BIAS_PERIOD = TREND_WINDOW
RSI_NEUTRAL_LEVEL = 40
BIAS_NEUTRAL_LEVEL = -0.05
MIN_WINDOW_RETURN = 0.0
STRATEGY_VARIANT_DIR = "daily_rotation"
SELECTION_VARIANT_DIR = f"top_{TOP_PERCENT * 100:g}pct__r_gt_{MIN_WINDOW_RETURN:g}"
ACCOUNT_VARIANT_DIR = (
    f"min_hold_{MIN_HOLD_DAYS}d__fixed_stop_{STOP_LOSS_PCT * 100:g}pct"
    f"__profit_trigger_{PROFIT_TRAILING_TRIGGER_PCT * 100:g}pct"
    f"__peak_drawdown_{PROFIT_TRAILING_DRAWDOWN_PCT * 100:g}pct"
)
INDICATOR_VARIANT_DIR = (
    f"rsi_{RSI_PERIOD}_gt_{RSI_NEUTRAL_LEVEL:g}"
    f"__bias_{BIAS_PERIOD}_gt_{BIAS_NEUTRAL_LEVEL:g}"
)
TRANSACTION_COST_RATE = 0.001
ANNUAL_TRADING_DAYS = 252
ANNUAL_RISK_FREE_RATE = 0.015
INITIAL_NAV = 1.0
CAPACITY_DAILY_AMOUNT_RATIO = 0.10
CAPACITY_DESCENDING_QUANTILE = 0.95
# ====================================================================

ALLOWED_CLUSTER_THRESHOLDS = (0.7, 0.8, 0.9)
ALLOWED_TREND_WINDOWS = (5, 10, 15)
SCORE_COLUMNS = {
    "return_r2": "趋势质量因子",
    "return_vol": "风险调整趋势得分",
}
SCORE_LABELS = {
    "return_r2": "收益率×R平方",
    "return_vol": "收益率÷波动率",
}
FACTOR_DIR = (
    PROJECT_ROOT
    / "outputs"
    / "etf_trend_strategy"
    / f"threshold_{CLUSTER_CORRELATION_THRESHOLD:g}"
    / "factors"
    / f"window_{TREND_WINDOW}"
)
INDEX_PRICE_DIR = (
    PROJECT_ROOT
    / "outputs"
    / "etf_trend_strategy"
    / f"threshold_{CLUSTER_CORRELATION_THRESHOLD:g}"
    / "index_prices"
)
ETF_DATA_FILE = PROJECT_ROOT / "outputs" / "etf_data" / "etf_data.csv"
BENCHMARK_DIR = PROJECT_ROOT / "outputs" / "benchmark_data"
BACKTEST_DIR = (
    PROJECT_ROOT
    / "outputs"
    / "etf_trend_strategy"
    / f"threshold_{CLUSTER_CORRELATION_THRESHOLD:g}"
    / "backtests"
    / f"window_{TREND_WINDOW}"
)

FACTOR_REQUIRED_COLUMNS = {
    "日期",
    "对标指数代码",
    "对标指数",
    "窗口结束日",
    "窗口收益率",
} | set(SCORE_COLUMNS.values())
INDEX_PRICE_REQUIRED_COLUMNS = {"收益日期", "对标指数代码", "收盘价"}
ETF_REQUIRED_COLUMNS = {
    "日期",
    "代码",
    "名称",
    "上市日期",
    "对标指数代码",
    "对标指数",
    "规模",
    "成交量",
    "成交额",
    "收盘价",
    "VWAP",
}
BENCHMARK_REQUIRED_COLUMNS = {"日期", "代码", "名称", "收盘价"}


@dataclass(frozen=True)
class FactorMember:
    signal_date: date
    window_end_date: date
    index_code: str
    index_name: str
    trend_factor: float
    window_return: float
    factor_rank: int = 0


@dataclass(frozen=True)
class DailyFactorSelection:
    signal_date: date
    planned_index_count: int
    members: tuple[FactorMember, ...]

    @property
    def filtered_index_count(self) -> int:
        return self.planned_index_count - len(self.members)


@dataclass(frozen=True)
class CandidateEtf:
    code: str
    name: str
    index_code: str
    index_name: str
    volume: float
    amount: float
    scale: float


@dataclass(frozen=True)
class TargetMember:
    signal_date: date
    index_code: str
    index_name: str
    etf_code: str
    etf_name: str
    target_weight: float
    trend_factor: float
    factor_rank: int
    selection_volume: float
    selection_amount: float
    selection_scale: float


@dataclass(frozen=True)
class DailyTarget:
    signal_date: date
    planned_index_count: int
    selected_index_count: int
    members: tuple[TargetMember, ...]

    @property
    def filtered_index_count(self) -> int:
        return self.planned_index_count - self.selected_index_count

    @property
    def unmapped_index_count(self) -> int:
        return self.selected_index_count - len(self.members)


@dataclass(frozen=True)
class PricePoint:
    close: float | None
    vwap: float | None
    amount: float | None


@dataclass
class Position:
    shares: float
    index_code: str
    index_name: str
    etf_name: str
    signal_date: date
    trend_factor: float
    factor_rank: int
    selection_volume: float
    target_weight: float
    etf_code: str
    batch_id: int
    buy_date: date
    buy_price: float
    buy_date_position: int
    highest_close: float
    pending_exit_reason: str = ""
    exit_signal_date: date | None = None
    exit_signal_price: float | None = None
    exit_signal_age: int = 0
    exit_signal_peak_price: float | None = None
    exit_risk_price: float | None = None


@dataclass(frozen=True)
class HoldingRecord:
    current_date: date
    signal_date: date
    index_code: str
    index_name: str
    trend_factor: float
    factor_rank: int
    etf_code: str
    etf_name: str
    selection_volume: float
    target_weight: float
    actual_weight: float


@dataclass(frozen=True)
class AccountDailyRecord:
    current_date: date
    account_id: int
    account_nav: float
    etf_market_value: float
    cash: float
    cash_weight: float
    rebalance_attempted: bool
    rebalance_succeeded: bool
    signal_date: date | None
    last_rebalance_date: date | None
    next_rebalance_date: date | None


@dataclass(frozen=True)
class AccountHoldingRecord:
    current_date: date
    account_id: int
    signal_date: date
    etf_code: str
    etf_name: str
    market_value: float
    account_weight: float
    total_portfolio_weight: float
    batch_id: int
    buy_date: date
    buy_price: float
    shares: float
    holding_days: int
    stop_price: float
    highest_close: float
    max_return: float
    trailing_active: bool
    trailing_stop_price: float | None
    pending_exit_reason: str


@dataclass(frozen=True)
class AccountTradeRecord:
    execution_date: date
    account_id: int
    signal_date: date | None
    etf_code: str
    etf_name: str
    direction: str
    before_value: float
    target_value: float
    trade_amount: float
    transaction_cost: float
    batch_id: int
    buy_date: date
    buy_price: float
    shares: float
    execution_price: float
    reason: str
    signal_price: float | None
    signal_holding_days: int
    signal_peak_price: float | None
    risk_exit_price: float | None


@dataclass(frozen=True)
class NavRecord:
    current_date: date
    signal_date: date | None
    pre_trade_nav: float
    nav: float
    gross_return: float
    daily_return: float
    drawdown: float
    buy_ratio: float
    sell_ratio: float
    bilateral_ratio: float
    one_way_turnover: float
    transaction_cost: float
    transaction_cost_rate: float
    cumulative_cost_rate: float
    capacity: float | None
    holding_count: int
    cash_weight: float
    initial_build: bool
    rebalance_account_id: int | None
    rebalance_attempted: bool
    rebalance_succeeded: bool
    skip_reason: str
    selected_index_count: int
    unmapped_index_count: int
    missing_price_index_count: int


@dataclass(frozen=True)
class BenchmarkData:
    code: str
    name: str
    closes: Mapping[date, float]


@dataclass(frozen=True)
class BenchmarkRecord:
    current_date: date
    close: float
    daily_return: float
    nav: float
    active_return: float
    relative_return: float
    relative_nav: float


@dataclass(frozen=True)
class TradeDetail:
    etf_code: str
    etf_name: str
    direction: str
    before_value: float
    target_value: float
    trade_amount: float
    transaction_cost: float
    batch_id: int
    buy_date: date
    buy_price: float
    shares: float
    execution_price: float
    reason: str
    signal_price: float | None
    signal_holding_days: int
    signal_date: date | None
    signal_peak_price: float | None
    risk_exit_price: float | None


@dataclass(frozen=True)
class RebalanceResult:
    positions: dict[int, Position]
    cash: float
    pre_trade_nav: float
    buy_amount: float
    sell_amount: float
    transaction_cost: float
    trade_details: tuple[TradeDetail, ...]
    succeeded: bool
    skip_reason: str


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text in {"", "--", "None", "null", "NULL"} else text


def parse_date(value: object, field_name: str) -> date:
    text = clean_text(value)[:10]
    for date_format in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            continue
    raise ValueError(f"{field_name}不是有效日期：{value!r}")


def finite_float(value: object) -> float | None:
    text = clean_text(value).replace(",", "")
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return number if math.isfinite(number) else None


def positive_float(value: object) -> float | None:
    number = finite_float(value)
    return number if number is not None and number > 0 else None


def calculate_wilder_rsi(
    closes: Sequence[float],
    period: int,
) -> list[float | None]:
    values: list[float | None] = [None] * len(closes)
    if len(closes) <= period:
        return values

    gains: list[float] = []
    losses: list[float] = []
    for position in range(1, period + 1):
        change = closes[position] - closes[position - 1]
        gains.append(max(change, 0.0))
        losses.append(max(-change, 0.0))
    average_gain = statistics.fmean(gains)
    average_loss = statistics.fmean(losses)

    def rsi_value() -> float:
        if average_gain == 0 and average_loss == 0:
            return 50.0
        if average_loss == 0:
            return 100.0
        relative_strength = average_gain / average_loss
        return 100.0 - 100.0 / (1.0 + relative_strength)

    values[period] = rsi_value()
    for position in range(period + 1, len(closes)):
        change = closes[position] - closes[position - 1]
        gain = max(change, 0.0)
        loss = max(-change, 0.0)
        average_gain = (average_gain * (period - 1) + gain) / period
        average_loss = (average_loss * (period - 1) + loss) / period
        values[position] = rsi_value()
    return values


def calculate_bias(
    closes: Sequence[float],
    period: int,
) -> list[float | None]:
    values: list[float | None] = [None] * len(closes)
    rolling_sum = 0.0
    for position, close in enumerate(closes):
        rolling_sum += close
        if position >= period:
            rolling_sum -= closes[position - period]
        if position >= period - 1:
            moving_average = rolling_sum / period
            values[position] = close / moving_average - 1.0
    return values


def load_defensive_indicators(
    required_index_codes: set[str],
) -> dict[tuple[str, date], tuple[float, float]]:
    if not INDEX_PRICE_DIR.exists():
        raise FileNotFoundError(f"找不到指数行情目录：{INDEX_PRICE_DIR}")
    files = sorted(
        path
        for path in INDEX_PRICE_DIR.glob("*.csv")
        if not path.name.startswith(".")
    )
    if not files:
        raise FileNotFoundError(f"指数行情目录没有CSV：{INDEX_PRICE_DIR}")

    prices_by_index: dict[str, dict[date, float]] = defaultdict(dict)
    for path in files:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            missing = INDEX_PRICE_REQUIRED_COLUMNS - set(reader.fieldnames or [])
            if missing:
                raise ValueError(f"{path.name}缺少列：{sorted(missing)}")
            for row_number, row in enumerate(reader, start=2):
                index_code = clean_text(row.get("对标指数代码"))
                if index_code not in required_index_codes:
                    continue
                close = positive_float(row.get("收盘价"))
                if close is None:
                    continue
                price_date = parse_date(row.get("收益日期"), "收益日期")
                existing_close = prices_by_index[index_code].get(price_date)
                if existing_close is not None and not math.isclose(
                    existing_close,
                    close,
                    rel_tol=1e-10,
                    abs_tol=1e-8,
                ):
                    raise ValueError(
                        f"{path.name}第{row_number}行与其他文件的指数收盘价冲突："
                        f"{price_date} {index_code}"
                    )
                prices_by_index[index_code][price_date] = close

    missing_indexes = sorted(required_index_codes - set(prices_by_index))
    if missing_indexes:
        raise ValueError(
            "指数行情中缺少因子指数：" + ",".join(missing_indexes[:10])
        )

    indicators: dict[tuple[str, date], tuple[float, float]] = {}
    for index_code, daily_prices in prices_by_index.items():
        price_dates = sorted(daily_prices)
        closes = [daily_prices[price_date] for price_date in price_dates]
        rsi_values = calculate_wilder_rsi(closes, RSI_PERIOD)
        bias_values = calculate_bias(closes, BIAS_PERIOD)
        for price_date, rsi, bias in zip(price_dates, rsi_values, bias_values):
            if rsi is not None and bias is not None:
                indicators[(index_code, price_date)] = (rsi, bias)
    return indicators


def passes_defensive_filter(
    member: FactorMember,
    indicators: Mapping[tuple[str, date], tuple[float, float]],
) -> bool:
    values = indicators.get((member.index_code, member.window_end_date))
    if values is None:
        raise ValueError(
            "无法计算防守指标："
            f"{member.window_end_date} {member.index_code}"
        )
    rsi, bias = values
    return rsi > RSI_NEUTRAL_LEVEL and bias > BIAS_NEUTRAL_LEVEL


def load_benchmark_data() -> BenchmarkData:
    """按BENCHMARK_CODE读取基准文件，并严格校验基础口径。"""

    files = sorted(
        path
        for path in BENCHMARK_DIR.glob("*.csv")
        if not path.name.startswith(".")
    )
    if not files:
        raise FileNotFoundError(
            f"{BENCHMARK_DIR} 中没有基准CSV。"
        )

    target_code = clean_text(BENCHMARK_CODE).upper()
    matched_files: list[Path] = []
    available_benchmarks: list[str] = []
    for candidate in files:
        with candidate.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            missing = BENCHMARK_REQUIRED_COLUMNS - set(reader.fieldnames or [])
            if missing:
                raise ValueError(f"{candidate.name} 缺少列：{sorted(missing)}")
            codes = {
                clean_text(row.get("代码")).upper()
                for row in reader
                if clean_text(row.get("代码"))
            }
        if len(codes) != 1:
            raise ValueError(f"{candidate.name} 必须且只能包含一个基准代码")
        code = next(iter(codes))
        available_benchmarks.append(f"{candidate.name}（{code}）")
        if code == target_code:
            matched_files.append(candidate)

    if not matched_files:
        available = "、".join(available_benchmarks)
        raise FileNotFoundError(
            f"{BENCHMARK_DIR} 中未找到代码为 {target_code} 的基准CSV；"
            f"现有基准：{available}"
        )
    if len(matched_files) > 1:
        matched_names = "、".join(path.name for path in matched_files)
        raise ValueError(
            f"代码 {target_code} 匹配到多个基准CSV：{matched_names}"
        )

    path = matched_files[0]
    closes: dict[date, float] = {}
    codes: set[str] = set()
    names: set[str] = set()
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = BENCHMARK_REQUIRED_COLUMNS - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"{path.name} 缺少列：{sorted(missing)}")
        for line_number, row in enumerate(reader, start=2):
            current_date = parse_date(
                row.get("日期"),
                f"{path.name}第{line_number}行日期",
            )
            code = clean_text(row.get("代码")).upper()
            name = clean_text(row.get("名称"))
            close = positive_float(row.get("收盘价"))
            if not code or not name or close is None:
                raise ValueError(f"{path.name} 第 {line_number} 行存在无效代码、名称或收盘价")
            if current_date in closes:
                raise ValueError(f"{path.name} 存在重复日期：{current_date}")
            closes[current_date] = close
            codes.add(code)
            names.add(name)

    if not closes:
        raise ValueError(f"{path.name} 没有有效数据")
    if len(codes) != 1 or len(names) != 1:
        raise ValueError(f"{path.name} 必须且只能包含一个基准代码和名称")
    return BenchmarkData(
        code=next(iter(codes)),
        name=next(iter(names)),
        closes=dict(sorted(closes.items())),
    )


def build_benchmark_records(
    nav_records: Sequence[NavRecord],
    benchmark: BenchmarkData,
) -> tuple[date, list[BenchmarkRecord]]:
    """按策略交易日对齐基准；以前一交易日为共同初始净值1。"""

    if not nav_records:
        raise ValueError("没有策略净值，无法对齐基准")
    first_date = nav_records[0].current_date
    previous_dates = [value for value in benchmark.closes if value < first_date]
    if not previous_dates:
        raise ValueError(
            f"基准 {benchmark.name} 缺少回测首日 {first_date} 之前的收盘价，"
            "无法建立初始净值1。"
        )
    baseline_date = previous_dates[-1]
    previous_close = benchmark.closes[baseline_date]
    benchmark_nav = 1.0
    relative_nav = 1.0
    result: list[BenchmarkRecord] = []

    for nav_record in nav_records:
        current_date = nav_record.current_date
        close = benchmark.closes.get(current_date)
        if close is None:
            raise ValueError(
                f"基准 {benchmark.name} 缺少策略交易日 {current_date} 的收盘价；"
                "为避免错位，回测不会向前填充。"
            )
        benchmark_return = close / previous_close - 1.0
        active_return = nav_record.daily_return - benchmark_return
        relative_return = (
            (1.0 + nav_record.daily_return) / (1.0 + benchmark_return) - 1.0
        )
        benchmark_nav *= 1.0 + benchmark_return
        relative_nav *= 1.0 + relative_return
        result.append(
            BenchmarkRecord(
                current_date=current_date,
                close=close,
                daily_return=benchmark_return,
                nav=benchmark_nav,
                active_return=active_return,
                relative_return=relative_return,
                relative_nav=relative_nav,
            )
        )
        previous_close = close

    return baseline_date, result


def validate_parameters() -> None:
    if not clean_text(BENCHMARK_CODE):
        raise ValueError("BENCHMARK_CODE不能为空")
    if not any(
        math.isclose(CLUSTER_CORRELATION_THRESHOLD, allowed, abs_tol=1e-12)
        for allowed in ALLOWED_CLUSTER_THRESHOLDS
    ):
        raise ValueError("CLUSTER_CORRELATION_THRESHOLD只能设为0.7、0.8或0.9")
    if TREND_WINDOW not in ALLOWED_TREND_WINDOWS:
        raise ValueError("TREND_WINDOW只能设为20、40或60")
    invalid_score_methods = [
        method for method in SCORE_METHODS_TO_RUN if method not in SCORE_COLUMNS
    ]
    if not SCORE_METHODS_TO_RUN or invalid_score_methods:
        raise ValueError(
            "SCORE_METHODS_TO_RUN包含无效得分方法："
            f"{invalid_score_methods}"
        )
    if len(set(SCORE_METHODS_TO_RUN)) != len(SCORE_METHODS_TO_RUN):
        raise ValueError("SCORE_METHODS_TO_RUN不能包含重复得分方法")
    if not 0 < TOP_PERCENT <= 1:
        raise ValueError("TOP_PERCENT必须在0到1之间")
    if type(MIN_HOLD_DAYS) is not int or MIN_HOLD_DAYS < 1:
        raise ValueError("MIN_HOLD_DAYS必须是正整数，买入日记为第0日")
    if not math.isfinite(STOP_LOSS_PCT) or not 0 < STOP_LOSS_PCT < 1:
        raise ValueError("STOP_LOSS_PCT必须在0与1之间，例如0.05表示5%")
    if (
        not math.isfinite(PROFIT_TRAILING_TRIGGER_PCT)
        or not 0 < PROFIT_TRAILING_TRIGGER_PCT < 1
    ):
        raise ValueError(
            "PROFIT_TRAILING_TRIGGER_PCT必须在0与1之间，例如0.08表示浮盈8%后启用"
        )
    if (
        not math.isfinite(PROFIT_TRAILING_DRAWDOWN_PCT)
        or not 0 < PROFIT_TRAILING_DRAWDOWN_PCT < 1
    ):
        raise ValueError(
            "PROFIT_TRAILING_DRAWDOWN_PCT必须在0与1之间，例如0.06表示从最高收盘价回撤6%"
        )
    if type(RSI_PERIOD) is not int or RSI_PERIOD < 2:
        raise ValueError("RSI_PERIOD必须是至少2的整数，与最短持有期独立")
    if not math.isfinite(MIN_WINDOW_RETURN):
        raise ValueError("MIN_WINDOW_RETURN必须是有限数值")
    if not math.isfinite(TRANSACTION_COST_RATE) or not 0 <= TRANSACTION_COST_RATE < 1:
        raise ValueError("TRANSACTION_COST_RATE必须在[0, 1)内")
    if not 0 < CAPACITY_DAILY_AMOUNT_RATIO <= 1:
        raise ValueError("CAPACITY_DAILY_AMOUNT_RATIO必须在0到1之间")
    if not 0 <= CAPACITY_DESCENDING_QUANTILE <= 1:
        raise ValueError("CAPACITY_DESCENDING_QUANTILE必须在0到1之间")
    if ANNUAL_TRADING_DAYS <= 0 or INITIAL_NAV <= 0:
        raise ValueError("年化交易日和初始净值必须大于0")


def read_daily_top_factors(
    score_column: str,
) -> dict[date, DailyFactorSelection]:
    if not FACTOR_DIR.exists():
        raise FileNotFoundError(f"找不到趋势因子目录：{FACTOR_DIR}")
    files = sorted(
        path
        for path in FACTOR_DIR.glob("*.csv")
        if not path.name.startswith(".") and path.stem.isdigit()
    )
    if not files:
        raise FileNotFoundError(f"趋势因子目录没有年度CSV：{FACTOR_DIR}")

    members_by_date: dict[date, dict[str, FactorMember]] = defaultdict(dict)
    for path in files:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            missing = FACTOR_REQUIRED_COLUMNS - set(reader.fieldnames or [])
            if missing:
                raise ValueError(f"{path.name}缺少列：{sorted(missing)}")
            for row_number, row in enumerate(reader, start=2):
                factor = finite_float(row.get(score_column))
                window_return = finite_float(row.get("窗口收益率"))
                index_code = clean_text(row.get("对标指数代码"))
                if factor is None or window_return is None or not index_code:
                    continue
                signal_date = parse_date(row.get("日期"), "日期")
                window_end_date = parse_date(row.get("窗口结束日"), "窗口结束日")
                if index_code in members_by_date[signal_date]:
                    raise ValueError(
                        f"{path.name}第{row_number}行出现重复指数："
                        f"{signal_date} {index_code}"
                    )
                members_by_date[signal_date][index_code] = FactorMember(
                    signal_date=signal_date,
                    window_end_date=window_end_date,
                    index_code=index_code,
                    index_name=clean_text(row.get("对标指数")),
                    trend_factor=factor,
                    window_return=window_return,
                )

    required_index_codes = {
        member.index_code
        for daily_members in members_by_date.values()
        for member in daily_members.values()
    }
    defensive_indicators = load_defensive_indicators(required_index_codes)

    daily_selections: dict[date, DailyFactorSelection] = {}
    for signal_date in sorted(members_by_date):
        valid_members = sorted(
            members_by_date[signal_date].values(),
            key=lambda member: (-member.trend_factor, member.index_code),
        )
        if not valid_members:
            continue
        planned_count = max(1, math.ceil(len(valid_members) * TOP_PERCENT))
        ranked_members = tuple(
            FactorMember(
                signal_date=member.signal_date,
                window_end_date=member.window_end_date,
                index_code=member.index_code,
                index_name=member.index_name,
                trend_factor=member.trend_factor,
                window_return=member.window_return,
                factor_rank=factor_rank,
            )
            for factor_rank, member in enumerate(valid_members, start=1)
        )
        top_members = ranked_members[:planned_count]
        filtered_members = tuple(
            member
            for member in top_members
            if member.window_return > MIN_WINDOW_RETURN
            and passes_defensive_filter(member, defensive_indicators)
        )
        daily_selections[signal_date] = DailyFactorSelection(
            signal_date=signal_date,
            planned_index_count=planned_count,
            members=filtered_members,
        )
    if not daily_selections:
        raise ValueError("趋势因子文件中没有有效因子")
    return daily_selections


def candidate_sort_key(candidate: CandidateEtf) -> tuple[float, float, float, str]:
    return (-candidate.volume, -candidate.amount, -candidate.scale, candidate.code)


def build_daily_targets(
    daily_selections: Mapping[date, DailyFactorSelection],
    etf_data_file: Path = ETF_DATA_FILE,
) -> tuple[dict[date, DailyTarget], list[date]]:
    """扫描一次ETF总表，同时建立交易日历和信号日的ETF目标。"""

    if not etf_data_file.exists():
        raise FileNotFoundError(f"找不到ETF数据：{etf_data_file}")
    required_indexes = {
        signal_date: {member.index_code for member in selection.members}
        for signal_date, selection in daily_selections.items()
    }
    best_by_date_index: dict[tuple[date, str], CandidateEtf] = {}
    trading_calendar: set[date] = set()

    with etf_data_file.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = ETF_REQUIRED_COLUMNS - set(reader.fieldnames or [])
        if missing:
            raise ValueError(
                f"etf_data.csv缺少列：{sorted(missing)}；"
                "请先完成ETF收盘价和VWAP下载。"
            )
        for row in reader:
            row_date_text = clean_text(row.get("日期"))
            try:
                current_date = date.fromisoformat(row_date_text[:10])
            except ValueError:
                continue
            trading_calendar.add(current_date)
            index_code = clean_text(row.get("对标指数代码"))
            if index_code not in required_indexes.get(current_date, set()):
                continue
            code = clean_text(row.get("代码")).upper()
            volume = positive_float(row.get("成交量"))
            if not code or volume is None:
                continue
            listed_text = clean_text(row.get("上市日期"))
            if listed_text:
                try:
                    if parse_date(listed_text, "上市日期") > current_date:
                        continue
                except ValueError:
                    continue
            candidate = CandidateEtf(
                code=code,
                name=clean_text(row.get("名称")),
                index_code=index_code,
                index_name=clean_text(row.get("对标指数")),
                volume=volume,
                amount=finite_float(row.get("成交额")) or 0.0,
                scale=finite_float(row.get("规模")) or 0.0,
            )
            key = (current_date, index_code)
            existing = best_by_date_index.get(key)
            if existing is None or candidate_sort_key(candidate) < candidate_sort_key(existing):
                best_by_date_index[key] = candidate

    targets: dict[date, DailyTarget] = {}
    for signal_date, selection in sorted(daily_selections.items()):
        mapped: list[TargetMember] = []
        for member in selection.members:
            candidate = best_by_date_index.get((signal_date, member.index_code))
            if candidate is None:
                continue
            mapped.append(
                TargetMember(
                    signal_date=signal_date,
                    index_code=member.index_code,
                    index_name=member.index_name or candidate.index_name,
                    etf_code=candidate.code,
                    etf_name=candidate.name,
                    target_weight=0.0,
                    trend_factor=member.trend_factor,
                    factor_rank=member.factor_rank,
                    selection_volume=candidate.volume,
                    selection_amount=candidate.amount,
                    selection_scale=candidate.scale,
                )
            )
        # 此权重仅描述当日新资金内部的等权，不是整个组合的目标权重。
        mapped = [replace(member, target_weight=1.0 / len(mapped)) for member in mapped]
        targets[signal_date] = DailyTarget(
            signal_date=signal_date,
            planned_index_count=selection.planned_index_count,
            selected_index_count=len(selection.members),
            members=tuple(mapped),
        )
    if not trading_calendar:
        raise ValueError("etf_data.csv中没有有效交易日期")
    return targets, sorted(trading_calendar)


def load_selected_prices(
    selected_codes: set[str],
    required_dates: set[date],
    etf_data_file: Path = ETF_DATA_FILE,
) -> dict[date, dict[str, PricePoint]]:
    """第二次扫描总表，只把真正可能交易的ETF价格载入内存。"""

    prices: dict[date, dict[str, PricePoint]] = defaultdict(dict)
    with etf_data_file.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = ETF_REQUIRED_COLUMNS - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"etf_data.csv缺少列：{sorted(missing)}")
        for row in reader:
            code = clean_text(row.get("代码")).upper()
            if code not in selected_codes:
                continue
            row_date_text = clean_text(row.get("日期"))[:10]
            try:
                current_date = date.fromisoformat(row_date_text)
            except ValueError:
                continue
            if current_date not in required_dates:
                continue
            if code in prices[current_date]:
                raise ValueError(f"etf_data.csv存在重复ETF日期：{current_date} {code}")
            prices[current_date][code] = PricePoint(
                close=positive_float(row.get("收盘价")),
                vwap=positive_float(row.get("VWAP")),
                amount=positive_float(row.get("成交额")),
            )
    return dict(prices)


def linear_quantile(values: Sequence[float], quantile: float) -> float | None:
    """按NumPy默认的线性插值口径计算分位数。"""

    valid_values = sorted(value for value in values if math.isfinite(value))
    if not valid_values:
        return None
    if len(valid_values) == 1:
        return valid_values[0]
    position = (len(valid_values) - 1) * quantile
    lower_index = math.floor(position)
    upper_index = math.ceil(position)
    if lower_index == upper_index:
        return valid_values[lower_index]
    fraction = position - lower_index
    return (
        valid_values[lower_index] * (1.0 - fraction)
        + valid_values[upper_index] * fraction
    )


def trade_batches(
    positions: dict[int, Position],
    cash: float,
    target: DailyTarget | None,
    signal_date: date,
    signal_position: int,
    execution_date: date,
    execution_position: int,
    signal_prices: Mapping[str, PricePoint],
    execution_prices: Mapping[str, float],
    next_batch_id: int,
) -> tuple[RebalanceResult, int, bool]:
    """先逐批次卖出，再仅将现金等额投入本轮信号名单；旧批次不调权。"""
    selected = {member.etf_code: member for member in target.members} if target else {}
    blocked_codes: set[str] = set()
    survivors: dict[int, Position] = {}
    details: list[TradeDetail] = []
    skipped: list[str] = []
    buy_amount = sell_amount = fee = 0.0
    attempted = False

    for batch_id, position in positions.items():
        point = signal_prices.get(position.etf_code)
        signal_close = point.close if point else None
        age = signal_position - position.buy_date_position
        reason = position.pending_exit_reason
        fixed_stop_price = position.buy_price * (1.0 - STOP_LOSS_PCT)
        if signal_close is not None:
            position.highest_close = max(position.highest_close, signal_close)
        trailing_active = (
            position.highest_close
            >= position.buy_price * (1.0 + PROFIT_TRAILING_TRIGGER_PCT)
        )
        trailing_stop_price = (
            position.highest_close * (1.0 - PROFIT_TRAILING_DRAWDOWN_PCT)
            if trailing_active
            else None
        )
        fixed_stop_triggered = (
            signal_close is not None and signal_close <= fixed_stop_price
        )
        trailing_stop_triggered = (
            signal_close is not None
            and trailing_stop_price is not None
            and signal_close <= trailing_stop_price
        )
        risk_reason = (
            "固定止损" if fixed_stop_triggered
            else "浮盈回撤" if trailing_stop_triggered
            else ""
        )
        if risk_reason and reason != "固定止损":
            # 未成交的普通卖单或浮盈回撤卖单如再触发更高优先级风控，升级原因。
            if reason != risk_reason:
                position.pending_exit_reason = reason = risk_reason
                position.exit_signal_date = signal_date
                position.exit_signal_price = signal_close
                position.exit_signal_age = age
                position.exit_signal_peak_price = position.highest_close
                position.exit_risk_price = (
                    fixed_stop_price
                    if risk_reason == "固定止损"
                    else trailing_stop_price
                )
        elif not reason and target is not None and age >= MIN_HOLD_DAYS and position.etf_code not in selected:
            position.pending_exit_reason = reason = "到期未入选"
            position.exit_signal_date = signal_date
            position.exit_signal_price = signal_close
            position.exit_signal_age = age
            position.exit_signal_peak_price = position.highest_close
            position.exit_risk_price = None

        if reason in {"固定止损", "浮盈回撤"}:
            blocked_codes.add(position.etf_code)
        if not reason:
            survivors[batch_id] = position
            continue
        attempted = True
        price = execution_prices.get(position.etf_code)
        if price is None:
            # 卖单保留到后续交易日；不得把未成交卖单当成可用现金。
            survivors[batch_id] = position
            skipped.append(f"批次{batch_id}/{position.etf_code}卖出缺少成交价")
            continue
        value = position.shares * price
        cost = value * TRANSACTION_COST_RATE
        cash += value - cost
        sell_amount += value
        fee += cost
        details.append(TradeDetail(
            etf_code=position.etf_code, etf_name=position.etf_name,
            direction="卖出", before_value=value, target_value=0.0,
            trade_amount=value, transaction_cost=cost, batch_id=batch_id,
            buy_date=position.buy_date, buy_price=position.buy_price,
            shares=position.shares, execution_price=price, reason=reason,
            signal_date=position.exit_signal_date,
            signal_price=position.exit_signal_price,
            signal_holding_days=position.exit_signal_age,
            signal_peak_price=position.exit_signal_peak_price,
            risk_exit_price=position.exit_risk_price,
        ))

    candidates = [member for code, member in selected.items() if code not in blocked_codes]
    if candidates and cash > 1e-15:
        attempted = True
        # 在看到执行价格前固定等额预算。某只无法成交，其预算留现金，不事后补选。
        budget = cash / len(candidates)
        allocation_weight = 1.0 / len(candidates)
        for member in candidates:
            price = execution_prices.get(member.etf_code)
            if price is None:
                skipped.append(f"{member.etf_code}买入缺少成交价")
                continue
            value = budget / (1.0 + TRANSACTION_COST_RATE)
            cost = value * TRANSACTION_COST_RATE
            batch_id = next_batch_id
            next_batch_id += 1
            position = Position(
                shares=value / price, index_code=member.index_code,
                index_name=member.index_name, etf_name=member.etf_name,
                signal_date=signal_date, trend_factor=member.trend_factor,
                factor_rank=member.factor_rank, selection_volume=member.selection_volume,
                target_weight=allocation_weight, etf_code=member.etf_code,
                batch_id=batch_id, buy_date=execution_date, buy_price=price,
                buy_date_position=execution_position, highest_close=price,
            )
            survivors[batch_id] = position
            cash -= value + cost
            buy_amount += value
            fee += cost
            point = signal_prices.get(member.etf_code)
            details.append(TradeDetail(
                etf_code=member.etf_code, etf_name=member.etf_name,
                direction="买入", before_value=0.0, target_value=value,
                trade_amount=value, transaction_cost=cost, batch_id=batch_id,
                buy_date=execution_date, buy_price=price, shares=position.shares,
                execution_price=price, reason="可用现金等额买入", signal_date=signal_date,
                signal_price=point.close if point else None, signal_holding_days=0,
                signal_peak_price=None, risk_exit_price=None,
            ))
    if cash < -1e-12:
        raise RuntimeError(f"交易后现金为负：{cash}")
    return RebalanceResult(
        positions=survivors, cash=max(cash, 0.0), pre_trade_nav=0.0,
        buy_amount=buy_amount, sell_amount=sell_amount, transaction_cost=fee,
        trade_details=tuple(details), succeeded=attempted and not skipped,
        skip_reason="；".join(skipped),
    ), next_batch_id, attempted


def run_backtest(
    mode: str,
    trading_dates: Sequence[date],
    targets: Mapping[date, DailyTarget],
    prices: Mapping[date, Mapping[str, PricePoint]],
) -> tuple[
    list[HoldingRecord], list[NavRecord], list[AccountDailyRecord],
    list[AccountHoldingRecord], list[AccountTradeRecord],
]:
    if mode not in {"close", "next_day_vwap"}:
        raise ValueError(f"未知回测模式：{mode}")
    if not trading_dates or list(trading_dates) != sorted(set(trading_dates)):
        raise ValueError("交易日历必须非空、严格递增且无重复")
    positions: dict[int, Position] = {}
    cash = INITIAL_NAV
    next_batch_id = 1
    last_closes: dict[str, float] = {}
    previous_nav = running_peak = INITIAL_NAV
    cumulative_cost_rate = 0.0
    has_been_built = False
    last_trade_date: date | None = None
    holdings: list[HoldingRecord] = []
    nav_records: list[NavRecord] = []
    daily_records: list[AccountDailyRecord] = []
    batch_records: list[AccountHoldingRecord] = []
    trade_records: list[AccountTradeRecord] = []

    for date_position, current_date in enumerate(trading_dates):
        daily_prices = prices.get(current_date, {})
        close_prices = {code: p.close for code, p in daily_prices.items() if p.close is not None}
        execution_prices = {
            code: price for code, point in daily_prices.items()
            if (price := (point.close if mode == "close" else point.vwap)) is not None
            and point.amount is not None and point.amount > 0
        }
        # close保留为原脚本的同收盘理想化对照；主口径始终用前一日信号、次日VWAP。
        signal_position = date_position if mode == "close" else date_position - 1
        signal_date = trading_dates[signal_position] if signal_position >= 0 else None
        target = targets.get(signal_date) if signal_date is not None else None
        attempted = succeeded = initial_build = False
        buy_amount = sell_amount = transaction_cost = 0.0
        skip_reason = ""
        missing_price_count = 0
        if signal_date is not None:
            result, next_batch_id, attempted = trade_batches(
                positions, cash, target, signal_date, signal_position,
                current_date, date_position, prices.get(signal_date, {}),
                execution_prices, next_batch_id,
            )
            positions, cash = result.positions, result.cash
            buy_amount, sell_amount = result.buy_amount, result.sell_amount
            transaction_cost = result.transaction_cost
            succeeded, skip_reason = result.succeeded, result.skip_reason
            if target is None:
                skip_reason = "缺少当日因子信号，仅处理风控退出及未成交卖单" + ("；" + skip_reason if skip_reason else "")
            if target:
                missing_price_count = sum(m.etf_code not in execution_prices for m in target.members)
            initial_build = not has_been_built and buy_amount > 0
            has_been_built = has_been_built or initial_build
            if result.trade_details:
                last_trade_date = current_date
            for detail in result.trade_details:
                trade_records.append(AccountTradeRecord(
                    execution_date=current_date, account_id=1,
                    signal_date=detail.signal_date, etf_code=detail.etf_code,
                    etf_name=detail.etf_name, direction=detail.direction,
                    before_value=detail.before_value, target_value=detail.target_value,
                    trade_amount=detail.trade_amount, transaction_cost=detail.transaction_cost,
                    batch_id=detail.batch_id, buy_date=detail.buy_date,
                    buy_price=detail.buy_price, shares=detail.shares,
                    execution_price=detail.execution_price, reason=detail.reason,
                    signal_price=detail.signal_price,
                    signal_holding_days=detail.signal_holding_days,
                    signal_peak_price=detail.signal_peak_price,
                    risk_exit_price=detail.risk_exit_price,
                ))

        last_closes.update(close_prices)
        combined_values: dict[str, float] = defaultdict(float)
        combined_positions: dict[str, Position] = {}
        batch_values: dict[int, float] = {}
        for batch_id, position in positions.items():
            # 当日缺失收盘价时沿用最近收盘；新买入且无历史估值时用实际成交价。
            close = last_closes.get(position.etf_code, position.buy_price)
            current_close = close_prices.get(position.etf_code)
            if current_close is not None:
                position.highest_close = max(position.highest_close, current_close)
            value = position.shares * close
            batch_values[batch_id] = value
            combined_values[position.etf_code] += value
            combined_positions[position.etf_code] = position
        etf_value = sum(combined_values.values())
        nav = cash + etf_value
        daily_return = nav / previous_nav - 1.0
        cost_rate = transaction_cost / previous_nav
        cumulative_cost_rate = 1.0 - (1.0 - cumulative_cost_rate) * (1.0 - cost_rate)
        running_peak = max(running_peak, nav)
        capacity_samples = []
        for code, value in combined_values.items():
            point = daily_prices.get(code)
            weight = value / nav
            if point and point.amount and weight > 1e-6:
                capacity_samples.append(CAPACITY_DAILY_AMOUNT_RATIO * point.amount / weight)
        capacity = linear_quantile(capacity_samples, 1.0 - CAPACITY_DESCENDING_QUANTILE)
        nav_records.append(NavRecord(
            current_date=current_date, signal_date=signal_date,
            pre_trade_nav=previous_nav, nav=nav,
            gross_return=daily_return + cost_rate, daily_return=daily_return,
            drawdown=1.0 - nav / running_peak, buy_ratio=buy_amount / previous_nav,
            sell_ratio=sell_amount / previous_nav,
            bilateral_ratio=(buy_amount + sell_amount) / previous_nav,
            one_way_turnover=(buy_amount + sell_amount) / (2.0 * previous_nav),
            transaction_cost=transaction_cost, transaction_cost_rate=cost_rate,
            cumulative_cost_rate=cumulative_cost_rate, capacity=capacity,
            holding_count=len(combined_values), cash_weight=cash / nav,
            initial_build=initial_build, rebalance_account_id=1 if attempted else None,
            rebalance_attempted=attempted, rebalance_succeeded=succeeded,
            skip_reason=skip_reason,
            selected_index_count=target.selected_index_count if target else 0,
            unmapped_index_count=target.unmapped_index_count if target else 0,
            missing_price_index_count=missing_price_count,
        ))
        for code, value in sorted(combined_values.items()):
            p = combined_positions[code]
            holdings.append(HoldingRecord(
                current_date=current_date, signal_date=p.signal_date,
                index_code=p.index_code, index_name=p.index_name,
                trend_factor=p.trend_factor, factor_rank=p.factor_rank,
                etf_code=code, etf_name=p.etf_name, selection_volume=p.selection_volume,
                target_weight=value / nav, actual_weight=value / nav,
            ))
        daily_records.append(AccountDailyRecord(
            current_date=current_date, account_id=1, account_nav=nav,
            etf_market_value=etf_value, cash=cash, cash_weight=cash / nav,
            rebalance_attempted=attempted, rebalance_succeeded=succeeded,
            signal_date=signal_date, last_rebalance_date=last_trade_date,
            next_rebalance_date=trading_dates[date_position + 1] if date_position + 1 < len(trading_dates) else None,
        ))
        for batch_id, p in sorted(positions.items()):
            value = batch_values[batch_id]
            trailing_active = (
                p.highest_close
                >= p.buy_price * (1.0 + PROFIT_TRAILING_TRIGGER_PCT)
            )
            batch_records.append(AccountHoldingRecord(
                current_date=current_date, account_id=1, signal_date=p.signal_date,
                etf_code=p.etf_code, etf_name=p.etf_name, market_value=value,
                account_weight=value / nav, total_portfolio_weight=value / nav,
                batch_id=batch_id, buy_date=p.buy_date, buy_price=p.buy_price,
                shares=p.shares, holding_days=date_position - p.buy_date_position,
                stop_price=p.buy_price * (1.0 - STOP_LOSS_PCT),
                highest_close=p.highest_close,
                max_return=p.highest_close / p.buy_price - 1.0,
                trailing_active=trailing_active,
                trailing_stop_price=(
                    p.highest_close * (1.0 - PROFIT_TRAILING_DRAWDOWN_PCT)
                    if trailing_active
                    else None
                ),
                pending_exit_reason=p.pending_exit_reason,
            ))
        previous_nav = nav
    return holdings, nav_records, daily_records, batch_records, trade_records


def calculate_performance(nav_records: Sequence[NavRecord]) -> dict[str, object]:
    if not nav_records:
        raise ValueError("没有净值记录，无法计算绩效")
    periods = len(nav_records)
    daily_returns = [record.daily_return for record in nav_records]
    cumulative_growth = math.prod(1.0 + value for value in daily_returns)
    cumulative_return = cumulative_growth - 1.0
    annual_return = (
        cumulative_growth ** (ANNUAL_TRADING_DAYS / periods) - 1.0
        if cumulative_growth > 0
        else float("nan")
    )
    daily_std = statistics.stdev(daily_returns) if len(daily_returns) >= 2 else 0.0
    annual_volatility = daily_std * math.sqrt(ANNUAL_TRADING_DAYS)
    daily_risk_free_rate = (
        (1.0 + ANNUAL_RISK_FREE_RATE) ** (1.0 / ANNUAL_TRADING_DAYS) - 1.0
    )
    excess_returns = [value - daily_risk_free_rate for value in daily_returns]
    sharpe = (
        statistics.fmean(excess_returns)
        / daily_std
        * math.sqrt(ANNUAL_TRADING_DAYS)
        if daily_std > 0
        else 0.0
    )
    downside_deviation = math.sqrt(
        statistics.fmean(min(value, 0.0) ** 2 for value in excess_returns)
    )
    sortino = (
        statistics.fmean(excess_returns)
        / downside_deviation
        * math.sqrt(ANNUAL_TRADING_DAYS)
        if downside_deviation > 0
        else 0.0
    )

    local_nav = 1.0
    running_peak = 1.0
    peak_date = nav_records[0].current_date
    max_drawdown = 0.0
    max_drawdown_start = peak_date
    max_drawdown_end = peak_date
    for record in nav_records:
        local_nav *= 1.0 + record.daily_return
        if local_nav >= running_peak:
            running_peak = local_nav
            peak_date = record.current_date
        drawdown = 1.0 - local_nav / running_peak
        if drawdown > max_drawdown:
            max_drawdown = drawdown
            max_drawdown_start = peak_date
            max_drawdown_end = record.current_date

    turnover_records = [record for record in nav_records if not record.initial_build]
    average_turnover = (
        statistics.fmean(record.one_way_turnover for record in turnover_records)
        if turnover_records
        else 0.0
    )
    years = periods / ANNUAL_TRADING_DAYS
    annual_turnover = (
        sum(record.one_way_turnover for record in turnover_records) / years
        if years > 0
        else 0.0
    )
    cumulative_cost_rate = 1.0 - math.prod(
        1.0 - record.transaction_cost_rate for record in nav_records
    )

    return {
        "回测开始日": nav_records[0].current_date.isoformat(),
        "回测结束日": nav_records[-1].current_date.isoformat(),
        "交易日数量": periods,
        "初始净值": INITIAL_NAV,
        "期末净值": nav_records[-1].nav,
        "累计收益率": cumulative_return,
        "年化收益率": annual_return,
        "年化波动率": annual_volatility,
        "夏普比率": sharpe,
        "Sortino比率": sortino,
        "最大回撤": max_drawdown,
        "最大回撤开始日": max_drawdown_start.isoformat(),
        "最大回撤结束日": max_drawdown_end.isoformat(),
        "Calmar比率": annual_return / max_drawdown if max_drawdown > 0 else 0.0,
        "平均每日单边换手率": average_turnover,
        "年化单边换手率（倍）": annual_turnover,
        "累计交易成本率": cumulative_cost_rate,
        "平均持仓ETF数量": statistics.fmean(
            record.holding_count for record in nav_records
        ),
        "调仓次数": sum(record.rebalance_succeeded for record in nav_records),
        "跳过调仓次数": sum(
            record.rebalance_attempted and not record.rebalance_succeeded
            for record in nav_records
        ),
        "无法映射指数-日期记录数": sum(
            record.unmapped_index_count for record in nav_records
        ),
        "缺少成交价指数-日期记录数": sum(
            record.missing_price_index_count for record in nav_records
        ),
    }


def calculate_return_performance(
    dates: Sequence[date],
    daily_returns: Sequence[float],
    *,
    annual_risk_free_rate: float = ANNUAL_RISK_FREE_RATE,
) -> dict[str, object]:
    """按现有策略绩效口径计算一组纯收益序列。"""

    if not dates or len(dates) != len(daily_returns):
        raise ValueError("收益序列日期为空或长度不一致")
    periods = len(daily_returns)
    cumulative_growth = math.prod(1.0 + value for value in daily_returns)
    cumulative_return = cumulative_growth - 1.0
    annual_return = (
        cumulative_growth ** (ANNUAL_TRADING_DAYS / periods) - 1.0
        if cumulative_growth > 0
        else float("nan")
    )
    daily_std = statistics.stdev(daily_returns) if periods >= 2 else 0.0
    annual_volatility = daily_std * math.sqrt(ANNUAL_TRADING_DAYS)
    daily_risk_free_rate = (
        (1.0 + annual_risk_free_rate) ** (1.0 / ANNUAL_TRADING_DAYS) - 1.0
    )
    excess_returns = [value - daily_risk_free_rate for value in daily_returns]
    sharpe = (
        statistics.fmean(excess_returns)
        / daily_std
        * math.sqrt(ANNUAL_TRADING_DAYS)
        if daily_std > 0
        else 0.0
    )
    downside_deviation = math.sqrt(
        statistics.fmean(min(value, 0.0) ** 2 for value in excess_returns)
    )
    sortino = (
        statistics.fmean(excess_returns)
        / downside_deviation
        * math.sqrt(ANNUAL_TRADING_DAYS)
        if downside_deviation > 0
        else 0.0
    )

    local_nav = 1.0
    running_peak = 1.0
    peak_date = dates[0]
    max_drawdown = 0.0
    max_drawdown_start = peak_date
    max_drawdown_end = peak_date
    for current_date, daily_return in zip(dates, daily_returns):
        local_nav *= 1.0 + daily_return
        if local_nav >= running_peak:
            running_peak = local_nav
            peak_date = current_date
        drawdown = 1.0 - local_nav / running_peak
        if drawdown > max_drawdown:
            max_drawdown = drawdown
            max_drawdown_start = peak_date
            max_drawdown_end = current_date

    return {
        "累计收益率": cumulative_return,
        "年化收益率": annual_return,
        "年化波动率": annual_volatility,
        "夏普比率": sharpe,
        "Sortino比率": sortino,
        "最大回撤": max_drawdown,
        "最大回撤开始日": max_drawdown_start.isoformat(),
        "最大回撤结束日": max_drawdown_end.isoformat(),
        "Calmar比率": annual_return / max_drawdown if max_drawdown > 0 else 0.0,
    }


def calculate_relative_performance(
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
) -> dict[str, object]:
    """计算基准、几何超额、跟踪误差和信息比率。"""

    if len(nav_records) != len(benchmark_records) or not nav_records:
        raise ValueError("策略与基准记录为空或长度不一致")
    for nav_record, benchmark_record in zip(nav_records, benchmark_records):
        if nav_record.current_date != benchmark_record.current_date:
            raise ValueError("策略与基准日期没有逐日对齐")

    dates = [record.current_date for record in benchmark_records]
    benchmark_returns = [record.daily_return for record in benchmark_records]
    relative_returns = [record.relative_return for record in benchmark_records]
    active_returns = [record.active_return for record in benchmark_records]
    benchmark_performance = calculate_return_performance(dates, benchmark_returns)
    relative_performance = calculate_return_performance(
        dates,
        relative_returns,
        annual_risk_free_rate=0.0,
    )

    active_std = (
        statistics.stdev(active_returns) if len(active_returns) >= 2 else 0.0
    )
    tracking_error = active_std * math.sqrt(ANNUAL_TRADING_DAYS)
    information_ratio = (
        statistics.fmean(active_returns)
        / active_std
        * math.sqrt(ANNUAL_TRADING_DAYS)
        if active_std > 0
        else 0.0
    )

    return {
        "基准累计收益率": benchmark_performance["累计收益率"],
        "基准年化收益率": benchmark_performance["年化收益率"],
        "基准年化波动率": benchmark_performance["年化波动率"],
        "基准夏普比率": benchmark_performance["夏普比率"],
        "基准Sortino比率": benchmark_performance["Sortino比率"],
        "基准最大回撤": benchmark_performance["最大回撤"],
        "基准最大回撤开始日": benchmark_performance["最大回撤开始日"],
        "基准最大回撤结束日": benchmark_performance["最大回撤结束日"],
        "基准Calmar比率": benchmark_performance["Calmar比率"],
        "超额累计收益率": relative_performance["累计收益率"],
        "超额年化收益率": relative_performance["年化收益率"],
        "超额最大回撤": relative_performance["最大回撤"],
        "超额最大回撤开始日": relative_performance["最大回撤开始日"],
        "超额最大回撤结束日": relative_performance["最大回撤结束日"],
        "跟踪误差": tracking_error,
        "信息比率": information_ratio,
    }


MODE_LABELS = {
    "close": "收盘价成交",
    "next_day_vwap": "次日VWAP成交",
}


def build_annual_metrics(
    mode: str,
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
) -> list[dict[str, object]]:
    by_year: dict[int, list[NavRecord]] = defaultdict(list)
    benchmark_by_year: dict[int, list[BenchmarkRecord]] = defaultdict(list)
    for record in nav_records:
        by_year[record.current_date.year].append(record)
    for record in benchmark_records:
        benchmark_by_year[record.current_date.year].append(record)

    rows: list[dict[str, object]] = []
    for year in sorted(by_year):
        records = by_year[year]
        performance = calculate_performance(records)
        relative_performance = calculate_relative_performance(
            records,
            benchmark_by_year[year],
        )
        start = records[0].current_date
        end = records[-1].current_date
        if start.month > 1:
            year_label = f"{year}年{start:%m-%d}起"
        elif end.month < 12:
            year_label = f"{year}年截至{end:%m-%d}"
        else:
            year_label = f"{year}年"
        rows.append(
            {
                "成交方式": MODE_LABELS[mode],
                "年份": year_label,
                "交易日数量": performance["交易日数量"],
                "年度收益率": performance["累计收益率"],
                "基准收益率": relative_performance["基准累计收益率"],
                "超额收益率": relative_performance["超额累计收益率"],
                "年化收益率": performance["年化收益率"],
                "年化波动率": performance["年化波动率"],
                "跟踪误差": relative_performance["跟踪误差"],
                "夏普比率": performance["夏普比率"],
                "Sortino比率": performance["Sortino比率"],
                "信息比率": relative_performance["信息比率"],
                "最大回撤": performance["最大回撤"],
                "基准最大回撤": relative_performance["基准最大回撤"],
                "超额最大回撤": relative_performance["超额最大回撤"],
                "最大回撤开始日": performance["最大回撤开始日"],
                "最大回撤结束日": performance["最大回撤结束日"],
                "Calmar比率": performance["Calmar比率"],
                "平均每日单边换手率": performance["平均每日单边换手率"],
                "年化单边换手率（倍）": performance["年化单边换手率（倍）"],
                "累计交易成本率": performance["累计交易成本率"],
                "平均持仓ETF数量": performance["平均持仓ETF数量"],
            }
        )
    return rows


def build_validation_rows(
    mode: str,
    holding_records: Sequence[HoldingRecord],
    nav_records: Sequence[NavRecord],
) -> list[dict[str, object]]:
    label = mode_execution_label(mode)
    holdings_by_date: dict[date, float] = defaultdict(float)
    for record in holding_records:
        holdings_by_date[record.current_date] += record.actual_weight

    nav_rebuild_error = 0.0
    previous_nav = INITIAL_NAV
    weight_error = 0.0
    cost_error = 0.0
    signal_errors = 0
    finite_errors = 0
    for record in nav_records:
        rebuilt_nav = previous_nav * (1.0 + record.daily_return)
        nav_rebuild_error = max(nav_rebuild_error, abs(rebuilt_nav - record.nav))
        previous_nav = record.nav
        weight_error = max(
            weight_error,
            abs(holdings_by_date.get(record.current_date, 0.0) + record.cash_weight - 1.0),
        )
        cost_error = max(
            cost_error,
            abs(
                record.transaction_cost_rate
                - TRANSACTION_COST_RATE * record.bilateral_ratio
            ),
        )
        if record.signal_date is not None:
            if mode == "close" and record.signal_date != record.current_date:
                signal_errors += 1
            if mode == "next_day_vwap" and record.signal_date >= record.current_date:
                signal_errors += 1
        numeric_values = (
            record.nav,
            record.daily_return,
            record.one_way_turnover,
            record.transaction_cost_rate,
            record.cash_weight,
        )
        finite_errors += sum(not math.isfinite(value) for value in numeric_values)

    initial_build_accounts = [
        record.rebalance_account_id
        for record in nav_records
        if record.initial_build and record.rebalance_account_id is not None
    ]
    initial_build_count = len(initial_build_accounts)
    initial_build_valid = (
        len(initial_build_accounts) == len(set(initial_build_accounts))
        and initial_build_count <= ACCOUNT_COUNT
    )
    skipped_count = sum(
        record.rebalance_attempted and not record.rebalance_succeeded
        for record in nav_records
    )

    def check_row(
        item: str,
        actual: object,
        expected: object,
        difference: float,
        tolerance: float,
        passed: bool,
        notes: str,
    ) -> dict[str, object]:
        return {
            "成交方式": label,
            "检查项": item,
            "实际值": actual,
            "期望值": expected,
            "差异": difference,
            "容差": tolerance,
            "状态": "OK" if passed else "FAIL",
            "说明": notes,
        }

    return [
        check_row(
            "净值可由日收益重建",
            nav_rebuild_error,
            0.0,
            nav_rebuild_error,
            1e-12,
            nav_rebuild_error <= 1e-12,
            "逐日用前一日净值×(1+净收益率)重建",
        ),
        check_row(
            "持仓权重与现金权重合计为1",
            weight_error,
            0.0,
            weight_error,
            2e-10,
            weight_error <= 2e-10,
            "使用每日收盘估值后的实际权重",
        ),
        check_row(
            "交易成本率与双边成交额一致",
            cost_error,
            0.0,
            cost_error,
            1e-12,
            cost_error <= 1e-12,
            "成本率=0.1%×(买入比例+卖出比例)",
        ),
        check_row(
            "信号与成交日期符合模式约定",
            signal_errors,
            0,
            float(signal_errors),
            0.0,
            signal_errors == 0,
            "close是同收盘理想化对照；次日VWAP严格使用此前收盘信号",
        ),
        check_row(
            "关键日序列均为有限数值",
            finite_errors,
            0,
            float(finite_errors),
            0.0,
            finite_errors == 0,
            "净值、收益、换手、成本和现金权重",
        ),
        check_row(
            "各账户初始建仓单独标记",
            initial_build_count,
            "每个账户至多1次",
            0.0 if initial_build_valid else 1.0,
            0.0,
            initial_build_valid,
            f"{ACCOUNT_COUNT}个账户分别至多初始建仓1次；收费，但不计入平均和年化换手率",
        ),
        check_row(
            "跳过调仓次数（信息项）",
            skipped_count,
            "仅记录",
            0.0,
            0.0,
            True,
            "信息项：记录回测中未成功执行的调仓次数",
        ),
    ]


def style_worksheet(
    sheet: Any,
    percent_headers: set[str] | None = None,
    decimal_headers: set[str] | None = None,
    integer_headers: set[str] | None = None,
    max_width: int = 28,
) -> None:
    percent_headers = percent_headers or set()
    decimal_headers = decimal_headers or set()
    integer_headers = integer_headers or set()
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(name="Arial", size=11, color="FFFFFF", bold=True)
    body_font = Font(name="Arial", size=10)
    thin_gray = Side(style="thin", color="D9D9D9")
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray
        )
    sheet.row_dimensions[1].height = 24
    headers = {cell.value: cell.column for cell in sheet[1]}
    for header in percent_headers:
        column = headers.get(header)
        if column:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=column).number_format = "0.00%"
    for header in decimal_headers:
        column = headers.get(header)
        if column:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=column).number_format = "0.000000"
    for header in integer_headers:
        column = headers.get(header)
        if column:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=column).number_format = "#,##0"
    for column_number in range(1, sheet.max_column + 1):
        column_letter = get_column_letter(column_number)
        values = [sheet.cell(row=row, column=column_number).value for row in range(1, min(sheet.max_row, 300) + 1)]
        width = min(
            max((len(str(value)) for value in values if value is not None), default=10) + 2,
            max_width,
        )
        sheet.column_dimensions[column_letter].width = max(width, 11)
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
            cell.border = Border(
                left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray
            )
            cell.alignment = Alignment(
                horizontal="center" if isinstance(cell.value, str) else "right",
                vertical="center",
            )


def save_figure_atomic(figure: Any, path: Path) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    try:
        figure.savefig(
            temporary,
            format="png",
            dpi=180,
            bbox_inches="tight",
            facecolor="white",
        )
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
        plt.close(figure)



def save_workbook_atomic(workbook: Workbook, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.stem}.tmp.xlsx")
    try:
        workbook.save(temporary)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
        workbook.close()
    return path


def output_prefix(mode: str) -> str:
    return (
        f"daily_hold_{MIN_HOLD_DAYS}d_stop_{STOP_LOSS_PCT * 100:g}pct"
        f"_profit_{PROFIT_TRAILING_TRIGGER_PCT * 100:g}pct"
        f"_trail_{PROFIT_TRAILING_DRAWDOWN_PCT * 100:g}pct_{mode}"
    )


def mode_execution_label(mode: str) -> str:
    execution = "同收盘价（理想化对照）" if mode == "close" else "次日VWAP"
    return (
        f"{execution}；每日轮动 / 最短{MIN_HOLD_DAYS}日 / 固定止损{STOP_LOSS_PCT:.1%}"
        f" / 浮盈{PROFIT_TRAILING_TRIGGER_PCT:.1%}后回撤"
        f"{PROFIT_TRAILING_DRAWDOWN_PCT:.1%}退出"
    )


def write_annual_metrics_workbook(
    mode: str,
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
    output_dir: Path,
) -> Path:
    """按参考项目单独输出年度策略、基准和超额指标。"""

    headers = [
        "年份",
        "策略收益",
        "基准收益",
        "超额收益",
        "年化波动",
        "跟踪误差",
        "Sharpe",
        "Sortino",
        "信息比率",
        "策略最大回撤",
        "基准最大回撤",
        "超额最大回撤",
        "Calmar",
        "年化换手率（单边）",
    ]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "annual_metrics"
    sheet.append(headers)
    for row in build_annual_metrics(mode, nav_records, benchmark_records):
        sheet.append(
            [
                row["年份"],
                row["年度收益率"],
                row["基准收益率"],
                row["超额收益率"],
                row["年化波动率"],
                row["跟踪误差"],
                row["夏普比率"],
                row["Sortino比率"],
                row["信息比率"],
                row["最大回撤"],
                row["基准最大回撤"],
                row["超额最大回撤"],
                row["Calmar比率"],
                row["年化单边换手率（倍）"],
            ]
        )
    style_worksheet(
        sheet,
        percent_headers={
            "策略收益", "基准收益", "超额收益", "年化波动", "跟踪误差",
            "策略最大回撤", "基准最大回撤", "超额最大回撤",
        },
        decimal_headers={"Sharpe", "Sortino", "信息比率", "Calmar"},
        max_width=24,
    )
    turnover_column = headers.index("年化换手率（单边）") + 1
    for row_number in range(2, sheet.max_row + 1):
        sheet.cell(row=row_number, column=turnover_column).number_format = '0.00"倍"'
    return save_workbook_atomic(
        workbook,
        output_dir / f"{output_prefix(mode)}_annual_metrics.xlsx",
    )


def write_backtest_metrics_workbook(
    mode: str,
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
    benchmark: BenchmarkData,
    output_dir: Path,
    score_method: str,
) -> Path:
    """按参考项目单独输出总回测指标和参数。"""

    performance = calculate_performance(nav_records)
    relative_performance = calculate_relative_performance(
        nav_records,
        benchmark_records,
    )
    attempted = sum(record.rebalance_attempted for record in nav_records)
    succeeded = sum(record.rebalance_succeeded for record in nav_records)
    success_rate = succeeded / attempted if attempted else 0.0
    drawdown_range = (
        f"{datetime.fromisoformat(str(performance['最大回撤开始日'])).strftime('%Y年%m月')}"
        f"-{datetime.fromisoformat(str(performance['最大回撤结束日'])).strftime('%Y年%m月')}"
    )
    relative_drawdown_range = (
        f"{datetime.fromisoformat(str(relative_performance['超额最大回撤开始日'])).strftime('%Y年%m月')}"
        f"-{datetime.fromisoformat(str(relative_performance['超额最大回撤结束日'])).strftime('%Y年%m月')}"
    )

    workbook = Workbook()
    performance_sheet = workbook.active
    performance_sheet.title = "performance"
    performance_sheet.append(["类别", "指标", "ETF趋势策略"])
    performance_rows = [
        [
            "基本信息",
            "回测区间",
            f"{performance['回测开始日']}至{performance['回测结束日']}",
        ],
        ["基本信息", "交易日数量", performance["交易日数量"]],
        ["基本信息", "初始净值", performance["初始净值"]],
        ["基本信息", "期末净值", performance["期末净值"]],
        ["基本信息", "累计收益率", performance["累计收益率"]],
        ["收益表现", "策略年化收益率", performance["年化收益率"]],
        [
            "收益表现",
            f"{benchmark.name}年化收益率",
            relative_performance["基准年化收益率"],
        ],
        ["收益表现", "年化超额收益", relative_performance["超额年化收益率"]],
        ["绝对风险收益", "策略年化波动率", performance["年化波动率"]],
        ["绝对风险收益", "Sharpe", performance["夏普比率"]],
        ["绝对风险收益", "Sortino", performance["Sortino比率"]],
        ["绝对风险收益", "策略最大回撤", performance["最大回撤"]],
        ["绝对风险收益", "策略最大回撤区间", drawdown_range],
        ["绝对风险收益", "Calmar", performance["Calmar比率"]],
        ["相对基准表现", "年化跟踪误差", relative_performance["跟踪误差"]],
        ["相对基准表现", "IR", relative_performance["信息比率"]],
        ["相对基准表现", "超额最大回撤", relative_performance["超额最大回撤"]],
        ["相对基准表现", "超额最大回撤区间", relative_drawdown_range],
        ["交易与组合", "年化换手率（单边）", performance["年化单边换手率（倍）"]],
        ["交易与组合", "平均单次换仓比率（单边）", performance["平均每日单边换手率"]],
        ["交易与组合", "累计交易成本率", performance["累计交易成本率"]],
        ["交易与组合", "平均持仓ETF数量", performance["平均持仓ETF数量"]],
        ["交易与组合", "调仓次数", performance["调仓次数"]],
        ["交易与组合", "调仓成功率", success_rate],
    ]
    for row in performance_rows:
        performance_sheet.append(row)
    style_worksheet(performance_sheet, max_width=30)
    percentage_metrics = {
        "累计收益率",
        "策略年化收益率",
        f"{benchmark.name}年化收益率",
        "年化超额收益",
        "策略年化波动率",
        "策略最大回撤",
        "年化跟踪误差",
        "超额最大回撤",
        "平均单次换仓比率（单边）",
        "累计交易成本率",
        "调仓成功率",
    }
    for row_number in range(2, performance_sheet.max_row + 1):
        metric = performance_sheet.cell(row=row_number, column=2).value
        value_cell = performance_sheet.cell(row=row_number, column=3)
        if metric in percentage_metrics:
            value_cell.number_format = "0.00%"
        elif metric == "年化换手率（单边）":
            value_cell.number_format = '0.00"倍"'
        elif metric == "交易日数量":
            value_cell.number_format = '0"天"'
        elif metric == "调仓次数":
            value_cell.number_format = '0"次"'
        elif metric == "平均持仓ETF数量":
            value_cell.number_format = '0.00"只"'
        elif metric in {"初始净值", "期末净值", "Sharpe", "Sortino", "Calmar", "IR"}:
            value_cell.number_format = "0.0000"
    performance_sheet.column_dimensions["A"].width = 18
    performance_sheet.column_dimensions["B"].width = 34
    performance_sheet.column_dimensions["C"].width = 28

    parameter_sheet = workbook.create_sheet("parameters")
    parameter_sheet.append(["类别", "参数 / 约束", "本项目设置"])
    execution_label = mode_execution_label(mode)
    parameter_rows = [
        ["基本信息", "执行价格", execution_label],
        ["基准设置", "基准名称", benchmark.name],
        ["基准设置", "基准代码", benchmark.code],
        ["基准设置", "超额净值", "策略累计净值÷基准累计净值；共同起点为1"],
        ["趋势策略", "聚类相关性阈值", CLUSTER_CORRELATION_THRESHOLD],
        ["趋势策略", "趋势因子窗口", TREND_WINDOW],
        ["趋势策略", "排名得分公式", SCORE_LABELS[score_method]],
        ["趋势策略", "调仓日入选比例", TOP_PERCENT],
        ["趋势过滤", "过滤位置", "排名后"],
        ["趋势过滤", "正收益条件", f"当前趋势窗口收益率>{MIN_WINDOW_RETURN:g}"],
        ["买入过滤", "RSI周期（独立参数）", RSI_PERIOD],
        ["买入过滤", "RSI条件", f"RSI{RSI_PERIOD}>{RSI_NEUTRAL_LEVEL:g}"],
        ["买入过滤", "BIAS周期（绑定趋势窗口）", BIAS_PERIOD],
        ["买入过滤", "BIAS条件", f"BIAS{BIAS_PERIOD}>{BIAS_NEUTRAL_LEVEL:g}"],
        ["趋势过滤", "未通过处理", "先取全池Top比例再过滤；不补选；新资金在合格ETF中等额投入"],
        ["ETF选择", "代表ETF选择", "跟踪同一指数中当日成交量最大"],
        ["账户结构", "调仓模式", "单账户每日轮动，独立买入批次"],
        ["账户结构", "账户数量", ACCOUNT_COUNT],
        ["持有约束", "最短持有期（交易日）", MIN_HOLD_DAYS],
        ["持有约束", "计时方式", "成交日第0日；同ETF每次买入独立计时，续持不重置"],
        ["持有约束", "正常退出", "信号日持有天数达到H且不在过滤后名单，整批卖出"],
        ["止损设置", "固定跌幅止损比例", STOP_LOSS_PCT],
        ["止损设置", "浮盈回撤启动比例", PROFIT_TRAILING_TRIGGER_PCT],
        ["止损设置", "最高收盘价回撤比例", PROFIT_TRAILING_DRAWDOWN_PCT],
        ["止损设置", "固定止损触发", "信号日收盘价<=该批次买入成交价×(1-S)，优先于最短持有期"],
        ["止损设置", "浮盈回撤触发", "最高收盘浮盈达到启动比例后，信号日收盘价从最高收盘价回撤达到设定比例；优先于最短持有期"],
        ["止损设置", "重新买入", "同轮风控退出ETF禁买；下一轮重新评估；无额外冷静期"],
        ["交易设置", "未成交处理", "缺价或无成交额不成交；未成交卖单后续重试；不预支卖出资金"],
        ["交易设置", "缺失信号", "保留普通持仓，仅处理风控退出和未成交卖单；现金等待有效名单"],
        ["交易设置", "价格口径", "沿用ETF总表收盘价和VWAP的统一前复权口径"],
        ["交易设置", "初始净值 NAV₀", INITIAL_NAV],
        ["收益参数", "年化无风险利率 r_f", ANNUAL_RISK_FREE_RATE],
        ["交易成本", "买入成本 c_buy", TRANSACTION_COST_RATE],
        ["交易成本", "卖出成本 c_sell", TRANSACTION_COST_RATE],
        ["交易成本", "完整换仓成本 T_c", 2.0 * TRANSACTION_COST_RATE],
        ["容量参数", "当日成交额使用比例 ρ_amt", CAPACITY_DAILY_AMOUNT_RATIO],
        ["容量参数", "倒序分位 q_desc", CAPACITY_DESCENDING_QUANTILE],
        ["容量参数", "等价升序分位 q_asc", 1.0 - CAPACITY_DESCENDING_QUANTILE],
        ["组合约束", "资金分配", "现金与实际卖出净收入合并，仅新投入资金等额分配；旧批次不调权"],
        ["组合约束", "持仓数量", "不设Top数量上限，未到期旧批次可与新批次并存"],
        ["组合约束", "无法买入", "按信号名单预先等分预算，某ETF无法成交则对应预算留现金"],
        ["收益口径", "扣费后收益", "买入和卖出均扣除交易成本"],
        ["容量口径", "组合容量", "各持仓ETF容量的5%分位"],
    ]
    for row in parameter_rows:
        parameter_sheet.append(row)
    style_worksheet(parameter_sheet, max_width=50)
    percentage_parameters = {
        "调仓日入选比例",
        "固定跌幅止损比例",
        "浮盈回撤启动比例",
        "最高收盘价回撤比例",
        "年化无风险利率 r_f",
        "买入成本 c_buy",
        "卖出成本 c_sell",
        "完整换仓成本 T_c",
        "当日成交额使用比例 ρ_amt",
        "倒序分位 q_desc",
        "等价升序分位 q_asc",
    }
    for row_number in range(2, parameter_sheet.max_row + 1):
        parameter = parameter_sheet.cell(row=row_number, column=2).value
        if parameter in percentage_parameters:
            parameter_sheet.cell(row=row_number, column=3).number_format = "0.00%"
    parameter_sheet.column_dimensions["A"].width = 18
    parameter_sheet.column_dimensions["B"].width = 40
    parameter_sheet.column_dimensions["C"].width = 52

    return save_workbook_atomic(
        workbook,
        output_dir / f"{output_prefix(mode)}_backtest_metrics.xlsx",
    )


def write_holdings_workbook(
    mode: str,
    holding_records: Sequence[HoldingRecord],
    output_dir: Path,
) -> Path:
    """按参考项目单独输出每日持仓。"""

    workbook = Workbook(write_only=False)
    sheet = workbook.active
    sheet.title = "holdings"
    sheet.append(["日期", "ETF代码", "ETF名称", "权重"])
    for record in holding_records:
        sheet.append(
            [
                int(record.current_date.strftime("%Y%m%d")),
                record.etf_code,
                record.etf_name,
                record.actual_weight,
            ]
        )
    style_worksheet(sheet, decimal_headers={"权重"}, max_width=28)
    weight_column = 4
    for row_number in range(2, sheet.max_row + 1):
        sheet.cell(row=row_number, column=weight_column).number_format = "0.000000"
    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["B"].width = 16
    sheet.column_dimensions["C"].width = 38
    sheet.column_dimensions["D"].width = 16
    return save_workbook_atomic(
        workbook,
        output_dir / f"{output_prefix(mode)}_holdings.xlsx",
    )


def write_time_series_workbook(
    mode: str,
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
    baseline_date: date,
    output_dir: Path,
) -> Path:
    """按参考项目输出策略、基准和超额时序。"""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "time_series"
    headers = [
        "日期",
        "策略日收益",
        "策略累计净值",
        "基准日收益",
        "基准累计净值",
        "主动日收益",
        "超额净值",
        "每日换仓比率（单边）",
        "策略可容纳规模（亿元）",
    ]
    sheet.append(headers)
    sheet.append(
        [baseline_date, None, INITIAL_NAV, None, 1.0, None, 1.0, None, None]
    )
    for record, benchmark_record in zip(nav_records, benchmark_records):
        sheet.append(
            [
                record.current_date,
                record.daily_return,
                record.nav,
                benchmark_record.daily_return,
                benchmark_record.nav,
                benchmark_record.active_return,
                benchmark_record.relative_nav,
                record.one_way_turnover,
                record.capacity / 1e8 if record.capacity is not None else None,
            ]
        )
    style_worksheet(
        sheet,
        percent_headers={
            "策略日收益", "基准日收益", "主动日收益", "每日换仓比率（单边）",
        },
        decimal_headers={
            "策略累计净值", "基准累计净值", "超额净值", "策略可容纳规模（亿元）",
        },
        max_width=34,
    )
    for row_number in range(2, sheet.max_row + 1):
        sheet.cell(row=row_number, column=1).number_format = "yyyy-mm-dd"
        sheet.cell(row=row_number, column=3).number_format = "0.000000"
        sheet.cell(row=row_number, column=5).number_format = "0.000000"
        sheet.cell(row=row_number, column=7).number_format = "0.000000"
        sheet.cell(row=row_number, column=9).number_format = "0.0000"
    sheet.column_dimensions["A"].width = 15
    sheet.column_dimensions["B"].width = 20
    sheet.column_dimensions["C"].width = 20
    sheet.column_dimensions["D"].width = 28
    sheet.column_dimensions["E"].width = 20
    sheet.column_dimensions["F"].width = 20
    sheet.column_dimensions["G"].width = 20
    sheet.column_dimensions["H"].width = 28
    sheet.column_dimensions["I"].width = 30
    return save_workbook_atomic(
        workbook,
        output_dir / f"{output_prefix(mode)}_time_series.xlsx",
    )


def write_account_details_workbook(
    mode: str,
    daily_records: Sequence[AccountDailyRecord],
    holding_records: Sequence[AccountHoldingRecord],
    trade_records: Sequence[AccountTradeRecord],
    output_dir: Path,
) -> Path:
    """输出单账户状态及逐批次持仓/交易，保留入场价、计时与退出原因。"""

    workbook = Workbook()
    daily_sheet = workbook.active
    daily_sheet.title = "账户每日状态"
    daily_headers = [
        "日期", "账户编号", "账户净值", "ETF市值", "现金", "现金权重",
        "当天是否有交易指令", "指令是否全部成交", "当前信号日", "上次成交日", "下次检查日",
    ]
    daily_sheet.append(daily_headers)
    for record in daily_records:
        daily_sheet.append(
            [
                record.current_date,
                record.account_id,
                record.account_nav,
                record.etf_market_value,
                record.cash,
                record.cash_weight,
                "是" if record.rebalance_attempted else "否",
                "是" if record.rebalance_succeeded else "否",
                record.signal_date,
                record.last_rebalance_date,
                record.next_rebalance_date,
            ]
        )
    style_worksheet(
        daily_sheet,
        percent_headers={"现金权重"},
        decimal_headers={"账户净值", "ETF市值", "现金"},
        integer_headers={"账户编号"},
        max_width=24,
    )
    for row_number in range(2, daily_sheet.max_row + 1):
        for column_number in (1, 9, 10, 11):
            daily_sheet.cell(row=row_number, column=column_number).number_format = (
                "yyyy-mm-dd"
            )

    holding_sheet = workbook.create_sheet("批次持仓")
    holding_headers = [
        "日期", "账户编号", "信号日", "ETF代码", "ETF名称", "ETF市值",
        "账户内部权重", "对总组合贡献权重",
        "批次编号", "买入日期", "买入成交价", "持有数量", "持有交易日",
        "固定止损价", "最高收盘价", "最高浮盈", "浮盈回撤是否启用",
        "浮盈回撤退出价", "待执行卖出原因",
    ]
    holding_sheet.append(holding_headers)
    for record in holding_records:
        holding_sheet.append(
            [
                record.current_date,
                record.account_id,
                record.signal_date,
                record.etf_code,
                record.etf_name,
                record.market_value,
                record.account_weight,
                record.total_portfolio_weight,
                record.batch_id, record.buy_date, record.buy_price, record.shares,
                record.holding_days, record.stop_price, record.highest_close,
                record.max_return, "是" if record.trailing_active else "否",
                record.trailing_stop_price, record.pending_exit_reason,
            ]
        )
    style_worksheet(
        holding_sheet,
        percent_headers={"账户内部权重", "对总组合贡献权重", "最高浮盈"},
        decimal_headers={
            "ETF市值", "买入成交价", "持有数量", "固定止损价",
            "最高收盘价", "浮盈回撤退出价",
        },
        integer_headers={"账户编号"},
        max_width=32,
    )
    for row_number in range(2, holding_sheet.max_row + 1):
        holding_sheet.cell(row=row_number, column=1).number_format = "yyyy-mm-dd"
        holding_sheet.cell(row=row_number, column=3).number_format = "yyyy-mm-dd"
        holding_sheet.cell(row=row_number, column=10).number_format = "yyyy-mm-dd"

    trade_sheet = workbook.create_sheet("批次交易")
    trade_headers = [
        "成交日期", "账户编号", "信号日", "ETF代码", "ETF名称", "交易方向",
        "交易前市值", "最新目标市值", "实际成交金额", "交易成本",
        "批次编号", "买入日期", "买入成交价", "成交数量", "成交价格",
        "交易原因", "信号收盘价", "信号日持有交易日", "信号时最高收盘价", "风控退出价",
    ]
    trade_sheet.append(trade_headers)
    for record in trade_records:
        trade_sheet.append(
            [
                record.execution_date,
                record.account_id,
                record.signal_date,
                record.etf_code,
                record.etf_name,
                record.direction,
                record.before_value,
                record.target_value,
                record.trade_amount,
                record.transaction_cost,
                record.batch_id, record.buy_date, record.buy_price, record.shares,
                record.execution_price, record.reason, record.signal_price, record.signal_holding_days,
                record.signal_peak_price, record.risk_exit_price,
            ]
        )
    style_worksheet(
        trade_sheet,
        decimal_headers={
            "交易前市值", "最新目标市值", "实际成交金额", "交易成本",
            "买入成交价", "成交数量", "成交价格", "信号收盘价",
            "信号时最高收盘价", "风控退出价",
        },
        integer_headers={"账户编号"},
        max_width=32,
    )
    for row_number in range(2, trade_sheet.max_row + 1):
        trade_sheet.cell(row=row_number, column=1).number_format = "yyyy-mm-dd"
        trade_sheet.cell(row=row_number, column=3).number_format = "yyyy-mm-dd"
        trade_sheet.cell(row=row_number, column=12).number_format = "yyyy-mm-dd"

    return save_workbook_atomic(
        workbook,
        output_dir / f"{output_prefix(mode)}_batch_details.xlsx",
    )


def write_mode_workbooks(
    mode: str,
    holding_records: Sequence[HoldingRecord],
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
    baseline_date: date,
    benchmark: BenchmarkData,
    account_daily_records: Sequence[AccountDailyRecord],
    account_holding_records: Sequence[AccountHoldingRecord],
    account_trade_records: Sequence[AccountTradeRecord],
    score_method: str,
    score_backtest_dir: Path,
) -> list[Path]:
    output_dir = score_backtest_dir / mode
    output_dir.mkdir(parents=True, exist_ok=True)
    return [
        write_annual_metrics_workbook(
            mode,
            nav_records,
            benchmark_records,
            output_dir,
        ),
        write_backtest_metrics_workbook(
            mode,
            nav_records,
            benchmark_records,
            benchmark,
            output_dir,
            score_method,
        ),
        write_holdings_workbook(mode, holding_records, output_dir),
        write_time_series_workbook(
            mode,
            nav_records,
            benchmark_records,
            baseline_date,
            output_dir,
        ),
        write_account_details_workbook(
            mode,
            account_daily_records,
            account_holding_records,
            account_trade_records,
            output_dir,
        ),
    ]


def style_plot_axis(axis: Any) -> None:
    axis.grid(axis="y", color="#D9D9D9", linewidth=0.6)
    axis.spines[["top", "right"]].set_visible(False)


def write_mode_figures(
    mode: str,
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
    baseline_date: date,
    score_backtest_dir: Path,
) -> list[Path]:
    """每种交易模式独立生成图表，不和另一模式叠加。"""

    output_dir = score_backtest_dir / mode
    output_dir.mkdir(parents=True, exist_ok=True)
    plt.style.use("default")
    plt.rcParams["font.sans-serif"] = [
        "Arial Unicode MS", "PingFang SC", "Heiti SC", "SimHei", "DejaVu Sans"
    ]
    plt.rcParams["axes.unicode_minus"] = False
    color = "#1F4E79" if mode == "close" else "#C0504D"
    label = mode_execution_label(mode)
    dates = [record.current_date for record in nav_records]
    comparison_dates = [baseline_date, *dates]

    nav_path = output_dir / f"{output_prefix(mode)}_cumulative_nav.png"
    figure, axis = plt.subplots(figsize=(10.0, 4.8), facecolor="white")
    axis.plot(
        comparison_dates,
        [INITIAL_NAV, *[record.nav for record in nav_records]],
        color=color,
        linewidth=1.5,
        label="ETF趋势策略",
    )
    axis.set_title(
        f"Plot 1: Cumulative Net Value ({label})",
        loc="left",
        fontsize=13,
        fontweight="bold",
    )
    axis.set_ylabel("累计净值")
    axis.legend(frameon=False)
    axis.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=5, maxticks=9))
    axis.xaxis.set_major_formatter(
        mdates.ConciseDateFormatter(axis.xaxis.get_major_locator())
    )
    style_plot_axis(axis)
    save_figure_atomic(figure, nav_path)

    excess_path = output_dir / f"{output_prefix(mode)}_cumulative_excess.png"
    figure, axis = plt.subplots(figsize=(10.0, 4.8), facecolor="white")
    axis.plot(
        comparison_dates,
        [1.0, *[record.relative_nav for record in benchmark_records]],
        color="#2867C7",
        linewidth=1.5,
        label="累计超额净值",
    )
    axis.set_title(
        f"Plot 2: Cumulative Excess Return ({label})",
        loc="left",
        fontsize=13,
        fontweight="bold",
    )
    axis.axhline(1.0, color="#A6A6A6", linewidth=0.8, linestyle="--")
    axis.set_ylabel("超额净值")
    axis.legend(frameon=False)
    axis.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=5, maxticks=9))
    axis.xaxis.set_major_formatter(
        mdates.ConciseDateFormatter(axis.xaxis.get_major_locator())
    )
    style_plot_axis(axis)
    save_figure_atomic(figure, excess_path)

    turnover_path = output_dir / f"{output_prefix(mode)}_turnover.png"
    figure, axis = plt.subplots(figsize=(10.0, 4.5), facecolor="white")
    axis.bar(
        range(len(nav_records)),
        [record.one_way_turnover for record in nav_records],
        width=1.0,
        color=color,
        label="组合单次换仓比率（单边）",
    )
    axis.set_title(
        f"Plot 3: Rebalance Turnover ({label})",
        loc="left",
        fontsize=13,
        fontweight="bold",
    )
    axis.set_xlabel("调仓序号")
    axis.yaxis.set_major_formatter(PercentFormatter(1.0))
    axis.legend(frameon=False)
    style_plot_axis(axis)
    save_figure_atomic(figure, turnover_path)

    cost_path = output_dir / f"{output_prefix(mode)}_transaction_cost.png"
    figure, axis = plt.subplots(figsize=(10.0, 4.5), facecolor="white")
    axis.plot(
        dates,
        [record.cumulative_cost_rate for record in nav_records],
        color=color,
        linewidth=1.5,
        label="累计交易成本率",
    )
    axis.set_title(
        f"Cumulative Transaction Cost Rate ({label})",
        loc="left",
        fontsize=13,
        fontweight="bold",
    )
    axis.yaxis.set_major_formatter(PercentFormatter(1.0))
    axis.legend(frameon=False)
    axis.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=5, maxticks=9))
    axis.xaxis.set_major_formatter(
        mdates.ConciseDateFormatter(axis.xaxis.get_major_locator())
    )
    style_plot_axis(axis)
    save_figure_atomic(figure, cost_path)

    capacity_path = output_dir / f"{output_prefix(mode)}_capacity.png"
    capacity_dates = [
        record.current_date for record in nav_records if record.capacity is not None
    ]
    capacity_values = [
        record.capacity / 1e8
        for record in nav_records
        if record.capacity is not None
    ]
    figure, axis = plt.subplots(figsize=(10.0, 4.5), facecolor="white")
    axis.plot(
        capacity_dates,
        capacity_values,
        color=color,
        linewidth=1.3,
        label="ETF趋势策略",
    )
    axis.set_title(
        f"Plot 4: Strategy Capacity ({label})",
        loc="left",
        fontsize=13,
        fontweight="bold",
    )
    axis.set_ylabel("亿元")
    axis.legend(frameon=False)
    axis.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=5, maxticks=9))
    axis.xaxis.set_major_formatter(
        mdates.ConciseDateFormatter(axis.xaxis.get_major_locator())
    )
    style_plot_axis(axis)
    save_figure_atomic(figure, capacity_path)
    return [nav_path, excess_path, turnover_path, cost_path, capacity_path]


def validate_batch_accounting(
    mode: str,
    nav_records: Sequence[NavRecord],
    holding_records: Sequence[AccountHoldingRecord],
    trade_records: Sequence[AccountTradeRecord],
) -> None:
    """验证新增批次规则、最低持有期及现金账目；收益指标校验沿用原逻辑。"""
    date_positions = {row.current_date: i for i, row in enumerate(nav_records)}
    trades_by_date: dict[date, list[AccountTradeRecord]] = defaultdict(list)
    holdings_by_date: dict[date, list[AccountHoldingRecord]] = defaultdict(list)
    for row in trade_records:
        trades_by_date[row.execution_date].append(row)
    for row in holding_records:
        holdings_by_date[row.current_date].append(row)
    live: dict[int, AccountTradeRecord] = {}
    seen: set[int] = set()
    cash = INITIAL_NAV
    for nav in nav_records:
        trades = trades_by_date[nav.current_date]
        risk_exits = {
            r.etf_code for r in trades
            if r.reason in {"固定止损", "浮盈回撤"}
        }
        buys = [r for r in trades if r.direction == "买入"]
        if risk_exits.intersection(r.etf_code for r in buys):
            raise RuntimeError(f"{mode}同轮风控退出ETF被重新买入")
        if buys and not all(math.isclose(r.trade_amount, buys[0].trade_amount, abs_tol=1e-12) for r in buys):
            raise RuntimeError(f"{mode}当轮新资金未等额分配")
        for row in trades:
            if row.signal_date is None or row.signal_date > row.execution_date or (mode == "next_day_vwap" and row.signal_date == row.execution_date):
                raise RuntimeError(f"{mode}批次交易信号时序不正确")
            if not math.isclose(row.shares * row.execution_price, row.trade_amount, abs_tol=1e-12):
                raise RuntimeError(f"{mode}批次成交金额不能由数量和价格重建")
            if row.direction == "买入":
                if row.batch_id in seen or row.buy_date != row.execution_date:
                    raise RuntimeError(f"{mode}新增买入没有独立批次")
                seen.add(row.batch_id)
                live[row.batch_id] = row
                cash -= row.trade_amount + row.transaction_cost
            else:
                old = live.pop(row.batch_id, None)
                if old is None or old.shares != row.shares or old.buy_price != row.buy_price:
                    raise RuntimeError(f"{mode}卖出批次数量或买入基准被改变")
                age = date_positions[row.signal_date] - date_positions[old.buy_date]
                if row.signal_holding_days != age:
                    raise RuntimeError(f"{mode}批次持有期计数不正确")
                if row.reason == "到期未入选" and age < MIN_HOLD_DAYS:
                    raise RuntimeError(f"{mode}未满最短持有期发生普通卖出")
                if row.reason == "固定止损" and (row.signal_price is None or row.signal_price > old.buy_price * (1.0 - STOP_LOSS_PCT)):
                    raise RuntimeError(f"{mode}固定止损未达到阈值")
                if row.reason == "浮盈回撤":
                    if (
                        row.signal_price is None
                        or row.signal_peak_price is None
                        or row.risk_exit_price is None
                        or row.signal_peak_price
                        < old.buy_price * (1.0 + PROFIT_TRAILING_TRIGGER_PCT)
                        or row.signal_price
                        > row.signal_peak_price
                        * (1.0 - PROFIT_TRAILING_DRAWDOWN_PCT)
                        or not math.isclose(
                            row.risk_exit_price,
                            row.signal_peak_price
                            * (1.0 - PROFIT_TRAILING_DRAWDOWN_PCT),
                            abs_tol=1e-12,
                        )
                    ):
                        raise RuntimeError(f"{mode}浮盈回撤退出未达到阈值")
                cash += row.trade_amount - row.transaction_cost
        held = holdings_by_date[nav.current_date]
        if len(held) != len(live) or {r.batch_id for r in held} != set(live):
            raise RuntimeError(f"{mode}批次持仓与交易记录不一致")
        for row in held:
            old = live[row.batch_id]
            if row.shares != old.shares or row.buy_price != old.buy_price or row.buy_date != old.buy_date:
                raise RuntimeError(f"{mode}未卖出批次发生调权或买入基准重置")
            if row.holding_days != date_positions[nav.current_date] - date_positions[old.buy_date]:
                raise RuntimeError(f"{mode}批次持有期被重置")
        if not math.isclose(cash, nav.nav * nav.cash_weight, abs_tol=2e-10):
            raise RuntimeError(f"{mode}现金不能由实际成交与费用重建")
        if not math.isclose(cash + sum(r.market_value for r in held), nav.nav, abs_tol=2e-10):
            raise RuntimeError(f"{mode}批次市值与现金无法重建净值")


def validate_mode_outputs(
    mode: str,
    holding_records: Sequence[HoldingRecord],
    nav_records: Sequence[NavRecord],
    benchmark_records: Sequence[BenchmarkRecord],
    account_daily_records: Sequence[AccountDailyRecord],
    account_holding_records: Sequence[AccountHoldingRecord],
    account_trade_records: Sequence[AccountTradeRecord],
    workbook_paths: Sequence[Path],
    figure_paths: Sequence[Path],
) -> None:
    failures = [
        row
        for row in build_validation_rows(mode, holding_records, nav_records)
        if row["状态"] == "FAIL"
    ]
    if failures:
        raise RuntimeError(
            f"{mode}回测内部校验失败："
            + "、".join(str(row["检查项"]) for row in failures)
        )

    if len(nav_records) != len(benchmark_records):
        raise RuntimeError(f"{mode}基准校验失败：策略与基准记录数量不一致")
    for nav_record, benchmark_record in zip(nav_records, benchmark_records):
        if nav_record.current_date != benchmark_record.current_date:
            raise RuntimeError(f"{mode}基准校验失败：策略与基准日期不一致")
        expected_relative_nav = (
            nav_record.nav / INITIAL_NAV / benchmark_record.nav
        )
        if not math.isclose(
            expected_relative_nav,
            benchmark_record.relative_nav,
            rel_tol=1e-10,
            abs_tol=1e-12,
        ):
            raise RuntimeError(f"{mode}基准校验失败：超额净值不能由策略和基准重建")

    nav_by_date = {record.current_date: record for record in nav_records}
    account_rows_by_date: dict[date, list[AccountDailyRecord]] = defaultdict(list)
    for record in account_daily_records:
        account_rows_by_date[record.current_date].append(record)
    for nav_record in nav_records:
        rows = account_rows_by_date.get(nav_record.current_date, [])
        if len(rows) != 1 or rows[0].account_id != 1:
            raise RuntimeError(f"{mode}单账户记录不正确")
        if not math.isclose(rows[0].account_nav, nav_record.nav, abs_tol=2e-12):
            raise RuntimeError(f"{mode}账户净值与组合净值不一致")
    validate_batch_accounting(mode, nav_records, account_holding_records, account_trade_records)

    trade_cost_by_date: dict[date, float] = defaultdict(float)
    for record in account_trade_records:
        trade_cost_by_date[record.execution_date] += record.transaction_cost
    for current_date, nav_record in nav_by_date.items():
        if not math.isclose(
            trade_cost_by_date.get(current_date, 0.0),
            nav_record.transaction_cost,
            abs_tol=2e-12,
        ):
            raise RuntimeError(f"{mode}账户校验失败：{current_date}交易成本不一致")

    expected_workbooks = {
        f"{output_prefix(mode)}_annual_metrics.xlsx": ["annual_metrics"],
        f"{output_prefix(mode)}_backtest_metrics.xlsx": ["performance", "parameters"],
        f"{output_prefix(mode)}_holdings.xlsx": ["holdings"],
        f"{output_prefix(mode)}_time_series.xlsx": ["time_series"],
        f"{output_prefix(mode)}_batch_details.xlsx": ["账户每日状态", "批次持仓", "批次交易"],
    }
    expected_headers = {
        (f"{output_prefix(mode)}_annual_metrics.xlsx", "annual_metrics"): [
            "年份", "策略收益", "基准收益", "超额收益", "年化波动",
            "跟踪误差", "Sharpe", "Sortino", "信息比率", "策略最大回撤",
            "基准最大回撤", "超额最大回撤", "Calmar", "年化换手率（单边）",
        ],
        (f"{output_prefix(mode)}_backtest_metrics.xlsx", "performance"): [
            "类别", "指标", "ETF趋势策略",
        ],
        (f"{output_prefix(mode)}_backtest_metrics.xlsx", "parameters"): [
            "类别", "参数 / 约束", "本项目设置",
        ],
        (f"{output_prefix(mode)}_holdings.xlsx", "holdings"): [
            "日期", "ETF代码", "ETF名称", "权重",
        ],
        (f"{output_prefix(mode)}_time_series.xlsx", "time_series"): [
            "日期", "策略日收益", "策略累计净值", "基准日收益",
            "基准累计净值", "主动日收益", "超额净值",
            "每日换仓比率（单边）", "策略可容纳规模（亿元）",
        ],
        (f"{output_prefix(mode)}_batch_details.xlsx", "账户每日状态"): [
            "日期", "账户编号", "账户净值", "ETF市值", "现金", "现金权重",
            "当天是否有交易指令", "指令是否全部成交", "当前信号日", "上次成交日", "下次检查日",
        ],
        (f"{output_prefix(mode)}_batch_details.xlsx", "批次持仓"): [
            "日期", "账户编号", "信号日", "ETF代码", "ETF名称", "ETF市值",
            "账户内部权重", "对总组合贡献权重",
            "批次编号", "买入日期", "买入成交价", "持有数量", "持有交易日",
            "固定止损价", "最高收盘价", "最高浮盈", "浮盈回撤是否启用",
            "浮盈回撤退出价", "待执行卖出原因",
        ],
        (f"{output_prefix(mode)}_batch_details.xlsx", "批次交易"): [
            "成交日期", "账户编号", "信号日", "ETF代码", "ETF名称", "交易方向",
            "交易前市值", "最新目标市值", "实际成交金额", "交易成本",
            "批次编号", "买入日期", "买入成交价", "成交数量", "成交价格",
            "交易原因", "信号收盘价", "信号日持有交易日", "信号时最高收盘价", "风控退出价",
        ],
    }
    actual_workbooks = {path.name: path for path in workbook_paths}
    if set(actual_workbooks) != set(expected_workbooks):
        raise RuntimeError(f"{mode}回测Excel文件结构不正确")
    for filename, expected_sheets in expected_workbooks.items():
        path = actual_workbooks[filename]
        workbook = load_workbook(path, read_only=True, data_only=False)
        try:
            if workbook.sheetnames != expected_sheets:
                raise RuntimeError(
                    f"{filename}的sheet不正确：{workbook.sheetnames}"
                )
            for sheet in workbook.worksheets:
                allow_header_only = (
                    (filename == f"{output_prefix(mode)}_batch_details.xlsx"
                     and sheet.title in {"批次持仓", "批次交易"})
                    or filename == f"{output_prefix(mode)}_holdings.xlsx"
                )
                if (sheet.max_row < 2 and not allow_header_only) or sheet.max_column < 1:
                    raise RuntimeError(f"{filename}/{sheet.title}没有有效数据")
                actual_headers = [
                    cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))
                ]
                if actual_headers != expected_headers[(filename, sheet.title)]:
                    raise RuntimeError(
                        f"{filename}/{sheet.title}的列结构不正确"
                    )
        finally:
            workbook.close()

    for path in [*workbook_paths, *figure_paths]:
        if not path.exists() or path.stat().st_size == 0:
            raise RuntimeError(f"输出文件缺失或为空：{path}")


def main() -> None:
    validate_parameters()
    benchmark = load_benchmark_data()
    rebalance_description = (
        f"单账户每日轮动；最短持有{MIN_HOLD_DAYS}个交易日，"
        f"固定止损{STOP_LOSS_PCT:.1%}，浮盈达到"
        f"{PROFIT_TRAILING_TRIGGER_PCT:.1%}后从最高收盘价回撤"
        f"{PROFIT_TRAILING_DRAWDOWN_PCT:.1%}退出；"
        "按买入批次独立管理，仅现金和卖出资金等额投入，原持仓不调权"
    )
    print(
        f"聚类阈值 {CLUSTER_CORRELATION_THRESHOLD:g}，"
        f"趋势窗口 {TREND_WINDOW}，选择调仓信号日得分前 {TOP_PERCENT:.0%}；"
        f"排名后过滤窗口收益率不大于 {MIN_WINDOW_RETURN:g}、"
        f"RSI{RSI_PERIOD}不大于 {RSI_NEUTRAL_LEVEL:g} 或"
        f"BIAS{BIAS_PERIOD}不大于 {BIAS_NEUTRAL_LEVEL:g} 的指数，"
        f"{rebalance_description}；"
        f"本次回测{len(SCORE_METHODS_TO_RUN)}种得分公式。",
        flush=True,
    )
    print(
        f"评价基准：{benchmark.name}（{benchmark.code}），"
        f"共 {len(benchmark.closes)} 个交易日。",
        flush=True,
    )
    for score_method in SCORE_METHODS_TO_RUN:
        score_column = SCORE_COLUMNS[score_method]
        score_label = SCORE_LABELS[score_method]
        score_backtest_dir = (
            BACKTEST_DIR
            / score_method
            / STRATEGY_VARIANT_DIR
            / SELECTION_VARIANT_DIR
            / ACCOUNT_VARIANT_DIR
            / INDICATOR_VARIANT_DIR
        )
        print(f"\n开始回测：{score_label}（{score_column}）", flush=True)

        daily_selections = read_daily_top_factors(score_column)
        targets, full_trading_calendar = build_daily_targets(daily_selections)
        first_signal_date = min(targets)
        last_signal_date = max(targets)
        benchmark_end_date = max(benchmark.closes)
        effective_last_date = min(last_signal_date, benchmark_end_date)
        if effective_last_date < first_signal_date:
            raise ValueError(
                f"基准数据截至 {benchmark_end_date}，早于首个趋势信号日 "
                f"{first_signal_date}。"
            )
        if effective_last_date < last_signal_date:
            print(
                f"基准数据截至 {benchmark_end_date}，本次回测比较区间同步截止到该日。",
                flush=True,
            )
        trading_dates = [
            current_date
            for current_date in full_trading_calendar
            if first_signal_date <= current_date <= effective_last_date
        ]
        targets_in_range = {
            current_date: target
            for current_date, target in targets.items()
            if first_signal_date <= current_date <= effective_last_date
        }
        missing_calendar_dates = sorted(set(targets_in_range) - set(trading_dates))
        if missing_calendar_dates:
            raise ValueError(
                "趋势信号日期不在ETF交易日历中："
                + ",".join(value.isoformat() for value in missing_calendar_dates[:10])
            )
        selected_codes = {
            member.etf_code
            for target in targets_in_range.values()
            for member in target.members
        }
        unmapped_count = sum(
            target.unmapped_index_count for target in targets_in_range.values()
        )
        filtered_count = sum(
            target.filtered_index_count for target in targets_in_range.values()
        )
        print(
            f"共 {len(trading_dates)} 个实际ETF交易日，"
            f"实际涉及 {len(selected_codes)} 只ETF，"
            f"正收益/RSI/BIAS过滤指数-日期记录 {filtered_count} 条，"
            f"无法映射的指数-日期记录 {unmapped_count} 条。",
            flush=True,
        )
        prices = load_selected_prices(selected_codes, set(trading_dates))

        for mode in ("close", "next_day_vwap"):
            (
                holdings,
                nav_records,
                account_daily_records,
                account_holding_records,
                account_trade_records,
            ) = run_backtest(
                mode,
                trading_dates,
                targets_in_range,
                prices,
            )
            baseline_date, benchmark_records = build_benchmark_records(
                nav_records,
                benchmark,
            )
            workbook_paths = write_mode_workbooks(
                mode,
                holdings,
                nav_records,
                benchmark_records,
                baseline_date,
                benchmark,
                account_daily_records,
                account_holding_records,
                account_trade_records,
                score_method,
                score_backtest_dir,
            )
            figure_paths = write_mode_figures(
                mode,
                nav_records,
                benchmark_records,
                baseline_date,
                score_backtest_dir,
            )
            validate_mode_outputs(
                mode,
                holdings,
                nav_records,
                benchmark_records,
                account_daily_records,
                account_holding_records,
                account_trade_records,
                workbook_paths,
                figure_paths,
            )
            performance = calculate_performance(nav_records)
            relative_performance = calculate_relative_performance(
                nav_records,
                benchmark_records,
            )
            print(
                f"✅ {score_label} / {mode} 回测完成：年化收益率 "
                f"{float(performance['年化收益率']):.2%}，"
                f"Sharpe {float(performance['夏普比率']):.2f}，"
                f"最大回撤 {float(performance['最大回撤']):.2%}，"
                f"累计超额 {float(relative_performance['超额累计收益率']):.2%}，"
                f"信息比率 {float(relative_performance['信息比率']):.2f}，"
                f"平均每日单边换手率 "
                f"{float(performance['平均每日单边换手率']):.2%}；"
                f"输出目录：{score_backtest_dir / mode}\n"
                + "\n".join(
                    f"  {path.name}"
                    for path in [*workbook_paths, *figure_paths]
                ),
                flush=True,
            )


if __name__ == "__main__":
    main()
