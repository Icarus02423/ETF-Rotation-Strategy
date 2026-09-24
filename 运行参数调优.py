#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""安全地预览、筛选和串行运行参数调优任务。

不带 ``--execute`` 时，本入口只显示计划，不运行任何参数网格。
完整矩阵需要显式使用 ``--all --execute``。
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Sequence

from 实验配置 import (
    CORRELATION_THRESHOLDS,
    STAGGER_DAYS,
    TREND_WINDOWS,
    UPDATE_FREQUENCIES,
)
from 输出路径 import OUTPUT_ROOT, factor_dir, index_price_dir
from scripts.ETF趋势策略回测 import 交易策略回测, 趋势因子计算
from scripts.策略参数调优 import BIAS参数调优, 波动率参数调优


BIAS_TUNING = "bias"
VOLATILITY_TUNING = "volatility"
TUNING_TYPES = (BIAS_TUNING, VOLATILITY_TUNING)
TUNING_LABELS = {
    BIAS_TUNING: "BIAS参数调优",
    VOLATILITY_TUNING: "波动率参数调优",
}
TUNING_PROFILES = {
    BIAS_TUNING: tuple(BIAS参数调优.ALLOWED_FILTER_PROFILES),
    VOLATILITY_TUNING: tuple(波动率参数调优.ALLOWED_FILTER_PROFILES),
}
OUTPUT_NAMES = {
    BIAS_TUNING: tuple(BIAS参数调优.OUTPUT_FILES.values()),
    VOLATILITY_TUNING: tuple(波动率参数调优.OUTPUT_FILES.values()),
}
SIMULATIONS_PER_CASE = {
    BIAS_TUNING: (
        len(BIAS参数调优.BIAS_WINDOWS)
        * len(BIAS参数调优.BIAS_UPPERS)
        * len(BIAS参数调优.SCORE_METHODS)
        * len(BIAS参数调优.TRADE_MODES)
    ),
    VOLATILITY_TUNING: (
        len(波动率参数调优.VOL_WINDOWS)
        * len(波动率参数调优.VOL_KEEP_RATIOS)
        * len(波动率参数调优.SCORE_METHODS)
        * len(波动率参数调优.TRADE_MODES)
    ),
}

ALL_FILTER_PROFILES = tuple(
    dict.fromkeys(
        profile
        for tuning_type in TUNING_TYPES
        for profile in TUNING_PROFILES[tuning_type]
    )
)
ETF_DATA_FILE = OUTPUT_ROOT / "etf_data" / "etf_data.csv"
REQUIRED_FACTOR_YEARS = tuple(
    range(趋势因子计算.START_YEAR, 趋势因子计算.END_YEAR + 1)
)
INCOMPLETE_MARKER_NAME = ".incomplete"


@dataclass(frozen=True)
class TuningCase:
    tuning_type: str
    update_frequency: str
    filter_profile: str
    correlation_threshold: float
    trend_window: int
    stagger_days: int


@dataclass(frozen=True)
class BatchTuningSummary:
    total_cases: int
    executed_cases: int
    skipped_cases: int


def iter_tuning_cases() -> Iterable[TuningCase]:
    for tuning_type in TUNING_TYPES:
        for update_frequency in UPDATE_FREQUENCIES:
            for filter_profile in TUNING_PROFILES[tuning_type]:
                for correlation_threshold in CORRELATION_THRESHOLDS:
                    for trend_window in TREND_WINDOWS:
                        for stagger_days in STAGGER_DAYS:
                            yield TuningCase(
                                tuning_type=tuning_type,
                                update_frequency=update_frequency,
                                filter_profile=filter_profile,
                                correlation_threshold=correlation_threshold,
                                trend_window=trend_window,
                                stagger_days=stagger_days,
                            )


def describe_case(tuning_case: TuningCase) -> str:
    return (
        f"{TUNING_LABELS.get(tuning_case.tuning_type, tuning_case.tuning_type)} / "
        f"{tuning_case.update_frequency} / "
        f"{tuning_case.filter_profile} / "
        f"阈值{tuning_case.correlation_threshold:g} / "
        f"窗口{tuning_case.trend_window}日 / "
        f"错峰{tuning_case.stagger_days}日"
    )


def validate_tuning_case(tuning_case: TuningCase) -> None:
    if tuning_case.tuning_type not in TUNING_TYPES:
        raise ValueError(f"未知调优类型：{tuning_case.tuning_type}")
    if tuning_case.update_frequency not in UPDATE_FREQUENCIES:
        raise ValueError(f"未知更新频率：{tuning_case.update_frequency}")
    if tuning_case.filter_profile not in TUNING_PROFILES[tuning_case.tuning_type]:
        raise ValueError(
            f"{TUNING_LABELS[tuning_case.tuning_type]}不支持过滤状态："
            f"{tuning_case.filter_profile}"
        )
    if tuning_case.correlation_threshold not in CORRELATION_THRESHOLDS:
        raise ValueError(
            f"未知聚类阈值：{tuning_case.correlation_threshold}"
        )
    if tuning_case.trend_window not in TREND_WINDOWS:
        raise ValueError(f"未知趋势窗口：{tuning_case.trend_window}")
    if tuning_case.stagger_days not in STAGGER_DAYS:
        raise ValueError(f"未知错峰天数：{tuning_case.stagger_days}")
    names = OUTPUT_NAMES[tuning_case.tuning_type]
    if len(names) != 4 or len(set(names)) != 4:
        raise ValueError(
            f"{TUNING_LABELS[tuning_case.tuning_type]}必须定义4张不同的热力图"
        )


def validate_case_sequence(tuning_cases: Sequence[TuningCase]) -> None:
    if not tuning_cases:
        raise ValueError("参数调优至少需要一个任务")
    seen: set[TuningCase] = set()
    for tuning_case in tuning_cases:
        validate_tuning_case(tuning_case)
        if tuning_case in seen:
            raise ValueError(f"参数调优出现重复任务：{describe_case(tuning_case)}")
        seen.add(tuning_case)


def select_tuning_cases(
    tuning_cases: Iterable[TuningCase],
    *,
    tuning_types: Optional[Sequence[str]] = None,
    update_frequencies: Optional[Sequence[str]] = None,
    filter_profiles: Optional[Sequence[str]] = None,
    correlation_thresholds: Optional[Sequence[float]] = None,
    trend_windows: Optional[Sequence[int]] = None,
    stagger_days: Optional[Sequence[int]] = None,
) -> tuple[TuningCase, ...]:
    type_filter = None if tuning_types is None else set(tuning_types)
    frequency_filter = (
        None if update_frequencies is None else set(update_frequencies)
    )
    profile_filter = None if filter_profiles is None else set(filter_profiles)
    threshold_filter = (
        None
        if correlation_thresholds is None
        else set(correlation_thresholds)
    )
    window_filter = None if trend_windows is None else set(trend_windows)
    stagger_filter = None if stagger_days is None else set(stagger_days)

    return tuple(
        tuning_case
        for tuning_case in tuning_cases
        if (type_filter is None or tuning_case.tuning_type in type_filter)
        and (
            frequency_filter is None
            or tuning_case.update_frequency in frequency_filter
        )
        and (
            profile_filter is None
            or tuning_case.filter_profile in profile_filter
        )
        and (
            threshold_filter is None
            or tuning_case.correlation_threshold in threshold_filter
        )
        and (
            window_filter is None
            or tuning_case.trend_window in window_filter
        )
        and (
            stagger_filter is None
            or tuning_case.stagger_days in stagger_filter
        )
    )


def tuning_output_directory(tuning_case: TuningCase) -> Path:
    arguments = (
        tuning_case.update_frequency,
        tuning_case.filter_profile,
        tuning_case.correlation_threshold,
        tuning_case.trend_window,
        tuning_case.stagger_days,
    )
    if tuning_case.tuning_type == BIAS_TUNING:
        return BIAS参数调优.build_tuning_settings(*arguments).output_directory
    return 波动率参数调优.build_tuning_settings(*arguments).output_directory


def expected_output_paths(tuning_case: TuningCase) -> tuple[Path, ...]:
    output_directory = tuning_output_directory(tuning_case)
    return tuple(
        output_directory / name
        for name in OUTPUT_NAMES[tuning_case.tuning_type]
    )


def incomplete_output_paths(tuning_case: TuningCase) -> tuple[Path, ...]:
    return tuple(
        path
        for path in expected_output_paths(tuning_case)
        if not path.is_file() or path.stat().st_size == 0
    )


def incomplete_marker_path(tuning_case: TuningCase) -> Path:
    return tuning_output_directory(tuning_case) / INCOMPLETE_MARKER_NAME


def case_is_complete(tuning_case: TuningCase) -> bool:
    return (
        not incomplete_marker_path(tuning_case).exists()
        and not incomplete_output_paths(tuning_case)
    )


def mark_case_incomplete(tuning_case: TuningCase) -> None:
    marker_path = incomplete_marker_path(tuning_case)
    marker_path.parent.mkdir(parents=True, exist_ok=True)
    marker_path.touch()


def clear_case_incomplete(tuning_case: TuningCase) -> None:
    incomplete_marker_path(tuning_case).unlink(missing_ok=True)


def directory_has_nonempty_csv(directory: Path) -> bool:
    if not directory.is_dir():
        return False
    return any(
        path.is_file() and path.stat().st_size > 0
        for path in directory.glob("*.csv")
    )


def find_missing_inputs(
    tuning_cases: Sequence[TuningCase],
) -> tuple[str, ...]:
    missing: list[str] = []
    if not ETF_DATA_FILE.is_file() or ETF_DATA_FILE.stat().st_size == 0:
        missing.append(f"ETF总表缺失或为空：{ETF_DATA_FILE}")
    try:
        交易策略回测.load_benchmark_data()
    except Exception as error:
        missing.append(f"基准数据不可用：{error}")

    factor_keys = sorted(
        {
            (
                tuning_case.update_frequency,
                tuning_case.correlation_threshold,
                tuning_case.trend_window,
            )
            for tuning_case in tuning_cases
        }
    )
    for update_frequency, correlation_threshold, trend_window in factor_keys:
        directory = factor_dir(
            update_frequency,
            correlation_threshold,
            trend_window,
        )
        missing_year_files = [
            path.name
            for year in REQUIRED_FACTOR_YEARS
            for path in (directory / f"{year}.csv",)
            if not path.is_file() or path.stat().st_size == 0
        ]
        if missing_year_files:
            missing.append(
                f"趋势因子目录缺少非空年度CSV：{directory}（"
                + "、".join(missing_year_files)
                + "）"
            )

    index_price_keys = sorted(
        {
            (
                tuning_case.update_frequency,
                tuning_case.correlation_threshold,
            )
            for tuning_case in tuning_cases
        }
    )
    for update_frequency, correlation_threshold in index_price_keys:
        directory = index_price_dir(update_frequency, correlation_threshold)
        if not directory_has_nonempty_csv(directory):
            missing.append(f"指数过滤行情目录没有非空CSV：{directory}")

    return tuple(missing)


def validate_required_inputs(tuning_cases: Sequence[TuningCase]) -> None:
    missing = find_missing_inputs(tuning_cases)
    if missing:
        raise FileNotFoundError(
            "参数调优输入预检失败：\n- " + "\n- ".join(missing)
        )


def run_tuning_case(tuning_case: TuningCase) -> None:
    arguments = {
        "update_frequency": tuning_case.update_frequency,
        "filter_profile": tuning_case.filter_profile,
        "correlation_threshold": tuning_case.correlation_threshold,
        "trend_window": tuning_case.trend_window,
        "stagger_days": tuning_case.stagger_days,
    }
    if tuning_case.tuning_type == BIAS_TUNING:
        BIAS参数调优.main(**arguments)
        return
    波动率参数调优.main(**arguments)


def run_cases(
    tuning_cases: Iterable[TuningCase],
    *,
    force: bool = False,
) -> BatchTuningSummary:
    cases = tuple(tuning_cases)
    validate_case_sequence(cases)

    if force:
        pending_cases = cases
    else:
        pending_cases = tuple(
            tuning_case
            for tuning_case in cases
            if not case_is_complete(tuning_case)
        )
    pending_set = set(pending_cases)
    skipped_count = len(cases) - len(pending_cases)

    if pending_cases:
        validate_required_inputs(pending_cases)

    print(
        f"参数调优共选择 {len(cases)} 组："
        f"待运行 {len(pending_cases)} 组，跳过 {skipped_count} 组；"
        f"force={force}。",
        flush=True,
    )

    executed_count = 0
    for position, tuning_case in enumerate(cases, start=1):
        description = describe_case(tuning_case)
        if tuning_case not in pending_set:
            print(
                f"[{position}/{len(cases)}] ⏭️ 已完整，跳过：{description}",
                flush=True,
            )
            continue

        print(
            f"[{position}/{len(cases)}] 开始：{description}",
            flush=True,
        )
        try:
            mark_case_incomplete(tuning_case)
            run_tuning_case(tuning_case)
            missing_outputs = incomplete_output_paths(tuning_case)
            if missing_outputs:
                raise RuntimeError(
                    "调优返回后仍有热力图缺失或为空："
                    + "、".join(path.name for path in missing_outputs)
                )
            clear_case_incomplete(tuning_case)
        except Exception as error:
            print(
                f"[{position}/{len(cases)}] ❌ 失败：{description}",
                flush=True,
            )
            raise RuntimeError(
                f"参数调优在第 {position}/{len(cases)} 组失败："
                f"{description}"
            ) from error
        executed_count += 1
        print(
            f"[{position}/{len(cases)}] ✅ 完成：{description}",
            flush=True,
        )

    summary = BatchTuningSummary(
        total_cases=len(cases),
        executed_cases=executed_count,
        skipped_cases=skipped_count,
    )
    print(
        f"参数调优处理完成：总计 {summary.total_cases} 组，"
        f"本次运行 {summary.executed_cases} 组，"
        f"跳过 {summary.skipped_cases} 组。",
        flush=True,
    )
    return summary


def simulation_count(tuning_cases: Sequence[TuningCase]) -> int:
    return sum(
        SIMULATIONS_PER_CASE[tuning_case.tuning_type]
        for tuning_case in tuning_cases
    )


def print_plan(tuning_cases: Sequence[TuningCase]) -> None:
    bias_count = sum(
        tuning_case.tuning_type == BIAS_TUNING
        for tuning_case in tuning_cases
    )
    volatility_count = len(tuning_cases) - bias_count
    print(
        f"本次选择 {len(tuning_cases)} 个外层调优任务："
        f"BIAS {bias_count} 个，波动率 {volatility_count} 个。",
        flush=True,
    )
    print(
        f"预计执行 {simulation_count(tuning_cases):,} 次逐日收益模拟，"
        f"完整时输出 {len(tuning_cases) * 4:,} 张热力图。",
        flush=True,
    )


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="预览或运行BIAS与波动率参数调优任务",
        allow_abbrev=False,
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="选择全部1200个外层任务；仍需--execute才会运行",
    )
    parser.add_argument(
        "--tuning-type",
        nargs="+",
        choices=TUNING_TYPES,
        help="选择bias、volatility或两者",
    )
    parser.add_argument(
        "--frequency",
        nargs="+",
        choices=UPDATE_FREQUENCIES,
        help="选择monthly、daily或两者",
    )
    parser.add_argument(
        "--profile",
        nargs="+",
        choices=ALL_FILTER_PROFILES,
        help="选择过滤状态",
    )
    parser.add_argument(
        "--threshold",
        nargs="+",
        type=float,
        choices=CORRELATION_THRESHOLDS,
        help="选择聚类相关性阈值",
    )
    parser.add_argument(
        "--trend-window",
        nargs="+",
        type=int,
        choices=TREND_WINDOWS,
        help="选择趋势窗口",
    )
    parser.add_argument(
        "--stagger-days",
        nargs="+",
        type=int,
        choices=STAGGER_DAYS,
        help="选择错峰天数",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="确认执行选中的任务；缺少此参数时仅预览",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="强制重跑选中的完整任务；只能与--execute一起使用",
    )
    return parser


def selector_was_provided(arguments: argparse.Namespace) -> bool:
    return any(
        value is not None
        for value in (
            arguments.tuning_type,
            arguments.frequency,
            arguments.profile,
            arguments.threshold,
            arguments.trend_window,
            arguments.stagger_days,
        )
    )


def resolve_cli_cases(
    arguments: argparse.Namespace,
    parser: argparse.ArgumentParser,
) -> tuple[TuningCase, ...]:
    has_selector = selector_was_provided(arguments)
    if arguments.all and has_selector:
        parser.error("--all不能与具体筛选参数同时使用")
    if not arguments.all and not has_selector:
        if arguments.execute or arguments.force:
            parser.error("执行调优前必须使用--all或至少一个筛选参数")
        return ()
    if arguments.force and not arguments.execute:
        parser.error("--force只能与--execute一起使用")

    all_cases = tuple(iter_tuning_cases())
    if arguments.all:
        return all_cases

    selected_types = tuple(arguments.tuning_type or TUNING_TYPES)
    incompatible_profiles = [
        profile
        for profile in arguments.profile or ()
        if not any(
            profile in TUNING_PROFILES[tuning_type]
            for tuning_type in selected_types
        )
    ]
    if incompatible_profiles:
        parser.error(
            "过滤状态与调优类型不兼容："
            + "、".join(incompatible_profiles)
        )

    selected = select_tuning_cases(
        all_cases,
        tuning_types=arguments.tuning_type,
        update_frequencies=arguments.frequency,
        filter_profiles=arguments.profile,
        correlation_thresholds=arguments.threshold,
        trend_windows=arguments.trend_window,
        stagger_days=arguments.stagger_days,
    )
    if not selected:
        parser.error("筛选条件没有匹配到任何调优任务")
    if len(selected) == len(all_cases):
        parser.error(
            "筛选条件覆盖了全部1200个任务；全量运行必须明确使用--all"
        )
    return selected


def cli(
    command_line_arguments: Optional[Sequence[str]] = None,
) -> Optional[BatchTuningSummary]:
    parser = build_argument_parser()
    arguments = parser.parse_args(command_line_arguments)
    selected_cases = resolve_cli_cases(arguments, parser)
    if not selected_cases:
        parser.print_help()
        print(
            "\n安全规则：先使用--all或具体筛选参数预览；"
            "确认后再增加--execute。",
            flush=True,
        )
        return None

    print_plan(selected_cases)
    if not arguments.execute:
        print(
            "当前仅预览，没有运行参数调优。确认范围后增加--execute。",
            flush=True,
        )
        return None
    return run_cases(selected_cases, force=arguments.force)


if __name__ == "__main__":
    cli()
