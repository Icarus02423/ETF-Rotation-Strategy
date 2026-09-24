#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""串行运行全部正式回测实验。

本入口只编排 ``ExperimentCase``，不下载或重建上游数据，也不运行参数调优。
默认跳过两种成交方式均已完整输出的实验；算法改动后可用 ``force=True``
强制重跑。任何实验失败都会立即停止，便于修复后从完整结果处继续。
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

from 实验配置 import FILTER_PROFILE_SETTINGS, ExperimentCase, iter_experiment_cases
from 输出路径 import backtest_result_dir, factor_dir, index_price_dir
from scripts.ETF趋势策略回测 import 交易策略回测, 趋势因子计算


EXECUTION_MODES = ("close", "next_day_vwap")
MODE_OUTPUT_SUFFIXES = (
    "annual_metrics.xlsx",
    "backtest_metrics.xlsx",
    "holdings.xlsx",
    "time_series.xlsx",
    "account_details.xlsx",
    "cumulative_nav.png",
    "cumulative_excess.png",
    "turnover.png",
    "transaction_cost.png",
    "capacity.png",
)
INCOMPLETE_MARKER_NAME = ".incomplete"
REQUIRED_FACTOR_YEARS = tuple(
    range(趋势因子计算.START_YEAR, 趋势因子计算.END_YEAR + 1)
)


@dataclass(frozen=True)
class BatchRunSummary:
    total_cases: int
    executed_cases: int
    skipped_cases: int


def describe_case(experiment_case: ExperimentCase) -> str:
    return (
        f"{experiment_case.update_frequency} / "
        f"{experiment_case.filter_profile} / "
        f"阈值{experiment_case.correlation_threshold:g} / "
        f"{experiment_case.trend_factor} / "
        f"窗口{experiment_case.trend_window}日 / "
        f"错峰{experiment_case.stagger_days}日"
    )


def expected_mode_output_paths(
    experiment_case: ExperimentCase,
    execution_mode: str,
) -> tuple[Path, ...]:
    if execution_mode not in EXECUTION_MODES:
        raise ValueError(f"未知成交方式：{execution_mode}")
    output_directory = backtest_result_dir(
        experiment_case,
        execution_mode,
    )
    return tuple(
        output_directory / f"{execution_mode}_{suffix}"
        for suffix in MODE_OUTPUT_SUFFIXES
    )


def incomplete_output_paths(
    experiment_case: ExperimentCase,
) -> tuple[Path, ...]:
    return tuple(
        path
        for execution_mode in EXECUTION_MODES
        for path in expected_mode_output_paths(
            experiment_case,
            execution_mode,
        )
        if not path.is_file() or path.stat().st_size == 0
    )


def case_incomplete_marker_path(
    experiment_case: ExperimentCase,
) -> Path:
    return (
        backtest_result_dir(experiment_case, EXECUTION_MODES[0]).parent
        / INCOMPLETE_MARKER_NAME
    )


def case_outputs_are_complete(experiment_case: ExperimentCase) -> bool:
    return not incomplete_output_paths(experiment_case)


def case_is_complete(experiment_case: ExperimentCase) -> bool:
    """20个文件均非空，且上次运行未中断，才算完成。"""

    return (
        not case_incomplete_marker_path(experiment_case).exists()
        and case_outputs_are_complete(experiment_case)
    )


def mark_case_incomplete(experiment_case: ExperimentCase) -> None:
    marker_path = case_incomplete_marker_path(experiment_case)
    marker_path.parent.mkdir(parents=True, exist_ok=True)
    marker_path.touch()


def clear_case_incomplete(experiment_case: ExperimentCase) -> None:
    case_incomplete_marker_path(experiment_case).unlink(missing_ok=True)


def validate_case_sequence(
    experiment_cases: Sequence[ExperimentCase],
) -> None:
    if not experiment_cases:
        raise ValueError("批量回测至少需要一个ExperimentCase")

    seen: set[ExperimentCase] = set()
    for experiment_case in experiment_cases:
        交易策略回测.validate_experiment_case(experiment_case)
        if experiment_case in seen:
            raise ValueError(
                "批量回测出现重复实验组合："
                f"{describe_case(experiment_case)}"
            )
        seen.add(experiment_case)


def directory_has_nonempty_csv(directory: Path) -> bool:
    if not directory.is_dir():
        return False
    return any(
        path.is_file() and path.stat().st_size > 0
        for path in directory.glob("*.csv")
    )


def find_missing_inputs(
    experiment_cases: Sequence[ExperimentCase],
) -> tuple[str, ...]:
    """汇总待运行case缺失的共享输入，不读取大文件内容。"""

    missing: list[str] = []
    if (
        not 交易策略回测.ETF_DATA_FILE.is_file()
        or 交易策略回测.ETF_DATA_FILE.stat().st_size == 0
    ):
        missing.append(f"ETF总表缺失或为空：{交易策略回测.ETF_DATA_FILE}")
    if not directory_has_nonempty_csv(交易策略回测.BENCHMARK_DIR):
        missing.append(
            f"基准目录没有非空CSV：{交易策略回测.BENCHMARK_DIR}"
        )

    factor_keys = sorted(
        {
            (
                case.update_frequency,
                case.correlation_threshold,
                case.trend_window,
            )
            for case in experiment_cases
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
            (case.update_frequency, case.correlation_threshold)
            for case in experiment_cases
            if (
                FILTER_PROFILE_SETTINGS[case.filter_profile]["bias_mode"]
                != "none"
                or FILTER_PROFILE_SETTINGS[case.filter_profile][
                    "vol_filter_enabled"
                ]
            )
        }
    )
    for update_frequency, correlation_threshold in index_price_keys:
        directory = index_price_dir(
            update_frequency,
            correlation_threshold,
        )
        if not directory_has_nonempty_csv(directory):
            missing.append(f"指数过滤行情目录没有非空CSV：{directory}")

    return tuple(missing)


def validate_required_inputs(
    experiment_cases: Sequence[ExperimentCase],
) -> None:
    missing = find_missing_inputs(experiment_cases)
    if missing:
        raise FileNotFoundError(
            "批量回测输入预检失败：\n- " + "\n- ".join(missing)
        )


def run_cases(
    experiment_cases: Iterable[ExperimentCase],
    *,
    force: bool = False,
) -> BatchRunSummary:
    cases = tuple(experiment_cases)
    validate_case_sequence(cases)

    if force:
        pending_cases = cases
    else:
        pending_cases = tuple(
            case for case in cases if not case_is_complete(case)
        )
    pending_set = set(pending_cases)
    skipped_count = len(cases) - len(pending_cases)

    if pending_cases:
        validate_required_inputs(pending_cases)

    print(
        f"正式回测计划共 {len(cases)} 组："
        f"待运行 {len(pending_cases)} 组，跳过 {skipped_count} 组；"
        f"force={force}。",
        flush=True,
    )

    executed_count = 0
    for position, experiment_case in enumerate(cases, start=1):
        description = describe_case(experiment_case)
        if experiment_case not in pending_set:
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
            mark_case_incomplete(experiment_case)
            交易策略回测.main(
                update_frequency=experiment_case.update_frequency,
                filter_profile=experiment_case.filter_profile,
                correlation_threshold=experiment_case.correlation_threshold,
                trend_factor=experiment_case.trend_factor,
                trend_window=experiment_case.trend_window,
                stagger_days=experiment_case.stagger_days,
            )
            missing_outputs = incomplete_output_paths(experiment_case)
            if missing_outputs:
                raise RuntimeError(
                    "正式回测返回后仍有输出文件缺失或为空："
                    + "、".join(path.name for path in missing_outputs)
                )
            clear_case_incomplete(experiment_case)
        except Exception as error:
            print(
                f"[{position}/{len(cases)}] ❌ 失败：{description}",
                flush=True,
            )
            raise RuntimeError(
                f"批量回测在第 {position}/{len(cases)} 组失败："
                f"{description}"
            ) from error
        executed_count += 1
        print(
            f"[{position}/{len(cases)}] ✅ 完成：{description}",
            flush=True,
        )

    summary = BatchRunSummary(
        total_cases=len(cases),
        executed_cases=executed_count,
        skipped_cases=skipped_count,
    )
    print(
        f"全部正式回测处理完成：总计 {summary.total_cases} 组，"
        f"本次运行 {summary.executed_cases} 组，"
        f"跳过 {summary.skipped_cases} 组。",
        flush=True,
    )
    return summary


def main(force: bool = False) -> BatchRunSummary:
    return run_cases(iter_experiment_cases(), force=force)


def parse_cli_arguments(
    arguments: Sequence[str] | None = None,
) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="串行运行全部正式回测实验")
    parser.add_argument(
        "--force",
        action="store_true",
        help="忽略已有完整结果，强制重跑全部实验",
    )
    return parser.parse_args(arguments)


if __name__ == "__main__":
    cli_arguments = parse_cli_arguments()
    main(force=cli_arguments.force)
