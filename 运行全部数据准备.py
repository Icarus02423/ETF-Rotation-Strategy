#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按依赖顺序运行全部正式上游数据准备任务。

本入口从现有 ``outputs/etf_data/etf_data.csv`` 开始，只编排ETF初筛、
指数收益率、指数聚类、趋势行情和趋势因子。它不下载原始ETF或基准，
也不运行参数调优和正式回测。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator, Optional, Sequence

from 实验配置 import (
    CORRELATION_THRESHOLDS,
    TREND_WINDOWS,
    UPDATE_FREQUENCIES,
)
from scripts.ETF池筛选 import ETF初筛, 指数收益率准备, 指数聚类筛选
from scripts.ETF趋势策略回测 import 趋势因子计算, 趋势行情准备


ETF_SELECTION_STAGE = "etf_selection"
INDEX_RETURN_STAGE = "index_return"
INDEX_CLUSTER_STAGE = "index_cluster"
TREND_PRICE_STAGE = "trend_price"
TREND_FACTOR_STAGE = "trend_factor"

STAGE_ORDER = (
    ETF_SELECTION_STAGE,
    INDEX_RETURN_STAGE,
    INDEX_CLUSTER_STAGE,
    TREND_PRICE_STAGE,
    TREND_FACTOR_STAGE,
)
STAGE_LABELS = {
    ETF_SELECTION_STAGE: "ETF初筛",
    INDEX_RETURN_STAGE: "指数收益率准备",
    INDEX_CLUSTER_STAGE: "指数聚类",
    TREND_PRICE_STAGE: "趋势行情准备",
    TREND_FACTOR_STAGE: "趋势因子计算",
}

ETF_DATA_FILE = Path(ETF初筛.INPUT_FILE)


@dataclass(frozen=True)
class PreparationTask:
    stage: str
    update_frequency: str
    correlation_threshold: Optional[float] = None
    trend_window: Optional[int] = None


@dataclass(frozen=True)
class PreparationRunSummary:
    total_tasks: int
    completed_tasks: int


def iter_preparation_tasks() -> Iterator[PreparationTask]:
    for update_frequency in UPDATE_FREQUENCIES:
        yield PreparationTask(
            stage=ETF_SELECTION_STAGE,
            update_frequency=update_frequency,
        )

    for update_frequency in UPDATE_FREQUENCIES:
        yield PreparationTask(
            stage=INDEX_RETURN_STAGE,
            update_frequency=update_frequency,
        )

    for update_frequency in UPDATE_FREQUENCIES:
        for correlation_threshold in CORRELATION_THRESHOLDS:
            yield PreparationTask(
                stage=INDEX_CLUSTER_STAGE,
                update_frequency=update_frequency,
                correlation_threshold=correlation_threshold,
            )

    for update_frequency in UPDATE_FREQUENCIES:
        for correlation_threshold in CORRELATION_THRESHOLDS:
            yield PreparationTask(
                stage=TREND_PRICE_STAGE,
                update_frequency=update_frequency,
                correlation_threshold=correlation_threshold,
            )

    for update_frequency in UPDATE_FREQUENCIES:
        for correlation_threshold in CORRELATION_THRESHOLDS:
            for trend_window in TREND_WINDOWS:
                yield PreparationTask(
                    stage=TREND_FACTOR_STAGE,
                    update_frequency=update_frequency,
                    correlation_threshold=correlation_threshold,
                    trend_window=trend_window,
                )


def describe_task(task: PreparationTask) -> str:
    parts = [
        STAGE_LABELS.get(task.stage, task.stage),
        task.update_frequency,
    ]
    if task.correlation_threshold is not None:
        parts.append(f"阈值{task.correlation_threshold:g}")
    if task.trend_window is not None:
        parts.append(f"窗口{task.trend_window}日")
    return " / ".join(parts)


def validate_task(task: PreparationTask) -> None:
    if task.stage not in STAGE_ORDER:
        raise ValueError(f"未知数据准备阶段：{task.stage}")
    if task.update_frequency not in UPDATE_FREQUENCIES:
        raise ValueError(f"未知更新频率：{task.update_frequency}")

    if task.stage in {ETF_SELECTION_STAGE, INDEX_RETURN_STAGE}:
        if (
            task.correlation_threshold is not None
            or task.trend_window is not None
        ):
            raise ValueError(f"{STAGE_LABELS[task.stage]}不接受阈值或趋势窗口")
        return

    if task.correlation_threshold not in CORRELATION_THRESHOLDS:
        raise ValueError(
            f"{STAGE_LABELS[task.stage]}的聚类阈值不合法："
            f"{task.correlation_threshold}"
        )

    if task.stage in {INDEX_CLUSTER_STAGE, TREND_PRICE_STAGE}:
        if task.trend_window is not None:
            raise ValueError(f"{STAGE_LABELS[task.stage]}不接受趋势窗口")
        return

    if task.trend_window not in TREND_WINDOWS:
        raise ValueError(f"趋势因子窗口不合法：{task.trend_window}")


def validate_task_sequence(tasks: Sequence[PreparationTask]) -> None:
    if not tasks:
        raise ValueError("数据准备至少需要一个任务")

    seen: set[PreparationTask] = set()
    previous_stage_position = -1
    for task in tasks:
        validate_task(task)
        if task in seen:
            raise ValueError(f"数据准备出现重复任务：{describe_task(task)}")
        seen.add(task)

        stage_position = STAGE_ORDER.index(task.stage)
        if stage_position < previous_stage_position:
            raise ValueError("数据准备任务没有按照依赖阶段排列")
        previous_stage_position = stage_position


def validate_source_data() -> None:
    if not ETF_DATA_FILE.is_file() or ETF_DATA_FILE.stat().st_size == 0:
        raise FileNotFoundError(f"ETF总表缺失或为空：{ETF_DATA_FILE}")


def run_task(task: PreparationTask) -> None:
    if task.stage == ETF_SELECTION_STAGE:
        ETF初筛.main(update_frequency=task.update_frequency)
        return

    if task.stage == INDEX_RETURN_STAGE:
        指数收益率准备.main(update_frequency=task.update_frequency)
        return

    correlation_threshold = task.correlation_threshold
    if correlation_threshold is None:
        raise ValueError(f"{STAGE_LABELS[task.stage]}缺少聚类阈值")

    if task.stage == INDEX_CLUSTER_STAGE:
        指数聚类筛选.main(
            update_frequency=task.update_frequency,
            correlation_threshold=correlation_threshold,
        )
        return

    if task.stage == TREND_PRICE_STAGE:
        趋势行情准备.main(
            update_frequency=task.update_frequency,
            correlation_threshold=correlation_threshold,
        )
        return

    trend_window = task.trend_window
    if trend_window is None:
        raise ValueError("趋势因子计算缺少趋势窗口")
    趋势因子计算.main(
        update_frequency=task.update_frequency,
        correlation_threshold=correlation_threshold,
        trend_window=trend_window,
    )


def run_tasks(
    preparation_tasks: Iterable[PreparationTask],
) -> PreparationRunSummary:
    tasks = tuple(preparation_tasks)
    validate_task_sequence(tasks)
    validate_source_data()

    print(f"全部数据准备计划共 {len(tasks)} 个任务。", flush=True)
    completed_count = 0
    for position, task in enumerate(tasks, start=1):
        description = describe_task(task)
        print(
            f"[{position}/{len(tasks)}] 开始：{description}",
            flush=True,
        )
        try:
            run_task(task)
        except Exception as error:
            print(
                f"[{position}/{len(tasks)}] ❌ 失败：{description}",
                flush=True,
            )
            raise RuntimeError(
                f"全部数据准备在第 {position}/{len(tasks)} 个任务失败："
                f"{description}"
            ) from error
        completed_count += 1
        print(
            f"[{position}/{len(tasks)}] ✅ 完成：{description}",
            flush=True,
        )

    summary = PreparationRunSummary(
        total_tasks=len(tasks),
        completed_tasks=completed_count,
    )
    print(
        f"全部数据准备完成：共 {summary.total_tasks} 个任务。",
        flush=True,
    )
    return summary


def main() -> PreparationRunSummary:
    return run_tasks(iter_preparation_tasks())


if __name__ == "__main__":
    main()
