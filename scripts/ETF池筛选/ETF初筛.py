#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按日或按月筛选ETF池，并按对标指数或 benchmark 去重。

筛选规则：
1. 截至当日，ETF上市满1年；
2. 当日基金规模 > 1亿元；
3. 包含当日在内的过去20个市场交易日平均成交额 > 2000万元；
4. 对标指数代码不能为空，缺失的ETF在去重前排除；
5. 对标指数相同或 benchmark 相同的ETF归为一类，每类只保留规模最大的ETF。

说明：
- 只筛选 SELECTED_MAJOR_CATEGORIES 参数指定的ETF大类；
- 停牌日或成交额空值按0计入20日平均成交额，分母固定为20；
- 每次运行只清理当前更新频率输出目录中的CSV，再生成本次结果；
- daily模式为每个满足20日成交额回看要求的交易日生成一个CSV；
- monthly模式只为每个自然月最后一个实际交易日生成一个CSV；
- CSV文件名使用当日交易日期，例如2021年1月29日使用2021_01_29.csv；
- CSV中的“日期”仍保留当日交易日期；
- 输出列与 etf_data.csv 完全一致，包含“对标指数代码”；
- 对标指数代码用于排除缺失ETF，但不作为归类依据。
"""

from __future__ import annotations

import calendar
import csv
import sys
from bisect import bisect_right
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import date, datetime
from itertools import groupby
from pathlib import Path
from typing import Iterable, Iterator, Mapping, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from 实验配置 import UPDATE_FREQUENCIES
from 输出路径 import etf_pool_initial_dir


INPUT_FILE = PROJECT_ROOT / "outputs" / "etf_data" / "etf_data.csv"

# ============================== 筛选参数 ==============================
# 直接运行本脚本时使用的频率；总运行器以后会显式传入daily或monthly。
UPDATE_FREQUENCY = "daily"
START_YEAR = 2021
END_YEAR = 2026
# 最后一个已经完整结束的自然月；不要填写仍在进行中的月份。
END_MONTH = 7
# 只处理这里指定的大类；以后增删类型只需要修改这个元组。
SELECTED_MAJOR_CATEGORIES = ("股票型ETF",)
MIN_LISTED_YEARS = 1
MIN_FUND_SCALE = 100_000_000.0
TURNOVER_LOOKBACK_DAYS = 20
MIN_AVERAGE_AMOUNT = 20_000_000.0
# ====================================================================

REQUIRED_COLUMNS = {
    "日期",
    "代码",
    "上市日期",
    "类别大类",
    "对标指数",
    "benchmark",
    "对标指数代码",
    "规模",
    "成交额",
}


@dataclass(frozen=True)
class SelectionPlan:
    """一个调仓快照对应的筛选日和20个交易日窗口。"""

    selection_date: date
    turnover_dates: tuple[date, ...]

    @property
    def key(self) -> date:
        return self.selection_date

    @property
    def file_name(self) -> str:
        return self.selection_date.strftime("%Y_%m_%d.csv")


@dataclass(frozen=True)
class Candidate:
    """通过三个基础筛选条件的ETF。"""

    row: Mapping[str, str]
    scale: float

    @property
    def code(self) -> str:
        return clean_text(self.row.get("代码"))


class UnionFind:
    """合并共享对标指数或 benchmark 的ETF。"""

    def __init__(self, size: int) -> None:
        self.parent = list(range(size))

    def find(self, item: int) -> int:
        while self.parent[item] != item:
            self.parent[item] = self.parent[self.parent[item]]
            item = self.parent[item]
        return item

    def union(self, left: int, right: int) -> None:
        left_root = self.find(left)
        right_root = self.find(right)
        if left_root != right_root:
            self.parent[right_root] = left_root


def clean_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def parse_date(value: object, field_name: str) -> date:
    text = clean_text(value)
    for date_format in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text[:10], date_format).date()
        except ValueError:
            continue
    raise ValueError(f"{field_name}不是有效日期：{text!r}")


def parse_number(value: object) -> float | None:
    text = clean_text(value).replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalized_group_value(value: object) -> str:
    """只清理空白和大小写，不擅自改写指数名称。"""

    return " ".join(clean_text(value).split()).casefold()


def selected_major_category(value: object) -> bool:
    """判断一行数据是否属于参数指定的大类。"""

    row_categories = {
        category.strip()
        for category in clean_text(value).split("|")
        if category.strip()
    }
    return bool(row_categories.intersection(SELECTED_MAJOR_CATEGORIES))


def listed_anniversary(listed_date: date, years: int) -> date:
    try:
        return listed_date.replace(year=listed_date.year + years)
    except ValueError:
        # 2月29日上市时，非闰年的周年日按2月28日处理。
        return listed_date.replace(
            year=listed_date.year + years,
            month=2,
            day=28,
        )


def inspect_input_dates(path: Path) -> tuple[list[str], list[date]]:
    if not path.exists():
        raise FileNotFoundError(f"找不到ETF数据文件：{path}")

    available_dates: set[date] = set()
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = list(reader.fieldnames or [])
        missing_columns = REQUIRED_COLUMNS - set(fieldnames)
        if missing_columns:
            raise ValueError(
                f"etf_data.csv 缺少字段：{sorted(missing_columns)}"
            )
        for row_number, row in enumerate(reader, start=2):
            try:
                available_dates.add(parse_date(row.get("日期"), "日期"))
            except ValueError as exc:
                raise ValueError(f"第{row_number}行日期无效") from exc

    if not available_dates:
        raise ValueError(f"ETF数据文件没有有效数据：{path}")
    return fieldnames, sorted(available_dates)


def build_selection_plans(
    available_dates: Sequence[date],
    update_frequency: str,
) -> list[SelectionPlan]:
    if update_frequency not in UPDATE_FREQUENCIES:
        raise ValueError(
            f"UPDATE_FREQUENCY必须是{UPDATE_FREQUENCIES}之一，"
            f"当前为{update_frequency!r}"
        )

    latest_available_date = available_dates[-1]
    configured_start_date = date(START_YEAR, 1, 1)
    configured_end_date = date(
        END_YEAR,
        END_MONTH,
        calendar.monthrange(END_YEAR, END_MONTH)[1],
    )
    effective_end_date = min(configured_end_date, latest_available_date)
    if configured_start_date > effective_end_date:
        raise ValueError(
            f"ETF数据最晚只到{latest_available_date}，"
            f"无法从{START_YEAR}年开始筛选"
        )

    daily_plans: list[SelectionPlan] = []
    for selection_date in available_dates:
        if not configured_start_date <= selection_date <= effective_end_date:
            continue
        selection_position = bisect_right(available_dates, selection_date)
        turnover_dates = tuple(
            available_dates[
                selection_position - TURNOVER_LOOKBACK_DAYS : selection_position
            ]
        )
        if len(turnover_dates) != TURNOVER_LOOKBACK_DAYS:
            continue
        daily_plans.append(
            SelectionPlan(
                selection_date=selection_date,
                turnover_dates=turnover_dates,
            )
        )
    if not daily_plans:
        raise ValueError(
            f"{configured_start_date}至{effective_end_date}之间没有满足"
            f"{TURNOVER_LOOKBACK_DAYS}日成交额回看要求的交易日"
        )

    if update_frequency == "daily":
        return daily_plans

    last_plan_by_month: dict[tuple[int, int], SelectionPlan] = {}
    for plan in daily_plans:
        month_key = plan.selection_date.year, plan.selection_date.month
        last_plan_by_month[month_key] = plan
    return list(last_plan_by_month.values())


def iter_daily_data(
    path: Path,
    plans: Sequence[SelectionPlan],
) -> Iterator[
    tuple[SelectionPlan, dict[str, dict[str, str]], Mapping[str, float]]
]:
    """按日期流式读取CSV，并维护包含当日的20日成交额滚动合计。"""

    plans_by_date = {plan.selection_date: plan for plan in plans}
    first_plan_date = plans[0].selection_date
    last_plan_date = plans[-1].selection_date
    amount_window: deque[dict[str, float]] = deque()
    rolling_amount_sums: dict[str, float] = defaultdict(float)
    previous_date: date | None = None

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for current_date, date_rows in groupby(
            reader,
            key=lambda row: parse_date(row.get("日期"), "日期"),
        ):
            if previous_date is not None and current_date <= previous_date:
                raise ValueError(
                    "etf_data.csv 必须按日期严格升序排列，"
                    f"但 {current_date} 出现在 {previous_date} 之后"
                )
            previous_date = current_date

            snapshot: dict[str, dict[str, str]] = {}
            daily_amounts: dict[str, float] = defaultdict(float)
            for row in date_rows:
                if not selected_major_category(row.get("类别大类")):
                    continue
                code = clean_text(row.get("代码"))
                if not code:
                    continue
                snapshot[code] = dict(row)
                # 空成交额（包括停牌日）按0处理，只累计有效数字。
                daily_amounts[code] += parse_number(row.get("成交额")) or 0.0

            amount_window.append(dict(daily_amounts))
            for code, amount in daily_amounts.items():
                rolling_amount_sums[code] += amount
            if len(amount_window) > TURNOVER_LOOKBACK_DAYS:
                expired_amounts = amount_window.popleft()
                for code, amount in expired_amounts.items():
                    rolling_amount_sums[code] -= amount

            if current_date < first_plan_date:
                continue
            if current_date > last_plan_date:
                break
            plan = plans_by_date.get(current_date)
            if plan is None:
                continue
            if len(amount_window) != TURNOVER_LOOKBACK_DAYS:
                raise RuntimeError(
                    f"{current_date}没有完整的{TURNOVER_LOOKBACK_DAYS}日成交额窗口"
                )
            yield plan, snapshot, rolling_amount_sums


def base_filter_candidates(
    plan: SelectionPlan,
    snapshot: Mapping[str, Mapping[str, str]],
    amount_sums: Mapping[str, float],
) -> list[Candidate]:
    candidates: list[Candidate] = []
    for code in sorted(snapshot):
        row = snapshot[code]
        if not clean_text(row.get("对标指数代码")):
            continue
        try:
            listed_date = parse_date(row.get("上市日期"), "上市日期")
        except ValueError:
            continue
        if listed_anniversary(listed_date, MIN_LISTED_YEARS) > plan.selection_date:
            continue

        scale = parse_number(row.get("规模"))
        if scale is None or scale <= MIN_FUND_SCALE:
            continue

        average_amount = (
            amount_sums.get(code, 0.0) / TURNOVER_LOOKBACK_DAYS
        )
        if average_amount <= MIN_AVERAGE_AMOUNT:
            continue

        candidates.append(Candidate(row=row, scale=scale))
    return candidates


def deduplicate_by_benchmark(
    candidates: Sequence[Candidate],
) -> list[Candidate]:
    if not candidates:
        return []

    union_find = UnionFind(len(candidates))
    for field_name in ("对标指数", "benchmark"):
        first_position_by_value: dict[str, int] = {}
        for position, candidate in enumerate(candidates):
            value = normalized_group_value(candidate.row.get(field_name))
            if not value:
                # 空值不能作为分组依据，否则所有空值ETF会被错误合并。
                continue
            first_position = first_position_by_value.setdefault(value, position)
            union_find.union(first_position, position)

    positions_by_group: dict[int, list[int]] = defaultdict(list)
    for position in range(len(candidates)):
        positions_by_group[union_find.find(position)].append(position)

    winners: list[Candidate] = []
    for positions in positions_by_group.values():
        # 规模相同时按代码升序确定唯一结果。
        winner_position = min(
            positions,
            key=lambda position: (
                -candidates[position].scale,
                candidates[position].code,
            ),
        )
        winners.append(candidates[winner_position])
    return sorted(winners, key=lambda candidate: candidate.code)


def write_daily_file(
    path: Path,
    fieldnames: Sequence[str],
    candidates: Iterable[Candidate],
) -> int:
    rows = list(candidates)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=fieldnames,
            extrasaction="ignore",
        )
        writer.writeheader()
        for candidate in rows:
            writer.writerow(candidate.row)
    return len(rows)


def clear_previous_csv_outputs(output_dir: Path) -> int:
    """删除输出目录第一层已有的CSV，避免保留旧结果。"""

    old_files = sorted(path for path in output_dir.glob("*.csv") if path.is_file())
    for path in old_files:
        path.unlink()
    return len(old_files)


def main(update_frequency: str = UPDATE_FREQUENCY) -> None:
    if START_YEAR > END_YEAR:
        raise ValueError("START_YEAR不能晚于END_YEAR")
    if not 1 <= END_MONTH <= 12:
        raise ValueError("END_MONTH必须在1至12之间")
    if TURNOVER_LOOKBACK_DAYS <= 0:
        raise ValueError("TURNOVER_LOOKBACK_DAYS必须大于0")
    if not SELECTED_MAJOR_CATEGORIES:
        raise ValueError("SELECTED_MAJOR_CATEGORIES不能为空")

    print(f"读取ETF数据：{INPUT_FILE}", flush=True)
    print(
        f"筛选ETF大类：{'、'.join(SELECTED_MAJOR_CATEGORIES)}",
        flush=True,
    )
    print(f"更新频率：{update_frequency}", flush=True)
    fieldnames, available_dates = inspect_input_dates(INPUT_FILE)
    plans = build_selection_plans(available_dates, update_frequency)

    output_dir = etf_pool_initial_dir(update_frequency)
    output_dir.mkdir(parents=True, exist_ok=True)
    removed_count = clear_previous_csv_outputs(output_dir)
    if removed_count:
        print(f"已清理旧的初筛CSV：{removed_count} 个", flush=True)
    written_count = 0
    for plan, snapshot, amount_sums in iter_daily_data(INPUT_FILE, plans):
        candidates = base_filter_candidates(
            plan,
            snapshot,
            amount_sums,
        )
        selected = deduplicate_by_benchmark(candidates)
        output_path = output_dir / plan.file_name
        selected_count = write_daily_file(output_path, fieldnames, selected)
        written_count += 1
        print(
            f"筛选日 {plan.selection_date}："
            f"基础条件通过 {len(candidates)} 只，基准去重后 {selected_count} 只，"
            f"已保存到 {output_path}",
            flush=True,
        )

    expected_count = len(plans)
    if written_count != expected_count:
        raise RuntimeError(
            f"计划生成 {expected_count} 个CSV，实际只生成 {written_count} 个"
        )
    print(
        f"完成：共生成 {expected_count} 个CSV文件，输出目录：{output_dir}",
        flush=True,
    )


if __name__ == "__main__":
    main()
