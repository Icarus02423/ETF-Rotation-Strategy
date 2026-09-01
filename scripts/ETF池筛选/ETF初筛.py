#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按月筛选ETF池，并按对标指数或 benchmark 去重。

筛选规则：
1. 截至当月最后一个交易日，ETF上市满1年；
2. 当日基金规模 > 1亿元；
3. 包含当日在内的过去20个市场交易日平均成交额 > 2000万元；
4. 对标指数代码不能为空，缺失的ETF在去重前排除；
5. 对标指数相同或 benchmark 相同的ETF归为一类，每类只保留规模最大的ETF。

说明：
- 只筛选 SELECTED_MAJOR_CATEGORIES 参数指定的ETF大类；
- 停牌日或成交额空值按0计入20日平均成交额，分母固定为20；
- 每次运行都会清理输出目录中已有的CSV，再生成本次结果；
- CSV文件名使用该月最后一个实际交易日，例如2021年1月使用2021_01_29.csv；
- CSV中的“日期”仍保留该月最后一个实际交易日；
- 输出列与 etf_data.csv 完全一致，包含“对标指数代码”；
- 对标指数代码用于排除缺失ETF，但不作为归类依据。
"""

from __future__ import annotations

import calendar
import csv
from bisect import bisect_right
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Mapping, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[2]
INPUT_FILE = PROJECT_ROOT / "outputs" / "etf_data" / "etf_data.csv"
OUTPUT_DIR = PROJECT_ROOT / "outputs" / "etf_pool" / "initial"

# ============================== 筛选参数 ==============================
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
class MonthPlan:
    """一个自然月对应的实际筛选日和20个交易日窗口。"""

    year: int
    month: int
    file_date: date
    selection_date: date
    turnover_dates: tuple[date, ...]

    @property
    def key(self) -> tuple[int, int]:
        return self.year, self.month

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


def build_month_plans(available_dates: Sequence[date]) -> list[MonthPlan]:
    latest_available_date = available_dates[-1]
    configured_end = (END_YEAR, END_MONTH)
    available_end = (
        latest_available_date.year,
        latest_available_date.month,
    )
    effective_end_year, effective_end_month = min(
        configured_end,
        available_end,
    )
    if (START_YEAR, 1) > (effective_end_year, effective_end_month):
        raise ValueError(
            f"ETF数据最晚只到{latest_available_date}，"
            f"无法从{START_YEAR}年开始筛选"
        )

    plans: list[MonthPlan] = []
    for year in range(START_YEAR, effective_end_year + 1):
        last_month = effective_end_month if year == effective_end_year else 12
        for month in range(1, last_month + 1):
            file_date = date(year, month, calendar.monthrange(year, month)[1])
            month_dates = [
                current_date
                for current_date in available_dates
                if current_date.year == year and current_date.month == month
            ]
            if not month_dates:
                raise ValueError(f"{year}年{month}月没有可用交易日数据")

            selection_date = month_dates[-1]
            selection_position = bisect_right(
                available_dates,
                selection_date,
            )
            turnover_dates = tuple(
                available_dates[
                    selection_position - TURNOVER_LOOKBACK_DAYS : selection_position
                ]
            )
            if len(turnover_dates) != TURNOVER_LOOKBACK_DAYS:
                raise ValueError(
                    f"{selection_date}之前不足{TURNOVER_LOOKBACK_DAYS}个交易日"
                )

            plans.append(
                MonthPlan(
                    year=year,
                    month=month,
                    file_date=file_date,
                    selection_date=selection_date,
                    turnover_dates=turnover_dates,
                )
            )
    return plans


def collect_month_data(
    path: Path,
    plans: Sequence[MonthPlan],
) -> tuple[
    dict[tuple[int, int], dict[str, dict[str, str]]],
    dict[tuple[int, int], dict[str, float]],
]:
    """第二遍读取CSV，只保留60个月末截面和对应的成交额合计。"""

    date_to_turnover_months: dict[date, list[tuple[int, int]]] = defaultdict(list)
    selection_date_to_month: dict[date, tuple[int, int]] = {}
    for plan in plans:
        selection_date_to_month[plan.selection_date] = plan.key
        for turnover_date in plan.turnover_dates:
            date_to_turnover_months[turnover_date].append(plan.key)

    snapshots: dict[tuple[int, int], dict[str, dict[str, str]]] = {
        plan.key: {} for plan in plans
    }
    amount_sums: dict[tuple[int, int], dict[str, float]] = {
        plan.key: defaultdict(float) for plan in plans
    }

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if not selected_major_category(row.get("类别大类")):
                continue
            code = clean_text(row.get("代码"))
            if not code:
                continue
            current_date = parse_date(row.get("日期"), "日期")

            # 空成交额（包括停牌日）按0处理，因此只累计有效数字即可。
            amount = parse_number(row.get("成交额")) or 0.0
            for month_key in date_to_turnover_months.get(current_date, ()):
                amount_sums[month_key][code] += amount

            month_key = selection_date_to_month.get(current_date)
            if month_key is not None:
                snapshots[month_key][code] = dict(row)

    return snapshots, amount_sums


def base_filter_candidates(
    plan: MonthPlan,
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


def write_month_file(
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


def clear_previous_csv_outputs() -> int:
    """删除输出目录第一层已有的CSV，避免保留旧命名或旧月份结果。"""

    old_files = sorted(path for path in OUTPUT_DIR.glob("*.csv") if path.is_file())
    for path in old_files:
        path.unlink()
    return len(old_files)


def main() -> None:
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
    fieldnames, available_dates = inspect_input_dates(INPUT_FILE)
    plans = build_month_plans(available_dates)
    snapshots, amount_sums = collect_month_data(INPUT_FILE, plans)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    removed_count = clear_previous_csv_outputs()
    if removed_count:
        print(f"已清理旧的初筛CSV：{removed_count} 个", flush=True)
    for plan in plans:
        candidates = base_filter_candidates(
            plan,
            snapshots[plan.key],
            amount_sums[plan.key],
        )
        selected = deduplicate_by_benchmark(candidates)
        output_path = OUTPUT_DIR / plan.file_name
        selected_count = write_month_file(output_path, fieldnames, selected)
        print(
            f"{plan.year}-{plan.month:02d}：筛选日 {plan.selection_date}，"
            f"基础条件通过 {len(candidates)} 只，基准去重后 {selected_count} 只，"
            f"已保存到 {output_path}",
            flush=True,
        )

    expected_count = len(plans)
    print(
        f"完成：共生成 {expected_count} 个CSV文件，输出目录：{OUTPUT_DIR}",
        flush=True,
    )


if __name__ == "__main__":
    main()
