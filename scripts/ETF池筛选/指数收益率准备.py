#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按月下载初筛ETF对应指数的历史收盘价，并计算60个共同交易区间收益率。

本脚本只负责聚类前的数据准备，不计算相关矩阵，也不执行层次聚类。

数据来源：
- 从 outputs/etf_pool/initial 读取每月ETF初筛结果；
- 指数收盘价优先使用 iFinD 的 cmd_history_quotation；
- iFinD失败或数据不足时，仅对参数区明确配置的同一指数使用AKShare回退；
- 唯一下载指标为 close；
- 每月先找出全部指数共有的最近61个收盘日期；
- 日收益率 = 本共同日期收盘价 / 上一共同日期收盘价 - 1。

输出：
- outputs/etf_pool/index_returns/window_60/YYYY_MM_DD.csv；
- 每个月最后一个实际交易日对应一个独立CSV；
- 每个指数使用完全相同的61个共同收盘日期计算60个收益率；
- 不会把不同月份的数据合并成一张总表。

断点规则：
- 已存在且结构完整的月度CSV直接跳过，不调用 cmd_history_quotation；
- 输出不完整或当月指数池发生变化时，重新下载并原子覆盖该月CSV；
- 同一次运行内，同一指数的重叠月份区间只请求一次。
"""

from __future__ import annotations

import csv
import importlib
import importlib.util
import math
import sys
import warnings
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from types import ModuleType
from typing import Any, Iterable, Mapping, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[2]
INITIAL_SELECTION_DIR = PROJECT_ROOT / "outputs" / "etf_pool" / "initial"
ETF_DOWNLOADER_SCRIPT = (
    PROJECT_ROOT / "scripts" / "ETF数据下载" / "下载ETF数据.py"
)

# ============================== 下载参数 ==============================
RETURN_TRADING_DAYS = 60

# 60个收益率需要61个有效收盘价。首次向前预留180个自然日，通常足以
# 覆盖不同市场的周末和节假日；仍不足时自动扩展到730个自然日。
INITIAL_LOOKBACK_CALENDAR_DAYS = 180
MAX_LOOKBACK_CALENDAR_DAYS = 730

# 全部指数的共同收盘日期不足61个时，按以下自然日范围逐步向前扩展。
# 每次只请求之前没有覆盖的日期区间。
LOOKBACK_CALENDAR_DAY_STEPS = (180, 240, 300, 365, 540, 730)

# False：完整的月度输出直接跳过，避免重复调用历史行情接口。
# True：忽略已有月度输出并重新下载、覆盖全部月份。
OVERWRITE_EXISTING_MONTHS = False

# iFinD无法取得某个指数时使用AKShare的同一指数数据。
# 每项格式：iFinD代码: (AKShare接口名, AKShare标的, 日期列, 收盘价列)
# 以后新增回退指数时只需在这里增加一行，不需要修改下载流程。
# 必须确认两边是同一个指数，不能填写ETF、期货或其他替代指数。
AKSHARE_FALLBACK_CONFIG: dict[str, tuple[str, str, str, str]] = {
    "SPX.GI": ("index_us_stock_sina", ".INX", "date", "close"),
    "DJI.GI": ("index_us_stock_sina", ".DJI", "date", "close"),
}
# ====================================================================

OUTPUT_DIR = (
    PROJECT_ROOT
    / "outputs"
    / "etf_pool"
    / "index_returns"
    / f"window_{RETURN_TRADING_DAYS}"
)

HISTORY_ENDPOINT = "cmd_history_quotation"
HISTORY_INDICATOR = "close"

INPUT_REQUIRED_COLUMNS = {
    "日期",
    "对标指数",
    "对标指数代码",
}
OUTPUT_COLUMNS = [
    "月末交易日",
    "收益日期",
    "对标指数代码",
    "对标指数",
    "收盘价",
    "日收益率",
]


@dataclass(frozen=True)
class MonthInput:
    """一个月末交易日及当月需要下载的指数代码。"""

    selection_date: date
    index_names: Mapping[str, str]
    source_file: Path

    @property
    def output_file(self) -> Path:
        return OUTPUT_DIR / self.selection_date.strftime("%Y_%m_%d.csv")


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text in {"", "--", "None", "null", "NULL"} else text


def parse_date(value: object, field_name: str) -> date:
    text = clean_text(value)
    for date_format in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text[:10], date_format).date()
        except ValueError:
            continue
    raise ValueError(f"{field_name}不是有效日期：{text!r}")


def parse_positive_float(value: object) -> float | None:
    text = clean_text(value).replace(",", "")
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if not math.isfinite(number) or number <= 0:
        return None
    return number


def load_etf_downloader_module() -> ModuleType:
    """复用ETF下载脚本的鉴权和HTTP客户端，避免复制或分叉token配置。"""

    if not ETF_DOWNLOADER_SCRIPT.exists():
        raise FileNotFoundError(
            f"找不到ETF下载脚本，无法读取iFinD鉴权配置：{ETF_DOWNLOADER_SCRIPT}"
        )
    module_name = "etf_downloader_for_index_history"
    spec = importlib.util.spec_from_file_location(module_name, ETF_DOWNLOADER_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载ETF下载脚本：{ETF_DOWNLOADER_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)

    required_names = ("IFIND_REFRESH_TOKEN", "get_access_token", "post_json")
    missing_names = [name for name in required_names if not hasattr(module, name)]
    if missing_names:
        raise RuntimeError(
            f"ETF下载脚本缺少必要配置或函数：{'、'.join(missing_names)}"
        )
    return module


def read_one_month(path: Path) -> MonthInput:
    selection_dates: set[date] = set()
    index_names: dict[str, str] = {}
    blank_code_rows: list[int] = []

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fields = set(reader.fieldnames or [])
        missing_columns = INPUT_REQUIRED_COLUMNS - fields
        if missing_columns:
            raise ValueError(f"{path.name} 缺少字段：{sorted(missing_columns)}")

        for row_number, row in enumerate(reader, start=2):
            selection_dates.add(parse_date(row.get("日期"), "日期"))
            index_code = clean_text(row.get("对标指数代码")).upper()
            if not index_code:
                blank_code_rows.append(row_number)
                continue
            index_name = clean_text(row.get("对标指数"))
            current_name = index_names.get(index_code, "")
            if not current_name or (index_name and index_name < current_name):
                index_names[index_code] = index_name

    if blank_code_rows:
        preview = "、".join(str(row) for row in blank_code_rows[:5])
        suffix = "……" if len(blank_code_rows) > 5 else ""
        raise ValueError(
            f"{path.name} 仍有对标指数代码为空的ETF（CSV行：{preview}{suffix}）。"
            "请先重新运行 ETF初筛.py。"
        )
    if len(selection_dates) != 1:
        raise ValueError(
            f"{path.name} 的日期不唯一："
            f"{sorted(current_date.isoformat() for current_date in selection_dates)}"
        )
    if not index_names:
        raise ValueError(f"{path.name} 没有有效的对标指数代码")

    return MonthInput(
        selection_date=next(iter(selection_dates)),
        index_names=dict(sorted(index_names.items())),
        source_file=path,
    )


def discover_month_inputs() -> list[MonthInput]:
    if not INITIAL_SELECTION_DIR.exists():
        raise FileNotFoundError(f"找不到ETF初筛输出目录：{INITIAL_SELECTION_DIR}")

    files = sorted(
        path
        for path in INITIAL_SELECTION_DIR.glob("*.csv")
        if path.is_file()
    )
    if not files:
        raise FileNotFoundError(f"ETF初筛输出目录中没有CSV：{INITIAL_SELECTION_DIR}")

    months_by_date: dict[date, MonthInput] = {}
    for path in files:
        month = read_one_month(path)
        existing = months_by_date.get(month.selection_date)
        if existing is not None:
            if dict(existing.index_names) != dict(month.index_names):
                raise ValueError(
                    f"{existing.source_file.name} 与 {path.name} 对应同一交易日"
                    f" {month.selection_date}，但指数池不同。请先重新运行ETF初筛。"
                )
            continue
        months_by_date[month.selection_date] = month
    return [months_by_date[current_date] for current_date in sorted(months_by_date)]


def month_output_is_complete(month: MonthInput) -> bool:
    """只有全部指数具有相同的60个收益日期时才跳过。"""

    path = month.output_file
    if OVERWRITE_EXISTING_MONTHS or not path.exists():
        return False

    rows_by_code: dict[str, list[tuple[date, float, float]]] = defaultdict(list)
    seen_dates: dict[str, set[date]] = defaultdict(set)
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            if list(reader.fieldnames or []) != OUTPUT_COLUMNS:
                return False
            for row in reader:
                if clean_text(row.get("月末交易日")) != month.selection_date.isoformat():
                    return False
                index_code = clean_text(row.get("对标指数代码")).upper()
                if index_code not in month.index_names:
                    return False
                return_date = parse_date(row.get("收益日期"), "收益日期")
                if return_date > month.selection_date:
                    return False
                if return_date in seen_dates[index_code]:
                    return False
                close = parse_positive_float(row.get("收盘价"))
                if close is None:
                    return False
                try:
                    return_value = float(clean_text(row.get("日收益率")))
                except ValueError:
                    return False
                if not math.isfinite(return_value):
                    return False
                seen_dates[index_code].add(return_date)
                rows_by_code[index_code].append(
                    (return_date, close, return_value)
                )
    except (OSError, csv.Error, ValueError):
        return False

    expected_codes = set(month.index_names)
    if set(rows_by_code) != expected_codes or not all(
        len(rows_by_code[code]) == RETURN_TRADING_DAYS
        for code in expected_codes
    ):
        return False

    common_return_dates = {
        tuple(sorted(seen_dates[code])) for code in expected_codes
    }
    if len(common_return_dates) != 1:
        return False

    # 从第二个输出日期开始，可直接用相邻共同日期的收盘价复核区间收益率。
    for index_code in expected_codes:
        ordered = sorted(rows_by_code[index_code])
        for previous, current in zip(ordered, ordered[1:]):
            expected_return = current[1] / previous[1] - 1.0
            if not math.isclose(
                current[2],
                expected_return,
                rel_tol=1e-9,
                abs_tol=1e-12,
            ):
                return False
    return True


def merge_windows(
    windows: Iterable[tuple[date, date]],
) -> list[tuple[date, date]]:
    ordered = sorted(windows)
    if not ordered:
        return []

    merged: list[tuple[date, date]] = []
    current_start, current_end = ordered[0]
    for next_start, next_end in ordered[1:]:
        if next_start <= current_end + timedelta(days=1):
            current_end = max(current_end, next_end)
        else:
            merged.append((current_start, current_end))
            current_start, current_end = next_start, next_end
    merged.append((current_start, current_end))
    return merged


def merge_request_windows(selection_dates: Iterable[date]) -> list[tuple[date, date]]:
    return merge_windows(
        (
            current_date - timedelta(days=INITIAL_LOOKBACK_CALENDAR_DAYS),
            current_date,
        )
        for current_date in set(selection_dates)
    )


def uncovered_windows(
    start_date: date,
    end_date: date,
    covered_windows: Iterable[tuple[date, date]],
) -> list[tuple[date, date]]:
    """返回目标区间中尚未请求过的自然日范围。"""

    if start_date > end_date:
        return []
    uncovered: list[tuple[date, date]] = []
    cursor = start_date
    for covered_start, covered_end in merge_windows(covered_windows):
        if covered_end < cursor:
            continue
        if covered_start > end_date:
            break
        if covered_start > cursor:
            uncovered.append(
                (cursor, min(end_date, covered_start - timedelta(days=1)))
            )
        cursor = max(cursor, covered_end + timedelta(days=1))
        if cursor > end_date:
            break
    if cursor <= end_date:
        uncovered.append((cursor, end_date))
    return uncovered


def table_values_case_insensitive(
    table: Mapping[str, Any],
    indicator: str,
) -> list[Any]:
    target = indicator.casefold()
    for key, values in table.items():
        if str(key).casefold() == target and isinstance(values, list):
            return values
    return []


def parse_close_payload(
    payload: Mapping[str, Any],
    expected_code: str,
) -> dict[date, float]:
    result: dict[date, float] = {}
    tables = payload.get("tables")
    if not isinstance(tables, list):
        return result

    for item in tables:
        if not isinstance(item, dict):
            continue
        returned_code = clean_text(item.get("thscode")).upper()
        if returned_code != expected_code:
            continue
        times = item.get("time")
        table = item.get("table")
        if not isinstance(times, list) or not isinstance(table, dict):
            continue
        closes = table_values_case_insensitive(table, HISTORY_INDICATOR)
        for position, raw_time in enumerate(times):
            if position >= len(closes):
                break
            try:
                current_date = parse_date(raw_time, "历史行情日期")
            except ValueError:
                continue
            close = parse_positive_float(closes[position])
            if close is not None:
                result[current_date] = close
    return result


def request_index_close_range(
    downloader: ModuleType,
    access_token: str,
    index_code: str,
    start_date: date,
    end_date: date,
) -> dict[date, float]:
    """固定调用cmd_history_quotation的close指标。"""

    body = {
        "codes": index_code,
        "indicators": HISTORY_INDICATOR,
        "startdate": start_date.isoformat(),
        "enddate": end_date.isoformat(),
        "functionpara": {},
    }
    payload = downloader.post_json(
        HISTORY_ENDPOINT,
        body=body,
        access_token=access_token,
        progress_label=f"{index_code}指数收盘价",
    )
    return parse_close_payload(payload, index_code)


def request_akshare_fallback_closes(
    index_code: str,
    start_date: date,
    end_date: date,
) -> dict[date, float]:
    """按参数区映射从AKShare读取同一指数的收盘价。"""

    config = AKSHARE_FALLBACK_CONFIG.get(index_code)
    if config is None:
        raise KeyError(f"{index_code}未配置AKShare回退")
    api_name, symbol, date_column, close_column = config

    try:
        # 当前项目Python使用LibreSSL；urllib3的兼容性提示不影响AKShare取数，
        # 仅在延迟导入AKShare时屏蔽这一条提示，避免污染下载日志。
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message="urllib3 v2 only supports OpenSSL.*",
            )
            akshare = importlib.import_module("akshare")
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "当前虚拟环境未安装akshare，请先执行：python -m pip install akshare"
        ) from exc

    api = getattr(akshare, api_name, None)
    if not callable(api):
        raise RuntimeError(f"当前akshare版本不存在接口：{api_name}")

    frame = api(symbol=symbol)
    if frame is None or not hasattr(frame, "to_dict"):
        raise RuntimeError(f"AKShare {api_name}({symbol})未返回表格数据")

    result: dict[date, float] = {}
    for row in frame.to_dict(orient="records"):
        if not isinstance(row, Mapping):
            continue
        try:
            current_date = parse_date(row.get(date_column), "AKShare历史行情日期")
        except ValueError:
            continue
        if current_date < start_date or current_date > end_date:
            continue
        close = parse_positive_float(row.get(close_column))
        if close is not None:
            result[current_date] = close
    if not result:
        raise RuntimeError(
            f"AKShare {api_name}({symbol})在 {start_date} 至 {end_date} 无有效收盘价"
        )
    return result


def common_close_dates(
    month: MonthInput,
    closes_by_code: Mapping[str, Mapping[date, float]],
) -> list[date]:
    """返回当月全部指数在月末前共同拥有收盘价的日期。"""

    common_dates: set[date] | None = None
    for index_code in month.index_names:
        available_dates = {
            current_date
            for current_date in closes_by_code.get(index_code, {})
            if current_date <= month.selection_date
        }
        if common_dates is None:
            common_dates = available_dates
        else:
            common_dates.intersection_update(available_dates)
        if not common_dates:
            return []
    return sorted(common_dates or set())


def months_without_enough_common_dates(
    months: Sequence[MonthInput],
    closes_by_code: Mapping[str, Mapping[date, float]],
) -> list[MonthInput]:
    required_close_count = RETURN_TRADING_DAYS + 1
    return [
        month
        for month in months
        if len(common_close_dates(month, closes_by_code))
        < required_close_count
    ]


def download_required_closes(
    downloader: ModuleType,
    access_token: str,
    pending_months: Sequence[MonthInput],
) -> tuple[dict[str, dict[date, float]], dict[str, str]]:
    selection_dates_by_code: dict[str, set[date]] = defaultdict(set)
    for month in pending_months:
        for index_code in month.index_names:
            selection_dates_by_code[index_code].add(month.selection_date)

    all_closes: dict[str, dict[date, float]] = defaultdict(dict)
    requested_windows: dict[str, list[tuple[date, date]]] = defaultdict(list)
    errors: dict[str, str] = {}
    codes = sorted(selection_dates_by_code)
    total = len(codes)
    for position, index_code in enumerate(codes, start=1):
        print(
            f"正在通过 {HISTORY_ENDPOINT} 下载 {index_code} 指数收盘价... "
            f"({position}/{total})",
            flush=True,
        )
        try:
            windows = merge_request_windows(selection_dates_by_code[index_code])
            for start_date, end_date in windows:
                requested_windows[index_code].append((start_date, end_date))
                all_closes[index_code].update(
                    request_index_close_range(
                        downloader,
                        access_token,
                        index_code,
                        start_date,
                        end_date,
                    )
                )
            print(
                f"✅ {index_code} {HISTORY_ENDPOINT} 请求完成，"
                f"取得 {len(all_closes[index_code])} 个有效收盘价",
                flush=True,
            )
        except Exception as exc:
            errors[index_code] = str(exc)
            print(f"❌ {index_code} 指数收盘价下载失败：{exc}", flush=True)

    # iFinD首次请求失败时，先按显式映射回退到AKShare的同一指数。
    for index_code in sorted(set(errors) & set(AKSHARE_FALLBACK_CONFIG)):
        selection_dates = selection_dates_by_code[index_code]
        api_name, symbol, _, _ = AKSHARE_FALLBACK_CONFIG[index_code]
        fallback_start = min(selection_dates) - timedelta(
            days=MAX_LOOKBACK_CALENDAR_DAYS
        )
        fallback_end = max(selection_dates)
        print(
            f"⚠️ {index_code} iFinD失败：{errors[index_code]}，改用AKShare "
            f"{api_name}({symbol})读取同一指数...",
            flush=True,
        )
        try:
            all_closes[index_code].update(
                request_akshare_fallback_closes(
                    index_code,
                    fallback_start,
                    fallback_end,
                )
            )
            requested_windows[index_code].append(
                (fallback_start, fallback_end)
            )
            errors.pop(index_code, None)
            print(
                f"✅ {index_code} 已通过AKShare {api_name}({symbol})取得 "
                f"{len(all_closes[index_code])} 个有效收盘价",
                flush=True,
            )
        except Exception as exc:
            errors[index_code] = f"AKShare回退失败：{exc}"
            print(f"❌ {index_code} AKShare回退失败：{exc}", flush=True)

    # 全部指数的共同日期不足61个时，逐步向前扩展；每次只请求尚未覆盖的区间。
    for lookback_days in LOOKBACK_CALENDAR_DAY_STEPS[1:]:
        unresolved_months = months_without_enough_common_dates(
            pending_months,
            all_closes,
        )
        if not unresolved_months:
            break
        unresolved_months = [
            month
            for month in unresolved_months
            if not any(index_code in errors for index_code in month.index_names)
        ]
        if not unresolved_months:
            break

        extra_windows_by_code: dict[str, list[tuple[date, date]]] = (
            defaultdict(list)
        )
        for month in unresolved_months:
            target_start = month.selection_date - timedelta(days=lookback_days)
            for index_code in month.index_names:
                if index_code in errors:
                    continue
                extra_windows_by_code[index_code].extend(
                    uncovered_windows(
                        target_start,
                        month.selection_date,
                        requested_windows[index_code],
                    )
                )

        for index_code in sorted(extra_windows_by_code):
            windows = merge_windows(extra_windows_by_code[index_code])
            if not windows:
                continue
            print(
                f"共同日期不足，正在通过 {HISTORY_ENDPOINT} 将 "
                f"{index_code} 向前扩展至 {lookback_days} 个自然日...",
                flush=True,
            )
            try:
                for start_date, end_date in windows:
                    requested_windows[index_code].append(
                        (start_date, end_date)
                    )
                    all_closes[index_code].update(
                        request_index_close_range(
                            downloader,
                            access_token,
                            index_code,
                            start_date,
                            end_date,
                        )
                    )
            except Exception as exc:
                errors[index_code] = str(exc)
                print(f"❌ {index_code} 向前扩展失败：{exc}", flush=True)

    # 向前扩展时iFinD失败的显式回退指数，再使用AKShare补足最大范围。
    for index_code in sorted(set(errors) & set(AKSHARE_FALLBACK_CONFIG)):
        selection_dates = selection_dates_by_code[index_code]
        api_name, symbol, _, _ = AKSHARE_FALLBACK_CONFIG[index_code]
        fallback_start = min(selection_dates) - timedelta(
            days=MAX_LOOKBACK_CALENDAR_DAYS
        )
        fallback_end = max(selection_dates)
        print(
            f"⚠️ {index_code} 向前扩展失败，改用AKShare "
            f"{api_name}({symbol})补充同一指数...",
            flush=True,
        )
        try:
            all_closes[index_code].update(
                request_akshare_fallback_closes(
                    index_code,
                    fallback_start,
                    fallback_end,
                )
            )
            errors.pop(index_code, None)
            print(
                f"✅ {index_code} AKShare补充完成，"
                f"共 {len(all_closes[index_code])} 个有效收盘价",
                flush=True,
            )
        except Exception as exc:
            errors[index_code] = f"AKShare回退失败：{exc}"
            print(f"❌ {index_code} AKShare回退失败：{exc}", flush=True)

    return dict(all_closes), errors


def monthly_return_rows(
    month: MonthInput,
    closes_by_code: Mapping[str, Mapping[date, float]],
) -> tuple[list[list[str]], list[str]]:
    required_close_count = RETURN_TRADING_DAYS + 1
    common_dates = common_close_dates(month, closes_by_code)
    selected_dates = common_dates[-required_close_count:]
    if len(selected_dates) < required_close_count:
        return [], sorted(month.index_names)

    rows: list[list[str]] = []
    for index_code in sorted(month.index_names):
        index_closes = closes_by_code[index_code]
        previous_close = index_closes[selected_dates[0]]
        for return_date in selected_dates[1:]:
            close = index_closes[return_date]
            daily_return = close / previous_close - 1.0
            rows.append(
                [
                    month.selection_date.isoformat(),
                    return_date.isoformat(),
                    index_code,
                    month.index_names[index_code],
                    format(close, ".15g"),
                    format(daily_return, ".15g"),
                ]
            )
            previous_close = close
    return rows, []


def write_month_output(path: Path, rows: Sequence[Sequence[str]]) -> None:
    temp_path = path.with_name(f".{path.name}.tmp")
    try:
        with temp_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(OUTPUT_COLUMNS)
            writer.writerows(rows)
        temp_path.replace(path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def validate_parameters() -> None:
    if RETURN_TRADING_DAYS <= 0:
        raise ValueError("RETURN_TRADING_DAYS必须大于0")
    if INITIAL_LOOKBACK_CALENDAR_DAYS <= RETURN_TRADING_DAYS:
        raise ValueError("INITIAL_LOOKBACK_CALENDAR_DAYS设置过短")
    if MAX_LOOKBACK_CALENDAR_DAYS < INITIAL_LOOKBACK_CALENDAR_DAYS:
        raise ValueError(
            "MAX_LOOKBACK_CALENDAR_DAYS不能小于INITIAL_LOOKBACK_CALENDAR_DAYS"
        )
    if (
        not LOOKBACK_CALENDAR_DAY_STEPS
        or LOOKBACK_CALENDAR_DAY_STEPS[0] != INITIAL_LOOKBACK_CALENDAR_DAYS
        or LOOKBACK_CALENDAR_DAY_STEPS[-1] != MAX_LOOKBACK_CALENDAR_DAYS
        or tuple(sorted(set(LOOKBACK_CALENDAR_DAY_STEPS)))
        != LOOKBACK_CALENDAR_DAY_STEPS
    ):
        raise ValueError(
            "LOOKBACK_CALENDAR_DAY_STEPS必须从初始回溯天数严格递增至最大回溯天数"
        )
    if HISTORY_ENDPOINT != "cmd_history_quotation":
        raise ValueError("指数数据必须使用历史行情接口cmd_history_quotation")
    if HISTORY_INDICATOR != "close":
        raise ValueError("指数收益率数据准备阶段只允许下载close")
    for index_code, config in AKSHARE_FALLBACK_CONFIG.items():
        if clean_text(index_code).upper() != index_code or len(config) != 4:
            raise ValueError(f"AKShare回退配置不合法：{index_code} -> {config}")
        if any(not clean_text(value) for value in config):
            raise ValueError(f"AKShare回退配置存在空值：{index_code} -> {config}")


def main() -> None:
    validate_parameters()
    months = discover_month_inputs()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    pending_months: list[MonthInput] = []
    for month in months:
        if month_output_is_complete(month):
            print(
                f"⏭️ {month.output_file.name} 的60日指数收益率已完整存在，"
                f"跳过 {HISTORY_ENDPOINT} 请求",
                flush=True,
            )
        else:
            pending_months.append(month)

    print(
        f"发现 {len(months)} 个月度指数池，其中 "
        f"{len(pending_months)} 个月需要下载或重建",
        flush=True,
    )
    if not pending_months:
        print(f"全部月度指数数据已存在：{OUTPUT_DIR}", flush=True)
        return

    downloader = load_etf_downloader_module()
    refresh_token = clean_text(downloader.IFIND_REFRESH_TOKEN)
    if not refresh_token:
        raise ValueError(f"IFIND_REFRESH_TOKEN为空：{ETF_DOWNLOADER_SCRIPT}")

    print("正在登录同花顺 iFinD...", flush=True)
    access_token = downloader.get_access_token(refresh_token, "iFinD账号")
    print(
        f"✅ Access Token已设置；本脚本只调用 {HISTORY_ENDPOINT} 请求指数close",
        flush=True,
    )

    closes_by_code, download_errors = download_required_closes(
        downloader,
        access_token,
        pending_months,
    )

    failed_months: dict[str, list[str]] = {}
    written_count = 0
    for month in pending_months:
        rows, insufficient_codes = monthly_return_rows(month, closes_by_code)
        failed_codes = sorted(
            set(insufficient_codes).union(
                code for code in month.index_names if code in download_errors
            )
        )
        if failed_codes:
            failed_months[month.output_file.name] = failed_codes
            print(
                f"❌ {month.output_file.name} 未输出："
                f"未取得全部指数共有的{RETURN_TRADING_DAYS + 1}个收盘日期，"
                f"代码：{'、'.join(failed_codes)}",
                flush=True,
            )
            continue

        write_month_output(month.output_file, rows)
        written_count += 1
        print(
            f"✅ 已保存 {month.output_file.name}："
            f"{len(month.index_names)} 个指数，"
            f"每个指数 {RETURN_TRADING_DAYS} 个共同区间收益率",
            flush=True,
        )

    print(
        f"完成：本次写入 {written_count} 个月度CSV，输出目录：{OUTPUT_DIR}",
        flush=True,
    )
    if failed_months:
        raise RuntimeError(
            f"仍有 {len(failed_months)} 个月度文件因共同收盘日期不足而未输出；"
            "查看上方日志中的具体指数代码。"
        )


if __name__ == "__main__":
    main()
