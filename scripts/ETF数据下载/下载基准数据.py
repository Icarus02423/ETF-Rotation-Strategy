#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从同花顺 iFinD 下载一个基准指数的日收盘数据。

输出目录中不生成 raw、总表或宽表等重复文件，只生成：
outputs/benchmark_data/基准名称.csv

指数点位直接使用历史行情接口返回的原始 close，不做前复权或后复权。
"""

from __future__ import annotations

import csv
import importlib.util
import math
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path
from types import ModuleType
from typing import Any, Iterator, Mapping


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent
ETF_DOWNLOADER_SCRIPT = SCRIPT_DIR / "下载ETF数据.py"


# ============================================================================
# 一、用户参数：以后通常只需要修改 BENCHMARK_CODE
# ============================================================================

# 同花顺指数代码，例如：中证A500指数为 000510.CSI。
BENCHMARK_CODE = "000510.CSI"

# 与现有 ETF 下载脚本使用相同的开始日期。
START_DATE = "2021-01-01"
END_DATE = "2026-07-30"

# ============================================================================
# 用户参数区域结束；下面是程序实现，通常不需要修改
# ============================================================================


OUTPUT_DIR = PROJECT_ROOT / "outputs" / "benchmark_data"
OUTPUT_COLUMNS = ["日期", "代码", "名称", "收盘价"]
NAME_INDICATOR = "ths_index_short_name_index"
HISTORY_INDICATOR = "close"
INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def load_etf_downloader() -> ModuleType:
    """复用 ETF 下载器已经验证过的鉴权、重试和请求实现。"""

    if not ETF_DOWNLOADER_SCRIPT.is_file():
        raise FileNotFoundError(f"未找到现有 ETF 下载脚本：{ETF_DOWNLOADER_SCRIPT}")

    module_name = "etf_downloader_for_benchmark"
    spec = importlib.util.spec_from_file_location(module_name, ETF_DOWNLOADER_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法载入现有 ETF 下载脚本：{ETF_DOWNLOADER_SCRIPT}")

    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def validate_config(downloader: ModuleType) -> tuple[str, date, date]:
    code = downloader.normalize_code(BENCHMARK_CODE)
    if not code:
        raise ValueError("BENCHMARK_CODE 不能为空。")

    start = downloader.parse_date(START_DATE, "START_DATE", allow_today=False)
    end = downloader.parse_date(END_DATE, "END_DATE", allow_today=False)
    if start > end:
        raise ValueError("START_DATE 不能晚于 END_DATE。")
    if not downloader.IFIND_REFRESH_TOKEN.strip():
        raise ValueError(
            f"未读取到 IFIND_REFRESH_TOKEN，请在 {downloader.ENV_FILE} 中填写该项。"
        )
    return code, start, end


def clean_name(value: Any) -> str:
    text = downloader_clean_value(value)
    return re.sub(r"\s+", " ", text).strip()


def downloader_clean_value(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text in {"", "--", "None", "null", "NULL"} else text


def safe_output_name(name: str) -> str:
    """保留同花顺名称，同时避免名称中的字符破坏文件路径。"""

    safe_name = INVALID_FILENAME_CHARS.sub("_", name).strip().rstrip(".")
    if not safe_name or safe_name in {".", ".."}:
        raise ValueError(f"同花顺返回的基准名称不能作为文件名：{name!r}")
    return safe_name


def fetch_benchmark_name(
    downloader: ModuleType,
    client: Any,
    code: str,
) -> str:
    """名称无法由历史行情返回，因此仅为名称使用一次基本信息接口。"""

    result = downloader.fetch_basic_fields(
        client,
        [code],
        fields=[("名称", NAME_INDICATOR)],
    )
    name = clean_name(result.get(code, {}).get("名称"))
    if not name:
        raise RuntimeError(
            f"同花顺未返回 {code} 的指数名称，请检查 BENCHMARK_CODE 是否正确。"
        )
    return name


def parse_market_date(value: Any) -> date:
    text = downloader_clean_value(value)
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"同花顺返回了无法识别的日期：{text!r}")


def parse_close(value: Any, current_date: date) -> float:
    text = downloader_clean_value(value).replace(",", "")
    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(f"{current_date} 的收盘价不是有效数字：{value!r}") from exc
    if not math.isfinite(number) or number <= 0:
        raise ValueError(f"{current_date} 的收盘价必须是正的有限数：{value!r}")
    return number


def iter_history_rows(
    downloader: ModuleType,
    payload: Mapping[str, Any],
    expected_code: str,
    range_start: date,
    range_end: date,
) -> Iterator[tuple[date, float]]:
    tables = payload.get("tables")
    if not isinstance(tables, list):
        raise RuntimeError("同花顺历史行情响应中缺少 tables。")

    found_code = False
    for item in tables:
        if not isinstance(item, dict):
            continue
        returned_code = downloader.normalize_code(item.get("thscode"))
        if returned_code != expected_code:
            continue
        found_code = True

        times = item.get("time")
        table = item.get("table")
        if not isinstance(times, list) or not isinstance(table, dict):
            raise RuntimeError(f"{expected_code} 的历史行情表结构不完整。")
        closes = downloader.table_values_case_insensitive(table, HISTORY_INDICATOR)
        if len(times) != len(closes):
            raise RuntimeError(
                f"{expected_code} 的日期数量与收盘价数量不一致："
                f"{len(times)} != {len(closes)}。"
            )

        for raw_date, raw_close in zip(times, closes):
            current_date = parse_market_date(raw_date)
            if range_start <= current_date <= range_end:
                yield current_date, parse_close(raw_close, current_date)

    if not found_code:
        raise RuntimeError(f"同花顺历史行情响应中未找到 {expected_code}。")


def download_history(
    downloader: ModuleType,
    client: Any,
    code: str,
    start: date,
    end: date,
) -> dict[date, float]:
    ranges = list(downloader.split_date_range(start, end))
    rows: dict[date, float] = {}

    for position, (range_start, range_end) in enumerate(ranges, start=1):
        print(
            f"正在下载 {code} 数据（本次指标：收盘价，"
            f"{range_start} 至 {range_end}）... ({position}/{len(ranges)})",
            flush=True,
        )
        payload = client.post(
            "cmd_history_quotation",
            {
                "codes": code,
                "indicators": HISTORY_INDICATOR,
                "startdate": range_start.isoformat(),
                "enddate": range_end.isoformat(),
                "functionpara": {"Interval": "D", "Fill": "Omit"},
            },
            progress_label=f"{code} 收盘价",
        )
        chunk_rows = dict(
            iter_history_rows(
                downloader,
                payload,
                code,
                range_start,
                range_end,
            )
        )
        rows.update(chunk_rows)
        print(f"✅ 本段取得 {len(chunk_rows)} 个交易日", flush=True)

    if not rows:
        raise RuntimeError(f"同花顺未返回 {code} 在指定日期区间内的收盘数据。")
    return rows


def write_output(
    downloader: ModuleType,
    output_path: Path,
    code: str,
    name: str,
    rows: Mapping[date, float],
) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    temp_path, handle = downloader.atomic_csv_writer(output_path)
    try:
        with handle:
            writer = csv.DictWriter(handle, fieldnames=OUTPUT_COLUMNS)
            writer.writeheader()
            for current_date in sorted(rows):
                writer.writerow(
                    {
                        "日期": current_date.isoformat(),
                        "代码": code,
                        "名称": name,
                        "收盘价": format(rows[current_date], ".15g"),
                    }
                )
        temp_path.replace(output_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def main() -> None:
    started_at = time.perf_counter()
    downloader = load_etf_downloader()
    code, start, end = validate_config(downloader)

    print("同花顺基准数据下载工具", flush=True)
    print("=" * 60, flush=True)
    print(f"基准代码：{code}", flush=True)
    print(f"下载日期：{start} 至 {end}", flush=True)
    print("行情接口：cmd_history_quotation / close", flush=True)
    print("复权方式：指数原始点位，不做前复权或后复权", flush=True)
    print("正在登录同花顺...", flush=True)

    access_token = downloader.get_access_token(
        downloader.IFIND_REFRESH_TOKEN,
        "同花顺",
    )
    print("✅ Access Token 已设置", flush=True)

    basic_client = downloader.IFindClient("基本面数据token", access_token)
    history_client = downloader.IFindClient("历史行情token", access_token)

    print("正在从同花顺获取基准名称...", flush=True)
    name = fetch_benchmark_name(downloader, basic_client, code)
    print(f"✅ 已识别基准：{name}（{code}）", flush=True)

    output_path = OUTPUT_DIR / f"{safe_output_name(name)}.csv"
    rows = download_history(downloader, history_client, code, start, end)
    write_output(downloader, output_path, code, name, rows)
    print(f"✅ 成功保存 {name} 数据到 {output_path}", flush=True)

    print("\n" + "=" * 60, flush=True)
    print("全部基准数据处理完成！", flush=True)
    print(f"基准：{name}（{code}）", flush=True)
    print(f"文件保存在：{output_path}", flush=True)
    print(f"共 {len(rows)} 个交易日", flush=True)
    print(f"用时：{time.perf_counter() - started_at:.1f}秒", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"❌ 基准数据下载失败：{exc}", flush=True)
        raise SystemExit(1) from exc
