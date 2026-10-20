#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从同花顺 iFinD 下载单个基准指数的日线，用于大盘择时研究。

直接运行：python scripts/下载基准数据.py
修改下方三个参数即可更换基准和日期；也支持 --code、--start、--end。
项目根目录 .env 中需有 IFIND_REFRESH_TOKEN，系统环境变量优先。
输出：outputs/benchmark_data/指数代码.csv（UTF-8-SIG，指数原始点位）。
导入此模块不会读取凭据、请求接口或写入文件。
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import socket
import ssl
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from http.client import HTTPException
from pathlib import Path
from typing import Any, Iterator, Mapping
from urllib.error import HTTPError, URLError
from urllib.request import HTTPSHandler, HTTPRedirectHandler, Request, build_opener


# 用户参数：通常只需要修改这里。
BENCHMARK_CODE = "000300.SH"  # 沪深300；请使用同花顺的完整指数代码。
START_DATE = "2017-01-01"    # 包含2021年回测开始前的RSRS预热数据。
END_DATE = "2026-09-30"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_ROOT / "outputs" / "benchmark_data"
API_BASE = "https://quantapi.51ifind.com/api/v1"
NAME_INDICATOR = "ths_index_short_name_index"
PRICE_FIELDS = ("open", "high", "low", "close")
OUTPUT_COLUMNS = ("日期", "代码", "名称", "开盘价", "最高价", "最低价", "收盘价")
TOKEN_KEY = "IFIND_REFRESH_TOKEN"


@dataclass(frozen=True)
class Config:
    benchmark_code: str = BENCHMARK_CODE
    start_date: str = START_DATE
    end_date: str = END_DATE


def parse_date(value: str) -> date:
    """配置日期严格使用 YYYY-MM-DD。"""
    if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise ValueError("日期必须使用 YYYY-MM-DD 格式。")
    return date.fromisoformat(value)


def validate_config(config: Config) -> tuple[str, date, date]:
    code = config.benchmark_code.strip().upper() if isinstance(config.benchmark_code, str) else ""
    if not re.fullmatch(r"[A-Z0-9][A-Z0-9_-]*\.[A-Z0-9]+", code):
        raise ValueError("请填写一个完整的基准代码，例如 000300.SH。")
    start, end = parse_date(config.start_date), parse_date(config.end_date)
    if start > end:
        raise ValueError("开始日期不能晚于结束日期。")
    if end > date.today():
        raise ValueError("结束日期不能晚于今天。")
    return code, start, end


def load_refresh_token(env_file: Path) -> str:
    """只解析所需变量，不执行 .env，也不打印其中的内容。"""
    token = os.environ.get(TOKEN_KEY, "").strip()
    if not token and env_file.is_file():
        for line in env_file.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line.startswith("export "):
                line = line[7:].lstrip()
            key, separator, value = line.partition("=")
            if not separator or key.strip() != TOKEN_KEY:
                continue
            value = value.strip()
            if value.startswith(("'", '"')):
                match = re.fullmatch(r"(['\"])(.*?)\1(?:\s*#.*)?", value)
                if match is None:
                    raise ValueError(".env 中 IFIND_REFRESH_TOKEN 的引号不完整。")
                token = match.group(2).strip()
            else:
                token = re.split(r"\s+#", value, maxsplit=1)[0].strip()
    if not token:
        raise ValueError("未找到 IFIND_REFRESH_TOKEN，请在项目根目录 .env 中配置。")
    if "\r" in token or "\n" in token:
        raise ValueError("IFIND_REFRESH_TOKEN 不能包含换行。")
    return token


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        # 令牌仅发送到固定的同花顺接口，不随重定向转发。
        return None


class IFindClient:
    """标准库 HTTP 客户端；仅对网络错误和 HTTP 429/5xx 有限重试。"""

    def __init__(self, timeout: float = 60, attempts: int = 3, retry_delay: float = 3):
        if timeout <= 0 or attempts < 1 or retry_delay < 0:
            raise ValueError("超时和尝试次数必须为正，重试间隔不能为负。")
        self.timeout = timeout
        self.attempts = attempts
        self.retry_delay = retry_delay
        context = ssl.create_default_context()
        # macOS 上沿用系统证书库，始终验证服务端证书。
        for cafile in (Path("/etc/ssl/cert.pem"), Path("/private/etc/ssl/cert.pem")):
            if cafile.is_file():
                context.load_verify_locations(cafile=str(cafile))
                break
        self.opener = build_opener(HTTPSHandler(context=context), _NoRedirect())
        self.access_token = ""
        self._secrets: list[str] = []

    def _post(self, endpoint: str, body: Mapping[str, Any], headers: Mapping[str, str]) -> dict:
        request = Request(
            f"{API_BASE}/{endpoint}",
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json", "ifindlang": "cn", **headers},
            method="POST",
        )
        for attempt in range(1, self.attempts + 1):
            try:
                with self.opener.open(request, timeout=self.timeout) as response:
                    raw = response.read()
            except HTTPError as exc:
                retryable = exc.code == 429 or 500 <= exc.code < 600
                if exc.fp is not None:
                    exc.close()
                if not retryable or attempt == self.attempts:
                    raise RuntimeError(f"同花顺请求失败（HTTP {exc.code}），请检查接口权限及网络。") from None
            except (URLError, TimeoutError, socket.timeout, ConnectionError, HTTPException) as exc:
                reason = getattr(exc, "reason", exc)
                if isinstance(reason, ssl.SSLCertVerificationError):
                    raise RuntimeError("HTTPS证书验证失败，请检查Python或系统CA证书配置。") from None
                if attempt == self.attempts:
                    raise RuntimeError("同花顺请求多次网络失败，请稍后重试。") from None
            else:
                try:
                    payload = json.loads(raw.decode("utf-8-sig"))
                except (UnicodeDecodeError, json.JSONDecodeError):
                    raise RuntimeError("同花顺响应不是有效的JSON，已有数据文件未修改。") from None
                if not isinstance(payload, dict) or "errorcode" not in payload:
                    raise RuntimeError("同花顺响应缺少有效的 errorcode。")
                if str(payload["errorcode"]) != "0":
                    detail = str(payload.get("errmsg", "请检查参数、令牌和数据权限。"))
                    for secret in self._secrets:
                        detail = detail.replace(secret, "[REDACTED]")
                    detail = " ".join(detail.split())[:200]
                    errorcode = str(payload["errorcode"])
                    if not re.fullmatch(r"-?\d+", errorcode):
                        errorcode = "unknown"
                    raise RuntimeError(f"同花顺业务错误 {errorcode}：{detail}")
                return payload
            print(
                f"⚠️ 同花顺第 {attempt}/{self.attempts} 次请求暂时失败，"
                f"{self.retry_delay * attempt:g} 秒后重试...",
                flush=True,
            )
            time.sleep(self.retry_delay * attempt)
        raise RuntimeError("同花顺请求未完成。")

    def login(self, refresh_token: str) -> None:
        if not refresh_token or "\r" in refresh_token or "\n" in refresh_token:
            raise ValueError("IFIND_REFRESH_TOKEN 不能为空或包含换行。")
        self._secrets.append(refresh_token)
        payload = self._post("get_access_token", {}, {"refresh_token": refresh_token})
        data = payload.get("data")
        access_token = data.get("access_token") if isinstance(data, dict) else None
        if not isinstance(access_token, str) or not access_token.strip() or "\r" in access_token or "\n" in access_token:
            raise RuntimeError("同花顺鉴权响应缺少有效的 access_token。")
        self.access_token = access_token.strip()
        self._secrets.append(self.access_token)

    def post(self, endpoint: str, body: Mapping[str, Any]) -> dict:
        if not self.access_token:
            raise RuntimeError("请先完成同花顺鉴权。")
        return self._post(endpoint, body, {"access_token": self.access_token})


def _table_field(table: Mapping[str, Any], field: str) -> list:
    for key, value in table.items():
        if str(key).lower() == field.lower():
            if isinstance(value, list):
                return value
            break
    raise RuntimeError(f"同花顺响应字段 {field} 缺失或不是列表。")


def fetch_benchmark_name(client: IFindClient, code: str) -> str:
    payload = client.post("basic_data_service", {
        "codes": code,
        "indipara": [{"indicator": NAME_INDICATOR}],
    })
    tables = payload.get("tables")
    if isinstance(tables, list):
        for item in tables:
            if not isinstance(item, dict) or str(item.get("thscode", "")).strip().upper() != code:
                continue
            table = item.get("table")
            if isinstance(table, dict):
                values = _table_field(table, NAME_INDICATOR)
                if values and isinstance(values[0], str):
                    name = " ".join(values[0].split())
                    if name and name not in {"--", "None", "null", "NULL"}:
                        return name
    raise RuntimeError(f"未取得 {code} 的指数名称，请检查代码及指数数据权限。")


def split_date_range(start: date, end: date, days: int = 365) -> Iterator[tuple[date, date]]:
    if days < 1 or start > end:
        raise ValueError("分段天数必须为正，且开始日期不能晚于结束日期。")
    current = start
    while current <= end:
        last = min(current + timedelta(days=days - 1), end)
        yield current, last
        current = last + timedelta(days=1)


def _market_date(value: Any) -> date:
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError("同花顺行情中包含无法识别的日期。")


def _price(value: Any, field: str, day: date) -> float:
    try:
        if isinstance(value, bool):
            raise ValueError
        number = float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        raise ValueError(f"{day} 的 {field} 缺失或不是有效数字。") from None
    if not math.isfinite(number) or number <= 0:
        raise ValueError(f"{day} 的 {field} 必须为有限正数。")
    return number


def merge_history_rows(existing: Mapping, new: Mapping) -> dict:
    """检查整批冲突后返回合并结果，不修改输入或覆盖同日不同价格。"""
    for day, row in new.items():
        if day in existing and existing[day] != row:
            raise ValueError(f"{day} 出现互相冲突的重复行情。")
    return {**existing, **new}


def parse_history_rows(
    payload: Mapping[str, Any], expected_code: str, start: date, end: date,
) -> dict[date, dict[str, float]]:
    tables = payload.get("tables")
    if not isinstance(tables, list):
        raise RuntimeError("同花顺行情响应缺少有效的 tables。")
    rows: dict[date, dict[str, float]] = {}
    for item in tables:
        if not isinstance(item, dict):
            raise RuntimeError("同花顺行情表结构不完整。")
        code = str(item.get("thscode", "")).strip().upper()
        if code != expected_code:
            raise RuntimeError("同花顺返回的基准代码与请求不一致。")
        times, table = item.get("time"), item.get("table")
        if not isinstance(times, list) or not isinstance(table, dict):
            raise RuntimeError("同花顺行情缺少日期列表或价格表。")
        columns = {field: _table_field(table, field) for field in PRICE_FIELDS}
        if any(len(values) != len(times) for values in columns.values()):
            raise RuntimeError("同花顺日期与四项价格的数量不一致。")
        for position, raw_day in enumerate(times):
            day = _market_date(raw_day)
            if not start <= day <= end:
                continue
            row = {field: _price(values[position], field, day) for field, values in columns.items()}
            if row["low"] > min(row["open"], row["close"]) or row["high"] < max(row["open"], row["close"]):
                raise ValueError(f"{day} 的高低价与开收盘价关系不合理。")
            rows = merge_history_rows(rows, {day: row})
    return rows


def download_history(client: IFindClient, code: str, start: date, end: date) -> dict:
    ranges = list(split_date_range(start, end))
    rows: dict[date, dict[str, float]] = {}
    empty_ranges = []
    for position, (first, last) in enumerate(ranges, start=1):
        print(
            f"正在下载 {code} 数据（本次指标：开盘价、最高价、最低价、收盘价）... "
            f"({position}/{len(ranges)})",
            flush=True,
        )
        print(f"本段日期：{first} 至 {last}", flush=True)
        payload = client.post("cmd_history_quotation", {
            "codes": code,
            "indicators": ",".join(PRICE_FIELDS),
            "startdate": first.isoformat(),
            "enddate": last.isoformat(),
            "functionpara": {"Interval": "D", "Fill": "Omit", "CPS": "1"},
        })
        chunk = parse_history_rows(payload, code, first, last)
        rows = merge_history_rows(rows, chunk)
        print(f"✅ 本段取得 {len(chunk)} 个交易日", flush=True)
        if not chunk:
            empty_ranges.append((first, last))
            print("⚠️ 本段未返回行情，请留意是否在指数有数据的期间内。", flush=True)
    if not rows:
        raise RuntimeError("指定区间未取得任何行情，已有数据文件未修改。")
    if any(first <= max(rows) and last >= min(rows) for first, last in empty_ranges):
        raise RuntimeError("已取得行情的中间存在整段空缺，停止保存，请检查数据完整性。")
    return rows


def write_csv_atomic(path: Path, code: str, name: str, rows: Mapping[date, Mapping[str, float]]) -> None:
    if not rows:
        raise ValueError("没有可保存的行情，已有数据文件未修改。")
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8-sig", newline="", dir=path.parent,
            prefix=f".{path.name}.", suffix=".tmp", delete=False,
        ) as stream:
            temp_path = Path(stream.name)
            writer = csv.writer(stream)
            writer.writerow(OUTPUT_COLUMNS)
            for day in sorted(rows):
                writer.writerow([day.isoformat(), code, name, *(
                    format(rows[day][field], ".15g") for field in PRICE_FIELDS
                )])
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_path, path)
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="从同花顺iFinD下载大盘基准日线。")
    parser.add_argument("--code", default=BENCHMARK_CODE, help="基准指数代码（默认沪深300）。")
    parser.add_argument("--start", default=START_DATE, help="开始日期 YYYY-MM-DD。")
    parser.add_argument("--end", default=END_DATE, help="结束日期 YYYY-MM-DD。")
    args = parser.parse_args(argv)
    started_at = time.perf_counter()
    try:
        code, start, end = validate_config(Config(args.code, args.start, args.end))
        token = load_refresh_token(PROJECT_ROOT / ".env")
        print("同花顺 iFinD 基准数据下载工具", flush=True)
        print("=" * 60, flush=True)
        print(f"基准代码：{code}", flush=True)
        print(f"下载日期：{start} 至 {end}", flush=True)
        print("数据来源：同花顺 iFinD", flush=True)
        print("日线原始点位，不复权；非交易日不补行情。", flush=True)
        client = IFindClient()
        print("正在登录同花顺...", flush=True)
        client.login(token)
        print("✅ 同花顺鉴权成功", flush=True)
        print("正在从同花顺获取基准名称...", flush=True)
        name = fetch_benchmark_name(client, code)
        print(f"✅ 已识别基准：{name}（{code}）", flush=True)
        print("\n开始下载基准指数日线数据...", flush=True)
        rows = download_history(client, code, start, end)
        output = OUTPUT_DIR / f"{code}.csv"
        write_csv_atomic(output, code, name, rows)
        print(f"✅ 成功保存 {name}（{code}）数据到 {output}", flush=True)
        print("\n" + "=" * 60, flush=True)
        print("✅ 基准数据下载完成！", flush=True)
        print(f"基准：{name}（{code}）", flush=True)
        print(f"文件保存在：{output}", flush=True)
        print(f"实际日期：{min(rows)} 至 {max(rows)}", flush=True)
        print(f"共 {len(rows)} 个交易日", flush=True)
        print(f"用时：{time.perf_counter() - started_at:.1f}秒", flush=True)
        return 0
    except (ValueError, RuntimeError, OSError) as exc:
        print(f"❌ 同花顺基准数据下载失败：{exc}", file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
