#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
同花顺 iFinD 全市场 ETF 一体化下载脚本。
1. 每次使用同花顺板块 ID 获取最新 ETF 成员，并与本地历史 ETF 合并；
2. 获取名称、上市日期、对标指数、benchmark、管理人、托管人等基本面；
3. 获取日频规模、成交量、成交额、收盘价、VWAP和溢价率；
4. 每取得一只 ETF 某个交易日的缺失指标，立即合并保存到
   etf_raw_data/代码.csv；
5. 重跑时逐列跳过 raw 文件中已有的“ETF + 日期 + 指标”，同时自动
   请求每只 ETF 本地最后日期之后至 END_DATE 的新区间；
6. 最后完全从 raw 文件重建总长表和“日期 × ETF代码”宽表。

重要口径：
- 历史行情 premiumRatio 虽在手册中被称为“贴水率”，实测与日期序列
  ths_premium_rate_fund 完全一致，正数表示溢价，负数表示折价；
- 基本面和类别是本次运行时取得的快照。如果下载一段历史日期，脚本会把这份
  快照写入新增日期，不应把它误认为严格的历史版本基本面；
- etf_raw_data 是唯一基础数据。etf_data.csv 和各字段宽表只是它的派生结果，
  可以随时从 raw 文件重建。

脚本只生成 CSV
"""

from __future__ import annotations

import csv
import json
import os
import re
import sqlite3
import ssl
import tempfile
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterator, Mapping, Sequence
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent
ENV_FILE = PROJECT_ROOT / ".env"


def load_project_env() -> None:
    """读取项目根目录的.env；已存在的系统环境变量优先。"""

    if not ENV_FILE.is_file():
        return
    for raw_line in ENV_FILE.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        key, separator, value = line.partition("=")
        key = key.strip()
        if not separator or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key):
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        os.environ.setdefault(key, value)


load_project_env()


# ============================================================================
# 一、用户参数：以后通常只需要修改本区域
# ============================================================================

# 支持 YYYY-MM-DD、YYYYMMDD、YYYY/MM/DD；结束日期也可以写 TODAY。
# 开始日期与结束日期相同，就只下载一天；不同则下载完整日期区间。
START_DATE = "2021-01-01"
END_DATE = "2026-08-25"

# 前复权基点日期（同花顺历史行情参数 BaseDate）。
# 它和下载开始日期不是一回事：START_DATE 决定下载哪些日期，
# 这个参数决定开盘价、收盘价等价格字段以哪一天为前复权基点。
# 收盘价和VWAP会使用 CPS=2 和这个 BaseDate 进行前复权。修改本参数后，
# 已保存的受复权影响数据会自动重新下载，不会被断点逻辑误跳过。
FORWARD_ADJUSTMENT_BASE_DATE = "2026-07-31"

# 凭据只从环境变量读取，禁止写入代码或提交到版本库。
# iFinD HTTP API 不接受用户名和密码直接登录，实际 HTTP 鉴权使用 refresh token。
IFIND_USERNAME = os.environ.get("IFIND_USERNAME", "")
IFIND_PASSWORD = os.environ.get("IFIND_PASSWORD", "")

# ---------------------------- 账号 refresh token ----------------------------
IFIND_REFRESH_TOKEN = os.environ.get("IFIND_REFRESH_TOKEN", "")

# ------------------------------- 板块 ID -------------------------------------
# 除股票型ETF外的大类板块。股票型ETF由下方五个细类直接构成，
# 所以不请求股票型ETF大类板块。
MAJOR_CATEGORY_BLOCKS: dict[str, str] = {
    "债券型ETF": "051001006002",
    "商品型ETF": "051001006003",
    "货币型ETF": "051001006004",
    "跨境型ETF": "051001006006001",
    "其他ETF": "051025001003",
}

# 股票型ETF细类。顺序就是最终“类别细类”的优先显示顺序。
MINOR_CATEGORY_BLOCKS: dict[str, str] = {
    "规模指数ETF": "051001006006001001",
    "行业指数ETF": "051001006006001002",
    "主题指数ETF": "051001006006001005",
    "策略指数ETF": "051001006006001003",
    "风格指数ETF": "051001006006001004",
}

# ----------------------- 使用“基本面数据token”的指标 -------------------------
# 这里的“token”仅表示接口路由组；实际鉴权仍使用上方唯一的账号 refresh token。
# 每项格式为：最终列名: (接口名称, 同花顺指标代码)。修改指标代码会改变输出内容。
BASIC_DATA_TOKEN_INDICATORS: dict[str, tuple[str, str]] = {
    "上市日期": ("basic_data_service", "ths_lof_listed_date_fund"),
    "对标指数": ("basic_data_service", "ths_name_of_tracking_index_fund"),
    "benchmark": ("basic_data_service", "ths_perf_comparative_benchmark_fund"),
    "管理人": ("basic_data_service", "ths_fund_supervisor_fullname_fund"),
    "托管人": ("basic_data_service", "ths_fund_mandator_fullname_fund"),
    "对标指数代码": ("basic_data_service", "ths_tracking_index_thscode_fund"),
    "规模": ("date_sequence", "ths_fund_scale_fund"),
}

# ----------------------- 使用“历史行情token”的指标 ---------------------------
# 这里同样只是接口路由组。成交量、成交额、价格和溢价率从历史行情接口读取。
HISTORY_QUOTE_TOKEN_INDICATORS: dict[str, tuple[str, str]] = {
    "成交量": ("cmd_history_quotation", "volume"),
    "成交额": ("cmd_history_quotation", "amount"),
    "收盘价": ("cmd_history_quotation", "close"),
    "VWAP": ("cmd_history_quotation", "avgPrice"),
    "溢价率": ("cmd_history_quotation", "premiumRatio"),
}

# ----------------------------- 请求性能参数 ----------------------------------
# 底层每次最多同时请求多少只ETF。日志和raw仍然按单只ETF显示、保存。
# 数值越大，首次全量下载越快，但单次返回数据也越大。
ETF_CODES_PER_REQUEST = 100

# 历史区间过长时，底层自动分成不超过该天数的请求，避免响应过大。
DATE_RANGE_DAYS_PER_REQUEST = 730

# 失败后最多尝试3次，每次等待3秒。
HTTP_TIMEOUT_SECONDS = 180
MAX_RETRIES = 3
RETRY_WAIT_SECONDS = 3

# ============================================================================
# 用户参数区域结束；下面是程序实现，通常不需要修改
# ============================================================================


API_BASE = "https://quantapi.51ifind.com/api/v1"
ALLOWED_HOSTS = {"quantapi.51ifind.com", "quantapi.10jqka.com.cn"}

OUTPUT_DIR = PROJECT_ROOT / "outputs" / "etf_data"
RAW_OUTPUT_DIR = OUTPUT_DIR / "etf_raw_data"
TOTAL_OUTPUT_FILE = OUTPUT_DIR / "etf_data.csv"
# 只有配置了受复权影响的历史行情字段时才会生成这个隐藏状态文件。
# 它用于识别 BaseDate 是否变更，不是数据输出。
ADJUSTMENT_STATE_FILE = OUTPUT_DIR / ".forward_adjustment_state.csv"

# 板块成分使用同花顺数据池 p03291。
BOARD_REPORT = "p03291"
BOARD_CODE_FIELD = "p03291_f002"
BOARD_NAME_FIELD = "p03291_f003"

# 最终总表和每只ETF长表的列顺序：日期最前、基本面居中、日频指标最后。
OUTPUT_COLUMNS = [
    "日期",
    "代码",
    "名称",
    "上市日期",
    "对标指数",
    "benchmark",
    "管理人",
    "托管人",
    "类别大类",
    "类别细类",
    "对标指数代码",
    "规模",
    "成交量",
    "成交额",
    "收盘价",
    "VWAP",
    "溢价率",
]

# 每一个字段对应一张“日期 × ETF代码”宽表，第一列固定为“日期”。
# 日期和代码本身是宽表的行索引与列名，因此不再单独输出 date.csv/code.csv。
WIDE_FILE_NAMES: dict[str, str] = {
    "名称": "name.csv",
    "上市日期": "listed_date.csv",
    "对标指数": "tracking_index.csv",
    "benchmark": "benchmark.csv",
    "管理人": "manager.csv",
    "托管人": "custodian.csv",
    "类别大类": "category_major.csv",
    "类别细类": "category_minor.csv",
    "对标指数代码": "tracking_index_code.csv",
    "规模": "fund_scale.csv",
    "成交量": "volume.csv",
    "成交额": "amount.csv",
    "收盘价": "close.csv",
    "VWAP": "vwap.csv",
    "溢价率": "premium_ratio.csv",
}

# SQLite 临时表使用英文列名，避免SQL中频繁处理中文标识符。
DAILY_DB_COLUMNS: dict[str, str] = {
    "规模": "fund_scale",
    "成交量": "volume",
    "成交额": "amount",
    "收盘价": "close",
    "VWAP": "vwap",
    "溢价率": "premium_ratio",
}

# 同花顺历史行情中会受 CPS/BaseDate 影响的价格类指标。
# close 和 avgPrice 会触发前复权；复权基点改变时只重下价格字段。
ADJUSTMENT_SENSITIVE_HISTORY_INDICATORS = {
    "preclose",
    "open",
    "high",
    "low",
    "close",
    "avgprice",
    "change",
    "changeratio",
}

# raw CSV 在临时 SQLite 中的列名映射。SQLite 只用来加速断点索引和派生表生成，
# 真正持久保存的基础数据仍然是 etf_raw_data 中的 CSV。
RAW_DB_COLUMNS: dict[str, str] = {
    "日期": "date",
    "代码": "code",
    "名称": "name",
    "上市日期": "listed_date",
    "对标指数": "tracking_index",
    "benchmark": "benchmark",
    "管理人": "manager",
    "托管人": "custodian",
    "类别大类": "category_major",
    "类别细类": "category_minor",
    "对标指数代码": "tracking_index_code",
    "规模": "fund_scale",
    "成交量": "volume",
    "成交额": "amount",
    "收盘价": "close",
    "VWAP": "vwap",
    "溢价率": "premium_ratio",
}

CODE_PATTERN = re.compile(r"^[A-Za-z0-9_.-]+$")


@dataclass(frozen=True)
class RuntimeConfig:
    """统一保存解析后的日期，避免程序不同位置重复解释 TODAY。"""

    start_date: date
    end_date: date
    universe_date: date
    forward_adjustment_base_date: date


@dataclass(frozen=True)
class DownloadTask:
    """一项日频下载任务：接口类别、ETF代码、日期区间和指标映射。"""

    route_type: str
    endpoint: str
    codes: tuple[str, ...]
    start_date: date
    end_date: date
    fields: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class FieldRequest:
    """一个日频指标及其真正缺失的日期段。"""

    route_type: str
    endpoint: str
    output_name: str
    indicator: str
    ranges: tuple[tuple[date, date], ...]


@dataclass(frozen=True)
class EtfDownloadPlan:
    """一只ETF本次只需要补充的固定字段和日频字段。"""

    basic_fields: tuple[tuple[str, str], ...]
    daily_requests: tuple[FieldRequest, ...]

    def indicator_names(self) -> tuple[str, ...]:
        names = [name for name, _ in self.basic_fields]
        names.extend(request.output_name for request in self.daily_requests)
        return tuple(dict.fromkeys(names))

    def is_empty(self) -> bool:
        return not self.basic_fields and not self.daily_requests


def parse_date(value: str, label: str, *, allow_today: bool = True) -> date:
    """把三种常见日期格式转换为 date；END_DATE 额外支持 TODAY。"""

    text = str(value).strip()
    if allow_today and text.upper() == "TODAY":
        return date.today()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    today_hint = " 或 TODAY" if allow_today else ""
    raise ValueError(f"{label} 必须是 YYYY-MM-DD、YYYYMMDD、YYYY/MM/DD{today_hint}。")


def validate_config() -> RuntimeConfig:
    """运行前一次性检查参数，避免下载到一半才发现配置问题。"""

    start = parse_date(START_DATE, "START_DATE")
    end = parse_date(END_DATE, "END_DATE")
    adjustment_base_date = parse_date(
        FORWARD_ADJUSTMENT_BASE_DATE,
        "FORWARD_ADJUSTMENT_BASE_DATE",
        allow_today=False,
    )
    if start > end:
        raise ValueError("START_DATE 不能晚于 END_DATE。")

    if not IFIND_REFRESH_TOKEN.strip():
        raise ValueError(
            f"未读取到 IFIND_REFRESH_TOKEN，请在 {ENV_FILE} 中填写该项。"
        )
    if ETF_CODES_PER_REQUEST <= 0 or DATE_RANGE_DAYS_PER_REQUEST <= 0:
        raise ValueError("每次请求的ETF数量和日期天数必须大于0。")
    if HTTP_TIMEOUT_SECONDS <= 0 or MAX_RETRIES <= 0 or RETRY_WAIT_SECONDS < 0:
        raise ValueError("超时、重试次数和重试等待参数不合法。")

    if set(OUTPUT_COLUMNS) != {
        "日期", "代码", "名称", "上市日期", "对标指数", "benchmark",
        "管理人", "托管人", "类别大类", "类别细类", "对标指数代码",
        "规模", "成交量", "成交额", "收盘价", "VWAP", "溢价率",
    }:
        raise ValueError("OUTPUT_COLUMNS 被误改，缺少或增加了预期之外的列。")

    # 板块日期使用结束日期：下载历史区间时，以区间末日的ETF分类为准。
    return RuntimeConfig(
        start_date=start,
        end_date=end,
        universe_date=end,
        forward_adjustment_base_date=adjustment_base_date,
    )


def make_ssl_context() -> ssl.SSLContext:
    """兼容部分 macOS Python 没有自动找到系统CA证书的情况。"""

    for candidate in (Path("/etc/ssl/cert.pem"), Path("/private/etc/ssl/cert.pem")):
        if candidate.exists():
            return ssl.create_default_context(cafile=str(candidate))
    return ssl.create_default_context()


SSL_CONTEXT = make_ssl_context()


def normalize_request_url(endpoint_or_url: str) -> str:
    """补全接口地址，并强制请求只能发往同花顺官方量化接口域名。"""

    text = str(endpoint_or_url).strip().strip("'\"")
    if not text.startswith(("http://", "https://")):
        text = f"{API_BASE}/{text.lstrip('/')}"
    parsed = urlparse(text)
    if parsed.hostname not in ALLOWED_HOSTS:
        raise ValueError(f"只允许访问同花顺官方量化接口，当前地址为：{text}")
    if "/api/v1/" not in parsed.path:
        raise ValueError(f"不是有效的 iFinD HTTP API 地址：{text}")
    return text


def post_json(
    endpoint_or_url: str,
    *,
    body: Mapping[str, Any] | None = None,
    access_token: str | None = None,
    refresh_token: str | None = None,
    progress_label: str = "",
) -> dict[str, Any]:
    """发送一次 iFinD POST 请求，处理网络重试并检查业务错误码。"""

    headers = {
        "Content-Type": "application/json",
        "ifindlang": "cn",
        "User-Agent": "GuangfaInternship-StandaloneETFDownloader/1.0",
    }
    if access_token:
        headers["access_token"] = access_token
    if refresh_token:
        headers["refresh_token"] = refresh_token

    encoded = b"" if body is None else json.dumps(dict(body)).encode("utf-8")
    last_error: Exception | None = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            request = Request(
                normalize_request_url(endpoint_or_url),
                data=encoded,
                headers=headers,
                method="POST",
            )
            with urlopen(
                request,
                timeout=HTTP_TIMEOUT_SECONDS,
                context=SSL_CONTEXT,
            ) as response:
                payload = json.loads(response.read().decode("utf-8"))

            if not isinstance(payload, dict):
                raise RuntimeError("iFinD 返回的 JSON 根节点不是对象。")
            error_code = payload.get("errorcode")
            if error_code is not None and int(error_code) != 0:
                raise RuntimeError(
                    f"iFinD error {error_code}: {payload.get('errmsg', '未知错误')}"
                )
            return payload
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
            last_error = RuntimeError(f"HTTP {exc.code}: {detail}")
        except (URLError, TimeoutError, json.JSONDecodeError, RuntimeError) as exc:
            last_error = exc

        if attempt < MAX_RETRIES:
            label = f"{progress_label}，" if progress_label else ""
            print(
                f"⚠️ {label}第 {attempt}/{MAX_RETRIES} 次下载失败："
                f"{last_error}；{RETRY_WAIT_SECONDS} 秒后重新下载",
                flush=True,
            )
            time.sleep(RETRY_WAIT_SECONDS)

    assert last_error is not None
    raise last_error


def get_access_token(refresh_token: str, account_label: str) -> str:
    """使用账号 refresh token 换取 access token；不会打印任何 token。"""

    payload = post_json(
        "get_access_token",
        refresh_token=refresh_token.strip(),
        progress_label=f"{account_label}鉴权",
    )
    token = payload.get("data", {}).get("access_token")
    if not token:
        raise RuntimeError(f"{account_label}鉴权成功但未返回 access_token。")
    return str(token)


class IFindClient:
    """用共享 access token 建立接口组客户端，避免指标路由混用。"""

    def __init__(self, route_label: str, access_token: str) -> None:
        self.route_label = route_label
        self.access_token = access_token

    def post(
        self,
        endpoint: str,
        body: Mapping[str, Any],
        *,
        progress_label: str = "",
    ) -> dict[str, Any]:
        return post_json(
            endpoint,
            access_token=self.access_token,
            body=body,
            progress_label=progress_label,
        )


def clean_value(value: Any) -> str:
    """统一清理同花顺常见空值，同时保留数字0。"""

    if value is None:
        return ""
    text = str(value).strip()
    return "" if text in {"", "--", "None", "null", "NULL"} else text


def normalize_code(value: Any) -> str:
    """标准化证券代码，并阻止异常代码成为输出文件名。"""

    code = clean_value(value).upper()
    if code and not CODE_PATTERN.fullmatch(code):
        raise ValueError(f"证券代码包含不安全字符：{code!r}")
    return code


def payload_column(payload: Mapping[str, Any], field: str) -> list[Any]:
    """从同花顺列式 tables 结构中收集指定字段。"""

    result: list[Any] = []
    tables = payload.get("tables")
    if not isinstance(tables, list):
        return result
    for item in tables:
        if not isinstance(item, dict):
            continue
        table = item.get("table")
        if not isinstance(table, dict):
            continue
        values = table.get(field)
        if isinstance(values, list):
            result.extend(values)
    return result


def fetch_board(
    client: IFindClient,
    board_id: str,
    board_date: date,
    *,
    include_name: bool,
) -> dict[str, str]:
    """通过板块ID取得成员，可按调用方需要同时取得ETF名称。"""

    output_fields = [BOARD_CODE_FIELD]
    if include_name:
        output_fields.append(BOARD_NAME_FIELD)
    payload = client.post(
        "data_pool",
        {
            "reportname": BOARD_REPORT,
            "functionpara": {
                "date": board_date.strftime("%Y%m%d"),
                "blockname": board_id,
                "iv_type": "allcontract",
            },
            "outputpara": ",".join(output_fields),
        },
        progress_label=f"板块{board_id}",
    )

    codes = payload_column(payload, BOARD_CODE_FIELD)
    names = payload_column(payload, BOARD_NAME_FIELD) if include_name else []
    result: dict[str, str] = {}
    for index, raw_code in enumerate(codes):
        code = normalize_code(raw_code)
        if not code:
            continue
        name = clean_value(names[index]) if index < len(names) else ""
        result[code] = name
    if not result:
        raise RuntimeError(f"板块 {board_id} 没有返回任何ETF代码。")
    return dict(sorted(result.items()))


def fetch_category_boards(
    client: IFindClient,
    board_date: date,
) -> tuple[dict[str, str], dict[str, set[str]], dict[str, set[str]]]:
    """读取五个股票细类和其他五个ETF大类。

    股票型ETF直接由五个股票细类组成，不额外请求股票型大类板块。
    """

    results: dict[str, dict[str, str]] = {}

    # 先读取并显示五个股票细类，它们直接组成股票型ETF。
    for label, board_id in MINOR_CATEGORY_BLOCKS.items():
        results[label] = fetch_board(
            client,
            board_id,
            board_date,
            include_name=True,
        )
        print(f"  {label}：{len(results[label])} 只", flush=True)

    stock_members: set[str] = set()
    for label in MINOR_CATEGORY_BLOCKS:
        stock_members.update(results[label])
    stock_total = sum(len(results[label]) for label in MINOR_CATEGORY_BLOCKS)
    print(f"  股票型ETF合计：{stock_total} 只", flush=True)

    # 再读取并显示其他五个大类。板块请求全部按顺序执行。
    for label, board_id in MAJOR_CATEGORY_BLOCKS.items():
        results[label] = fetch_board(
            client,
            board_id,
            board_date,
            include_name=True,
        )
        print(f"  {label}：{len(results[label])} 只", flush=True)

    etfs: dict[str, str] = {}
    for category_label in (*MINOR_CATEGORY_BLOCKS, *MAJOR_CATEGORY_BLOCKS):
        etfs.update(results[category_label])

    major: dict[str, set[str]] = {"股票型ETF": stock_members}
    for label in MAJOR_CATEGORY_BLOCKS:
        major[label] = set(results[label])

    minor = {
        label: set(results[label])
        for label in MINOR_CATEGORY_BLOCKS
    }
    return dict(sorted(etfs.items())), major, minor


def category_labels(
    code: str,
    major_members: Mapping[str, set[str]],
    minor_members: Mapping[str, set[str]],
) -> tuple[str, str]:
    """按板块实际成员关系生成大类和细类标签。"""

    major_labels = [label for label, members in major_members.items() if code in members]
    minor_labels = [label for label, members in minor_members.items() if code in members]

    # 直接保留板块返回的标签；多个标签用竖线连接，不做额外分类校验。
    return "|".join(major_labels), "|".join(minor_labels)


def build_master_data(
    client: IFindClient,
    config: RuntimeConfig,
) -> tuple[dict[str, dict[str, str]], list[str]]:
    """读取ETF大类和细类，生成基本面主表骨架。"""

    print(f"正在从同花顺 iFinD 获取 {config.universe_date} 的ETF列表...", flush=True)
    etfs, major_members, minor_members = fetch_category_boards(
        client,
        config.universe_date,
    )

    master: dict[str, dict[str, str]] = {}
    for code, name in etfs.items():
        major, minor = category_labels(
            code,
            major_members,
            minor_members,
        )
        master[code] = {
            "代码": code,
            "名称": name,
            "上市日期": "",
            "对标指数": "",
            "benchmark": "",
            "管理人": "",
            "托管人": "",
            "类别大类": major,
            "类别细类": minor,
            "对标指数代码": "",
        }

    codes = sorted(master)
    print(f"✅ ETF总数：{len(codes)} 只", flush=True)
    if codes:
        examples = [f"'{code}'" for code in codes[:5]]
        suffix = ", ..." if len(codes) > 5 else ""
        print(f"  示例ETF: [{', '.join(examples)}{suffix}]", flush=True)
    return master, codes


def first_table_value(table: Mapping[str, Any], indicator: str) -> str:
    """基础函数每个指标只取当前快照的第一个值。"""

    values = table.get(indicator)
    if not isinstance(values, list) or not values:
        return ""
    return clean_value(values[0])


def fields_for_endpoint(
    indicators: Mapping[str, tuple[str, str]],
    endpoint: str,
) -> tuple[tuple[str, str], ...]:
    """从某个token路由组中取出指定接口的“输出列名、指标代码”。"""

    return tuple(
        (output_name, indicator)
        for output_name, (configured_endpoint, indicator) in indicators.items()
        if configured_endpoint == endpoint
    )


def fetch_basic_fields(
    client: IFindClient,
    codes: Sequence[str],
    fields: Sequence[tuple[str, str]] | None = None,
) -> dict[str, dict[str, str]]:
    """一次请求多只ETF真正缺失的固定字段，返回后仍按代码分开。"""

    if not codes:
        return {}

    requested_fields = tuple(fields) if fields is not None else fields_for_endpoint(
        BASIC_DATA_TOKEN_INDICATORS,
        "basic_data_service",
    )
    if not requested_fields:
        return {}
    payload = client.post(
        "basic_data_service",
        {
            "codes": ",".join(codes),
            "indipara": [
                {"indicator": indicator, "indiparams": []}
                for _, indicator in requested_fields
            ],
        },
        progress_label=(
            codes[0]
            if len(codes) == 1
            else f"{codes[0]}等{len(codes)}只ETF"
        ),
    )

    result: dict[str, dict[str, str]] = {}
    expected_codes = set(codes)
    tables = payload.get("tables")
    if not isinstance(tables, list):
        return result
    for item in tables:
        if not isinstance(item, dict):
            continue
        returned_code = normalize_code(item.get("thscode"))
        table = item.get("table")
        if returned_code not in expected_codes or not isinstance(table, dict):
            continue
        result[returned_code] = {
            output_name: first_table_value(table, indicator)
            for output_name, indicator in requested_fields
        }
    return result


def split_date_range(range_start: date, range_end: date) -> Iterator[tuple[date, date]]:
    """长区间分段请求，不改变最终raw的ETF-日期断点逻辑。"""

    current = range_start
    step = timedelta(days=DATE_RANGE_DAYS_PER_REQUEST - 1)
    while current <= range_end:
        current_end = min(range_end, current + step)
        yield current, current_end
        current = current_end + timedelta(days=1)


def consecutive_missing_ranges(
    known_dates: Sequence[date],
    completed_dates: set[date],
) -> Iterator[tuple[date, date]]:
    """把某只ETF缺少的已知交易日合并成尽可能少的日期段。

    这里的“连续”指在已知交易日序列中连续，因此周末和节假日不会
    把两个相邻交易日拆成两次请求。
    """

    missing_positions = [
        index
        for index, current_date in enumerate(known_dates)
        if current_date not in completed_dates
    ]
    if not missing_positions:
        return

    range_start = missing_positions[0]
    previous = range_start
    for position in missing_positions[1:]:
        if position != previous + 1:
            yield known_dates[range_start], known_dates[previous]
            range_start = position
        previous = position
    yield known_dates[range_start], known_dates[previous]


def task_uses_forward_adjustment(task: DownloadTask) -> bool:
    """判断该历史行情请求是否包含受复权影响的价格指标。"""

    return any(
        indicator.lower() in ADJUSTMENT_SENSITIVE_HISTORY_INDICATORS
        for _, indicator in task.fields
    )


def fetch_daily_task(
    task: DownloadTask,
    basic_client: IFindClient,
    history_client: IFindClient,
    forward_adjustment_base_date: date,
) -> tuple[DownloadTask, dict[str, Any]]:
    """为单只ETF、单个日期段请求一类日频数据。"""

    start = task.start_date.strftime("%Y-%m-%d")
    end = task.end_date.strftime("%Y-%m-%d")
    request_label = (
        task.codes[0]
        if len(task.codes) == 1
        else f"{task.codes[0]}等{len(task.codes)}只ETF"
    )

    if task.route_type == "基本面数据token":
        # 日期序列放在基本面数据token路由组；规模由这个接口获取。
        body = {
            "codes": ",".join(task.codes),
            "startdate": start,
            "enddate": end,
            "functionpara": {
                "Days": "Tradedays",
                "Fill": "Blank",
                "Interval": "D",
            },
            "indipara": [
                {"indicator": indicator, "indiparams": []}
                for _, indicator in task.fields
            ],
        }
        payload = basic_client.post(task.endpoint, body, progress_label=request_label)
    elif task.route_type == "历史行情token":
        # 成交量、成交额、收盘价、VWAP和溢价率放在历史行情token路由组。
        body = {
            "codes": ",".join(task.codes),
            "indicators": ",".join(indicator for _, indicator in task.fields),
            "startdate": start,
            "enddate": end,
            "functionpara": (
                {
                    "CPS": "2",
                    "baseDate": forward_adjustment_base_date.isoformat(),
                }
                if task_uses_forward_adjustment(task)
                else {}
            ),
        }
        payload = history_client.post(task.endpoint, body, progress_label=request_label)
    else:
        raise ValueError(f"未知token路由类型：{task.route_type}")

    return task, payload


def create_database(path: Path) -> sqlite3.Connection:
    """创建临时数据库，用于逐指标查缺、合并接口结果和生成派生表。"""

    connection = sqlite3.connect(str(path))
    connection.execute(
        "CREATE TABLE daily ("
        "date TEXT NOT NULL, "
        "code TEXT NOT NULL, "
        "fund_scale TEXT, "
        "volume TEXT, "
        "amount TEXT, "
        "close TEXT, "
        "vwap TEXT, "
        "premium_ratio TEXT, "
        "raw_saved INTEGER NOT NULL DEFAULT 0, "
        "PRIMARY KEY (date, code))"
    )
    connection.execute("CREATE INDEX daily_code_date ON daily(code, date)")
    connection.execute("CREATE INDEX daily_date_code ON daily(date, code)")

    connection.execute(
        "CREATE TABLE raw_data ("
        "date TEXT NOT NULL, "
        "code TEXT NOT NULL, "
        "name TEXT, "
        "listed_date TEXT, "
        "tracking_index TEXT, "
        "benchmark TEXT, "
        "manager TEXT, "
        "custodian TEXT, "
        "category_major TEXT, "
        "category_minor TEXT, "
        "tracking_index_code TEXT, "
        "fund_scale TEXT, "
        "volume TEXT, "
        "amount TEXT, "
        "close TEXT, "
        "vwap TEXT, "
        "premium_ratio TEXT, "
        "PRIMARY KEY (date, code))"
    )
    connection.execute("CREATE INDEX raw_code_date ON raw_data(code, date)")
    connection.execute("CREATE INDEX raw_date_code ON raw_data(date, code)")
    return connection


def load_existing_raw_data(
    connection: sqlite3.Connection,
    allowed_codes: set[str],
) -> int:
    """读取所有单ETF raw CSV，建立断点索引和派生表数据源。

    除日期和代码外，任何缺列或空值都允许载入；后续按“ETF + 日期 + 指标”
    单独计算缺口，只请求真正缺失的指标。
    """

    raw_files = sorted(
        path
        for path in RAW_OUTPUT_DIR.glob("*.csv")
        if path.stem.upper() in allowed_codes
    )
    if not raw_files:
        return 0

    raw_columns = [RAW_DB_COLUMNS[column] for column in OUTPUT_COLUMNS]
    raw_placeholders = ",".join("?" for _ in raw_columns)
    raw_sql = (
        f"INSERT INTO raw_data ({','.join(raw_columns)}) VALUES ({raw_placeholders}) "
        "ON CONFLICT(date,code) DO UPDATE SET "
        + ",".join(
            f"{column}=excluded.{column}"
            for column in raw_columns
            if column not in {"date", "code"}
        )
    )
    daily_sql = (
        "INSERT INTO daily "
        "(date,code,fund_scale,volume,amount,close,vwap,premium_ratio,raw_saved) "
        "VALUES (?,?,?,?,?,?,?,?,?) "
        "ON CONFLICT(date,code) DO UPDATE SET "
        "fund_scale=excluded.fund_scale,volume=excluded.volume,"
        "amount=excluded.amount,close=excluded.close,vwap=excluded.vwap,"
        "premium_ratio=excluded.premium_ratio,"
        "raw_saved=excluded.raw_saved"
    )

    row_count = 0
    for path in raw_files:
        expected_code = normalize_code(path.stem)
        if not expected_code or not CODE_PATTERN.fullmatch(expected_code):
            raise RuntimeError(f"raw文件名不是合法ETF代码：{path.name}")

        seen_dates: set[str] = set()
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            fieldnames = reader.fieldnames or []
            missing_columns = [column for column in ("日期", "代码") if column not in fieldnames]
            if missing_columns:
                raise RuntimeError(
                    f"raw文件 {path.name} 缺少列：{','.join(missing_columns)}"
                )

            raw_buffer: list[tuple[str, ...]] = []
            daily_buffer: list[tuple[str, ...]] = []
            for row in reader:
                current_date = clean_value(row.get("日期"))[:10]
                code = normalize_code(row.get("代码"))
                try:
                    date.fromisoformat(current_date)
                except ValueError as exc:
                    raise RuntimeError(
                        f"raw文件 {path.name} 包含非法日期：{current_date}"
                    ) from exc
                if code != expected_code:
                    raise RuntimeError(
                        f"raw文件 {path.name} 中出现了其他代码：{code or '空'}"
                    )
                if current_date in seen_dates:
                    raise RuntimeError(
                        f"raw文件 {path.name} 中日期重复：{current_date}"
                    )
                seen_dates.add(current_date)

                values = [clean_value(row.get(column)) for column in OUTPUT_COLUMNS]
                values[0] = current_date
                values[1] = code
                raw_buffer.append(tuple(values))
                scale = clean_value(row.get("规模"))
                volume = clean_value(row.get("成交量"))
                amount = clean_value(row.get("成交额"))
                close = clean_value(row.get("收盘价"))
                vwap = clean_value(row.get("VWAP"))
                premium_ratio = clean_value(row.get("溢价率"))
                daily_buffer.append(
                    (
                        current_date,
                        code,
                        scale,
                        volume,
                        amount,
                        close,
                        vwap,
                        premium_ratio,
                        1,
                    )
                )

                if len(raw_buffer) >= 10_000:
                    connection.executemany(raw_sql, raw_buffer)
                    connection.executemany(daily_sql, daily_buffer)
                    row_count += len(raw_buffer)
                    raw_buffer.clear()
                    daily_buffer.clear()

            if raw_buffer:
                connection.executemany(raw_sql, raw_buffer)
                connection.executemany(daily_sql, daily_buffer)
                row_count += len(raw_buffer)
            connection.commit()
    return row_count


def discover_local_universe_codes(config: RuntimeConfig) -> list[str]:
    """从现有raw识别全部已上市ETF大类，不消耗接口额度。"""

    eligible_labels = (
        "股票型ETF",
        "债券型ETF",
        "商品型ETF",
        "货币型ETF",
        "跨境型ETF",
        "其他ETF",
    )
    codes: list[str] = []
    for path in sorted(RAW_OUTPUT_DIR.glob("*.csv")):
        category_major = ""
        listed_date = ""
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                for row in csv.DictReader(handle):
                    current_category = clean_value(row.get("类别大类"))
                    current_listed_date = clean_value(row.get("上市日期"))[:10]
                    if current_category:
                        category_major = current_category
                    if current_listed_date:
                        listed_date = current_listed_date
        except (OSError, csv.Error):
            continue

        if not any(label in category_major for label in eligible_labels):
            continue
        if listed_date:
            try:
                parsed_listed_date = parse_date(
                    listed_date,
                    "上市日期",
                    allow_today=False,
                )
            except ValueError:
                parsed_listed_date = None
            if parsed_listed_date and parsed_listed_date > config.end_date:
                continue
        code = normalize_code(path.stem)
        if code:
            codes.append(code)
    return sorted(set(codes))


def discover_all_local_raw_codes() -> list[str]:
    """读取全部单ETF raw 文件代码，仅用于最终汇总，不参与接口下载。"""

    codes: set[str] = set()
    for path in RAW_OUTPUT_DIR.glob("*.csv"):
        code = normalize_code(path.stem)
        if code and CODE_PATTERN.fullmatch(code):
            codes.add(code)
    return sorted(codes)


def build_master_from_raw_data(
    connection: sqlite3.Connection,
    codes: Sequence[str],
) -> dict[str, dict[str, str]]:
    """从raw中每个固定字段最新的非空值恢复主表，不调用接口。"""

    fields = [
        "名称",
        "上市日期",
        "对标指数",
        "benchmark",
        "管理人",
        "托管人",
        "类别大类",
        "类别细类",
        "对标指数代码",
    ]
    master: dict[str, dict[str, str]] = {}
    for code in sorted(codes):
        row = {"代码": code}
        for field in fields:
            db_field = RAW_DB_COLUMNS[field]
            value_row = connection.execute(
                f"SELECT {db_field} FROM raw_data "
                f"WHERE code=? AND TRIM(COALESCE({db_field},''))<>'' "
                "ORDER BY date DESC LIMIT 1",
                (code,),
            ).fetchone()
            row[field] = clean_value(value_row[0]) if value_row else ""
        master[code] = row
    return master


def merge_master_snapshots(
    local_master: Mapping[str, Mapping[str, str]],
    current_master: Mapping[str, Mapping[str, str]],
    codes: Sequence[str],
) -> dict[str, dict[str, str]]:
    """合并本地历史快照和本次同花顺快照。

    本地非空字段用于保留已退市 ETF 的历史信息；同花顺本次返回的非空字段
    优先，用于让新增日期采用最新名称和分类。两边都没有的字段保留为空，
    后续仍由基本面接口按原有断点逻辑补充。
    """

    merged: dict[str, dict[str, str]] = {}
    master_fields = [column for column in OUTPUT_COLUMNS if column not in DAILY_DB_COLUMNS]
    master_fields.remove("日期")
    for code in sorted(set(codes)):
        row = {field: "" for field in master_fields}
        row["代码"] = code
        for snapshot in (local_master.get(code, {}), current_master.get(code, {})):
            for field in master_fields:
                value = clean_value(snapshot.get(field))
                if value:
                    row[field] = value
        merged[code] = row
    return merged


def raw_dates_in_range(
    connection: sqlite3.Connection,
    config: RuntimeConfig,
) -> list[date]:
    """使用raw中实际存在的日期作为历史补缺日期，不制造非交易日缺口。"""

    rows = connection.execute(
        "SELECT DISTINCT date FROM raw_data WHERE date BETWEEN ? AND ? ORDER BY date",
        (config.start_date.isoformat(), config.end_date.isoformat()),
    ).fetchall()
    return [date.fromisoformat(str(row[0])) for row in rows]


def listed_date_or_none(master_row: Mapping[str, str]) -> date | None:
    """读取ETF上市日期；raw为空或格式异常时不限制请求范围。"""

    text = clean_value(master_row.get("上市日期"))[:10]
    try:
        return parse_date(text, "上市日期", allow_today=False)
    except ValueError:
        return None


def build_field_download_plans(
    connection: sqlite3.Connection,
    master: Mapping[str, Mapping[str, str]],
    codes: Sequence[str],
    target_dates: Sequence[date],
    config: RuntimeConfig,
    *,
    force_fields: set[str] | None = None,
) -> dict[str, EtfDownloadPlan]:
    """逐ETF、逐日期、逐指标检查raw，只生成真正缺失的请求。"""

    force_fields = force_fields or set()
    fixed_fields = fields_for_endpoint(
        BASIC_DATA_TOKEN_INDICATORS,
        "basic_data_service",
    )
    daily_specs = [
        (
            "基本面数据token",
            endpoint,
            output_name,
            indicator,
        )
        for output_name, (endpoint, indicator) in BASIC_DATA_TOKEN_INDICATORS.items()
        if endpoint == "date_sequence"
    ]
    daily_specs.extend(
        (
            "历史行情token",
            endpoint,
            output_name,
            indicator,
        )
        for output_name, (endpoint, indicator) in HISTORY_QUOTE_TOKEN_INDICATORS.items()
        if endpoint == "cmd_history_quotation"
    )

    plans: dict[str, EtfDownloadPlan] = {}
    range_start_text = config.start_date.isoformat()
    range_end_text = config.end_date.isoformat()
    sorted_target_dates = sorted(set(target_dates))
    last_raw_dates = {
        str(code): date.fromisoformat(str(last_date))
        for code, last_date in connection.execute(
            "SELECT code,MAX(date) FROM raw_data "
            "WHERE date BETWEEN ? AND ? GROUP BY code",
            (range_start_text, range_end_text),
        )
        if last_date
    }

    for code in sorted(codes):
        master_row = master[code]
        listed_date = listed_date_or_none(master_row)
        if listed_date and listed_date > config.end_date:
            plans[code] = EtfDownloadPlan((), ())
            continue

        raw_count = int(
            connection.execute(
                "SELECT COUNT(*) FROM raw_data "
                "WHERE code=? AND date BETWEEN ? AND ?",
                (code, range_start_text, range_end_text),
            ).fetchone()[0]
        )
        missing_fixed: list[tuple[str, str]] = []
        for output_name, indicator in fixed_fields:
            db_field = RAW_DB_COLUMNS[output_name]
            blank_count = int(
                connection.execute(
                    f"SELECT COUNT(*) FROM raw_data WHERE code=? "
                    "AND date BETWEEN ? AND ? "
                    f"AND TRIM(COALESCE({db_field},''))=''",
                    (code, range_start_text, range_end_text),
                ).fetchone()[0]
            )
            if blank_count or (not raw_count and not clean_value(master_row.get(output_name))):
                missing_fixed.append((output_name, indicator))

        last_raw_date = last_raw_dates.get(code)
        # 已有日期只用于检查历史空值；本地最后日期之后的区间在下方单独追加。
        # 这样既不会把周末制造成历史缺口，也不会再遗漏全新的年份或月份。
        relevant_dates = [
            current_date
            for current_date in sorted_target_dates
            if last_raw_date is not None
            and current_date <= last_raw_date
            and (listed_date is None or current_date >= listed_date)
        ]
        daily_requests: list[FieldRequest] = []
        for route_type, endpoint, output_name, indicator in daily_specs:
            missing_ranges: list[tuple[date, date]] = []
            if relevant_dates:
                if output_name in force_fields:
                    missing_ranges = [(relevant_dates[0], relevant_dates[-1])]
                else:
                    db_field = DAILY_DB_COLUMNS[output_name]
                    completed_rows = connection.execute(
                        f"SELECT date FROM daily WHERE code=? "
                        "AND date BETWEEN ? AND ? "
                        f"AND TRIM(COALESCE({db_field},''))<>''",
                        (code, range_start_text, range_end_text),
                    ).fetchall()
                    completed_dates = {
                        date.fromisoformat(str(row[0])) for row in completed_rows
                    }
                    missing_ranges = list(
                        consecutive_missing_ranges(relevant_dates, completed_dates)
                    )

            # 向后更新不能依赖 raw 中已经出现过的交易日。直接请求连续日历区间，
            # iFinD 会只返回真实交易日；新ETF则从上市日期（若已取得）开始。
            extension_start = (
                last_raw_date + timedelta(days=1)
                if last_raw_date is not None
                else config.start_date
            )
            extension_start = max(extension_start, config.start_date)
            if listed_date is not None:
                extension_start = max(extension_start, listed_date)
            if extension_start <= config.end_date:
                missing_ranges.append((extension_start, config.end_date))

            if missing_ranges:
                daily_requests.append(
                    FieldRequest(
                        route_type=route_type,
                        endpoint=endpoint,
                        output_name=output_name,
                        indicator=indicator,
                        ranges=tuple(missing_ranges),
                    )
                )

        plans[code] = EtfDownloadPlan(
            basic_fields=tuple(missing_fixed),
            daily_requests=tuple(daily_requests),
        )
    return plans


def update_basic_fields_in_raw(
    connection: sqlite3.Connection,
    master: Mapping[str, Mapping[str, str]],
    codes: Sequence[str],
    fields: Sequence[tuple[str, str]],
    config: RuntimeConfig,
) -> dict[str, set[str]]:
    """只把本次取得的固定字段写进原来为空的对应日期单元格。"""

    changed_fields: dict[str, set[str]] = {code: set() for code in codes}
    for code in codes:
        for output_name, _ in fields:
            value = clean_value(master[code].get(output_name))
            if not value:
                continue
            db_field = RAW_DB_COLUMNS[output_name]
            cursor = connection.execute(
                f"UPDATE raw_data SET {db_field}=? WHERE code=? "
                "AND date BETWEEN ? AND ? "
                f"AND TRIM(COALESCE({db_field},''))=''",
                (
                    value,
                    code,
                    config.start_date.isoformat(),
                    config.end_date.isoformat(),
                ),
            )
            if cursor.rowcount:
                changed_fields[code].add(output_name)
    connection.commit()
    return changed_fields


def save_accumulated_plan_results(
    connection: sqlite3.Connection,
    master: Mapping[str, Mapping[str, str]],
    codes: Sequence[str],
    basic_changed_codes: set[str],
) -> None:
    """先保存成功返回的日频字段，再保存仅固定字段发生变化的raw。"""

    candidate_codes = set(codes)
    dirty_codes = {
        str(row[0])
        for row in connection.execute(
            "SELECT DISTINCT code FROM daily WHERE raw_saved=0"
        )
        if str(row[0]) in candidate_codes
    }
    persist_completed_raw_rows(connection, master, codes)
    for code in sorted(basic_changed_codes - dirty_codes):
        write_one_raw_file(connection, code)
    connection.commit()


def execute_field_plan_group(
    connection: sqlite3.Connection,
    basic_client: IFindClient,
    history_client: IFindClient,
    codes: tuple[str, ...],
    plan: EtfDownloadPlan,
    master: dict[str, dict[str, str]],
    config: RuntimeConfig,
) -> dict[str, set[str]]:
    """批量执行一组完全相同的指标缺口，并保存已成功返回的字段。"""

    basic_changed_codes: set[str] = set()
    written_fields: dict[str, set[str]] = {code: set() for code in codes}
    try:
        if plan.basic_fields:
            basic_rows = fetch_basic_fields(
                basic_client,
                codes,
                plan.basic_fields,
            )
            for code, fields in basic_rows.items():
                for output_name, value in fields.items():
                    cleaned = clean_value(value)
                    if cleaned:
                        master[code][output_name] = cleaned
            basic_changes = update_basic_fields_in_raw(
                connection,
                master,
                codes,
                plan.basic_fields,
                config,
            )
            for code, fields in basic_changes.items():
                written_fields[code].update(fields)
            basic_changed_codes = {
                code for code, fields in basic_changes.items() if fields
            }

        for request in plan.daily_requests:
            for range_start, range_end in request.ranges:
                for request_start, request_end in split_date_range(range_start, range_end):
                    task = DownloadTask(
                        route_type=request.route_type,
                        endpoint=request.endpoint,
                        codes=codes,
                        start_date=request_start,
                        end_date=request_end,
                        fields=((request.output_name, request.indicator),),
                    )
                    _, payload = fetch_daily_task(
                        task,
                        basic_client,
                        history_client,
                        config.forward_adjustment_base_date,
                    )
                    written_counts = insert_daily_payload(
                        connection,
                        task,
                        payload,
                        set(codes),
                    )
                    for code, count in written_counts.items():
                        if count:
                            written_fields[code].add(request.output_name)
    finally:
        save_accumulated_plan_results(
            connection,
            master,
            codes,
            basic_changed_codes,
        )
    return written_fields


def formatted_plan_indicators(plan: EtfDownloadPlan) -> str:
    """按最终CSV列顺序显示本次真正请求的指标名称。"""

    names = sorted(
        plan.indicator_names(),
        key=lambda name: OUTPUT_COLUMNS.index(name),
    )
    return "、".join(names)


def print_download_start(
    code: str,
    indicators: str,
    position: int,
    total: int,
) -> None:
    print(
        f"正在下载 {code} 数据（本次指标：{indicators}）... ({position}/{total})",
        flush=True,
    )


def print_download_success(
    code: str,
    requested_indicators: str,
    written_fields: set[str],
) -> None:
    output_path = RAW_OUTPUT_DIR / f"{code}.csv"
    if written_fields:
        saved_indicators = "、".join(
            sorted(written_fields, key=lambda name: OUTPUT_COLUMNS.index(name))
        )
        print(
            f"✅ 成功保存 {code} 数据（本次指标：{saved_indicators}）到 {output_path}",
            flush=True,
        )
    else:
        print(
            f"⚠️  {code} 本次指标（{requested_indicators}）请求完成，"
            "但没有写入新数据",
            flush=True,
        )


def download_missing_fields(
    connection: sqlite3.Connection,
    basic_client: IFindClient | None,
    history_client: IFindClient | None,
    plans: Mapping[str, EtfDownloadPlan],
    master: dict[str, dict[str, str]],
    target_dates: Sequence[date],
    config: RuntimeConfig,
    *,
    force_fields: set[str] | None = None,
) -> list[str]:
    """按代码升序处理全部ETF，日志序号严格连续且不按缺口分组跳号。"""

    codes = sorted(plans)
    total = len(codes)
    failed_codes: list[str] = []
    print("\n开始检查并下载ETF数据...", flush=True)

    index = 0
    while index < total:
        code = codes[index]
        plan = plans[code]
        position = index + 1
        if plan.is_empty():
            print(f"正在检查 {code} 数据... ({position}/{total})", flush=True)
            print(
                f"⏭️ {code} 指定日期的全部指标已存在，跳过下载",
                flush=True,
            )
            index += 1
            continue

        if basic_client is None or history_client is None:
            raise RuntimeError("存在指标缺口，但同花顺客户端尚未登录。")

        group_codes = [code]
        next_index = index + 1
        while (
            next_index < total
            and len(group_codes) < ETF_CODES_PER_REQUEST
            and plans[codes[next_index]] == plan
        ):
            group_codes.append(codes[next_index])
            next_index += 1

        indicators = formatted_plan_indicators(plan)
        try:
            written_by_code = execute_field_plan_group(
                connection,
                basic_client,
                history_client,
                tuple(group_codes),
                plan,
                master,
                config,
            )
            for offset, current_code in enumerate(group_codes):
                current_position = position + offset
                print_download_start(
                    current_code,
                    indicators,
                    current_position,
                    total,
                )
                print_download_success(
                    current_code,
                    indicators,
                    written_by_code.get(current_code, set()),
                )
        except Exception as batch_exc:
            if len(group_codes) > 1:
                print(
                    f"⚠️  本组请求未完成，正在按ETF升序逐只重试：{batch_exc}",
                    flush=True,
                )
            for offset, current_code in enumerate(group_codes):
                current_position = position + offset
                refreshed_plan = build_field_download_plans(
                    connection,
                    master,
                    (current_code,),
                    target_dates,
                    config,
                    force_fields=force_fields,
                )[current_code]
                retry_plan = refreshed_plan if not refreshed_plan.is_empty() else plan
                retry_indicators = formatted_plan_indicators(retry_plan)
                print_download_start(
                    current_code,
                    retry_indicators,
                    current_position,
                    total,
                )
                try:
                    if not refreshed_plan.is_empty():
                        retry_written = execute_field_plan_group(
                            connection,
                            basic_client,
                            history_client,
                            (current_code,),
                            refreshed_plan,
                            master,
                            config,
                        )
                        written_fields = retry_written.get(current_code, set())
                    else:
                        written_fields = set(retry_plan.indicator_names())
                    print_download_success(
                        current_code,
                        retry_indicators,
                        written_fields,
                    )
                except Exception as exc:
                    failed_codes.append(current_code)
                    print(
                        f"❌ 下载 {current_code} 失败"
                        f"（本次指标：{retry_indicators}）：{exc}",
                        flush=True,
                    )
        index = next_index

    return failed_codes


def configured_forward_adjustment_fields() -> list[str]:
    """返回当前已配置且会受前复权基点影响的历史行情输出列。"""

    return [
        output_name
        for output_name, (endpoint, indicator) in HISTORY_QUOTE_TOKEN_INDICATORS.items()
        if endpoint == "cmd_history_quotation"
        and indicator.lower() in ADJUSTMENT_SENSITIVE_HISTORY_INDICATORS
    ]


def read_saved_adjustment_base_date() -> str:
    """读取上次完整下载时使用的前复权 BaseDate。"""

    if not ADJUSTMENT_STATE_FILE.exists():
        return ""
    with ADJUSTMENT_STATE_FILE.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        row = next(reader, None)
    if not row:
        return ""
    return clean_value(row.get("前复权基点日期"))


def save_adjustment_base_date(current_date: date) -> None:
    """下载全部成功后，原子保存本次使用的 BaseDate。"""

    temp_path, handle = atomic_csv_writer(ADJUSTMENT_STATE_FILE)
    try:
        with handle:
            writer = csv.DictWriter(handle, fieldnames=["前复权基点日期"])
            writer.writeheader()
            writer.writerow({"前复权基点日期": current_date.isoformat()})
        temp_path.replace(ADJUSTMENT_STATE_FILE)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def table_values_case_insensitive(
    table: Mapping[str, Any],
    indicator: str,
) -> list[Any]:
    """历史行情少数字段可能改变大小写，因此按不区分大小写查找。"""

    target = indicator.lower()
    for key, values in table.items():
        if str(key).lower() == target and isinstance(values, list):
            return values
    return []


def iter_daily_rows(
    payload: Mapping[str, Any],
    fields: Sequence[tuple[str, str]],
    valid_codes: set[str],
) -> Iterator[tuple[str, str, list[str]]]:
    """把一个接口响应展开为 date、code 和对应指标值。"""

    tables = payload.get("tables")
    if not isinstance(tables, list):
        return
    for item in tables:
        if not isinstance(item, dict):
            continue
        code = normalize_code(item.get("thscode"))
        times = item.get("time")
        table = item.get("table")
        if code not in valid_codes or not isinstance(times, list) or not isinstance(table, dict):
            continue

        value_lists = [
            table_values_case_insensitive(table, indicator)
            for _, indicator in fields
        ]
        for index, raw_time in enumerate(times):
            current_date = clean_value(raw_time)[:10]
            if not current_date:
                continue
            values = [
                clean_value(items[index]) if index < len(items) else ""
                for items in value_lists
            ]
            yield current_date, code, values


def insert_daily_payload(
    connection: sqlite3.Connection,
    task: DownloadTask,
    payload: Mapping[str, Any],
    valid_codes: set[str],
) -> dict[str, int]:
    """把一个指标任务增量写入SQLite，并把对应raw行标记为待保存。"""

    if task.route_type not in {"基本面数据token", "历史行情token"}:
        raise ValueError(f"未知token路由类型：{task.route_type}")

    db_columns = [DAILY_DB_COLUMNS[output_name] for output_name, _ in task.fields]
    insert_columns = ["date", "code", *db_columns, "raw_saved"]
    placeholders = ",".join("?" for _ in insert_columns)
    updates = ",".join(
        [
            *(f"{column}=excluded.{column}" for column in db_columns),
            "raw_saved=0",
        ]
    )
    sql = (
        f"INSERT INTO daily ({','.join(insert_columns)}) VALUES ({placeholders}) "
        f"ON CONFLICT(date,code) DO UPDATE SET {updates}"
    )

    written_counts = {code: 0 for code in valid_codes}
    buffer: list[tuple[Any, ...]] = []
    for current_date, code, values in iter_daily_rows(payload, task.fields, valid_codes):
        buffer.append((current_date, code, *values, 0))
        written_counts[code] += sum(bool(value) for value in values)
        if len(buffer) >= 10_000:
            connection.executemany(sql, buffer)
            buffer.clear()
    if buffer:
        connection.executemany(sql, buffer)
    connection.commit()
    return written_counts


def database_dates(connection: sqlite3.Connection) -> list[str]:
    """取得全部 raw 基础数据的日期。"""

    rows = connection.execute("SELECT DISTINCT date FROM raw_data ORDER BY date").fetchall()
    dates = [str(row[0]) for row in rows]
    if not dates:
        raise RuntimeError(
            "etf_raw_data中没有任何已完成的ETF-日期数据；"
            "请检查日期是否为交易日及指标权限。"
        )
    return dates


def database_codes(connection: sqlite3.Connection) -> list[str]:
    """取得全部 raw 基础数据的ETF代码。"""

    rows = connection.execute("SELECT DISTINCT code FROM raw_data ORDER BY code").fetchall()
    return [str(row[0]) for row in rows]


def combined_row(
    current_date: str,
    master_row: Mapping[str, str],
    daily_row: Mapping[str, str] | None,
) -> dict[str, str]:
    """按固定列顺序合并一天、一只ETF的基本面与日频数据。"""

    daily_row = daily_row or {}
    return {
        "日期": current_date,
        "代码": master_row.get("代码", ""),
        "名称": master_row.get("名称", ""),
        "上市日期": master_row.get("上市日期", ""),
        "对标指数": master_row.get("对标指数", ""),
        "benchmark": master_row.get("benchmark", ""),
        "管理人": master_row.get("管理人", ""),
        "托管人": master_row.get("托管人", ""),
        "类别大类": master_row.get("类别大类", ""),
        "类别细类": master_row.get("类别细类", ""),
        "对标指数代码": master_row.get("对标指数代码", ""),
        "规模": daily_row.get("规模", ""),
        "成交量": daily_row.get("成交量", ""),
        "成交额": daily_row.get("成交额", ""),
        "收盘价": daily_row.get("收盘价", ""),
        "VWAP": daily_row.get("VWAP", ""),
        "溢价率": daily_row.get("溢价率", ""),
    }


def atomic_csv_writer(path: Path) -> tuple[Path, Any]:
    """创建同目录临时CSV；调用方写完后再 replace，避免留下半个文件。"""

    temp_path = path.with_name(f".{path.name}.tmp")
    handle = temp_path.open("w", encoding="utf-8-sig", newline="")
    return temp_path, handle


def raw_select_columns() -> str:
    """按 OUTPUT_COLUMNS 的顺序生成 raw_data 查询列。"""

    return ",".join(RAW_DB_COLUMNS[column] for column in OUTPUT_COLUMNS)


def write_one_raw_file(connection: sqlite3.Connection, code: str) -> None:
    """从临时合并表原子重写某只ETF的 raw CSV。"""

    output_path = RAW_OUTPUT_DIR / f"{code}.csv"
    temp_path, handle = atomic_csv_writer(output_path)
    try:
        with handle:
            writer = csv.writer(handle)
            writer.writerow(OUTPUT_COLUMNS)
            cursor = connection.execute(
                f"SELECT {raw_select_columns()} FROM raw_data "
                "WHERE code=? ORDER BY date",
                (code,),
            )
            for record in cursor:
                writer.writerow([clean_value(value) for value in record])
        temp_path.replace(output_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def persist_completed_raw_rows(
    connection: sqlite3.Connection,
    master: Mapping[str, Mapping[str, str]],
    candidate_codes: Sequence[str],
) -> int:
    """把本次成功补到任意指标的ETF-日期合并进单ETF raw CSV。"""

    raw_columns = [RAW_DB_COLUMNS[column] for column in OUTPUT_COLUMNS]
    placeholders = ",".join("?" for _ in raw_columns)
    upsert_sql = (
        f"INSERT INTO raw_data ({','.join(raw_columns)}) VALUES ({placeholders}) "
        "ON CONFLICT(date,code) DO UPDATE SET "
        + ",".join(
            f"{column}=excluded.{column}"
            for column in raw_columns
            if column not in {"date", "code"}
        )
    )

    saved_count = 0
    for code in sorted(set(candidate_codes)):
        if code not in master:
            continue
        pending = connection.execute(
            "SELECT date,fund_scale,volume,amount,close,vwap,premium_ratio FROM daily "
            "WHERE code=? AND raw_saved=0 "
            "ORDER BY date",
            (code,),
        ).fetchall()
        if not pending:
            continue

        raw_rows: list[tuple[str, ...]] = []
        for current_date, scale, volume, amount, close, vwap, premium_ratio in pending:
            row = combined_row(
                str(current_date),
                master[code],
                {
                    "规模": clean_value(scale),
                    "成交量": clean_value(volume),
                    "成交额": clean_value(amount),
                    "收盘价": clean_value(close),
                    "VWAP": clean_value(vwap),
                    "溢价率": clean_value(premium_ratio),
                },
            )
            raw_rows.append(tuple(clean_value(row[column]) for column in OUTPUT_COLUMNS))

        connection.executemany(upsert_sql, raw_rows)
        # CSV 替换成功后才标记 raw_saved。即使写文件时中断，旧文件也不会损坏。
        write_one_raw_file(connection, code)
        connection.executemany(
            "UPDATE daily SET raw_saved=1 WHERE date=? AND code=?",
            [(str(current_date), code) for current_date, *_ in pending],
        )
        connection.commit()
        saved_count += len(pending)

    return saved_count


def save_total_table(connection: sqlite3.Connection) -> None:
    """完全从 raw_data 输出总长表 etf_data.csv。"""

    temp_path, handle = atomic_csv_writer(TOTAL_OUTPUT_FILE)
    try:
        with handle:
            writer = csv.writer(handle)
            writer.writerow(OUTPUT_COLUMNS)
            cursor = connection.execute(
                f"SELECT {raw_select_columns()} FROM raw_data ORDER BY date,code"
            )
            for record in cursor:
                writer.writerow([clean_value(value) for value in record])
        temp_path.replace(TOTAL_OUTPUT_FILE)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def save_one_wide_table(
    connection: sqlite3.Connection,
    codes: Sequence[str],
    dates: Sequence[str],
    field: str,
    file_name: str,
) -> None:
    """完全从 raw_data 输出某一字段的日期×ETF代码宽表。"""

    output_path = OUTPUT_DIR / file_name
    temp_path, handle = atomic_csv_writer(output_path)
    db_field = RAW_DB_COLUMNS[field]
    try:
        with handle:
            writer = csv.writer(handle)
            writer.writerow(["日期", *codes])
            for current_date in dates:
                values_by_code = {
                    str(code): clean_value(value)
                    for code, value in connection.execute(
                        f"SELECT code,{db_field} FROM raw_data "
                        "WHERE date=? ORDER BY code",
                        (current_date,),
                    )
                }
                writer.writerow(
                    [current_date, *(values_by_code.get(code, "") for code in codes)]
                )
        temp_path.replace(output_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def save_wide_tables(
    connection: sqlite3.Connection,
    codes: Sequence[str],
    dates: Sequence[str],
) -> None:
    """从 raw 依次输出名称、基本面、类别和日频指标宽表。"""

    for field, file_name in WIDE_FILE_NAMES.items():
        save_one_wide_table(
            connection,
            codes,
            dates,
            field,
            file_name,
        )


def validate_master(
    master: Mapping[str, Mapping[str, str]],
    codes: Sequence[str],
) -> None:
    """输出前只检查代码列表与主表是否一致。"""

    if set(codes) != set(master):
        raise RuntimeError("ETF代码列表与基本面主表不一致。")


def prepare_output_directories() -> None:
    """仅创建用户指定的两个输出目录，不创建JSON目录或其他旁路输出。"""

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    RAW_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def main() -> None:
    """主流程：先本地逐指标查缺，再只为真实缺口鉴权和请求。"""

    started_at = time.perf_counter()
    config = validate_config()
    print("同花顺ETF数据下载工具", flush=True)
    print("=" * 60, flush=True)
    print(f"下载日期：{config.start_date} 至 {config.end_date}", flush=True)
    print(f"前复权基点日期：{config.forward_adjustment_base_date}", flush=True)
    prepare_output_directories()

    basic_client: IFindClient | None = None
    history_client: IFindClient | None = None

    def ensure_clients() -> tuple[IFindClient, IFindClient]:
        nonlocal basic_client, history_client
        if basic_client is not None and history_client is not None:
            return basic_client, history_client
        print("正在登录同花顺...", flush=True)
        access_token = get_access_token(IFIND_REFRESH_TOKEN, "iFinD账号")
        print("✅ Access Token 已设置", flush=True)
        basic_client = IFindClient("基本面数据token", access_token)
        history_client = IFindClient("历史行情token", access_token)
        return basic_client, history_client

    with tempfile.TemporaryDirectory(prefix="ifind_etf_data_") as temp_dir:
        database_path = Path(temp_dir) / "daily.sqlite3"
        connection = create_database(database_path)
        try:
            # 每次都刷新 END_DATE 对应的ETF名单，再与本地历史ETF取并集。
            # 当前名单负责发现新上市ETF；本地名单负责保留已退市或已移出板块的ETF。
            try:
                current_basic_client, _ = ensure_clients()
                current_master, current_codes = build_master_data(
                    current_basic_client,
                    config,
                )
            except Exception as exc:
                print(f"❌ 获取ETF列表失败：{exc}", flush=True)
                return

            local_codes = discover_local_universe_codes(config)
            new_codes = sorted(set(current_codes) - set(local_codes))
            if new_codes:
                print(
                    f"正在获取 {len(new_codes)} 只新增ETF的上市日期和基本面...",
                    flush=True,
                )
                try:
                    for start_index in range(0, len(new_codes), ETF_CODES_PER_REQUEST):
                        code_batch = new_codes[
                            start_index:start_index + ETF_CODES_PER_REQUEST
                        ]
                        basic_rows = fetch_basic_fields(
                            current_basic_client,
                            code_batch,
                        )
                        for code, fields in basic_rows.items():
                            for output_name, value in fields.items():
                                cleaned = clean_value(value)
                                if cleaned:
                                    current_master[code][output_name] = cleaned
                except Exception as exc:
                    print(
                        f"⚠️ 新增ETF基本面预取未完成：{exc}；"
                        "后续将按完整配置区间请求，不会漏掉行情",
                        flush=True,
                    )

            codes = sorted(set(current_codes) | set(local_codes))
            existing_rows = load_existing_raw_data(connection, set(codes))
            local_master = build_master_from_raw_data(connection, codes)
            master = merge_master_snapshots(
                local_master,
                current_master,
                codes,
            )
            print(
                f"✅ ETF名单合并完成：同花顺当前 {len(current_codes)} 只，"
                f"本地历史 {len(local_codes)} 只，合并后 {len(codes)} 只",
                flush=True,
            )

            validate_master(master, codes)
            target_dates = raw_dates_in_range(connection, config)
            if target_dates and target_dates[-1] < config.end_date:
                print(
                    f"✅ 本地raw最晚日期为 {target_dates[-1]}，"
                    f"本次将向后请求至 {config.end_date}",
                    flush=True,
                )
            elif not target_dates:
                print(
                    f"本地raw在指定区间内没有数据，本次将请求 "
                    f"{config.start_date} 至 {config.end_date}",
                    flush=True,
                )

            adjustment_fields = set(configured_forward_adjustment_fields())
            saved_base_date = read_saved_adjustment_base_date()
            current_base_date = config.forward_adjustment_base_date.isoformat()
            force_fields = (
                adjustment_fields
                if adjustment_fields
                and existing_rows
                and saved_base_date != current_base_date
                else set()
            )
            if force_fields:
                print(
                    f"前复权基点已变更为 {current_base_date}，"
                    f"本次仅重下受影响指标：{'、'.join(sorted(force_fields))}",
                    flush=True,
                )

            plans = build_field_download_plans(
                connection,
                master,
                codes,
                target_dates,
                config,
                force_fields=force_fields,
            )
            has_downloads = any(not plan.is_empty() for plan in plans.values())
            if has_downloads and (basic_client is None or history_client is None):
                try:
                    ensure_clients()
                except Exception as exc:
                    print(f"❌ 登录失败：{exc}", flush=True)
                    return

            failed_codes = download_missing_fields(
                connection,
                basic_client,
                history_client,
                plans,
                master,
                target_dates,
                config,
                force_fields=force_fields,
            )

            # 下载补缺处理raw中已识别的全部ETF大类；最终总表和宽表还会
            # 汇总 etf_raw_data 中的全部ETF，且这一步只读本地文件、不调用接口。
            additional_raw_codes = set(discover_all_local_raw_codes()) - set(codes)
            if additional_raw_codes:
                additional_rows = load_existing_raw_data(
                    connection,
                    additional_raw_codes,
                )
                print(
                    f"✅ 汇总阶段从本地raw补充载入ETF："
                    f"{len(additional_raw_codes)} 只，{additional_rows} 行",
                    flush=True,
                )

            dates = database_dates(connection)
            output_codes = database_codes(connection)
            print("\n开始整理ETF数据...", flush=True)
            save_total_table(connection)
            save_wide_tables(connection, output_codes, dates)
            print("✅ ETF数据整理完成", flush=True)

            if adjustment_fields and not failed_codes:
                save_adjustment_base_date(config.forward_adjustment_base_date)
        finally:
            connection.close()

    print("\n" + "=" * 60, flush=True)
    print("全部ETF数据处理完成！", flush=True)
    print(f"总表保存在: {TOTAL_OUTPUT_FILE}", flush=True)
    print(f"单ETF原始数据保存在: {RAW_OUTPUT_DIR}", flush=True)
    print(f"共 {len(output_codes)} 只ETF，{len(dates)} 个交易日", flush=True)
    if failed_codes:
        print(f"❌ 仍有 {len(failed_codes)} 只ETF下载失败，重新运行会自动续传", flush=True)
    else:
        print("✅ 所有ETF均已下载成功", flush=True)
    print(f"用时: {time.perf_counter() - started_at:.1f}秒", flush=True)


if __name__ == "__main__":
    main()
