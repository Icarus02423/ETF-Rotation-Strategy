"""分别检验两种ETF趋势得分的截面Rank IC。

口径：
1. 每个信号日分别按两种趋势得分降序选择前10%的有效指数，数量向上取整；随后
   仅保留窗口收益率大于0的指数，不向后补选。
2. 按现有回测口径，在信号日为每个指数选择成交量最大的ETF；成交量相同则依次
   比较成交额、规模和ETF代码。
3. 下一ETF交易日按VWAP买入，持有1至20个交易日后按VWAP卖出。
4. 每个信号日、每个持有期计算一次Spearman秩相关系数（Rank IC）。

脚本按得分公式分别生成两个文件；每个文件都包含“汇总”和“逐日RankIC”：
outputs/etf_strategy_test/trend_rank_ic/01_收益率乘R平方.xlsx
outputs/etf_strategy_test/trend_rank_ic/02_收益率除以波动率.xlsx
"""

from __future__ import annotations

import csv
import math
from bisect import bisect_right
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path(__file__).resolve().parents[2]
FACTOR_ROOT = PROJECT_ROOT / "outputs" / "etf_trend_strategy"
ETF_DATA_FILE = PROJECT_ROOT / "outputs" / "etf_data" / "etf_data.csv"
OUTPUT_DIR = PROJECT_ROOT / "outputs" / "etf_strategy_test" / "trend_rank_ic"

HOLDING_DAYS = tuple(range(1, 21))
TOP_PERCENT = 0.20
MIN_WINDOW_RETURN = 0.0

SCORE_COLUMNS = {
    "收益率×R平方": "趋势质量因子",
    "收益率÷波动率": "风险调整趋势得分",
}
OUTPUT_FILES = {
    "收益率×R平方": OUTPUT_DIR / "01_收益率乘R平方.xlsx",
    "收益率÷波动率": OUTPUT_DIR / "02_收益率除以波动率.xlsx",
}
FACTOR_COLUMNS = {"日期", "对标指数代码", "窗口收益率"} | set(SCORE_COLUMNS.values())
ETF_COLUMNS = {
    "日期",
    "代码",
    "上市日期",
    "对标指数代码",
    "规模",
    "成交量",
    "成交额",
    "VWAP",
}


@dataclass(frozen=True)
class CandidateETF:
    code: str
    volume: float
    amount: float
    scale: float


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none", "null"} else text


def parse_date(value: object) -> date | None:
    text = clean_text(value)
    if not text:
        return None
    compact_date = text[:8]
    if len(compact_date) == 8 and compact_date.isdigit():
        try:
            return date(
                int(compact_date[:4]),
                int(compact_date[4:6]),
                int(compact_date[6:8]),
            )
        except ValueError:
            return None
    try:
        return date.fromisoformat(text[:10].replace("/", "-"))
    except ValueError:
        return None


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


def candidate_sort_key(candidate: CandidateETF) -> tuple[float, float, float, str]:
    return (-candidate.volume, -candidate.amount, -candidate.scale, candidate.code)


def parse_configuration(window_dir: Path) -> tuple[float, int]:
    try:
        threshold = float(window_dir.parents[1].name.removeprefix("threshold_"))
        trend_window = int(window_dir.name.removeprefix("window_"))
    except ValueError as exc:
        raise ValueError(f"无法识别因子目录参数：{window_dir}") from exc
    return threshold, trend_window


def load_factor_signals() -> pd.DataFrame:
    """读取所有已有参数组合的完整因子截面。"""

    if not FACTOR_ROOT.exists():
        raise FileNotFoundError(f"找不到趋势因子目录：{FACTOR_ROOT}")

    frames: list[pd.DataFrame] = []
    window_dirs = sorted(FACTOR_ROOT.glob("threshold_*/factors/window_*"))
    for window_dir in window_dirs:
        threshold, trend_window = parse_configuration(window_dir)
        for factor_file in sorted(window_dir.glob("*.csv")):
            frame = pd.read_csv(
                factor_file,
                encoding="utf-8-sig",
                usecols=lambda column: column in FACTOR_COLUMNS,
                dtype={"对标指数代码": "string"},
                low_memory=False,
            )
            missing = FACTOR_COLUMNS - set(frame.columns)
            if missing:
                raise ValueError(
                    f"{factor_file} 缺少列：{sorted(missing)}。"
                    "请先用趋势因子计算.py重新生成这个参数组合。"
                )

            frame = frame.rename(
                columns={
                    "日期": "signal_date",
                    "对标指数代码": "index_code",
                    "窗口收益率": "window_return",
                }
            )
            frame["signal_date"] = pd.to_datetime(
                frame["signal_date"], errors="coerce"
            ).dt.date
            frame["index_code"] = frame["index_code"].str.strip()
            score_columns = list(SCORE_COLUMNS.values())
            numeric_columns = score_columns + ["window_return"]
            for score_column in numeric_columns:
                frame[score_column] = pd.to_numeric(
                    frame[score_column],
                    errors="coerce",
                )
            frame = frame.dropna(subset=["window_return"])
            frame = frame[frame["window_return"].map(math.isfinite)]

            frame = frame.melt(
                id_vars=["signal_date", "index_code", "window_return"],
                value_vars=score_columns,
                var_name="score_column",
                value_name="score",
            )
            score_labels = {
                column: label for label, column in SCORE_COLUMNS.items()
            }
            frame["score_method"] = frame["score_column"].map(score_labels)
            frame = frame.drop(columns="score_column")
            frame["threshold"] = threshold
            frame["trend_window"] = trend_window
            frames.append(frame)

    if not frames:
        raise FileNotFoundError(f"没有在 {FACTOR_ROOT} 下找到因子CSV")

    signals = pd.concat(frames, ignore_index=True)
    signals = signals.dropna(subset=["signal_date", "index_code", "score"])
    signals = signals[signals["index_code"].ne("")]
    signals = signals[signals["score"].map(math.isfinite)]
    signals = signals.drop_duplicates(
        subset=[
            "threshold",
            "trend_window",
            "signal_date",
            "index_code",
            "score_method",
        ],
        keep="last",
    )
    signals = signals.sort_values(
        [
            "threshold",
            "trend_window",
            "signal_date",
            "score_method",
            "index_code",
        ]
    ).reset_index(drop=True)
    if signals.empty:
        raise ValueError("因子CSV中没有可用于Rank IC检验的有效记录")
    return signals


def scan_etf_mapping(
    signals: pd.DataFrame,
) -> tuple[dict[tuple[date, str], CandidateETF], list[date]]:
    """第一次扫描ETF总表：建立交易日历和信号日的指数到ETF映射。"""

    if not ETF_DATA_FILE.exists():
        raise FileNotFoundError(f"找不到ETF数据：{ETF_DATA_FILE}")

    required_indexes: dict[date, set[str]] = defaultdict(set)
    for row in signals[["signal_date", "index_code"]].itertuples(index=False):
        required_indexes[row.signal_date].add(row.index_code)

    trading_dates: set[date] = set()
    best_by_date_index: dict[tuple[date, str], CandidateETF] = {}

    with ETF_DATA_FILE.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = ETF_COLUMNS - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"etf_data.csv缺少列：{sorted(missing)}")

        for row in reader:
            current_date = parse_date(row.get("日期"))
            if current_date is None:
                continue
            trading_dates.add(current_date)

            index_code = clean_text(row.get("对标指数代码"))
            if index_code not in required_indexes.get(current_date, set()):
                continue

            code = clean_text(row.get("代码")).upper()
            volume = positive_float(row.get("成交量"))
            if not code or volume is None:
                continue

            listed_text = clean_text(row.get("上市日期"))
            if listed_text:
                listed_date = parse_date(listed_text)
                if listed_date is None or listed_date > current_date:
                    continue

            candidate = CandidateETF(
                code=code,
                volume=volume,
                amount=finite_float(row.get("成交额")) or 0.0,
                scale=finite_float(row.get("规模")) or 0.0,
            )
            key = (current_date, index_code)
            existing = best_by_date_index.get(key)
            if existing is None or candidate_sort_key(candidate) < candidate_sort_key(existing):
                best_by_date_index[key] = candidate

    if not trading_dates:
        raise ValueError("ETF总表中没有有效交易日期")
    return best_by_date_index, sorted(trading_dates)


def build_trade_schedule(
    signal_dates: set[date], trading_dates: list[date]
) -> dict[tuple[date, int], tuple[date, date]]:
    """生成信号日对应的下一交易日买入及1至20日后的卖出日期。"""

    schedule: dict[tuple[date, int], tuple[date, date]] = {}
    for signal_date in signal_dates:
        entry_position = bisect_right(trading_dates, signal_date)
        if entry_position >= len(trading_dates):
            continue
        entry_date = trading_dates[entry_position]
        for holding_day in HOLDING_DAYS:
            exit_position = entry_position + holding_day
            if exit_position < len(trading_dates):
                schedule[(signal_date, holding_day)] = (
                    entry_date,
                    trading_dates[exit_position],
                )
    return schedule


def collect_required_prices(
    signals: pd.DataFrame,
    mapping: dict[tuple[date, str], CandidateETF],
    schedule: dict[tuple[date, int], tuple[date, date]],
) -> set[tuple[date, str]]:
    required: set[tuple[date, str]] = set()
    for row in signals[["signal_date", "index_code"]].itertuples(index=False):
        candidate = mapping.get((row.signal_date, row.index_code))
        if candidate is None:
            continue
        for holding_day in HOLDING_DAYS:
            dates = schedule.get((row.signal_date, holding_day))
            if dates is None:
                continue
            entry_date, exit_date = dates
            required.add((entry_date, candidate.code))
            required.add((exit_date, candidate.code))
    return required


def load_required_vwap(
    required: set[tuple[date, str]],
) -> dict[tuple[date, str], float]:
    """第二次扫描ETF总表，只保留计算未来收益所需的VWAP。"""

    prices: dict[tuple[date, str], float] = {}
    if not required:
        return prices

    required_codes = {code for _, code in required}
    with ETF_DATA_FILE.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            code = clean_text(row.get("代码")).upper()
            if code not in required_codes:
                continue
            current_date = parse_date(row.get("日期"))
            key = (current_date, code) if current_date is not None else None
            if key not in required:
                continue
            vwap = positive_float(row.get("VWAP"))
            if vwap is not None:
                prices[key] = vwap
    return prices


def spearman_rank_ic(scores: list[float], returns: list[float]) -> float | None:
    if len(scores) < 2:
        return None
    score_ranks = pd.Series(scores, dtype="float64").rank(method="average")
    return_ranks = pd.Series(returns, dtype="float64").rank(method="average")
    if score_ranks.nunique() < 2 or return_ranks.nunique() < 2:
        return None
    rank_ic = score_ranks.corr(return_ranks)
    return float(rank_ic) if pd.notna(rank_ic) and math.isfinite(rank_ic) else None


def calculate_daily_rank_ic(
    signals: pd.DataFrame,
    mapping: dict[tuple[date, str], CandidateETF],
    schedule: dict[tuple[date, int], tuple[date, date]],
    prices: dict[tuple[date, str], float],
) -> pd.DataFrame:
    records: list[dict[str, object]] = []
    group_columns = [
        "threshold",
        "trend_window",
        "signal_date",
        "score_method",
    ]

    for (
        threshold,
        trend_window,
        signal_date,
        score_method,
    ), group in signals.groupby(group_columns, sort=True):
        planned_count = max(1, math.ceil(len(group) * TOP_PERCENT))
        group = group.sort_values(
            ["score", "index_code"],
            ascending=[False, True],
        ).head(planned_count)
        group = group[group["window_return"] > MIN_WINDOW_RETURN]
        for holding_day in HOLDING_DAYS:
            dates = schedule.get((signal_date, holding_day))
            if dates is None:
                continue
            entry_date, exit_date = dates
            scores: list[float] = []
            future_returns: list[float] = []

            for row in group[["index_code", "score"]].itertuples(index=False):
                candidate = mapping.get((signal_date, row.index_code))
                if candidate is None:
                    continue
                entry_vwap = prices.get((entry_date, candidate.code))
                exit_vwap = prices.get((exit_date, candidate.code))
                if entry_vwap is None or exit_vwap is None:
                    continue
                scores.append(float(row.score))
                future_returns.append(exit_vwap / entry_vwap - 1.0)

            rank_ic = spearman_rank_ic(scores, future_returns)
            if rank_ic is None:
                continue
            records.append(
                {
                    "信号日": signal_date,
                    "买入日": entry_date,
                    "卖出日": exit_date,
                    "聚类阈值": threshold,
                    "趋势窗口": int(trend_window),
                    "得分公式": score_method,
                    "持有期(交易日)": holding_day,
                    "RankIC": rank_ic,
                }
            )

    daily = pd.DataFrame.from_records(records)
    if daily.empty:
        raise ValueError(
            "没有得到有效Rank IC，请检查因子日期、指数代码映射及ETF的VWAP数据"
        )
    return daily.sort_values(
        [
            "聚类阈值",
            "趋势窗口",
            "得分公式",
            "持有期(交易日)",
            "信号日",
        ]
    ).reset_index(drop=True)


def build_summary(daily: pd.DataFrame) -> pd.DataFrame:
    return (
        daily.groupby(
            ["聚类阈值", "趋势窗口", "持有期(交易日)"],
            sort=True,
        )
        .agg(RankIC均值=("RankIC", "mean"))
        .reset_index()
    )


def set_sheet_style(worksheet, frame: pd.DataFrame) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column_number, column_name in enumerate(frame.columns, start=1):
        values = frame[column_name].astype(str)
        width = min(max(len(str(column_name)), values.map(len).max()) + 2, 24)
        worksheet.column_dimensions[get_column_letter(column_number)].width = width


def write_output(daily: pd.DataFrame) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for score_method, output_file in OUTPUT_FILES.items():
        formula_daily = (
            daily.loc[daily["得分公式"].eq(score_method)]
            .drop(columns="得分公式")
            .reset_index(drop=True)
        )
        if formula_daily.empty:
            raise ValueError(f"没有得到{score_method}的有效Rank IC")

        summary = build_summary(formula_daily)
        with pd.ExcelWriter(
            output_file,
            engine="openpyxl",
            date_format="yyyy-mm-dd",
        ) as writer:
            summary.to_excel(writer, sheet_name="汇总", index=False)
            formula_daily.to_excel(
                writer,
                sheet_name="逐日RankIC",
                index=False,
            )
            set_sheet_style(writer.sheets["汇总"], summary)
            set_sheet_style(writer.sheets["逐日RankIC"], formula_daily)


def main() -> None:
    print("读取趋势因子……")
    signals = load_factor_signals()
    print("扫描ETF数据并建立信号日映射……")
    mapping, trading_dates = scan_etf_mapping(signals)
    schedule = build_trade_schedule(set(signals["signal_date"]), trading_dates)
    required_prices = collect_required_prices(signals, mapping, schedule)
    print("读取计算未来收益所需的VWAP……")
    prices = load_required_vwap(required_prices)
    print("比较两种得分前10%指数在1至20个交易日的逐日Rank IC……")
    daily = calculate_daily_rank_ic(signals, mapping, schedule, prices)
    write_output(daily)
    print(f"完成：{OUTPUT_DIR}")


if __name__ == "__main__":
    main()
