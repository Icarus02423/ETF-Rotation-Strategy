import csv
from datetime import date, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook

from scripts.ETF趋势策略回测 import 趋势因子计算


def write_pool_workbook(
    path: Path,
    selection_date: date,
    etf_code: str,
    index_code: str,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 趋势因子计算.POOL_SHEET_NAME
    sheet.append(
        [
            "日期",
            "代码",
            "名称",
            "对标指数",
            "对标指数代码",
            "过去20个交易日平均成交额",
        ]
    )
    sheet.append(
        [
            selection_date.isoformat(),
            etf_code,
            f"ETF{etf_code}",
            f"指数{index_code}",
            index_code,
            1000000,
        ]
    )
    workbook.save(path)
    workbook.close()


def write_index_prices(
    path: Path,
    index_code: str,
    price_date: date,
    close: float,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "月末交易日",
                "收益日期",
                "对标指数代码",
                "对标指数",
                "收盘价",
                "日收益率",
            ],
        )
        writer.writeheader()
        writer.writerow(
            {
                "月末交易日": price_date.isoformat(),
                "收益日期": price_date.isoformat(),
                "对标指数代码": index_code,
                "对标指数": f"指数{index_code}",
                "收盘价": close,
                "日收益率": 0.001,
            }
        )


def build_member(
    etf_code: str = "510001.SH",
    index_code: str = "000001.SH",
) -> 趋势因子计算.PoolMember:
    return 趋势因子计算.PoolMember(
        etf_code=etf_code,
        etf_name=f"ETF{etf_code}",
        index_code=index_code,
        index_name=f"指数{index_code}",
    )


class TrendFactorCalculationTests(unittest.TestCase):
    def test_output_columns_remain_unchanged(self) -> None:
        self.assertEqual(
            趋势因子计算.OUTPUT_COLUMNS,
            [
                "日期",
                "指数池日期",
                "ETF代码",
                "ETF名称",
                "对标指数代码",
                "对标指数",
                "趋势窗口",
                "窗口起始日",
                "窗口结束日",
                "窗口起始收盘价",
                "窗口结束收盘价",
                "log价格回归斜率",
                "窗口收益率",
                "窗口波动率",
                "R平方",
                "趋势质量因子",
                "风险调整趋势得分",
                "因子排名",
                "排名百分比",
            ],
        )

    def test_all_windows_use_exact_history_and_ignore_future_prices(
        self,
    ) -> None:
        expected_windows = (10, 15, 20, 40, 60)
        self.assertEqual(趋势因子计算.TREND_WINDOWS, expected_windows)
        price_dates = [
            date(2021, 1, 1) + timedelta(days=offset)
            for offset in range(80)
        ]
        prices = [
            100.0 + 0.4 * offset + 0.03 * (offset % 5) ** 2
            for offset in range(80)
        ]
        current_date = price_dates[69]
        prices_with_extreme_future = prices[:70] + [1000000.0] * 10

        for trend_window in expected_windows:
            with self.subTest(trend_window=trend_window):
                趋势因子计算.validate_parameters(
                    "daily",
                    0.9,
                    trend_window,
                )
                result, status = 趋势因子计算.calculate_trend(
                    current_date,
                    price_dates[:70],
                    prices[:70],
                    trend_window,
                )
                future_result, future_status = (
                    趋势因子计算.calculate_trend(
                        current_date,
                        price_dates,
                        prices_with_extreme_future,
                        trend_window,
                    )
                )

                self.assertEqual(status, "有效")
                self.assertEqual(future_status, "有效")
                self.assertIsNotNone(result)
                self.assertEqual(result, future_result)
                assert result is not None
                self.assertEqual(
                    result.window_start_date,
                    price_dates[70 - trend_window],
                )
                self.assertEqual(result.window_end_date, current_date)
                self.assertAlmostEqual(
                    result.trend_factor,
                    result.window_return * result.r_squared,
                )
                self.assertAlmostEqual(
                    result.risk_adjusted_trend_factor,
                    result.window_return / result.window_volatility,
                )

    def test_daily_row_contains_both_factor_values(self) -> None:
        current_date = date(2021, 3, 31)
        price_dates = [
            current_date - timedelta(days=24 - offset)
            for offset in range(25)
        ]
        prices = [
            100.0 + 0.5 * offset + 0.02 * (offset % 4) ** 2
            for offset in range(25)
        ]
        member = build_member()
        snapshot = 趋势因子计算.PoolSnapshot(
            selection_date=date(2021, 2, 26),
            members=(member,),
            source_file=Path("2021_02_26.xlsx"),
        )
        result, _ = 趋势因子计算.calculate_trend(
            current_date,
            price_dates,
            prices,
            20,
        )

        rows = 趋势因子计算.build_daily_rows(
            current_date,
            snapshot,
            {member.index_code: price_dates},
            {member.index_code: prices},
            20,
        )

        self.assertIsNotNone(result)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["趋势窗口"], 20)
        assert result is not None
        self.assertAlmostEqual(
            float(rows[0]["趋势质量因子"]),
            result.trend_factor,
        )
        self.assertAlmostEqual(
            float(rows[0]["风险调整趋势得分"]),
            result.risk_adjusted_trend_factor,
        )

    def test_daily_snapshot_becomes_active_next_trading_day(self) -> None:
        friday = 趋势因子计算.PoolSnapshot(
            selection_date=date(2021, 1, 29),
            members=(build_member(),),
            source_file=Path("2021_01_29.xlsx"),
        )
        monday = 趋势因子计算.PoolSnapshot(
            selection_date=date(2021, 2, 1),
            members=(build_member(index_code="000002.SH"),),
            source_file=Path("2021_02_01.xlsx"),
        )
        snapshots = [friday, monday]
        snapshot_dates = [item.selection_date for item in snapshots]

        self.assertIsNone(
            趋势因子计算.active_pool_snapshot(
                friday.selection_date,
                snapshots,
                snapshot_dates,
            )
        )
        self.assertEqual(
            趋势因子计算.active_pool_snapshot(
                monday.selection_date,
                snapshots,
                snapshot_dates,
            ),
            friday,
        )
        self.assertEqual(
            趋势因子计算.active_pool_snapshot(
                date(2021, 2, 2),
                snapshots,
                snapshot_dates,
            ),
            monday,
        )

    def test_monthly_snapshot_stays_active_until_next_one_takes_effect(
        self,
    ) -> None:
        january = 趋势因子计算.PoolSnapshot(
            selection_date=date(2021, 1, 29),
            members=(build_member(),),
            source_file=Path("2021_01_29.xlsx"),
        )
        february = 趋势因子计算.PoolSnapshot(
            selection_date=date(2021, 2, 26),
            members=(build_member(index_code="000002.SH"),),
            source_file=Path("2021_02_26.xlsx"),
        )
        snapshots = [january, february]
        snapshot_dates = [item.selection_date for item in snapshots]

        self.assertEqual(
            趋势因子计算.active_pool_snapshot(
                date(2021, 2, 1),
                snapshots,
                snapshot_dates,
            ),
            january,
        )
        self.assertEqual(
            趋势因子计算.active_pool_snapshot(
                february.selection_date,
                snapshots,
                snapshot_dates,
            ),
            january,
        )
        self.assertEqual(
            趋势因子计算.active_pool_snapshot(
                date(2021, 3, 1),
                snapshots,
                snapshot_dates,
            ),
            february,
        )

    def test_pool_discovery_is_limited_to_requested_directory(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            daily = root / "daily" / "reports"
            monthly = root / "monthly" / "reports"
            selection_date = date(2021, 1, 29)
            file_name = "2021_01_29.xlsx"
            write_pool_workbook(
                daily / file_name,
                selection_date,
                "510001.SH",
                "000001.SH",
            )
            write_pool_workbook(
                monthly / file_name,
                selection_date,
                "510002.SH",
                "000002.SH",
            )

            daily_snapshots = 趋势因子计算.discover_pool_snapshots(daily)
            monthly_snapshots = 趋势因子计算.discover_pool_snapshots(
                monthly
            )

            self.assertEqual(
                daily_snapshots[0].members[0].index_code,
                "000001.SH",
            )
            self.assertEqual(
                monthly_snapshots[0].members[0].index_code,
                "000002.SH",
            )

    def test_index_prices_are_limited_to_requested_directory(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            daily = root / "daily" / "index_prices"
            monthly = root / "monthly" / "index_prices"
            price_date = date(2021, 1, 29)
            write_index_prices(
                daily / "2021_01_29.csv",
                "000001.SH",
                price_date,
                101.0,
            )
            write_index_prices(
                monthly / "2021_01_29.csv",
                "000002.SH",
                price_date,
                202.0,
            )
            requested_codes = {"000001.SH", "000002.SH"}

            daily_dates, daily_prices, _ = (
                趋势因子计算.load_index_close_history(
                    requested_codes,
                    daily,
                )
            )
            monthly_dates, monthly_prices, _ = (
                趋势因子计算.load_index_close_history(
                    requested_codes,
                    monthly,
                )
            )

            self.assertEqual(daily_dates["000001.SH"], [price_date])
            self.assertEqual(daily_prices["000001.SH"], [101.0])
            self.assertEqual(daily_dates["000002.SH"], [])
            self.assertEqual(monthly_dates["000001.SH"], [])
            self.assertEqual(monthly_dates["000002.SH"], [price_date])
            self.assertEqual(monthly_prices["000002.SH"], [202.0])

    def test_cleanup_is_limited_to_requested_output_directory(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            target = root / "daily" / "threshold_0.8" / "window_20"
            other_frequency = (
                root / "monthly" / "threshold_0.8" / "window_20"
            )
            other_threshold = (
                root / "daily" / "threshold_0.9" / "window_20"
            )
            other_window = (
                root / "daily" / "threshold_0.8" / "window_40"
            )
            for directory in (
                target,
                other_frequency,
                other_threshold,
                other_window,
            ):
                directory.mkdir(parents=True)

            keep = target / "2021.csv"
            stale = target / "2020.csv"
            note = target / "说明.txt"
            outside_files = [
                other_frequency / "2020.csv",
                other_threshold / "2020.csv",
                other_window / "2020.csv",
            ]
            for path in (keep, stale, note, *outside_files):
                path.write_text("test", encoding="utf-8")

            removed_count = 趋势因子计算.remove_stale_outputs(
                target,
                {keep.name},
            )

            self.assertEqual(removed_count, 1)
            self.assertTrue(keep.exists())
            self.assertFalse(stale.exists())
            self.assertTrue(note.exists())
            self.assertTrue(all(path.exists() for path in outside_files))

    def test_main_routes_frequency_threshold_and_window_to_paths(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            cluster_directory = root / "monthly_cluster_reports"
            index_price_directory = root / "monthly_index_prices"
            output_directory = root / "monthly_factors"
            member = build_member()
            previous_snapshot = 趋势因子计算.PoolSnapshot(
                selection_date=date(2021, 1, 28),
                members=(member,),
                source_file=cluster_directory / "2021_01_28.xlsx",
            )
            latest_snapshot = 趋势因子计算.PoolSnapshot(
                selection_date=date(2021, 1, 29),
                members=(member,),
                source_file=cluster_directory / "2021_01_29.xlsx",
            )

            def fake_write_year_output(
                directory: Path,
                year: int,
                rows: object,
            ) -> Path:
                self.assertEqual(directory, output_directory)
                self.assertEqual(rows, [])
                return directory / f"{year}.csv"

            with patch.object(
                趋势因子计算,
                "cluster_report_dir",
                return_value=cluster_directory,
            ) as cluster_path, patch.object(
                趋势因子计算,
                "index_price_dir",
                return_value=index_price_directory,
            ) as price_path, patch.object(
                趋势因子计算,
                "factor_dir",
                return_value=output_directory,
            ) as output_path, patch.object(
                趋势因子计算,
                "discover_pool_snapshots",
                return_value=[previous_snapshot, latest_snapshot],
            ) as discover, patch.object(
                趋势因子计算,
                "read_etf_trading_calendar",
                return_value=[latest_snapshot.selection_date],
            ), patch.object(
                趋势因子计算,
                "load_index_close_history",
                return_value=(
                    {member.index_code: []},
                    {member.index_code: []},
                    {},
                ),
            ) as load_prices, patch.object(
                趋势因子计算,
                "build_daily_rows",
                return_value=[],
            ) as build_rows, patch.object(
                趋势因子计算,
                "write_year_output",
                side_effect=fake_write_year_output,
            ) as write_output, patch.object(
                趋势因子计算,
                "remove_stale_outputs",
            ) as cleanup, patch("builtins.print"):
                趋势因子计算.main("monthly", 0.8, 40)

            cluster_path.assert_called_once_with("monthly", 0.8)
            price_path.assert_called_once_with("monthly", 0.8)
            output_path.assert_called_once_with("monthly", 0.8, 40)
            discover.assert_called_once_with(cluster_directory)
            load_prices.assert_called_once_with(
                {member.index_code},
                index_price_directory,
            )
            build_rows.assert_called_once_with(
                latest_snapshot.selection_date,
                previous_snapshot,
                {member.index_code: []},
                {member.index_code: []},
                40,
            )
            self.assertEqual(write_output.call_count, 6)
            cleanup.assert_called_once_with(
                output_directory,
                {f"{year}.csv" for year in range(2021, 2027)},
            )

    def test_invalid_parameters_fail_before_path_resolution(self) -> None:
        invalid_cases = (
            ("weekly", 0.9, 20),
            ("daily", 0.75, 20),
            ("daily", 0.9, 30),
        )
        for update_frequency, correlation_threshold, trend_window in (
            invalid_cases
        ):
            with self.subTest(
                update_frequency=update_frequency,
                correlation_threshold=correlation_threshold,
                trend_window=trend_window,
            ), patch.object(
                趋势因子计算,
                "cluster_report_dir",
                side_effect=AssertionError("不应计算聚类路径"),
            ), patch.object(
                趋势因子计算,
                "index_price_dir",
                side_effect=AssertionError("不应计算行情路径"),
            ), patch.object(
                趋势因子计算,
                "factor_dir",
                side_effect=AssertionError("不应计算因子路径"),
            ):
                with self.assertRaises(ValueError):
                    趋势因子计算.main(
                        update_frequency,
                        correlation_threshold,
                        trend_window,
                    )


if __name__ == "__main__":
    unittest.main()
