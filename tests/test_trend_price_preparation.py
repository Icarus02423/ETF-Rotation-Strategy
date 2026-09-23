import csv
from datetime import date, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook

from scripts.ETF趋势策略回测 import 趋势行情准备


def write_pool_workbook(
    path: Path,
    selection_date: date,
    index_code: str,
    index_name: str,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 趋势行情准备.POOL_SHEET_NAME
    sheet.append(["日期", "对标指数", "对标指数代码"])
    sheet.append(
        [
            selection_date.isoformat(),
            index_name,
            index_code,
        ]
    )
    workbook.save(path)
    workbook.close()


def build_pool(
    output_file: Path,
    selection_date: date = date(2021, 3, 31),
    end_date: date = date(2021, 3, 31),
) -> 趋势行情准备.TrendPoolInput:
    return 趋势行情准备.TrendPoolInput(
        selection_date=selection_date,
        end_date=end_date,
        index_names={"000001.SH": "指数一"},
        output_file=output_file,
    )


def write_complete_output(pool: 趋势行情准备.TrendPoolInput) -> None:
    pool.output_file.parent.mkdir(parents=True, exist_ok=True)
    with pool.output_file.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        writer = csv.writer(handle)
        writer.writerow(趋势行情准备.OUTPUT_COLUMNS)
        first_date = pool.selection_date - timedelta(days=60)
        for offset in range(61):
            writer.writerow(
                [
                    pool.selection_date.isoformat(),
                    (first_date + timedelta(days=offset)).isoformat(),
                    "000001.SH",
                    "指数一",
                    format(100.0 + offset, ".15g"),
                    "0.001",
                ]
            )


class TrendPricePreparationTests(unittest.TestCase):
    def test_output_columns_remain_unchanged(self) -> None:
        self.assertEqual(
            趋势行情准备.OUTPUT_COLUMNS,
            [
                "月末交易日",
                "收益日期",
                "对标指数代码",
                "对标指数",
                "收盘价",
                "日收益率",
            ],
        )

    def test_discovery_is_limited_to_requested_frequency_directory(
        self,
    ) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            daily_reports = root / "daily" / "reports"
            monthly_reports = root / "monthly" / "reports"
            daily_prices = root / "daily" / "index_prices"
            monthly_prices = root / "monthly" / "index_prices"
            first_date = date(2021, 1, 29)
            next_date = date(2021, 2, 1)

            write_pool_workbook(
                daily_reports / "2021_01_29.xlsx",
                first_date,
                "000001.SH",
                "日度指数一",
            )
            write_pool_workbook(
                daily_reports / "2021_02_01.xlsx",
                next_date,
                "000002.SH",
                "日度指数二",
            )
            write_pool_workbook(
                monthly_reports / "2021_01_29.xlsx",
                first_date,
                "000003.SH",
                "月度指数",
            )

            daily_pools = 趋势行情准备.discover_trend_pool_inputs(
                daily_reports,
                daily_prices,
            )
            monthly_pools = 趋势行情准备.discover_trend_pool_inputs(
                monthly_reports,
                monthly_prices,
            )

            self.assertEqual(len(daily_pools), 2)
            self.assertEqual(daily_pools[0].end_date, next_date)
            self.assertEqual(daily_pools[1].end_date, next_date)
            self.assertEqual(
                daily_pools[0].output_file,
                daily_prices / "2021_01_29.csv",
            )
            self.assertEqual(
                set(daily_pools[0].index_names),
                {"000001.SH"},
            )
            self.assertEqual(len(monthly_pools), 1)
            self.assertEqual(monthly_pools[0].end_date, first_date)
            self.assertEqual(
                monthly_pools[0].output_file,
                monthly_prices / "2021_01_29.csv",
            )
            self.assertEqual(
                set(monthly_pools[0].index_names),
                {"000003.SH"},
            )

    def test_complete_output_still_skips_downloader(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            output_directory = Path(temporary_directory) / "index_prices"
            pool = build_pool(output_directory / "2021_03_31.csv")
            write_complete_output(pool)

            with patch.object(
                趋势行情准备,
                "load_index_downloader_module",
                side_effect=AssertionError("完整输出不应加载下载模块"),
            ), patch("builtins.print"):
                趋势行情准备.download_trend_index_data(
                    [pool],
                    output_directory,
                )

            self.assertTrue(pool.output_file.exists())
            self.assertTrue(趋势行情准备.output_is_complete(pool))

    def test_cleanup_is_limited_to_requested_output_directory(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            target = root / "daily" / "threshold_0.7" / "index_prices"
            other_frequency = (
                root / "monthly" / "threshold_0.7" / "index_prices"
            )
            other_threshold = (
                root / "daily" / "threshold_0.9" / "index_prices"
            )
            for directory in (target, other_frequency, other_threshold):
                directory.mkdir(parents=True)

            keep = target / "keep.csv"
            stale = target / "stale.csv"
            note = target / "说明.txt"
            other_frequency_file = other_frequency / "stale.csv"
            other_threshold_file = other_threshold / "stale.csv"
            for path in (
                keep,
                stale,
                note,
                other_frequency_file,
                other_threshold_file,
            ):
                path.write_text("test", encoding="utf-8")

            removed_count = 趋势行情准备.remove_stale_outputs(
                target,
                {keep.name},
            )

            self.assertEqual(removed_count, 1)
            self.assertTrue(keep.exists())
            self.assertFalse(stale.exists())
            self.assertTrue(note.exists())
            self.assertTrue(other_frequency_file.exists())
            self.assertTrue(other_threshold_file.exists())

    def test_main_routes_frequency_and_threshold_to_paths(self) -> None:
        cluster_directory = Path("monthly_cluster_reports")
        output_directory = Path("monthly_index_prices")
        with patch.object(
            趋势行情准备,
            "cluster_report_dir",
            return_value=cluster_directory,
        ) as cluster_path, patch.object(
            趋势行情准备,
            "index_price_dir",
            return_value=output_directory,
        ) as output_path, patch.object(
            趋势行情准备,
            "discover_trend_pool_inputs",
            return_value=[],
        ) as discover, patch.object(
            趋势行情准备,
            "download_trend_index_data",
        ) as download, patch("builtins.print"):
            趋势行情准备.main("monthly", 0.7)

        cluster_path.assert_called_once_with("monthly", 0.7)
        output_path.assert_called_once_with("monthly", 0.7)
        discover.assert_called_once_with(
            cluster_directory,
            output_directory,
        )
        download.assert_called_once_with([], output_directory)

    def test_invalid_parameters_fail_before_path_resolution(self) -> None:
        invalid_cases = (("weekly", 0.9), ("daily", 0.75))
        for update_frequency, correlation_threshold in invalid_cases:
            with self.subTest(
                update_frequency=update_frequency,
                correlation_threshold=correlation_threshold,
            ), patch.object(
                趋势行情准备,
                "cluster_report_dir",
                side_effect=AssertionError("不应计算聚类路径"),
            ), patch.object(
                趋势行情准备,
                "index_price_dir",
                side_effect=AssertionError("不应计算行情路径"),
            ), patch.object(
                趋势行情准备,
                "load_index_downloader_module",
                side_effect=AssertionError("不应加载下载模块"),
            ):
                with self.assertRaises(ValueError):
                    趋势行情准备.main(
                        update_frequency,
                        correlation_threshold,
                    )


if __name__ == "__main__":
    unittest.main()
