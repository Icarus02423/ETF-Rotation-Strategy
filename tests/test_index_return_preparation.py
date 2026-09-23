import csv
from datetime import date, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from scripts.ETF池筛选 import 指数收益率准备


def write_initial_selection(
    path: Path,
    selection_date: date,
    index_names: dict[str, str],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["日期", "对标指数", "对标指数代码"],
        )
        writer.writeheader()
        for index_code, index_name in index_names.items():
            writer.writerow(
                {
                    "日期": selection_date.isoformat(),
                    "对标指数": index_name,
                    "对标指数代码": index_code,
                }
            )


def build_close_history(
    selection_date: date,
    index_names: dict[str, str],
) -> dict[str, dict[date, float]]:
    first_date = selection_date - timedelta(days=60)
    close_dates = [
        first_date + timedelta(days=offset)
        for offset in range(61)
    ]
    closes_by_code: dict[str, dict[date, float]] = {}
    for code_position, index_code in enumerate(sorted(index_names)):
        base_close = 100.0 + code_position * 100.0
        closes_by_code[index_code] = {
            close_date: base_close + date_position
            for date_position, close_date in enumerate(close_dates)
        }
        closes_by_code[index_code][selection_date + timedelta(days=1)] = 1.0
    return closes_by_code


class IndexReturnPreparationTests(unittest.TestCase):
    def test_discovery_is_limited_to_requested_directory(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            daily_initial = root / "daily" / "initial"
            monthly_initial = root / "monthly" / "initial"
            daily_output = root / "daily" / "index_returns" / "window_60"
            monthly_output = root / "monthly" / "index_returns" / "window_60"
            file_name = "snapshot.csv"
            expected_output_name = "2021_01_29.csv"

            write_initial_selection(
                daily_initial / file_name,
                date(2021, 1, 29),
                {"000001.SH": "日度指数"},
            )
            write_initial_selection(
                monthly_initial / file_name,
                date(2021, 1, 29),
                {"000002.SH": "月度指数"},
            )
            (daily_initial / "忽略说明.txt").write_text(
                "不是输入CSV",
                encoding="utf-8",
            )

            daily_selections = 指数收益率准备.discover_selection_inputs(
                daily_initial,
                daily_output,
            )
            monthly_selections = 指数收益率准备.discover_selection_inputs(
                monthly_initial,
                monthly_output,
            )

            self.assertEqual(len(daily_selections), 1)
            self.assertEqual(len(monthly_selections), 1)
            self.assertEqual(
                set(daily_selections[0].index_names),
                {"000001.SH"},
            )
            self.assertEqual(
                set(monthly_selections[0].index_names),
                {"000002.SH"},
            )
            self.assertEqual(
                daily_selections[0].output_file,
                daily_output / expected_output_name,
            )
            self.assertEqual(
                monthly_selections[0].output_file,
                monthly_output / expected_output_name,
            )

    def test_return_rows_use_61_common_closes_without_future_data(self) -> None:
        selection_date = date(2021, 3, 2)
        index_names = {
            "000001.SH": "指数一",
            "000002.SH": "指数二",
        }
        selection = 指数收益率准备.SelectionInput(
            selection_date=selection_date,
            index_names=index_names,
            source_file=Path("input.csv"),
            output_file=Path("output.csv"),
        )
        closes_by_code = build_close_history(selection_date, index_names)

        rows, insufficient_codes = 指数收益率准备.selection_return_rows(
            selection,
            closes_by_code,
        )

        self.assertEqual(insufficient_codes, [])
        self.assertEqual(len(rows), 120)
        self.assertEqual(
            max(date.fromisoformat(row[1]) for row in rows),
            selection_date,
        )
        self.assertTrue(
            all(date.fromisoformat(row[1]) <= selection_date for row in rows)
        )
        first_row = rows[0]
        self.assertAlmostEqual(float(first_row[5]), 101.0 / 100.0 - 1.0)

    def test_complete_output_requires_full_non_future_data(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            selection_date = date(2021, 3, 2)
            index_names = {
                "000001.SH": "指数一",
                "000002.SH": "指数二",
            }
            selection = 指数收益率准备.SelectionInput(
                selection_date=selection_date,
                index_names=index_names,
                source_file=root / "input.csv",
                output_file=root / "output.csv",
            )
            rows, insufficient_codes = 指数收益率准备.selection_return_rows(
                selection,
                build_close_history(selection_date, index_names),
            )
            self.assertEqual(insufficient_codes, [])

            指数收益率准备.write_selection_output(selection.output_file, rows)
            self.assertTrue(
                指数收益率准备.selection_output_is_complete(selection)
            )

            指数收益率准备.write_selection_output(
                selection.output_file,
                rows[:-1],
            )
            self.assertFalse(
                指数收益率准备.selection_output_is_complete(selection)
            )

            future_rows = [list(row) for row in rows]
            future_rows[-1][1] = (selection_date + timedelta(days=1)).isoformat()
            指数收益率准备.write_selection_output(
                selection.output_file,
                future_rows,
            )
            self.assertFalse(
                指数收益率准备.selection_output_is_complete(selection)
            )

    def test_complete_output_skips_downloader(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            initial_dir = root / "initial"
            output_dir = root / "index_returns" / "window_60"
            output_dir.mkdir(parents=True)
            selection_date = date(2021, 3, 2)
            index_names = {
                "000001.SH": "指数一",
                "000002.SH": "指数二",
            }
            file_name = selection_date.strftime("%Y_%m_%d.csv")
            write_initial_selection(
                initial_dir / file_name,
                selection_date,
                index_names,
            )
            selection = 指数收益率准备.SelectionInput(
                selection_date=selection_date,
                index_names=index_names,
                source_file=initial_dir / file_name,
                output_file=output_dir / file_name,
            )
            rows, _ = 指数收益率准备.selection_return_rows(
                selection,
                build_close_history(selection_date, index_names),
            )
            指数收益率准备.write_selection_output(selection.output_file, rows)

            with patch.object(
                指数收益率准备,
                "etf_pool_initial_dir",
                return_value=initial_dir,
            ), patch.object(
                指数收益率准备,
                "index_return_dir",
                return_value=output_dir,
            ), patch.object(
                指数收益率准备,
                "load_etf_downloader_module",
                side_effect=AssertionError("不应调用下载器"),
            ):
                指数收益率准备.main("daily")

    def test_invalid_frequency_fails_before_path_or_download_work(self) -> None:
        with patch.object(
            指数收益率准备,
            "etf_pool_initial_dir",
            side_effect=AssertionError("不应计算路径"),
        ), patch.object(
            指数收益率准备,
            "index_return_dir",
            side_effect=AssertionError("不应计算路径"),
        ), patch.object(
            指数收益率准备,
            "load_etf_downloader_module",
            side_effect=AssertionError("不应调用下载器"),
        ):
            with self.assertRaises(ValueError):
                指数收益率准备.main("weekly")


if __name__ == "__main__":
    unittest.main()
