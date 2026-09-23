from datetime import date, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from scripts.ETF池筛选 import ETF初筛


def consecutive_dates(start: date, end: date) -> list[date]:
    return [
        start + timedelta(days=offset)
        for offset in range((end - start).days + 1)
    ]


class ETFInitialSelectionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.available_dates = consecutive_dates(
            date(2020, 12, 1),
            date(2021, 2, 28),
        )

    def test_daily_uses_every_eligible_date(self) -> None:
        plans = ETF初筛.build_selection_plans(
            self.available_dates,
            "daily",
        )
        expected_dates = consecutive_dates(
            date(2021, 1, 1),
            date(2021, 2, 28),
        )
        self.assertEqual(
            [plan.selection_date for plan in plans],
            expected_dates,
        )

    def test_monthly_uses_last_available_date_of_each_month(self) -> None:
        plans = ETF初筛.build_selection_plans(
            self.available_dates,
            "monthly",
        )
        self.assertEqual(
            [plan.selection_date for plan in plans],
            [date(2021, 1, 31), date(2021, 2, 28)],
        )

        for plan in plans:
            self.assertEqual(len(plan.turnover_dates), 20)
            self.assertEqual(plan.turnover_dates[-1], plan.selection_date)
            self.assertTrue(
                all(
                    turnover_date <= plan.selection_date
                    for turnover_date in plan.turnover_dates
                )
            )

    def test_invalid_update_frequency_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            ETF初筛.build_selection_plans(
                self.available_dates,
                "weekly",
            )

    def test_output_cleanup_is_limited_to_one_frequency(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            daily_dir = root / "daily" / "initial"
            monthly_dir = root / "monthly" / "initial"
            daily_dir.mkdir(parents=True)
            monthly_dir.mkdir(parents=True)

            daily_csv = daily_dir / "daily.csv"
            daily_note = daily_dir / "保留说明.txt"
            monthly_csv = monthly_dir / "monthly.csv"
            daily_csv.write_text("daily", encoding="utf-8")
            daily_note.write_text("keep", encoding="utf-8")
            monthly_csv.write_text("monthly", encoding="utf-8")

            removed_count = ETF初筛.clear_previous_csv_outputs(daily_dir)

            self.assertEqual(removed_count, 1)
            self.assertFalse(daily_csv.exists())
            self.assertTrue(daily_note.exists())
            self.assertTrue(monthly_csv.exists())


if __name__ == "__main__":
    unittest.main()
