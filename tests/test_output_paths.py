from pathlib import Path
import unittest

from 实验配置 import ExperimentCase
from 输出路径 import (
    PROJECT_ROOT,
    backtest_result_dir,
    cluster_report_dir,
    etf_pool_initial_dir,
    factor_dir,
    index_price_dir,
    index_return_dir,
    tuning_result_dir,
)


def relative_path(path: Path) -> str:
    return path.relative_to(PROJECT_ROOT).as_posix()


class OutputPathTests(unittest.TestCase):
    def test_etf_pool_paths(self) -> None:
        self.assertEqual(
            relative_path(etf_pool_initial_dir("daily")),
            "outputs/etf_pool/daily/initial",
        )
        self.assertEqual(
            relative_path(index_return_dir("monthly")),
            "outputs/etf_pool/monthly/index_returns/window_60",
        )
        self.assertEqual(
            relative_path(cluster_report_dir("daily", 0.7)),
            "outputs/etf_pool/daily/clusters/threshold_0.7/reports",
        )

    def test_trend_data_paths(self) -> None:
        self.assertEqual(
            relative_path(index_price_dir("monthly", 0.9)),
            "outputs/etf_trend_strategy/monthly/threshold_0.9/index_prices",
        )
        self.assertEqual(
            relative_path(factor_dir("daily", 0.8, 40)),
            "outputs/etf_trend_strategy/daily/threshold_0.8/factors/window_40",
        )

    def test_backtest_path(self) -> None:
        experiment_case = ExperimentCase(
            update_frequency="daily",
            filter_profile="bias_only",
            correlation_threshold=0.8,
            trend_factor="return_r2",
            trend_window=40,
            stagger_days=3,
        )
        self.assertEqual(
            relative_path(backtest_result_dir(experiment_case, "next_day_vwap")),
            (
                "outputs/etf_trend_strategy/daily/threshold_0.8/backtest/"
                "bias_only/window_40/return_r2/staggered_3d/next_day_vwap"
            ),
        )

    def test_tuning_paths(self) -> None:
        self.assertEqual(
            relative_path(
                tuning_result_dir(
                    "monthly",
                    "bias_and_volatility",
                    "bias_parameter_grid",
                    0.9,
                    15,
                    4,
                )
            ),
            (
                "outputs/parameter_grid/monthly/bias_and_volatility/"
                "bias_parameter_grid/threshold_0.9/window_15/staggered_4d"
            ),
        )
        self.assertEqual(
            relative_path(
                tuning_result_dir(
                    "daily",
                    "bias_and_volatility",
                    "volatility_parameter_grid",
                    0.7,
                    20,
                    10,
                )
            ),
            (
                "outputs/parameter_grid/daily/bias_and_volatility/"
                "volatility_parameter_grid/threshold_0.7/window_20/staggered_10d"
            ),
        )

    def test_invalid_tuning_combinations_are_rejected(self) -> None:
        invalid_combinations = (
            ("base", "bias_parameter_grid"),
            ("base", "volatility_parameter_grid"),
            ("bias_only", "volatility_parameter_grid"),
            ("volatility_only", "bias_parameter_grid"),
        )
        for filter_profile, tuning_type in invalid_combinations:
            with self.subTest(
                filter_profile=filter_profile,
                tuning_type=tuning_type,
            ):
                with self.assertRaises(ValueError):
                    tuning_result_dir(
                        "daily",
                        filter_profile,
                        tuning_type,
                        0.9,
                        15,
                        4,
                    )


if __name__ == "__main__":
    unittest.main()
