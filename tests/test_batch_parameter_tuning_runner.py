from argparse import ArgumentParser
from collections import Counter
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from 实验配置 import (
    CORRELATION_THRESHOLDS,
    STAGGER_DAYS,
    TREND_WINDOWS,
    UPDATE_FREQUENCIES,
)
import 运行参数调优


def make_case(
    *,
    tuning_type: str = "bias",
    update_frequency: str = "monthly",
    filter_profile: str = "bias_only",
    correlation_threshold: float = 0.7,
    trend_window: int = 10,
    stagger_days: int = 1,
) -> 运行参数调优.TuningCase:
    return 运行参数调优.TuningCase(
        tuning_type=tuning_type,
        update_frequency=update_frequency,
        filter_profile=filter_profile,
        correlation_threshold=correlation_threshold,
        trend_window=trend_window,
        stagger_days=stagger_days,
    )


class BatchParameterTuningRunnerTests(unittest.TestCase):
    def test_full_matrix_has_1200_unique_compatible_cases(self) -> None:
        cases = tuple(运行参数调优.iter_tuning_cases())
        expected_cases = tuple(
            运行参数调优.TuningCase(
                tuning_type=tuning_type,
                update_frequency=frequency,
                filter_profile=profile,
                correlation_threshold=threshold,
                trend_window=window,
                stagger_days=stagger,
            )
            for tuning_type in ("bias", "volatility")
            for frequency in UPDATE_FREQUENCIES
            for profile in (
                ("bias_only", "bias_and_volatility")
                if tuning_type == "bias"
                else ("volatility_only", "bias_and_volatility")
            )
            for threshold in CORRELATION_THRESHOLDS
            for window in TREND_WINDOWS
            for stagger in STAGGER_DAYS
        )

        self.assertEqual(cases, expected_cases)
        self.assertEqual(len(cases), 1200)
        self.assertEqual(len(set(cases)), 1200)
        self.assertEqual(
            Counter(case.tuning_type for case in cases),
            {"bias": 600, "volatility": 600},
        )
        self.assertEqual(
            Counter(case.filter_profile for case in cases),
            {
                "bias_only": 300,
                "volatility_only": 300,
                "bias_and_volatility": 600,
            },
        )
        self.assertEqual(运行参数调优.simulation_count(cases), 7_221_600)

    def test_selectors_have_expected_counts_and_exact_case(self) -> None:
        cases = tuple(运行参数调优.iter_tuning_cases())
        self.assertEqual(
            len(
                运行参数调优.select_tuning_cases(
                    cases,
                    tuning_types=("bias",),
                )
            ),
            600,
        )
        self.assertEqual(
            len(
                运行参数调优.select_tuning_cases(
                    cases,
                    filter_profiles=("bias_and_volatility",),
                )
            ),
            600,
        )
        selected = 运行参数调优.select_tuning_cases(
            cases,
            tuning_types=("volatility",),
            update_frequencies=("daily",),
            filter_profiles=("volatility_only",),
            correlation_thresholds=(0.8,),
            trend_windows=(40,),
            stagger_days=(6,),
        )
        self.assertEqual(
            selected,
            (
                make_case(
                    tuning_type="volatility",
                    update_frequency="daily",
                    filter_profile="volatility_only",
                    correlation_threshold=0.8,
                    trend_window=40,
                    stagger_days=6,
                ),
            ),
        )

    def test_complete_case_requires_four_named_nonempty_pngs_and_no_marker(
        self,
    ) -> None:
        expected_names = {
            "return_r2_close_total_return_heatmap.png",
            "return_r2_next_day_vwap_total_return_heatmap.png",
            "return_vol_close_total_return_heatmap.png",
            "return_vol_next_day_vwap_total_return_heatmap.png",
        }
        for tuning_case in (
            make_case(),
            make_case(
                tuning_type="volatility",
                filter_profile="volatility_only",
            ),
        ):
            with self.subTest(tuning_type=tuning_case.tuning_type):
                with TemporaryDirectory() as temporary_directory:
                    output_directory = Path(temporary_directory)
                    with patch.object(
                        运行参数调优,
                        "tuning_output_directory",
                        return_value=output_directory,
                    ):
                        paths = 运行参数调优.expected_output_paths(
                            tuning_case
                        )
                        self.assertEqual(
                            {path.name for path in paths},
                            expected_names,
                        )
                        self.assertFalse(
                            运行参数调优.case_is_complete(tuning_case)
                        )
                        for path in paths:
                            path.write_bytes(b"png")
                        self.assertTrue(
                            运行参数调优.case_is_complete(tuning_case)
                        )

                        marker = 运行参数调优.incomplete_marker_path(
                            tuning_case
                        )
                        marker.touch()
                        self.assertFalse(
                            运行参数调优.case_is_complete(tuning_case)
                        )
                        marker.unlink()
                        paths[-1].write_bytes(b"")
                        self.assertFalse(
                            运行参数调优.case_is_complete(tuning_case)
                        )

    def test_run_tuning_case_routes_five_dimensions_to_correct_module(
        self,
    ) -> None:
        bias_case = make_case(
            update_frequency="daily",
            filter_profile="bias_and_volatility",
            correlation_threshold=0.9,
            trend_window=60,
            stagger_days=10,
        )
        volatility_case = make_case(
            tuning_type="volatility",
            update_frequency="daily",
            filter_profile="bias_and_volatility",
            correlation_threshold=0.8,
            trend_window=40,
            stagger_days=6,
        )
        with patch.object(
            运行参数调优.BIAS参数调优,
            "main",
        ) as bias_main, patch.object(
            运行参数调优.波动率参数调优,
            "main",
        ) as volatility_main:
            运行参数调优.run_tuning_case(bias_case)
            运行参数调优.run_tuning_case(volatility_case)

        bias_main.assert_called_once_with(
            update_frequency="daily",
            filter_profile="bias_and_volatility",
            correlation_threshold=0.9,
            trend_window=60,
            stagger_days=10,
        )
        volatility_main.assert_called_once_with(
            update_frequency="daily",
            filter_profile="bias_and_volatility",
            correlation_threshold=0.8,
            trend_window=40,
            stagger_days=6,
        )

    def test_default_and_preview_cli_never_execute_or_preflight(self) -> None:
        preview_commands = (
            (),
            ("--all",),
            (
                "--tuning-type",
                "bias",
                "--frequency",
                "monthly",
                "--profile",
                "bias_only",
            ),
        )
        for command in preview_commands:
            with self.subTest(command=command), patch.object(
                运行参数调优,
                "run_cases",
            ) as run_cases, patch.object(
                运行参数调优,
                "validate_required_inputs",
            ) as validate_inputs, patch.object(
                运行参数调优.BIAS参数调优,
                "main",
            ) as bias_main, patch.object(
                运行参数调优.波动率参数调优,
                "main",
            ) as volatility_main, patch.object(
                ArgumentParser,
                "print_help",
            ), patch("builtins.print"):
                result = 运行参数调优.cli(command)

            self.assertIsNone(result)
            run_cases.assert_not_called()
            validate_inputs.assert_not_called()
            bias_main.assert_not_called()
            volatility_main.assert_not_called()

    def test_cli_executes_one_exact_case_and_forwards_force(self) -> None:
        command = (
            "--tuning-type",
            "volatility",
            "--frequency",
            "daily",
            "--profile",
            "volatility_only",
            "--threshold",
            "0.8",
            "--trend-window",
            "40",
            "--stagger-days",
            "6",
            "--execute",
            "--force",
        )
        expected_summary = 运行参数调优.BatchTuningSummary(1, 1, 0)
        with patch.object(
            运行参数调优,
            "run_cases",
            return_value=expected_summary,
        ) as run_cases, patch("builtins.print"):
            result = 运行参数调优.cli(command)

        selected_cases = run_cases.call_args.args[0]
        self.assertEqual(
            selected_cases,
            (
                make_case(
                    tuning_type="volatility",
                    update_frequency="daily",
                    filter_profile="volatility_only",
                    correlation_threshold=0.8,
                    trend_window=40,
                    stagger_days=6,
                ),
            ),
        )
        self.assertTrue(run_cases.call_args.kwargs["force"])
        self.assertEqual(result, expected_summary)

    def test_cli_all_execute_routes_all_1200_cases(self) -> None:
        expected_summary = 运行参数调优.BatchTuningSummary(1200, 1200, 0)
        with patch.object(
            运行参数调优,
            "run_cases",
            return_value=expected_summary,
        ) as run_cases, patch("builtins.print"):
            result = 运行参数调优.cli(("--all", "--execute"))

        selected_cases = run_cases.call_args.args[0]
        self.assertEqual(len(selected_cases), 1200)
        self.assertEqual(
            selected_cases,
            tuple(运行参数调优.iter_tuning_cases()),
        )
        self.assertFalse(run_cases.call_args.kwargs["force"])
        self.assertEqual(result, expected_summary)

    def test_cli_rejects_unsafe_or_incompatible_commands(self) -> None:
        invalid_commands = (
            ("--execute",),
            ("--all", "--frequency", "daily"),
            ("--all", "--force"),
            ("--a", "--e"),
            ("--all", "--exec"),
            ("--tuning-type", "bias", "volatility", "--execute"),
            ("--frequency", "monthly", "daily", "--execute"),
            (
                "--tuning-type",
                "bias",
                "--profile",
                "volatility_only",
            ),
        )
        for command in invalid_commands:
            with self.subTest(command=command), patch.object(
                运行参数调优,
                "run_cases",
            ) as run_cases, patch("sys.stderr"):
                with self.assertRaises(SystemExit):
                    运行参数调优.cli(command)
            run_cases.assert_not_called()

    def test_failure_stops_before_later_cases(self) -> None:
        first = make_case(stagger_days=1)
        second = make_case(stagger_days=2)
        with patch.object(
            运行参数调优,
            "validate_required_inputs",
        ), patch.object(
            运行参数调优,
            "mark_case_incomplete",
        ), patch.object(
            运行参数调优,
            "run_tuning_case",
            side_effect=RuntimeError("grid failed"),
        ) as run_tuning_case, patch("builtins.print"):
            with self.assertRaisesRegex(
                RuntimeError,
                "第 1/2 组失败",
            ):
                运行参数调优.run_cases((first, second), force=True)

        run_tuning_case.assert_called_once_with(first)

    def test_completed_cases_skip_and_force_bypasses_completion_check(
        self,
    ) -> None:
        first = make_case(stagger_days=1)
        second = make_case(stagger_days=2)
        cases = (first, second)
        with patch.object(
            运行参数调优,
            "case_is_complete",
            side_effect=lambda tuning_case: tuning_case == first,
        ), patch.object(
            运行参数调优,
            "validate_required_inputs",
        ) as validate_inputs, patch.object(
            运行参数调优,
            "mark_case_incomplete",
        ), patch.object(
            运行参数调优,
            "incomplete_output_paths",
            return_value=(),
        ), patch.object(
            运行参数调优,
            "clear_case_incomplete",
        ), patch.object(
            运行参数调优,
            "run_tuning_case",
        ) as run_tuning_case, patch("builtins.print"):
            summary = 运行参数调优.run_cases(cases)

        validate_inputs.assert_called_once_with((second,))
        run_tuning_case.assert_called_once_with(second)
        self.assertEqual(summary, 运行参数调优.BatchTuningSummary(2, 1, 1))

        with patch.object(
            运行参数调优,
            "case_is_complete",
            side_effect=AssertionError("force不应做运行前完成检查"),
        ), patch.object(
            运行参数调优,
            "validate_required_inputs",
        ) as validate_inputs, patch.object(
            运行参数调优,
            "mark_case_incomplete",
        ), patch.object(
            运行参数调优,
            "incomplete_output_paths",
            return_value=(),
        ), patch.object(
            运行参数调优,
            "clear_case_incomplete",
        ), patch.object(
            运行参数调优,
            "run_tuning_case",
        ) as run_tuning_case, patch("builtins.print"):
            forced = 运行参数调优.run_cases(cases, force=True)

        validate_inputs.assert_called_once_with(cases)
        self.assertEqual(run_tuning_case.call_count, 2)
        self.assertEqual(forced, 运行参数调优.BatchTuningSummary(2, 2, 0))

    def test_failed_force_run_keeps_marker_and_default_retry_runs(self) -> None:
        tuning_case = make_case()
        with TemporaryDirectory() as temporary_directory:
            output_directory = Path(temporary_directory)
            with patch.object(
                运行参数调优,
                "tuning_output_directory",
                return_value=output_directory,
            ), patch.object(
                运行参数调优,
                "validate_required_inputs",
            ), patch("builtins.print"):
                for path in 运行参数调优.expected_output_paths(tuning_case):
                    path.write_bytes(b"old png")
                self.assertTrue(运行参数调优.case_is_complete(tuning_case))

                with patch.object(
                    运行参数调优,
                    "run_tuning_case",
                    side_effect=RuntimeError("grid failed"),
                ):
                    with self.assertRaises(RuntimeError):
                        运行参数调优.run_cases(
                            (tuning_case,),
                            force=True,
                        )

                self.assertTrue(
                    运行参数调优.incomplete_marker_path(tuning_case).is_file()
                )
                self.assertFalse(运行参数调优.case_is_complete(tuning_case))

                with patch.object(
                    运行参数调优,
                    "run_tuning_case",
                ) as retry:
                    summary = 运行参数调优.run_cases((tuning_case,))

                retry.assert_called_once_with(tuning_case)
                self.assertEqual(summary.executed_cases, 1)
                self.assertTrue(运行参数调优.case_is_complete(tuning_case))

    def test_missing_heatmap_after_return_fails_and_keeps_marker(self) -> None:
        tuning_case = make_case()
        with TemporaryDirectory() as temporary_directory:
            output_directory = Path(temporary_directory)
            with patch.object(
                运行参数调优,
                "tuning_output_directory",
                return_value=output_directory,
            ), patch.object(
                运行参数调优,
                "validate_required_inputs",
            ), patch.object(
                运行参数调优,
                "run_tuning_case",
            ), patch("builtins.print"):
                with self.assertRaisesRegex(
                    RuntimeError,
                    "BIAS参数调优 / monthly / bias_only",
                ):
                    运行参数调优.run_cases((tuning_case,))

                self.assertTrue(
                    运行参数调优.incomplete_marker_path(tuning_case).is_file()
                )

    def test_keyboard_interrupt_propagates_and_keeps_marker(self) -> None:
        tuning_case = make_case()
        with TemporaryDirectory() as temporary_directory:
            output_directory = Path(temporary_directory)
            with patch.object(
                运行参数调优,
                "tuning_output_directory",
                return_value=output_directory,
            ), patch.object(
                运行参数调优,
                "validate_required_inputs",
            ), patch.object(
                运行参数调优,
                "run_tuning_case",
                side_effect=KeyboardInterrupt,
            ), patch("builtins.print"):
                with self.assertRaises(KeyboardInterrupt):
                    运行参数调优.run_cases((tuning_case,))

                self.assertTrue(
                    运行参数调优.incomplete_marker_path(tuning_case).is_file()
                )

    def test_full_preflight_deduplicates_30_factor_and_6_price_paths(
        self,
    ) -> None:
        cases = tuple(运行参数调优.iter_tuning_cases())
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            etf_file = root / "etf_data.csv"
            etf_file.write_bytes(b"data")
            factor_directory = root / "factors"
            factor_directory.mkdir()
            for year in 运行参数调优.REQUIRED_FACTOR_YEARS:
                (factor_directory / f"{year}.csv").write_bytes(b"data")
            price_directory = root / "index_prices"
            price_directory.mkdir()
            (price_directory / "2021_01_04.csv").write_bytes(b"data")

            with patch.object(
                运行参数调优,
                "ETF_DATA_FILE",
                etf_file,
            ), patch.object(
                运行参数调优.交易策略回测,
                "load_benchmark_data",
            ), patch.object(
                运行参数调优,
                "factor_dir",
                return_value=factor_directory,
            ) as factor_path, patch.object(
                运行参数调优,
                "index_price_dir",
                return_value=price_directory,
            ) as price_path:
                missing = 运行参数调优.find_missing_inputs(cases)

        self.assertEqual(missing, ())
        self.assertEqual(factor_path.call_count, 30)
        self.assertEqual(price_path.call_count, 6)

    def test_missing_inputs_include_etf_benchmark_factors_and_prices(
        self,
    ) -> None:
        cases = (
            make_case(),
            make_case(
                tuning_type="volatility",
                filter_profile="volatility_only",
            ),
        )
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            with patch.object(
                运行参数调优,
                "ETF_DATA_FILE",
                root / "missing_etf.csv",
            ), patch.object(
                运行参数调优.交易策略回测,
                "load_benchmark_data",
                side_effect=FileNotFoundError("missing benchmark"),
            ), patch.object(
                运行参数调优,
                "factor_dir",
                return_value=root / "missing_factors",
            ), patch.object(
                运行参数调优,
                "index_price_dir",
                return_value=root / "missing_prices",
            ):
                missing = 运行参数调优.find_missing_inputs(cases)

        message = "\n".join(missing)
        self.assertIn("ETF总表缺失或为空", message)
        self.assertIn("基准数据不可用", message)
        self.assertIn("趋势因子目录缺少非空年度CSV", message)
        self.assertIn("指数过滤行情目录没有非空CSV", message)


if __name__ == "__main__":
    unittest.main()
