from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from 实验配置 import ExperimentCase, iter_experiment_cases
import 运行全部回测


def make_case(
    *,
    update_frequency: str = "daily",
    filter_profile: str = "base",
    correlation_threshold: float = 0.9,
    trend_factor: str = "return_r2",
    trend_window: int = 20,
    stagger_days: int = 1,
) -> ExperimentCase:
    return ExperimentCase(
        update_frequency=update_frequency,
        filter_profile=filter_profile,
        correlation_threshold=correlation_threshold,
        trend_factor=trend_factor,
        trend_window=trend_window,
        stagger_days=stagger_days,
    )


def case_from_keyword_arguments(values: dict[str, object]) -> ExperimentCase:
    return ExperimentCase(
        update_frequency=str(values["update_frequency"]),
        filter_profile=str(values["filter_profile"]),
        correlation_threshold=float(values["correlation_threshold"]),
        trend_factor=str(values["trend_factor"]),
        trend_window=int(values["trend_window"]),
        stagger_days=int(values["stagger_days"]),
    )


class BatchBacktestRunnerTests(unittest.TestCase):
    def test_complete_case_requires_twenty_nonempty_files(self) -> None:
        experiment_case = make_case()
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)

            def output_directory(
                _: ExperimentCase,
                execution_mode: str,
            ) -> Path:
                return root / execution_mode

            with patch.object(
                运行全部回测,
                "backtest_result_dir",
                side_effect=output_directory,
            ):
                expected_paths = [
                    path
                    for mode in 运行全部回测.EXECUTION_MODES
                    for path in 运行全部回测.expected_mode_output_paths(
                        experiment_case,
                        mode,
                    )
                ]
                self.assertEqual(len(expected_paths), 20)
                self.assertEqual(len(set(expected_paths)), 20)
                self.assertFalse(
                    运行全部回测.case_is_complete(experiment_case)
                )

                for path in expected_paths:
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(b"complete")
                self.assertTrue(
                    运行全部回测.case_is_complete(experiment_case)
                )

                marker_path = 运行全部回测.case_incomplete_marker_path(
                    experiment_case
                )
                marker_path.touch()
                self.assertFalse(
                    运行全部回测.case_is_complete(experiment_case)
                )
                marker_path.unlink()

                expected_paths[-1].unlink()
                self.assertFalse(
                    运行全部回测.case_is_complete(experiment_case)
                )
                expected_paths[-1].touch()
                self.assertFalse(
                    运行全部回测.case_is_complete(experiment_case)
                )

    def test_expected_output_names_match_both_formal_modes(self) -> None:
        experiment_case = make_case()
        with patch.object(
            运行全部回测,
            "backtest_result_dir",
            side_effect=lambda _case, mode: Path(mode),
        ):
            close_names = {
                path.name
                for path in 运行全部回测.expected_mode_output_paths(
                    experiment_case,
                    "close",
                )
            }
            vwap_names = {
                path.name
                for path in 运行全部回测.expected_mode_output_paths(
                    experiment_case,
                    "next_day_vwap",
                )
            }

        self.assertEqual(
            close_names,
            {
                "close_annual_metrics.xlsx",
                "close_backtest_metrics.xlsx",
                "close_holdings.xlsx",
                "close_time_series.xlsx",
                "close_account_details.xlsx",
                "close_cumulative_nav.png",
                "close_cumulative_excess.png",
                "close_turnover.png",
                "close_transaction_cost.png",
                "close_capacity.png",
            },
        )
        self.assertEqual(
            vwap_names,
            {
                "next_day_vwap_annual_metrics.xlsx",
                "next_day_vwap_backtest_metrics.xlsx",
                "next_day_vwap_holdings.xlsx",
                "next_day_vwap_time_series.xlsx",
                "next_day_vwap_account_details.xlsx",
                "next_day_vwap_cumulative_nav.png",
                "next_day_vwap_cumulative_excess.png",
                "next_day_vwap_turnover.png",
                "next_day_vwap_transaction_cost.png",
                "next_day_vwap_capacity.png",
            },
        )

    def test_main_routes_all_2400_unique_cases_once(self) -> None:
        expected_cases = tuple(iter_experiment_cases())
        with patch.object(
            运行全部回测,
            "case_is_complete",
            return_value=False,
        ), patch.object(
            运行全部回测,
            "incomplete_output_paths",
            return_value=(),
        ), patch.object(
            运行全部回测,
            "mark_case_incomplete",
        ), patch.object(
            运行全部回测,
            "clear_case_incomplete",
        ), patch.object(
            运行全部回测,
            "validate_required_inputs",
        ) as validate_inputs, patch.object(
            运行全部回测.交易策略回测,
            "main",
        ) as backtest_main, patch("builtins.print"):
            summary = 运行全部回测.main()

        actual_cases = tuple(
            case_from_keyword_arguments(call_item.kwargs)
            for call_item in backtest_main.call_args_list
        )
        self.assertEqual(len(expected_cases), 2400)
        self.assertEqual(len(set(expected_cases)), 2400)
        self.assertEqual(actual_cases, expected_cases)
        validate_inputs.assert_called_once_with(expected_cases)
        self.assertEqual(
            summary,
            运行全部回测.BatchRunSummary(
                total_cases=2400,
                executed_cases=2400,
                skipped_cases=0,
            ),
        )

    def test_complete_cases_skip_and_only_pending_cases_are_preflighted(
        self,
    ) -> None:
        first = make_case(stagger_days=1)
        second = make_case(stagger_days=2)
        third = make_case(stagger_days=3)
        cases = (first, second, third)

        with patch.object(
            运行全部回测,
            "case_is_complete",
            side_effect=lambda case: case == first,
        ), patch.object(
            运行全部回测,
            "incomplete_output_paths",
            return_value=(),
        ), patch.object(
            运行全部回测,
            "mark_case_incomplete",
        ), patch.object(
            运行全部回测,
            "clear_case_incomplete",
        ), patch.object(
            运行全部回测,
            "validate_required_inputs",
        ) as validate_inputs, patch.object(
            运行全部回测.交易策略回测,
            "main",
        ) as backtest_main, patch("builtins.print"):
            summary = 运行全部回测.run_cases(cases)

        validate_inputs.assert_called_once_with((second, third))
        actual_cases = tuple(
            case_from_keyword_arguments(call_item.kwargs)
            for call_item in backtest_main.call_args_list
        )
        self.assertEqual(actual_cases, (second, third))
        self.assertEqual(
            summary,
            运行全部回测.BatchRunSummary(
                total_cases=3,
                executed_cases=2,
                skipped_cases=1,
            ),
        )

    def test_force_runs_every_case_without_completion_checks(self) -> None:
        cases = (make_case(stagger_days=1), make_case(stagger_days=2))
        with patch.object(
            运行全部回测,
            "case_is_complete",
            side_effect=AssertionError("force模式不应检查已有结果"),
        ), patch.object(
            运行全部回测,
            "incomplete_output_paths",
            return_value=(),
        ), patch.object(
            运行全部回测,
            "mark_case_incomplete",
        ), patch.object(
            运行全部回测,
            "clear_case_incomplete",
        ), patch.object(
            运行全部回测,
            "validate_required_inputs",
        ) as validate_inputs, patch.object(
            运行全部回测.交易策略回测,
            "main",
        ) as backtest_main, patch("builtins.print"):
            summary = 运行全部回测.run_cases(cases, force=True)

        validate_inputs.assert_called_once_with(cases)
        self.assertEqual(backtest_main.call_count, 2)
        self.assertEqual(summary.executed_cases, 2)
        self.assertEqual(summary.skipped_cases, 0)

    def test_failure_stops_before_later_cases_and_preserves_cause(self) -> None:
        first = make_case(stagger_days=1)
        second = make_case(stagger_days=2)
        third = make_case(stagger_days=3)
        cases = (first, second, third)
        original_error = ValueError("broken case")

        with patch.object(
            运行全部回测,
            "case_is_complete",
            return_value=False,
        ), patch.object(
            运行全部回测,
            "incomplete_output_paths",
            return_value=(),
        ), patch.object(
            运行全部回测,
            "mark_case_incomplete",
        ), patch.object(
            运行全部回测,
            "clear_case_incomplete",
        ), patch.object(
            运行全部回测,
            "validate_required_inputs",
        ), patch.object(
            运行全部回测.交易策略回测,
            "main",
            side_effect=(None, original_error, None),
        ) as backtest_main, patch("builtins.print"):
            with self.assertRaisesRegex(
                RuntimeError,
                "第 2/3 组失败",
            ) as raised:
                运行全部回测.run_cases(cases)

        self.assertIs(raised.exception.__cause__, original_error)
        self.assertEqual(backtest_main.call_count, 2)
        actual_cases = tuple(
            case_from_keyword_arguments(call_item.kwargs)
            for call_item in backtest_main.call_args_list
        )
        self.assertEqual(actual_cases, (first, second))

    def test_keyboard_interrupt_is_not_swallowed(self) -> None:
        experiment_case = make_case()
        with patch.object(
            运行全部回测,
            "case_is_complete",
            return_value=False,
        ), patch.object(
            运行全部回测,
            "mark_case_incomplete",
        ), patch.object(
            运行全部回测,
            "validate_required_inputs",
        ), patch.object(
            运行全部回测.交易策略回测,
            "main",
            side_effect=KeyboardInterrupt,
        ), patch("builtins.print"):
            with self.assertRaises(KeyboardInterrupt):
                运行全部回测.run_cases((experiment_case,))

    def test_failed_force_run_stays_incomplete_and_is_retried(self) -> None:
        experiment_case = make_case()
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)

            def output_directory(
                _: ExperimentCase,
                execution_mode: str,
            ) -> Path:
                return root / execution_mode

            with patch.object(
                运行全部回测,
                "backtest_result_dir",
                side_effect=output_directory,
            ), patch.object(
                运行全部回测,
                "validate_required_inputs",
            ), patch("builtins.print"):
                for execution_mode in 运行全部回测.EXECUTION_MODES:
                    for path in 运行全部回测.expected_mode_output_paths(
                        experiment_case,
                        execution_mode,
                    ):
                        path.parent.mkdir(parents=True, exist_ok=True)
                        path.write_bytes(b"old complete result")

                self.assertTrue(
                    运行全部回测.case_is_complete(experiment_case)
                )
                with patch.object(
                    运行全部回测.交易策略回测,
                    "main",
                    side_effect=RuntimeError("validation failed"),
                ):
                    with self.assertRaises(RuntimeError):
                        运行全部回测.run_cases(
                            (experiment_case,),
                            force=True,
                        )

                self.assertTrue(
                    运行全部回测.case_incomplete_marker_path(
                        experiment_case
                    ).is_file()
                )
                self.assertFalse(
                    运行全部回测.case_is_complete(experiment_case)
                )

                with patch.object(
                    运行全部回测.交易策略回测,
                    "main",
                ) as retry_main:
                    summary = 运行全部回测.run_cases((experiment_case,))

                retry_main.assert_called_once()
                self.assertEqual(summary.executed_cases, 1)
                self.assertTrue(
                    运行全部回测.case_is_complete(experiment_case)
                )

    def test_cli_force_flag_is_explicit(self) -> None:
        self.assertFalse(运行全部回测.parse_cli_arguments([]).force)
        self.assertTrue(
            运行全部回测.parse_cli_arguments(["--force"]).force
        )

    def test_missing_inputs_are_aggregated_before_first_backtest(self) -> None:
        cases = (
            make_case(filter_profile="base"),
            make_case(filter_profile="bias_only"),
        )
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            with patch.object(
                运行全部回测,
                "factor_dir",
                return_value=root / "missing_factors",
            ), patch.object(
                运行全部回测,
                "index_price_dir",
                return_value=root / "missing_index_prices",
            ), patch.object(
                运行全部回测.交易策略回测,
                "ETF_DATA_FILE",
                root / "missing_etf_data.csv",
            ), patch.object(
                运行全部回测.交易策略回测,
                "BENCHMARK_DIR",
                root / "missing_benchmark",
            ), patch.object(
                运行全部回测,
                "case_is_complete",
                return_value=False,
            ), patch.object(
                运行全部回测.交易策略回测,
                "main",
            ) as backtest_main, patch("builtins.print"):
                with self.assertRaises(FileNotFoundError) as raised:
                    运行全部回测.run_cases(cases)

        message = str(raised.exception)
        self.assertIn("ETF总表缺失或为空", message)
        self.assertIn("基准目录没有非空CSV", message)
        self.assertIn("趋势因子目录缺少非空年度CSV", message)
        self.assertIn("指数过滤行情目录没有非空CSV", message)
        backtest_main.assert_not_called()

    def test_base_profile_does_not_require_filter_price_directory(self) -> None:
        experiment_case = make_case(filter_profile="base")
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            etf_file = root / "etf_data.csv"
            etf_file.write_bytes(b"data")
            benchmark_directory = root / "benchmark"
            benchmark_directory.mkdir()
            (benchmark_directory / "benchmark.csv").write_bytes(b"data")
            factor_directory = root / "factors"
            factor_directory.mkdir()
            for year in 运行全部回测.REQUIRED_FACTOR_YEARS:
                (factor_directory / f"{year}.csv").write_bytes(b"data")

            with patch.object(
                运行全部回测,
                "factor_dir",
                return_value=factor_directory,
            ), patch.object(
                运行全部回测,
                "index_price_dir",
                side_effect=AssertionError("base不应要求过滤行情"),
            ), patch.object(
                运行全部回测.交易策略回测,
                "ETF_DATA_FILE",
                etf_file,
            ), patch.object(
                运行全部回测.交易策略回测,
                "BENCHMARK_DIR",
                benchmark_directory,
            ):
                missing = 运行全部回测.find_missing_inputs(
                    (experiment_case,)
                )

        self.assertEqual(missing, ())

    def test_factor_preflight_requires_every_configured_year(self) -> None:
        experiment_case = make_case(filter_profile="base")
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            etf_file = root / "etf_data.csv"
            etf_file.write_bytes(b"data")
            benchmark_directory = root / "benchmark"
            benchmark_directory.mkdir()
            (benchmark_directory / "benchmark.csv").write_bytes(b"data")
            factor_directory = root / "factors"
            factor_directory.mkdir()
            first_year = 运行全部回测.REQUIRED_FACTOR_YEARS[0]
            (factor_directory / f"{first_year}.csv").write_bytes(b"data")

            with patch.object(
                运行全部回测,
                "factor_dir",
                return_value=factor_directory,
            ), patch.object(
                运行全部回测.交易策略回测,
                "ETF_DATA_FILE",
                etf_file,
            ), patch.object(
                运行全部回测.交易策略回测,
                "BENCHMARK_DIR",
                benchmark_directory,
            ):
                missing = 运行全部回测.find_missing_inputs(
                    (experiment_case,)
                )

        self.assertEqual(len(missing), 1)
        for year in 运行全部回测.REQUIRED_FACTOR_YEARS[1:]:
            self.assertIn(f"{year}.csv", missing[0])

    def test_full_preflight_deduplicates_shared_input_dimensions(self) -> None:
        cases = tuple(iter_experiment_cases())
        expected_factor_keys = {
            (
                case.update_frequency,
                case.correlation_threshold,
                case.trend_window,
            )
            for case in cases
        }
        expected_index_keys = {
            (case.update_frequency, case.correlation_threshold)
            for case in cases
            if case.filter_profile != "base"
        }
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            etf_file = root / "etf_data.csv"
            etf_file.write_bytes(b"data")
            benchmark_directory = root / "benchmark"
            benchmark_directory.mkdir()
            (benchmark_directory / "benchmark.csv").write_bytes(b"data")
            factor_directory = root / "factors"
            factor_directory.mkdir()
            for year in 运行全部回测.REQUIRED_FACTOR_YEARS:
                (factor_directory / f"{year}.csv").write_bytes(b"data")
            index_directory = root / "index_prices"
            index_directory.mkdir()
            (index_directory / "2021_01_04.csv").write_bytes(b"data")

            with patch.object(
                运行全部回测,
                "factor_dir",
                return_value=factor_directory,
            ) as factor_path, patch.object(
                运行全部回测,
                "index_price_dir",
                return_value=index_directory,
            ) as index_path, patch.object(
                运行全部回测.交易策略回测,
                "ETF_DATA_FILE",
                etf_file,
            ), patch.object(
                运行全部回测.交易策略回测,
                "BENCHMARK_DIR",
                benchmark_directory,
            ):
                missing = 运行全部回测.find_missing_inputs(cases)

        actual_factor_keys = {
            tuple(call_item.args) for call_item in factor_path.call_args_list
        }
        actual_index_keys = {
            tuple(call_item.args) for call_item in index_path.call_args_list
        }
        self.assertEqual(missing, ())
        self.assertEqual(len(expected_factor_keys), 30)
        self.assertEqual(factor_path.call_count, 30)
        self.assertEqual(actual_factor_keys, expected_factor_keys)
        self.assertEqual(len(expected_index_keys), 6)
        self.assertEqual(index_path.call_count, 6)
        self.assertEqual(actual_index_keys, expected_index_keys)

    def test_empty_duplicate_and_invalid_cases_fail_before_io(self) -> None:
        valid_case = make_case()
        invalid_case = make_case(update_frequency="weekly")

        with patch.object(
            运行全部回测,
            "case_is_complete",
            side_effect=AssertionError("非法计划不应检查输出"),
        ):
            with self.assertRaises(ValueError):
                运行全部回测.run_cases(())
            with self.assertRaisesRegex(ValueError, "重复实验组合"):
                运行全部回测.run_cases((valid_case, valid_case))
            with self.assertRaises(ValueError):
                运行全部回测.run_cases((invalid_case,))


if __name__ == "__main__":
    unittest.main()
