import csv
from datetime import date, timedelta
from itertools import product
import math
import os
from pathlib import Path
import statistics
from tempfile import TemporaryDirectory, gettempdir
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault(
    "MPLCONFIGDIR",
    str(Path(gettempdir()) / "etf_rotation_matplotlib_tests"),
)

import numpy as np

from 实验配置 import (
    CORRELATION_THRESHOLDS,
    STAGGER_DAYS,
    TREND_WINDOWS,
    UPDATE_FREQUENCIES,
)
from scripts.ETF趋势策略回测 import 交易策略回测
from scripts.策略参数调优 import BIAS参数调优


def make_settings(
    *,
    update_frequency: str = "daily",
    filter_profile: str = "bias_only",
    correlation_threshold: float = 0.9,
    trend_window: int = 20,
    stagger_days: int = 1,
    root: Path = Path("test_outputs"),
) -> BIAS参数调优.BiasTuningSettings:
    return BIAS参数调优.BiasTuningSettings(
        update_frequency=update_frequency,
        filter_profile=filter_profile,
        correlation_threshold=correlation_threshold,
        trend_window=trend_window,
        stagger_days=stagger_days,
        factor_directory=root / "factors",
        index_price_directory=root / "index_prices",
        output_directory=root / "parameter_grid",
        bias_mode="upper",
        vol_filter_enabled=(filter_profile == "bias_and_volatility"),
    )


def make_factor_member(index_code: str, signal_date: date) -> SimpleNamespace:
    return SimpleNamespace(
        signal_date=signal_date,
        index_code=index_code,
        index_name=f"指数{index_code}",
        trend_factor=1.0,
        window_return=0.1,
    )


class BiasParameterTuningTests(unittest.TestCase):
    def test_all_public_dimensions_route_the_three_directories(self) -> None:
        factor_path = Mock(
            side_effect=lambda frequency, threshold, window: Path(
                f"factor/{frequency}/{threshold:g}/{window}"
            )
        )
        index_path = Mock(
            side_effect=lambda frequency, threshold: Path(
                f"index/{frequency}/{threshold:g}"
            )
        )
        output_path = Mock(
            side_effect=lambda frequency, profile, tuning_type,
            threshold, window, stagger: Path(
                f"output/{frequency}/{profile}/{tuning_type}/"
                f"{threshold:g}/{window}/{stagger}"
            )
        )

        profiles = ("bias_only", "bias_and_volatility")
        combinations = tuple(
            product(
                UPDATE_FREQUENCIES,
                profiles,
                CORRELATION_THRESHOLDS,
                TREND_WINDOWS,
                STAGGER_DAYS,
            )
        )
        with patch.object(
            BIAS参数调优,
            "factor_dir",
            factor_path,
        ), patch.object(
            BIAS参数调优,
            "index_price_dir",
            index_path,
        ), patch.object(
            BIAS参数调优,
            "tuning_result_dir",
            output_path,
        ):
            for values in combinations:
                with self.subTest(values=values):
                    settings = BIAS参数调优.build_tuning_settings(*values)
                    frequency, profile, threshold, window, stagger = values
                    self.assertEqual(
                        settings.factor_directory,
                        Path(f"factor/{frequency}/{threshold:g}/{window}"),
                    )
                    self.assertEqual(
                        settings.index_price_directory,
                        Path(f"index/{frequency}/{threshold:g}"),
                    )
                    self.assertEqual(
                        settings.output_directory,
                        Path(
                            f"output/{frequency}/{profile}/"
                            f"bias_parameter_grid/{threshold:g}/"
                            f"{window}/{stagger}"
                        ),
                    )

        self.assertEqual(factor_path.call_count, len(combinations))
        self.assertEqual(index_path.call_count, len(combinations))
        self.assertEqual(output_path.call_count, len(combinations))

    def test_profiles_map_to_bias_and_volatility_switches(self) -> None:
        bias_only = BIAS参数调优.build_tuning_settings(
            "daily",
            "bias_only",
            0.9,
            20,
            1,
        )
        combined = BIAS参数调优.build_tuning_settings(
            "monthly",
            "bias_and_volatility",
            0.8,
            40,
            10,
        )

        self.assertEqual(
            (bias_only.bias_mode, bias_only.vol_filter_enabled),
            ("upper", False),
        )
        self.assertEqual(
            (combined.bias_mode, combined.vol_filter_enabled),
            ("upper", True),
        )

    def test_invalid_dimensions_fail_before_path_resolution(self) -> None:
        invalid_values = (
            ("weekly", "bias_only", 0.9, 20, 1),
            ("daily", "base", 0.9, 20, 1),
            ("daily", "volatility_only", 0.9, 20, 1),
            ("daily", "bias_only", 0.75, 20, 1),
            ("daily", "bias_only", 0.9, 30, 1),
            ("daily", "bias_only", 0.9, 20, 0),
        )

        for values in invalid_values:
            with self.subTest(values=values), patch.object(
                BIAS参数调优,
                "factor_dir",
                side_effect=AssertionError("不应计算因子路径"),
            ), patch.object(
                BIAS参数调优,
                "index_price_dir",
                side_effect=AssertionError("不应计算行情路径"),
            ), patch.object(
                BIAS参数调优,
                "tuning_result_dir",
                side_effect=AssertionError("不应计算输出路径"),
            ):
                with self.assertRaises(ValueError):
                    BIAS参数调优.build_tuning_settings(*values)

    def test_backtest_module_receives_only_current_case_values(self) -> None:
        settings = make_settings(
            update_frequency="monthly",
            filter_profile="bias_only",
            correlation_threshold=0.7,
            trend_window=60,
            stagger_days=9,
        )
        backtest = SimpleNamespace()

        BIAS参数调优.configure_backtest_module(backtest, settings)

        self.assertEqual(backtest.UPDATE_FREQUENCY, "monthly")
        self.assertEqual(backtest.FILTER_PROFILE, "bias_only")
        self.assertEqual(backtest.CLUSTER_CORRELATION_THRESHOLD, 0.7)
        self.assertEqual(backtest.TREND_WINDOW, 60)
        self.assertEqual(backtest.ACCOUNT_COUNT, 9)
        self.assertEqual(backtest.ACCOUNT_REBALANCE_INTERVAL, 9)
        self.assertEqual(backtest.ACCOUNT_VARIANT_DIR, "staggered_9d")
        self.assertEqual(backtest.FACTOR_DIR, settings.factor_directory)
        self.assertEqual(
            backtest.INDEX_RAW_DATA_DIR,
            settings.index_price_directory,
        )
        self.assertEqual(backtest.BIAS_MODE, "upper")
        self.assertFalse(backtest.VOL_FILTER_ENABLED)
        self.assertEqual(backtest.BIAS_WINDOW, 36)
        self.assertEqual(backtest.BIAS_UPPER, 0.09)

    def test_real_backtest_module_loads_and_validates_both_profiles(self) -> None:
        cases = (
            ("daily", "bias_only", 0.9, 20, 1),
            ("monthly", "bias_and_volatility", 0.7, 60, 10),
        )

        for values in cases:
            with self.subTest(values=values):
                settings = BIAS参数调优.build_tuning_settings(*values)
                backtest = BIAS参数调优.load_backtest_module(settings)
                BIAS参数调优.validate_grid(backtest, settings)

                self.assertEqual(backtest.FACTOR_DIR, settings.factor_directory)
                self.assertEqual(
                    backtest.INDEX_RAW_DATA_DIR,
                    settings.index_price_directory,
                )
                self.assertEqual(
                    backtest.VOL_FILTER_ENABLED,
                    settings.vol_filter_enabled,
                )
                self.assertEqual(backtest.ACCOUNT_COUNT, settings.stagger_days)

    def test_factor_reader_uses_the_case_factor_directory(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            factor_directory = Path(temporary_directory) / "case_factors"
            factor_directory.mkdir()
            factor_file = factor_directory / "2021.csv"
            with factor_file.open(
                "w",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                writer = csv.DictWriter(
                    handle,
                    fieldnames=(
                        "日期",
                        "对标指数代码",
                        "对标指数",
                        "窗口收益率",
                        "趋势质量因子",
                        "风险调整趋势得分",
                    ),
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "日期": "2021-01-04",
                        "对标指数代码": "INDEX",
                        "对标指数": "测试指数",
                        "窗口收益率": 0.1,
                        "趋势质量因子": 0.2,
                        "风险调整趋势得分": 0.3,
                    }
                )

            backtest = SimpleNamespace(
                FACTOR_DIR=factor_directory,
                SCORE_COLUMNS=交易策略回测.SCORE_COLUMNS,
                finite_float=交易策略回测.finite_float,
                clean_text=交易策略回测.clean_text,
                parse_date=交易策略回测.parse_date,
                FactorMember=交易策略回测.FactorMember,
            )
            universes = BIAS参数调优.load_factor_universes(backtest)

        signal_date = date(2021, 1, 4)
        self.assertEqual(
            universes["return_r2"][signal_date][0].trend_factor,
            0.2,
        )
        self.assertEqual(
            universes["return_vol"][signal_date][0].trend_factor,
            0.3,
        )

    def test_filter_metrics_ignore_prices_after_signal_date(self) -> None:
        signal_date = date(2021, 1, 3)
        members = (make_factor_member("A", signal_date),)
        factor_universes = {
            method: {signal_date: members}
            for method in BIAS参数调优.SCORE_METHODS
        }
        backtest = SimpleNamespace(
            VOL_RETURN_DAYS=2,
            VOL_FILTER_ENABLED=True,
            MAX_PRICE_STALENESS_CALENDAR_DAYS=7,
        )
        dates = [
            signal_date - timedelta(days=2),
            signal_date - timedelta(days=1),
            signal_date,
        ]
        prices = [100.0, 110.0, 105.0]

        original = BIAS参数调优.precompute_filter_metrics(
            backtest,
            factor_universes,
            {"A": (dates, prices)},
            (2,),
        )
        with_future = BIAS参数调优.precompute_filter_metrics(
            backtest,
            factor_universes,
            {
                "A": (
                    [*dates, signal_date + timedelta(days=1)],
                    [*prices, 1000000.0],
                )
            },
            (2,),
        )

        self.assertEqual(original, with_future)
        metric = original[(signal_date, "A")]
        self.assertAlmostEqual(
            metric.biases[0],
            105.0 / statistics.mean((110.0, 105.0)) - 1.0,
        )
        self.assertAlmostEqual(
            metric.volatility,
            statistics.stdev((110.0 / 100.0 - 1.0, 105.0 / 110.0 - 1.0)),
        )

    def test_volatility_filter_is_off_for_bias_only_and_on_for_combined(
        self,
    ) -> None:
        signal_date = date(2021, 1, 4)
        members = tuple(
            make_factor_member(code, signal_date) for code in ("A", "B", "C")
        )
        factor_universes = {
            method: {signal_date: members}
            for method in BIAS参数调优.SCORE_METHODS
        }
        metrics = {
            (signal_date, "A"): BIAS参数调优.FilterMetrics(0.1, (0.0,)),
            (signal_date, "B"): BIAS参数调优.FilterMetrics(0.3, (0.0,)),
            (signal_date, "C"): BIAS参数调优.FilterMetrics(0.2, (0.0,)),
        }
        bias_only = SimpleNamespace(
            VOL_FILTER_ENABLED=False,
            VOL_KEEP_TOP_RATIO=0.55,
            VOL_FILTER_MODE="high",
        )
        combined = SimpleNamespace(
            VOL_FILTER_ENABLED=True,
            VOL_KEEP_TOP_RATIO=0.55,
            VOL_FILTER_MODE="high",
        )

        no_vol = BIAS参数调优.precompute_volatility_membership(
            bias_only,
            factor_universes,
            metrics,
        )
        with_vol = BIAS参数调优.precompute_volatility_membership(
            combined,
            factor_universes,
            metrics,
        )

        for method in BIAS参数调优.SCORE_METHODS:
            self.assertEqual(
                no_vol[method][signal_date],
                frozenset(("A", "B", "C")),
            )
            self.assertEqual(
                with_vol[method][signal_date],
                frozenset(("B", "C")),
            )

    def test_bias_upper_boundary_is_inclusive_after_volatility_membership(
        self,
    ) -> None:
        signal_date = date(2021, 1, 4)
        members = tuple(
            make_factor_member(code, signal_date) for code in ("A", "B", "C")
        )
        metrics = {
            (signal_date, "A"): BIAS参数调优.FilterMetrics(None, (0.09,)),
            (signal_date, "B"): BIAS参数调优.FilterMetrics(
                None,
                (0.09 + 5e-13,),
            ),
            (signal_date, "C"): BIAS参数调优.FilterMetrics(None, (0.10,)),
        }
        backtest = SimpleNamespace(TOP_PERCENT=1.0)

        selected, eligible_count = BIAS参数调优.select_members(
            backtest,
            members,
            frozenset(("A", "B")),
            metrics,
            0,
            0.09,
        )

        self.assertEqual(eligible_count, 2)
        self.assertEqual(
            tuple(member.index_code for member, _ in selected),
            ("A", "B"),
        )

    def test_parameter_targets_intersect_volatility_and_bias_filters(
        self,
    ) -> None:
        signal_date = date(2021, 1, 4)
        members = tuple(
            make_factor_member(code, signal_date) for code in ("A", "B", "C")
        )
        metrics = {
            (signal_date, "A"): BIAS参数调优.FilterMetrics(None, (0.01,)),
            (signal_date, "B"): BIAS参数调优.FilterMetrics(None, (0.02,)),
            (signal_date, "C"): BIAS参数调优.FilterMetrics(None, (0.20,)),
        }
        candidates = {
            signal_date: {
                code: SimpleNamespace(
                    index_name=f"指数{code}",
                    etf_code=f"ETF{code}",
                    etf_name=f"ETF{code}",
                    selection_volume=100.0,
                    selection_amount=1000.0,
                    selection_scale=10000.0,
                )
                for code in ("A", "B", "C")
            }
        }
        backtest = SimpleNamespace(
            TOP_PERCENT=1.0,
            TargetMember=交易策略回测.TargetMember,
            DailyTarget=交易策略回测.DailyTarget,
        )

        targets = BIAS参数调优.build_parameter_targets(
            backtest,
            "return_r2",
            {signal_date: members},
            candidates,
            {signal_date: frozenset(("B", "C"))},
            metrics,
            0,
            0.09,
            (signal_date,),
        )

        target = targets[signal_date]
        self.assertEqual(
            tuple(member.index_code for member in target.members),
            ("B",),
        )
        self.assertEqual(target.planned_index_count, 1)
        self.assertEqual(target.selected_index_count, 1)
        self.assertEqual(target.filtered_index_count, 2)

    def test_all_stagger_days_drive_both_fast_backtest_schedules(self) -> None:
        trading_dates = tuple(
            date(2021, 1, 1) + timedelta(days=offset)
            for offset in range(25)
        )

        for stagger_days in STAGGER_DAYS:
            for mode, expected_calls in (
                ("close", len(trading_dates)),
                ("next_day_vwap", len(trading_dates) - 1),
            ):
                with self.subTest(stagger_days=stagger_days, mode=mode):
                    account_order: dict[int, int] = {}
                    observed: list[int] = []

                    def rebalance(
                        positions: dict[str, object],
                        cash: float,
                        *_: object,
                    ) -> SimpleNamespace:
                        identity = id(positions)
                        if identity not in account_order:
                            account_order[identity] = len(account_order)
                        observed.append(account_order[identity])
                        return SimpleNamespace(positions=positions, cash=cash)

                    backtest = SimpleNamespace(
                        AccountState=交易策略回测.AccountState,
                        INITIAL_NAV=1.0,
                        ACCOUNT_COUNT=stagger_days,
                        ACCOUNT_REBALANCE_INTERVAL=stagger_days,
                        rebalance_portfolio=rebalance,
                    )
                    result = BIAS参数调优.run_total_return(
                        backtest,
                        mode,
                        trading_dates,
                        {},
                        {},
                    )

                    self.assertTrue(math.isclose(result, 0.0, abs_tol=1e-12))
                    self.assertEqual(len(observed), expected_calls)
                    self.assertEqual(
                        observed,
                        [
                            position % stagger_days
                            for position in range(expected_calls)
                        ],
                    )

    def test_close_and_next_day_vwap_use_the_intended_signal_and_price(
        self,
    ) -> None:
        first_date = date(2021, 1, 4)
        second_date = date(2021, 1, 5)
        third_date = date(2021, 1, 6)
        trading_dates = (first_date, second_date, third_date)
        targets = {
            first_date: SimpleNamespace(
                signal_date=first_date,
                members=(SimpleNamespace(etf_code="ETF1"),),
            ),
            second_date: SimpleNamespace(
                signal_date=second_date,
                members=(SimpleNamespace(etf_code="ETF2"),),
            ),
        }
        prices = {
            first_date: {
                "ETF1": SimpleNamespace(close=100.0, vwap=101.0),
                "ETF2": SimpleNamespace(close=200.0, vwap=202.0),
            },
            second_date: {
                "ETF1": SimpleNamespace(close=110.0, vwap=111.0),
                "ETF2": SimpleNamespace(close=210.0, vwap=212.0),
            },
            third_date: {
                "ETF1": SimpleNamespace(close=120.0, vwap=121.0),
                "ETF2": SimpleNamespace(close=220.0, vwap=222.0),
            },
        }

        for mode, expected in (
            (
                "close",
                (
                    (first_date, {"ETF1": 100.0}),
                    (second_date, {"ETF2": 210.0}),
                ),
            ),
            (
                "next_day_vwap",
                (
                    (first_date, {"ETF1": 111.0}),
                    (second_date, {"ETF2": 222.0}),
                ),
            ),
        ):
            with self.subTest(mode=mode):
                observed: list[tuple[date, dict[str, float]]] = []

                def aggregate(
                    target: SimpleNamespace,
                    execution_prices: dict[str, float],
                    _: dict[str, float],
                ) -> tuple[dict[str, float], dict[str, object]]:
                    observed.append(
                        (target.signal_date, dict(execution_prices))
                    )
                    return {}, {}

                backtest = SimpleNamespace(
                    AccountState=交易策略回测.AccountState,
                    INITIAL_NAV=1.0,
                    ACCOUNT_COUNT=1,
                    ACCOUNT_REBALANCE_INTERVAL=1,
                    aggregate_target_weights=aggregate,
                    rebalance_portfolio=lambda positions, cash, *_: (
                        SimpleNamespace(positions=positions, cash=cash)
                    ),
                )

                BIAS参数调优.run_total_return(
                    backtest,
                    mode,
                    trading_dates,
                    targets,
                    prices,
                )

                self.assertEqual(observed, list(expected))

    def test_one_cell_grid_still_runs_two_factors_and_two_trade_modes(
        self,
    ) -> None:
        signal_date = date(2021, 1, 4)
        member = make_factor_member("INDEX", signal_date)
        factor_universes = {
            method: {signal_date: (member,)}
            for method in BIAS参数调优.SCORE_METHODS
        }
        contexts = {
            method: BIAS参数调优.MethodContext(
                trading_dates=(signal_date,),
                benchmark_end_date=signal_date,
            )
            for method in BIAS参数调优.SCORE_METHODS
        }
        backtest = SimpleNamespace(
            INDEX_RAW_DATA_DIR=Path("selected_index_prices"),
            build_daily_targets=Mock(return_value=({}, (signal_date,))),
            load_filter_price_history=Mock(return_value={}),
            load_selected_prices=Mock(return_value={}),
        )
        returns = iter((0.1, 0.2, 0.3, 0.4))

        with patch.object(
            BIAS参数调优,
            "load_factor_universes",
            return_value=factor_universes,
        ), patch.object(
            BIAS参数调优,
            "build_universal_selections",
            return_value={},
        ), patch.object(
            BIAS参数调优,
            "prepare_method_contexts",
            return_value=contexts,
        ), patch.object(
            BIAS参数调优,
            "precompute_filter_metrics",
            return_value={},
        ), patch.object(
            BIAS参数调优,
            "precompute_volatility_membership",
            return_value={
                method: {signal_date: frozenset(("INDEX",))}
                for method in BIAS参数调优.SCORE_METHODS
            },
        ), patch.object(
            BIAS参数调优,
            "build_parameter_targets",
            return_value={signal_date: SimpleNamespace(members=())},
        ) as build_targets, patch.object(
            BIAS参数调优,
            "run_total_return",
            side_effect=lambda *_: next(returns),
        ) as run_return:
            matrices, _ = BIAS参数调优.run_grid(
                backtest,
                (2,),
                (0.09,),
            )

        self.assertEqual(build_targets.call_count, 2)
        self.assertEqual(run_return.call_count, 4)
        self.assertEqual(
            set(matrices),
            {
                ("return_r2", "close"),
                ("return_r2", "next_day_vwap"),
                ("return_vol", "close"),
                ("return_vol", "next_day_vwap"),
            },
        )
        self.assertEqual(
            [float(matrices[key][0, 0]) for key in matrices],
            [0.1, 0.2, 0.3, 0.4],
        )
        backtest.load_filter_price_history.assert_called_once_with(
            {"INDEX"},
            Path("selected_index_prices"),
        )

    def test_main_routes_one_case_without_expanding_the_full_matrix(
        self,
    ) -> None:
        settings = make_settings(
            update_frequency="monthly",
            filter_profile="bias_and_volatility",
            correlation_threshold=0.8,
            trend_window=40,
            stagger_days=6,
        )
        backtest = SimpleNamespace()
        matrices = {
            key: np.array([[0.1]])
            for key in BIAS参数调优.OUTPUT_FILES
        }
        output_paths = [
            settings.output_directory / filename
            for filename in BIAS参数调优.OUTPUT_FILES.values()
        ]

        with patch.object(
            BIAS参数调优,
            "build_tuning_settings",
            return_value=settings,
        ) as build_settings, patch.object(
            BIAS参数调优,
            "load_backtest_module",
            return_value=backtest,
        ) as load_backtest, patch.object(
            BIAS参数调优,
            "validate_grid",
        ) as validate, patch.object(
            BIAS参数调优,
            "run_grid",
            return_value=(matrices, {}),
        ) as run_grid, patch.object(
            BIAS参数调优,
            "save_heatmaps",
            return_value=output_paths,
        ) as save_heatmaps:
            BIAS参数调优.main(
                "monthly",
                "bias_and_volatility",
                0.8,
                40,
                6,
            )

        build_settings.assert_called_once_with(
            "monthly",
            "bias_and_volatility",
            0.8,
            40,
            6,
        )
        load_backtest.assert_called_once_with(settings)
        validate.assert_called_once_with(backtest, settings)
        run_grid.assert_called_once_with(
            backtest,
            BIAS参数调优.BIAS_WINDOWS,
            BIAS参数调优.BIAS_UPPERS,
        )
        save_heatmaps.assert_called_once_with(
            backtest,
            matrices,
            BIAS参数调优.BIAS_WINDOWS,
            BIAS参数调优.BIAS_UPPERS,
            settings,
        )

    def test_heatmaps_keep_the_four_existing_file_names(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            output_directory = Path(temporary_directory)
            settings = make_settings(root=output_directory)
            backtest = SimpleNamespace(
                BIAS_WINDOW=36,
                BIAS_UPPER=0.09,
                VOL_KEEP_TOP_RATIO=0.55,
                VOL_RETURN_DAYS=18,
                CLUSTER_CORRELATION_THRESHOLD=0.9,
                TREND_WINDOW=20,
                TOP_PERCENT=0.10,
                ACCOUNT_COUNT=1,
                TRANSACTION_COST_RATE=0.001,
            )
            matrices = {
                key: np.array([[0.1]])
                for key in BIAS参数调优.OUTPUT_FILES
            }

            paths = BIAS参数调优.save_heatmaps(
                backtest,
                matrices,
                (36,),
                (0.09,),
                settings,
            )

            self.assertEqual(
                {path.name for path in paths},
                set(BIAS参数调优.OUTPUT_FILES.values()),
            )
            self.assertTrue(all(path.exists() for path in paths))
            self.assertTrue(
                all(path.parent == settings.output_directory for path in paths)
            )


if __name__ == "__main__":
    unittest.main()
