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
from scripts.策略参数调优 import 波动率参数调优


def make_settings(
    *,
    update_frequency: str = "daily",
    filter_profile: str = "volatility_only",
    correlation_threshold: float = 0.9,
    trend_window: int = 20,
    stagger_days: int = 1,
    root: Path = Path("test_outputs"),
) -> 波动率参数调优.VolatilityTuningSettings:
    return 波动率参数调优.VolatilityTuningSettings(
        update_frequency=update_frequency,
        filter_profile=filter_profile,
        correlation_threshold=correlation_threshold,
        trend_window=trend_window,
        stagger_days=stagger_days,
        factor_directory=root / "factors",
        index_price_directory=root / "index_prices",
        output_directory=root / "parameter_grid",
        bias_mode=(
            "upper" if filter_profile == "bias_and_volatility" else "none"
        ),
        vol_filter_enabled=True,
    )


def make_factor_member(index_code: str, signal_date: date) -> SimpleNamespace:
    return SimpleNamespace(
        signal_date=signal_date,
        index_code=index_code,
        index_name=f"指数{index_code}",
        trend_factor=1.0,
        window_return=0.1,
    )


class VolatilityParameterTuningTests(unittest.TestCase):
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

        profiles = ("volatility_only", "bias_and_volatility")
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
            波动率参数调优,
            "factor_dir",
            factor_path,
        ), patch.object(
            波动率参数调优,
            "index_price_dir",
            index_path,
        ), patch.object(
            波动率参数调优,
            "tuning_result_dir",
            output_path,
        ):
            for values in combinations:
                with self.subTest(values=values):
                    settings = 波动率参数调优.build_tuning_settings(*values)
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
                            f"volatility_parameter_grid/{threshold:g}/"
                            f"{window}/{stagger}"
                        ),
                    )

        self.assertEqual(factor_path.call_count, len(combinations))
        self.assertEqual(index_path.call_count, len(combinations))
        self.assertEqual(output_path.call_count, len(combinations))

    def test_profiles_map_to_volatility_and_bias_switches(self) -> None:
        volatility_only = 波动率参数调优.build_tuning_settings(
            "daily",
            "volatility_only",
            0.9,
            20,
            1,
        )
        combined = 波动率参数调优.build_tuning_settings(
            "monthly",
            "bias_and_volatility",
            0.8,
            40,
            10,
        )

        self.assertEqual(
            (volatility_only.bias_mode, volatility_only.vol_filter_enabled),
            ("none", True),
        )
        self.assertEqual(
            (combined.bias_mode, combined.vol_filter_enabled),
            ("upper", True),
        )

    def test_invalid_dimensions_fail_before_path_resolution(self) -> None:
        invalid_values = (
            ("weekly", "volatility_only", 0.9, 20, 1),
            ("daily", "base", 0.9, 20, 1),
            ("daily", "bias_only", 0.9, 20, 1),
            ("daily", "volatility_only", 0.75, 20, 1),
            ("daily", "volatility_only", 0.9, 30, 1),
            ("daily", "volatility_only", 0.9, 20, 0),
        )

        for values in invalid_values:
            with self.subTest(values=values), patch.object(
                波动率参数调优,
                "factor_dir",
                side_effect=AssertionError("不应计算因子路径"),
            ), patch.object(
                波动率参数调优,
                "index_price_dir",
                side_effect=AssertionError("不应计算行情路径"),
            ), patch.object(
                波动率参数调优,
                "tuning_result_dir",
                side_effect=AssertionError("不应计算输出路径"),
            ):
                with self.assertRaises(ValueError):
                    波动率参数调优.build_tuning_settings(*values)

    def test_backtest_module_receives_only_current_case_values(self) -> None:
        settings = make_settings(
            update_frequency="monthly",
            filter_profile="bias_and_volatility",
            correlation_threshold=0.7,
            trend_window=60,
            stagger_days=9,
        )
        backtest = SimpleNamespace()

        波动率参数调优.configure_backtest_module(backtest, settings)

        self.assertEqual(backtest.UPDATE_FREQUENCY, "monthly")
        self.assertEqual(backtest.FILTER_PROFILE, "bias_and_volatility")
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
        self.assertTrue(backtest.VOL_FILTER_ENABLED)
        self.assertEqual(backtest.BIAS_WINDOW, 36)
        self.assertEqual(backtest.BIAS_UPPER, 0.09)
        self.assertEqual(backtest.VOL_RETURN_DAYS, 18)
        self.assertEqual(backtest.VOL_KEEP_TOP_RATIO, 0.55)

    def test_real_backtest_module_loads_and_validates_both_profiles(self) -> None:
        cases = (
            ("daily", "volatility_only", 0.9, 20, 1),
            ("monthly", "bias_and_volatility", 0.7, 60, 10),
        )

        for values in cases:
            with self.subTest(values=values):
                settings = 波动率参数调优.build_tuning_settings(*values)
                backtest = 波动率参数调优.load_backtest_module(settings)
                波动率参数调优.validate_grid(backtest, settings)

                self.assertEqual(backtest.FACTOR_DIR, settings.factor_directory)
                self.assertEqual(
                    backtest.INDEX_RAW_DATA_DIR,
                    settings.index_price_directory,
                )
                self.assertEqual(backtest.BIAS_MODE, settings.bias_mode)
                self.assertTrue(backtest.VOL_FILTER_ENABLED)
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
            universes = 波动率参数调优.load_factor_universes(backtest)

        signal_date = date(2021, 1, 4)
        self.assertEqual(
            universes["return_r2"][signal_date][0].trend_factor,
            0.2,
        )
        self.assertEqual(
            universes["return_vol"][signal_date][0].trend_factor,
            0.3,
        )

    def test_metrics_use_exact_returns_and_ignore_future_prices(self) -> None:
        signal_date = date(2021, 2, 5)
        member = make_factor_member("A", signal_date)
        factor_universes = {
            method: {signal_date: (member,)}
            for method in 波动率参数调优.SCORE_METHODS
        }
        dates = [
            signal_date - timedelta(days=2 - offset)
            for offset in range(3)
        ]
        prices = [110.0, 105.0, 120.0]
        volatility_only = SimpleNamespace(
            BIAS_MODE="none",
            BIAS_WINDOW=36,
            MAX_PRICE_STALENESS_CALENDAR_DAYS=7,
        )

        original = 波动率参数调优.precompute_filter_metrics(
            volatility_only,
            factor_universes,
            {"A": (dates, prices)},
            (2,),
        )
        with_future = 波动率参数调优.precompute_filter_metrics(
            volatility_only,
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
        self.assertIsNone(metric.bias)
        self.assertAlmostEqual(
            metric.volatilities[0],
            statistics.stdev((105.0 / 110.0 - 1.0, 120.0 / 105.0 - 1.0)),
        )
        with self.assertRaisesRegex(ValueError, "历史不足3个收盘价"):
            波动率参数调优.precompute_filter_metrics(
                volatility_only,
                factor_universes,
                {"A": (dates[-2:], prices[-2:])},
                (2,),
            )

    def test_combined_metrics_include_signal_close_but_not_future(self) -> None:
        signal_date = date(2021, 2, 5)
        member = make_factor_member("A", signal_date)
        factor_universes = {
            method: {signal_date: (member,)}
            for method in 波动率参数调优.SCORE_METHODS
        }
        dates = [
            signal_date - timedelta(days=35 - offset)
            for offset in range(36)
        ]
        prices = [100.0 + offset for offset in range(36)]
        combined = SimpleNamespace(
            BIAS_MODE="upper",
            BIAS_WINDOW=36,
            MAX_PRICE_STALENESS_CALENDAR_DAYS=7,
        )

        original = 波动率参数调优.precompute_filter_metrics(
            combined,
            factor_universes,
            {"A": (dates, prices)},
            (2,),
        )
        with_future = 波动率参数调优.precompute_filter_metrics(
            combined,
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
            metric.bias,
            prices[-1] / statistics.mean(prices) - 1.0,
        )
        self.assertAlmostEqual(
            metric.volatilities[0],
            statistics.stdev(
                (
                    prices[-2] / prices[-3] - 1.0,
                    prices[-1] / prices[-2] - 1.0,
                )
            ),
        )

    def test_each_window_builds_its_own_high_volatility_order(self) -> None:
        signal_date = date(2021, 1, 4)
        members = tuple(
            make_factor_member(code, signal_date) for code in ("A", "B", "C")
        )
        factor_universes = {
            method: {signal_date: members}
            for method in 波动率参数调优.SCORE_METHODS
        }
        metrics = {
            (signal_date, "A"): 波动率参数调优.FilterMetrics((0.3, 0.1), None),
            (signal_date, "B"): 波动率参数调优.FilterMetrics((0.2, 0.4), None),
            (signal_date, "C"): 波动率参数调优.FilterMetrics((0.2, 0.3), None),
        }
        backtest = SimpleNamespace(VOL_FILTER_MODE="high")

        first = 波动率参数调优.build_volatility_orders(
            backtest,
            factor_universes,
            metrics,
            0,
        )
        second = 波动率参数调优.build_volatility_orders(
            backtest,
            factor_universes,
            metrics,
            1,
        )

        for method in 波动率参数调优.SCORE_METHODS:
            self.assertEqual(first[method][signal_date], ("A", "B", "C"))
            self.assertEqual(second[method][signal_date], ("B", "C", "A"))

    def test_volatility_ranking_precedes_bias_without_backfill(self) -> None:
        signal_date = date(2021, 1, 4)
        members = tuple(
            make_factor_member(code, signal_date) for code in ("A", "B", "C")
        )
        metrics = {
            (signal_date, "A"): 波动率参数调优.FilterMetrics((0.3,), 0.20),
            (signal_date, "B"): 波动率参数调优.FilterMetrics((0.2,), 0.01),
            (signal_date, "C"): 波动率参数调优.FilterMetrics((0.1,), 0.01),
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

        def build_target(bias_mode: str, keep_ratio: float) -> object:
            backtest = SimpleNamespace(
                BIAS_MODE=bias_mode,
                TOP_PERCENT=1.0,
                passes_bias_filter=lambda value: value <= 0.09,
                TargetMember=交易策略回测.TargetMember,
                DailyTarget=交易策略回测.DailyTarget,
            )
            return 波动率参数调优.build_parameter_targets(
                backtest,
                "return_r2",
                {signal_date: members},
                candidates,
                {signal_date: ("A", "B", "C")},
                metrics,
                keep_ratio,
                (signal_date,),
            )[signal_date]

        volatility_only = build_target("none", 0.5)
        combined = build_target("upper", 0.5)
        combined_at_100 = build_target("upper", 1.0)

        self.assertEqual(
            tuple(member.index_code for member in volatility_only.members),
            ("A", "B"),
        )
        self.assertEqual(
            tuple(member.index_code for member in combined.members),
            ("B",),
        )
        self.assertEqual(combined.filtered_index_count, 2)
        self.assertEqual(
            tuple(member.index_code for member in combined_at_100.members),
            ("B", "C"),
        )
        self.assertEqual(combined_at_100.filtered_index_count, 1)

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
                    result = 波动率参数调优.run_total_return(
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

                波动率参数调优.run_total_return(
                    backtest,
                    mode,
                    trading_dates,
                    targets,
                    prices,
                )

                self.assertEqual(observed, list(expected))

    def test_small_grid_propagates_both_axes_to_all_four_results(
        self,
    ) -> None:
        signal_date = date(2021, 1, 4)
        member = make_factor_member("INDEX", signal_date)
        factor_universes = {
            method: {signal_date: (member,)}
            for method in 波动率参数调优.SCORE_METHODS
        }
        contexts = {
            method: 波动率参数调优.MethodContext(
                trading_dates=(signal_date,),
                benchmark_end_date=signal_date,
            )
            for method in 波动率参数调优.SCORE_METHODS
        }
        backtest = SimpleNamespace(
            INDEX_RAW_DATA_DIR=Path("selected_index_prices"),
            build_daily_targets=Mock(return_value=({}, (signal_date,))),
            load_filter_price_history=Mock(return_value={}),
            load_selected_prices=Mock(return_value={}),
        )
        order_positions: list[int] = []
        target_parameters: list[tuple[str, int, float]] = []

        def build_orders(
            _: object,
            __: object,
            ___: object,
            window_position: int,
        ) -> dict[str, dict[date, tuple[str, ...]]]:
            order_positions.append(window_position)
            return {
                method: {signal_date: (f"WINDOW_{window_position}",)}
                for method in 波动率参数调优.SCORE_METHODS
            }

        def build_targets(
            _: object,
            method: str,
            __: object,
            ___: object,
            volatility_orders: dict[date, tuple[str, ...]],
            ____: object,
            keep_ratio: float,
            _____: object,
        ) -> dict[date, SimpleNamespace]:
            marker = volatility_orders[signal_date][0]
            window_position = int(marker.rsplit("_", 1)[1])
            target_parameters.append((method, window_position, keep_ratio))
            return {
                signal_date: SimpleNamespace(
                    members=(),
                    method=method,
                    window_position=window_position,
                    ratio_position=(0 if keep_ratio == 0.5 else 1),
                )
            }

        def run_return(
            _: object,
            mode: str,
            __: object,
            targets: dict[date, SimpleNamespace],
            ___: object,
        ) -> float:
            target = targets[signal_date]
            method_base = 0 if target.method == "return_r2" else 100
            mode_base = 0 if mode == "close" else 50
            return float(
                method_base
                + mode_base
                + target.window_position * 10
                + target.ratio_position
            )

        with patch.object(
            波动率参数调优,
            "load_factor_universes",
            return_value=factor_universes,
        ), patch.object(
            波动率参数调优,
            "build_universal_selections",
            return_value={},
        ), patch.object(
            波动率参数调优,
            "prepare_method_contexts",
            return_value=contexts,
        ), patch.object(
            波动率参数调优,
            "precompute_filter_metrics",
            return_value={},
        ), patch.object(
            波动率参数调优,
            "build_volatility_orders",
            side_effect=build_orders,
        ) as build_orders, patch.object(
            波动率参数调优,
            "build_parameter_targets",
            side_effect=build_targets,
        ) as build_targets, patch.object(
            波动率参数调优,
            "run_total_return",
            side_effect=run_return,
        ) as run_return:
            matrices, _ = 波动率参数调优.run_grid(
                backtest,
                (2, 3),
                (0.5, 1.0),
            )

        self.assertEqual(build_orders.call_count, 2)
        self.assertEqual(order_positions, [0, 1])
        self.assertEqual(build_targets.call_count, 8)
        self.assertEqual(
            target_parameters,
            [
                (method, window_position, keep_ratio)
                for window_position in (0, 1)
                for keep_ratio in (0.5, 1.0)
                for method in 波动率参数调优.SCORE_METHODS
            ],
        )
        self.assertEqual(run_return.call_count, 16)
        self.assertEqual(
            set(matrices),
            {
                ("return_r2", "close"),
                ("return_r2", "next_day_vwap"),
                ("return_vol", "close"),
                ("return_vol", "next_day_vwap"),
            },
        )
        for method in 波动率参数调优.SCORE_METHODS:
            for mode in 波动率参数调优.TRADE_MODES:
                method_base = 0 if method == "return_r2" else 100
                mode_base = 0 if mode == "close" else 50
                np.testing.assert_array_equal(
                    matrices[(method, mode)],
                    np.array(
                        [
                            [method_base + mode_base, method_base + mode_base + 10],
                            [method_base + mode_base + 1, method_base + mode_base + 11],
                        ],
                        dtype=float,
                    ),
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
        backtest = SimpleNamespace(
            BIAS_WINDOW=36,
            BIAS_MODE="upper",
            BIAS_UPPER=0.09,
        )
        matrices = {
            key: np.array([[0.1]])
            for key in 波动率参数调优.OUTPUT_FILES
        }
        output_paths = [
            settings.output_directory / filename
            for filename in 波动率参数调优.OUTPUT_FILES.values()
        ]

        with patch.object(
            波动率参数调优,
            "build_tuning_settings",
            return_value=settings,
        ) as build_settings, patch.object(
            波动率参数调优,
            "load_backtest_module",
            return_value=backtest,
        ) as load_backtest, patch.object(
            波动率参数调优,
            "validate_grid",
        ) as validate, patch.object(
            波动率参数调优,
            "run_grid",
            return_value=(matrices, {}),
        ) as run_grid, patch.object(
            波动率参数调优,
            "save_heatmaps",
            return_value=output_paths,
        ) as save_heatmaps:
            波动率参数调优.main(
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
            波动率参数调优.VOL_WINDOWS,
            波动率参数调优.VOL_KEEP_RATIOS,
        )
        save_heatmaps.assert_called_once_with(
            backtest,
            matrices,
            波动率参数调优.VOL_WINDOWS,
            波动率参数调优.VOL_KEEP_RATIOS,
            settings,
        )
        self.assertEqual(len(波动率参数调优.VOL_WINDOWS), 59)
        self.assertEqual(len(波动率参数调优.VOL_KEEP_RATIOS), 20)

    def test_heatmaps_keep_the_four_existing_file_names(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            output_directory = Path(temporary_directory)
            settings = make_settings(root=output_directory)
            backtest = SimpleNamespace(
                VOL_RETURN_DAYS=18,
                VOL_KEEP_TOP_RATIO=0.55,
                BIAS_WINDOW=36,
                BIAS_MODE="none",
                BIAS_UPPER=0.09,
                VOL_FILTER_MODE="high",
                CLUSTER_CORRELATION_THRESHOLD=0.9,
                TREND_WINDOW=20,
                TOP_PERCENT=0.10,
                ACCOUNT_COUNT=1,
                TRANSACTION_COST_RATE=0.001,
            )
            matrices = {
                key: np.array([[0.1]])
                for key in 波动率参数调优.OUTPUT_FILES
            }

            paths = 波动率参数调优.save_heatmaps(
                backtest,
                matrices,
                (18,),
                (0.55,),
                settings,
            )

            self.assertEqual(
                {path.name for path in paths},
                set(波动率参数调优.OUTPUT_FILES.values()),
            )
            self.assertTrue(all(path.exists() for path in paths))
            self.assertTrue(
                all(path.parent == settings.output_directory for path in paths)
            )


if __name__ == "__main__":
    unittest.main()
