from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from 实验配置 import (
    CORRELATION_THRESHOLDS,
    TREND_WINDOWS,
    UPDATE_FREQUENCIES,
)
import 运行全部数据准备


class BatchDataPreparationRunnerTests(unittest.TestCase):
    def test_plan_has_forty_six_unique_tasks_in_dependency_order(self) -> None:
        tasks = tuple(运行全部数据准备.iter_preparation_tasks())

        expected_tasks = (
            tuple(
                运行全部数据准备.PreparationTask(
                    stage=运行全部数据准备.ETF_SELECTION_STAGE,
                    update_frequency=frequency,
                )
                for frequency in UPDATE_FREQUENCIES
            )
            + tuple(
                运行全部数据准备.PreparationTask(
                    stage=运行全部数据准备.INDEX_RETURN_STAGE,
                    update_frequency=frequency,
                )
                for frequency in UPDATE_FREQUENCIES
            )
            + tuple(
                运行全部数据准备.PreparationTask(
                    stage=运行全部数据准备.INDEX_CLUSTER_STAGE,
                    update_frequency=frequency,
                    correlation_threshold=threshold,
                )
                for frequency in UPDATE_FREQUENCIES
                for threshold in CORRELATION_THRESHOLDS
            )
            + tuple(
                运行全部数据准备.PreparationTask(
                    stage=运行全部数据准备.TREND_PRICE_STAGE,
                    update_frequency=frequency,
                    correlation_threshold=threshold,
                )
                for frequency in UPDATE_FREQUENCIES
                for threshold in CORRELATION_THRESHOLDS
            )
            + tuple(
                运行全部数据准备.PreparationTask(
                    stage=运行全部数据准备.TREND_FACTOR_STAGE,
                    update_frequency=frequency,
                    correlation_threshold=threshold,
                    trend_window=window,
                )
                for frequency in UPDATE_FREQUENCIES
                for threshold in CORRELATION_THRESHOLDS
                for window in TREND_WINDOWS
            )
        )

        self.assertEqual(len(tasks), 46)
        self.assertEqual(len(set(tasks)), 46)
        self.assertEqual(tasks, expected_tasks)
        self.assertEqual(
            [task.stage for task in tasks],
            [运行全部数据准备.ETF_SELECTION_STAGE] * 2
            + [运行全部数据准备.INDEX_RETURN_STAGE] * 2
            + [运行全部数据准备.INDEX_CLUSTER_STAGE] * 6
            + [运行全部数据准备.TREND_PRICE_STAGE] * 6
            + [运行全部数据准备.TREND_FACTOR_STAGE] * 30,
        )
        expected_factor_dimensions = {
            (frequency, threshold, window)
            for frequency in UPDATE_FREQUENCIES
            for threshold in CORRELATION_THRESHOLDS
            for window in TREND_WINDOWS
        }
        actual_factor_dimensions = {
            (
                task.update_frequency,
                task.correlation_threshold,
                task.trend_window,
            )
            for task in tasks
            if task.stage == 运行全部数据准备.TREND_FACTOR_STAGE
        }
        self.assertEqual(actual_factor_dimensions, expected_factor_dimensions)

    def test_main_routes_all_tasks_to_the_five_existing_scripts(self) -> None:
        actual_calls: list[tuple[str, dict[str, object]]] = []

        def record(stage: str):
            return lambda **values: actual_calls.append((stage, values))

        with TemporaryDirectory() as temporary_directory:
            etf_file = Path(temporary_directory) / "etf_data.csv"
            etf_file.write_bytes(b"data")
            with patch.object(
                运行全部数据准备,
                "ETF_DATA_FILE",
                etf_file,
            ), patch.object(
                运行全部数据准备.ETF初筛,
                "main",
                side_effect=record(运行全部数据准备.ETF_SELECTION_STAGE),
            ), patch.object(
                运行全部数据准备.指数收益率准备,
                "main",
                side_effect=record(运行全部数据准备.INDEX_RETURN_STAGE),
            ), patch.object(
                运行全部数据准备.指数聚类筛选,
                "main",
                side_effect=record(运行全部数据准备.INDEX_CLUSTER_STAGE),
            ), patch.object(
                运行全部数据准备.趋势行情准备,
                "main",
                side_effect=record(运行全部数据准备.TREND_PRICE_STAGE),
            ), patch.object(
                运行全部数据准备.趋势因子计算,
                "main",
                side_effect=record(运行全部数据准备.TREND_FACTOR_STAGE),
            ), patch("builtins.print"):
                summary = 运行全部数据准备.main()

        expected_calls: list[tuple[str, dict[str, object]]] = []
        for task in 运行全部数据准备.iter_preparation_tasks():
            values: dict[str, object] = {
                "update_frequency": task.update_frequency,
            }
            if task.correlation_threshold is not None:
                values["correlation_threshold"] = task.correlation_threshold
            if task.trend_window is not None:
                values["trend_window"] = task.trend_window
            expected_calls.append((task.stage, values))

        self.assertEqual(actual_calls, expected_calls)
        self.assertEqual(
            summary,
            运行全部数据准备.PreparationRunSummary(
                total_tasks=46,
                completed_tasks=46,
            ),
        )

    def test_missing_etf_data_fails_before_first_task(self) -> None:
        task = next(运行全部数据准备.iter_preparation_tasks())
        with TemporaryDirectory() as temporary_directory:
            missing_file = Path(temporary_directory) / "missing.csv"
            with patch.object(
                运行全部数据准备,
                "ETF_DATA_FILE",
                missing_file,
            ), patch.object(
                运行全部数据准备,
                "run_task",
            ) as run_task:
                with self.assertRaisesRegex(
                    FileNotFoundError,
                    "ETF总表缺失或为空",
                ):
                    运行全部数据准备.run_tasks((task,))

        run_task.assert_not_called()

    def test_empty_etf_data_fails_before_first_task(self) -> None:
        task = next(运行全部数据准备.iter_preparation_tasks())
        with TemporaryDirectory() as temporary_directory:
            empty_file = Path(temporary_directory) / "etf_data.csv"
            empty_file.touch()
            with patch.object(
                运行全部数据准备,
                "ETF_DATA_FILE",
                empty_file,
            ), patch.object(
                运行全部数据准备,
                "run_task",
            ) as run_task:
                with self.assertRaisesRegex(
                    FileNotFoundError,
                    "ETF总表缺失或为空",
                ):
                    运行全部数据准备.run_tasks((task,))

        run_task.assert_not_called()

    def test_failure_stops_before_later_tasks_and_preserves_cause(self) -> None:
        tasks = tuple(运行全部数据准备.iter_preparation_tasks())[-3:]
        original_error = ValueError("broken preparation")
        with TemporaryDirectory() as temporary_directory:
            etf_file = Path(temporary_directory) / "etf_data.csv"
            etf_file.write_bytes(b"data")
            with patch.object(
                运行全部数据准备,
                "ETF_DATA_FILE",
                etf_file,
            ), patch.object(
                运行全部数据准备,
                "run_task",
                side_effect=(None, original_error, None, None),
            ) as run_task, patch("builtins.print"):
                with self.assertRaises(RuntimeError) as raised:
                    运行全部数据准备.run_tasks(tasks)

        self.assertIs(raised.exception.__cause__, original_error)
        self.assertIn("第 2/3 个任务失败", str(raised.exception))
        self.assertIn(
            "趋势因子计算 / daily / 阈值0.9 / 窗口40日",
            str(raised.exception),
        )
        self.assertEqual(run_task.call_count, 2)
        self.assertEqual(
            tuple(call.args[0] for call in run_task.call_args_list),
            tasks[:2],
        )

    def test_keyboard_interrupt_is_not_swallowed(self) -> None:
        task = next(运行全部数据准备.iter_preparation_tasks())
        with TemporaryDirectory() as temporary_directory:
            etf_file = Path(temporary_directory) / "etf_data.csv"
            etf_file.write_bytes(b"data")
            with patch.object(
                运行全部数据准备,
                "ETF_DATA_FILE",
                etf_file,
            ), patch.object(
                运行全部数据准备,
                "run_task",
                side_effect=KeyboardInterrupt,
            ), patch("builtins.print"):
                with self.assertRaises(KeyboardInterrupt):
                    运行全部数据准备.run_tasks((task,))

    def test_invalid_empty_duplicate_and_out_of_order_plans_fail_early(
        self,
    ) -> None:
        tasks = tuple(运行全部数据准备.iter_preparation_tasks())
        invalid = 运行全部数据准备.PreparationTask(
            stage="unknown",
            update_frequency=UPDATE_FREQUENCIES[0],
        )
        out_of_order = (tasks[2], tasks[0])

        with patch.object(
            运行全部数据准备,
            "validate_source_data",
            side_effect=AssertionError("非法计划不应读取输入"),
        ):
            with self.assertRaises(ValueError):
                运行全部数据准备.run_tasks(())
            with self.assertRaises(ValueError):
                运行全部数据准备.run_tasks((invalid,))
            with self.assertRaisesRegex(ValueError, "重复任务"):
                运行全部数据准备.run_tasks((tasks[0], tasks[0]))
            with self.assertRaisesRegex(ValueError, "依赖阶段"):
                运行全部数据准备.run_tasks(out_of_order)


if __name__ == "__main__":
    unittest.main()
