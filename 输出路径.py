from pathlib import Path

from 实验配置 import ExperimentCase


PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = PROJECT_ROOT / "outputs"

TUNING_TYPES_BY_PROFILE = {
    "bias_only": ("bias_parameter_grid",),
    "volatility_only": ("volatility_parameter_grid",),
    "bias_and_volatility": (
        "bias_parameter_grid",
        "volatility_parameter_grid",
    ),
}


def _threshold_directory(correlation_threshold: float) -> str:
    return f"threshold_{correlation_threshold:g}"


def _window_directory(window: int) -> str:
    return f"window_{window}"


def _stagger_directory(stagger_days: int) -> str:
    return f"staggered_{stagger_days}d"


def etf_pool_initial_dir(update_frequency: str) -> Path:
    return OUTPUT_ROOT / "etf_pool" / update_frequency / "initial"


def index_return_dir(update_frequency: str, return_window: int = 60) -> Path:
    return (
        OUTPUT_ROOT
        / "etf_pool"
        / update_frequency
        / "index_returns"
        / _window_directory(return_window)
    )


def cluster_report_dir(
    update_frequency: str,
    correlation_threshold: float,
) -> Path:
    return (
        OUTPUT_ROOT
        / "etf_pool"
        / update_frequency
        / "clusters"
        / _threshold_directory(correlation_threshold)
        / "reports"
    )


def index_price_dir(
    update_frequency: str,
    correlation_threshold: float,
) -> Path:
    return (
        OUTPUT_ROOT
        / "etf_trend_strategy"
        / update_frequency
        / _threshold_directory(correlation_threshold)
        / "index_prices"
    )


def factor_dir(
    update_frequency: str,
    correlation_threshold: float,
    trend_window: int,
) -> Path:
    return (
        OUTPUT_ROOT
        / "etf_trend_strategy"
        / update_frequency
        / _threshold_directory(correlation_threshold)
        / "factors"
        / _window_directory(trend_window)
    )


def backtest_result_dir(
    experiment_case: ExperimentCase,
    execution_price: str,
) -> Path:
    return (
        OUTPUT_ROOT
        / "etf_trend_strategy"
        / experiment_case.update_frequency
        / _threshold_directory(experiment_case.correlation_threshold)
        / "backtest"
        / experiment_case.filter_profile
        / _window_directory(experiment_case.trend_window)
        / experiment_case.trend_factor
        / _stagger_directory(experiment_case.stagger_days)
        / execution_price
    )


def tuning_result_dir(
    update_frequency: str,
    filter_profile: str,
    tuning_type: str,
    correlation_threshold: float,
    trend_window: int,
    stagger_days: int,
) -> Path:
    allowed_tuning_types = TUNING_TYPES_BY_PROFILE.get(filter_profile, ())
    if tuning_type not in allowed_tuning_types:
        raise ValueError(
            f"过滤状态 {filter_profile!r} 不允许使用调优类型 {tuning_type!r}"
        )

    return (
        OUTPUT_ROOT
        / "parameter_grid"
        / update_frequency
        / filter_profile
        / tuning_type
        / _threshold_directory(correlation_threshold)
        / _window_directory(trend_window)
        / _stagger_directory(stagger_days)
    )
