from __future__ import annotations

from dataclasses import dataclass
from itertools import product
from typing import Iterator


UPDATE_FREQUENCIES = ("monthly", "daily")

FILTER_PROFILE_SETTINGS = {
    "base": {
        "bias_mode": "none",
        "vol_filter_enabled": False,
    },
    "bias_only": {
        "bias_mode": "upper",
        "vol_filter_enabled": False,
    },
    "volatility_only": {
        "bias_mode": "none",
        "vol_filter_enabled": True,
    },
    "bias_and_volatility": {
        "bias_mode": "upper",
        "vol_filter_enabled": True,
    },
}

CORRELATION_THRESHOLDS = (0.7, 0.8, 0.9)
TREND_FACTORS = ("return_r2", "return_vol")
TREND_WINDOWS = (10, 15, 20, 40, 60)
STAGGER_DAYS = tuple(range(1, 11))


@dataclass(frozen=True)
class ExperimentCase:
    update_frequency: str
    filter_profile: str
    correlation_threshold: float
    trend_factor: str
    trend_window: int
    stagger_days: int


def iter_experiment_cases() -> Iterator[ExperimentCase]:
    for values in product(
        UPDATE_FREQUENCIES,
        FILTER_PROFILE_SETTINGS,
        CORRELATION_THRESHOLDS,
        TREND_FACTORS,
        TREND_WINDOWS,
        STAGGER_DAYS,
    ):
        yield ExperimentCase(*values)
