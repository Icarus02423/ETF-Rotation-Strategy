import unittest

from 实验配置 import FILTER_PROFILE_SETTINGS, iter_experiment_cases


class ExperimentConfigTests(unittest.TestCase):
    def test_filter_profiles_have_expected_switches(self) -> None:
        expected_settings = {
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
        self.assertEqual(FILTER_PROFILE_SETTINGS, expected_settings)

    def test_experiment_matrix_contains_2400_cases(self) -> None:
        cases = tuple(iter_experiment_cases())
        self.assertEqual(len(cases), 2400)

    def test_experiment_cases_are_unique(self) -> None:
        cases = tuple(iter_experiment_cases())
        self.assertEqual(len(set(cases)), len(cases))


if __name__ == "__main__":
    unittest.main()
