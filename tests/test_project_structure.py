from pathlib import Path
import unittest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PACKAGE_DIRECTORIES = (
    PROJECT_ROOT / "scripts",
    PROJECT_ROOT / "scripts" / "ETF数据下载",
    PROJECT_ROOT / "scripts" / "ETF池筛选",
    PROJECT_ROOT / "scripts" / "ETF趋势策略回测",
    PROJECT_ROOT / "scripts" / "策略参数调优",
)


class ProjectStructureTests(unittest.TestCase):
    def test_requirements_file_exists(self) -> None:
        self.assertTrue((PROJECT_ROOT / "requirements.txt").is_file())

    def test_package_markers_exist(self) -> None:
        for package_directory in PACKAGE_DIRECTORIES:
            with self.subTest(package_directory=package_directory):
                self.assertTrue((package_directory / "__init__.py").is_file())


if __name__ == "__main__":
    unittest.main()
