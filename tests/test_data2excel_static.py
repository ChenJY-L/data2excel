import unittest
from pathlib import Path
import re


DATA2EXCEL_PATH = Path(__file__).resolve().parents[1] / "src" / "data2excel.py"


class Data2ExcelStaticTests(unittest.TestCase):
    def test_chart_rendering_does_not_duplicate_excel_chart_objects(self):
        source = DATA2EXCEL_PATH.read_text(encoding="utf-8")

        self.assertNotIn("chartApi.Parent.Duplicate()", source)
        self.assertNotIn("duplicated_chart", source)
        self.assertNotIn("duplicateCheckBox", source)

    def test_chart_rendering_supports_ring_source_charts(self):
        source = DATA2EXCEL_PATH.read_text(encoding="utf-8")

        self.assertIn("ring_source_match", source)
        self.assertIn(r"^Ring-?(\d+)$", source)
        self.assertIn("4 + 7 * (ring_num - 1)", source)

    def test_chart_config_json_can_be_selected_from_ui(self):
        source = DATA2EXCEL_PATH.read_text(encoding="utf-8")

        self.assertIn("self.ChartConfigSelect.clicked.connect(self.ChartConfigSelectF)", source)
        self.assertIn("self.chart_config_path", source)
        self.assertIn("load_chart_config(self.chart_config_path)", source)

    def test_excel_cleanup_falls_back_to_kill_when_quit_fails(self):
        source = DATA2EXCEL_PATH.read_text(encoding="utf-8")
        cleanup = source.split("# 清理Excel应用程序", 1)[1]

        self.assertRegex(
            cleanup,
            re.compile(
                r"try:\s+form\.xwapp\.quit\(\)\s+"
                r"except Exception:\s+try:\s+form\.xwapp\.kill\(\)\s+"
                r"except Exception:\s+pass"
            ),
        )


if __name__ == "__main__":
    unittest.main()
