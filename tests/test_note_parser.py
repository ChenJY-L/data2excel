import ast
import re
import unittest
from pathlib import Path


SOURCE_PATH = Path(__file__).resolve().parents[1] / "src" / "data2excel.py"


class NoteParserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        module = ast.parse(SOURCE_PATH.read_text(encoding="utf-8"))
        dialog = next(node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "GUI_Dialog")
        method = next(node for node in dialog.body if isinstance(node, ast.FunctionDef) and node.name == "parseText")
        namespace = {"re": re}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(SOURCE_PATH), "exec"), namespace)
        cls.parse_text = staticmethod(namespace["parseText"])

    def test_blood_glucose_aliases_parse_values_and_times(self):
        for name in ("血糖", "血糖值"):
            for template in ("{}", "{}:", "{}：", "[{}]"):
                header = template.format(name)
                with self.subTest(header=header):
                    result = self.parse_text(None, f"  {header}  \n08:30:5.6\n09:15:6.1")
                    self.assertEqual(result["blood_glucose"], [
                        {"time": 8.5 / 24, "value": 5.6},
                        {"time": 9.25 / 24, "value": 6.1},
                    ])
                    self.assertEqual(result["schedule"], [])
                    self.assertIsNone(result["baseline_cycle"])

    def test_baseline_aliases_parse_values_and_time_ranges(self):
        for name in ("基准周期", "基准"):
            for template in ("{}", "{}:", "{}：", "[{}]"):
                header = template.format(name)
                with self.subTest(header=header):
                    result = self.parse_text(None, f"{header}\n08:00-09:00:3\n09:00:5")
                    self.assertEqual(result["baseline_cycle"], [
                        {"time": (8 / 24, 9 / 24), "value": 3},
                        {"time": 9 / 24, "value": 5},
                    ])
                    self.assertEqual(result["schedule"], [])
                    self.assertIsNone(result["blood_glucose"])

    def test_aliases_end_remark_block_and_switch_sections(self):
        result = self.parse_text(None, "备注:\n保持静坐\n血糖值:\n08:30:5.6\n基准:\n08:00-09:00:3")

        self.assertEqual(result["remark"], "保持静坐")
        self.assertEqual(result["blood_glucose"], [{"time": 8.5 / 24, "value": 5.6}])
        self.assertEqual(result["baseline_cycle"], [{"time": (8 / 24, 9 / 24), "value": 3}])
        self.assertEqual(result["schedule"], [])

    def test_ordinary_remark_text_is_not_mistaken_for_header(self):
        result = self.parse_text(None, "备注:\n血糖波动较小\n基准设置沿用上次\n血糖值:\n08:30:5.6")

        self.assertEqual(result["remark"], "血糖波动较小\n基准设置沿用上次")
        self.assertEqual(result["blood_glucose"], [{"time": 8.5 / 24, "value": 5.6}])
        self.assertIsNone(result["baseline_cycle"])


if __name__ == "__main__":
    unittest.main()
