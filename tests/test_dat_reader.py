import ast
import datetime
import os
import re
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import numpy as np


SOURCE_PATH = Path(__file__).resolve().parents[1] / "src" / "data2excel.py"


def load_processing_methods():
    # Load the actual methods without starting the Qt interface or Excel.
    module = ast.parse(SOURCE_PATH.read_text(encoding="utf-8"))
    dialog = next(node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "GUI_Dialog")
    methods = [node for node in dialog.body if isinstance(node, ast.FunctionDef) and node.name in ("load_and_validate_data", "DataProcess")]
    namespace = {"np": np, "os": os, "re": re, "datetime": datetime}
    exec(compile(ast.Module(body=methods, type_ignores=[]), str(SOURCE_PATH), "exec"), namespace)
    return namespace


class DatReaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.methods = load_processing_methods()

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.data_dir = Path(self.temp_dir.name)
        self.form = SimpleNamespace(
            Path=Mock(), Rings=Mock(), Status=Mock(), ErrorText=Mock(),
            Process=Mock(), xwapp=Mock(),
            GuiRefresh=lambda field, text: field.setPlainText(text),
        )
        self.form.Path.toPlainText.return_value = str(self.ring_path(1))
        self.form.Rings.currentText.return_value = "5 Rings"
        self.form.xwapp.books.open.side_effect = AssertionError("DAT reading must not call Excel")

    def ring_path(self, ring):
        return self.data_dir / f"measurement(校准后)-{ring}环.dat"

    def test_five_rings_preserve_values_and_order_without_excel(self):
        expected = []
        for ring in range(1, 6):
            # 2个周期，每周期6个波长；5列测量值和1列Excel时间。
            rows = [
                [ring * 0.1 + wave * 0.001 + sample * 0.0001 for sample in range(5)]
                + [46279.5 + cycle / 86400]
                for cycle in range(2)
                for wave in range(6)
            ]
            np.savetxt(self.ring_path(ring), rows, delimiter="\t", fmt="%.17g")
            expected.append(rows)

        values, rings, chinese_names, path = self.methods["load_and_validate_data"](self.form)

        np.testing.assert_array_equal(values, np.asarray(expected))
        self.assertEqual(values.shape, (5, 12, 6))
        self.assertEqual(rings, 5)
        self.assertTrue(chinese_names)
        self.assertEqual(Path(path), self.ring_path(1))
        self.form.xwapp.books.open.assert_not_called()

    def test_process_retains_full_exception_and_reenables_button(self):
        def failing_loader():
            raise RuntimeError("Excel connection failed during input reading")

        self.form.load_and_validate_data = failing_loader
        self.methods["DataProcess"](self.form)

        displayed_error = self.form.ErrorText.setPlainText.call_args.args[0]
        self.assertIn("Traceback (most recent call last)", displayed_error)
        self.assertIn("failing_loader", displayed_error)
        self.assertIn("RuntimeError: Excel connection failed during input reading", displayed_error)
        self.form.Process.setEnabled.assert_called_with(True)


if __name__ == "__main__":
    unittest.main()
