import ast
import datetime
import tempfile
import unittest
from pathlib import Path

import numpy as np
from scipy.signal import butter, sosfiltfilt


SOURCE_PATH = Path(__file__).resolve().parents[1] / "src" / "data2excel.py"


def load_filter_method():
    module = ast.parse(SOURCE_PATH.read_text(encoding="utf-8"))
    dialog = next(node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "GUI_Dialog")
    method = next(node for node in dialog.body if isinstance(node, ast.FunctionDef) and node.name == "filter_doppler_data")
    namespace = {"datetime": datetime, "np": np, "butter": butter, "sosfiltfilt": sosfiltfilt, "Path": Path}
    exec(compile(ast.Module(body=[method], type_ignores=[]), str(SOURCE_PATH), "exec"), namespace)
    return namespace["filter_doppler_data"]


class DopplerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.filter_doppler_data = staticmethod(load_filter_method())

    def test_filters_all_channels_and_interpolates_at_spectrum_times(self):
        start = datetime.datetime(2026, 10, 12, 15, 19, 34)
        excel_start = (start - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
        seconds = np.arange(1200)
        signals = np.column_stack((
            10 + 4 * np.sin(2 * np.pi * 0.2 * seconds),
            20 + 6 * np.sin(2 * np.pi * 0.2 * seconds),
            30 + 8 * np.sin(2 * np.pi * 0.2 * seconds),
        ))
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "doppler.txt"
            rows = [
                "moorVMS-PC Text Data File v4.1",
                "Monitoring Date: 12/10/2026",
                "Monitoring Time: 15:19:34",
                "Elapsed Time\tFlux1\tConc\tSpeed",
            ]
            rows.extend(
                f"{s // 3600:02d}:{s // 60 % 60:02d}:{s % 60:02d}\t"
                + "\t".join(f"{v:.6f}" for v in values)
                for s, values in zip(seconds, signals)
            )
            path.write_text("\n".join(rows), encoding="utf-8")
            spectrum_times = excel_start + np.array([-1, 300.5, 600.5, 1200]) / 86400
            result = self.filter_doppler_data(None, path, spectrum_times.reshape(-1, 1))

        self.assertEqual(result.shape, (4, 3))
        self.assertTrue(np.isnan(result[[0, 3]]).all())
        np.testing.assert_allclose(result[1:3], [[10, 20, 30], [10, 20, 30]], atol=0.1)


if __name__ == "__main__":
    unittest.main()
