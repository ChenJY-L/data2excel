import sys
import unittest
from pathlib import Path


SRC_DIR = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(SRC_DIR))

from chart_config import (  # noqa: E402
    build_chart_plan,
    get_bounds,
    load_chart_config,
    resolve_header_column,
    resolve_secondary_axis_value,
    validate_chart_config,
)


class ChartConfigTests(unittest.TestCase):
    def test_build_chart_plan_uses_classic_list_and_disables_temperature(self):
        bounds = {"left": 1, "top": 2, "width": 3, "height": 4}
        config = {
            "charts": [
                {"source": "Default", "title": "Default", "temp": 4, "info": False, "bounds": bounds},
            ],
            "classic_charts": [
                {
                    "source": "Classic",
                    "title": "Classic",
                    "temp": 5,
                    "info": True,
                    "bounds": bounds,
                    "duplicate": {
                        "enabled": True,
                        "bounds": {"left": 10, "top": 20, "width": 30, "height": 40},
                    },
                },
            ],
        }

        plan = build_chart_plan(
            config,
            classic=True,
            temp_enabled=False,
            ogtt=False,
            temp_correlation=False,
            rng_lcol=30,
        )

        self.assertEqual([chart["source"] for chart in plan], ["Classic", "Classic"])
        self.assertEqual(plan[0]["temp"], "0")
        self.assertEqual(plan[0]["secondary_axis"], {"enabled": False, "source": "none"})
        self.assertTrue(plan[0]["info"])
        self.assertEqual(plan[1]["temp"], "0")
        self.assertEqual(plan[1]["secondary_axis"], {"enabled": False, "source": "none"})
        self.assertTrue(plan[1]["info"])
        self.assertNotIn("duplicate", plan[1])
        self.assertEqual(get_bounds(plan[1]), (10.0, 20.0, 30.0, 40.0))

    def test_build_chart_plan_expands_duplicate_configs_into_chart_objects(self):
        config = {
            "charts": [
                {
                    "source": "A",
                    "title": "A",
                    "bounds": {"left": 1, "top": 2, "width": 3, "height": 4},
                    "duplicate": {
                        "enabled": True,
                        "delete_original": False,
                        "ignore_series": [2],
                        "bounds": {"left": 10, "top": 20, "width": 30, "height": 40},
                        "extra_series": [{"enabled": True, "target_id": "Diff35"}],
                    },
                },
                {
                    "source": "B",
                    "title": "B",
                    "bounds": {"left": 5, "top": 6, "width": 7, "height": 8},
                    "duplicate": {
                        "enabled": True,
                        "delete_original": True,
                        "ignore_series": [1, 3],
                        "bounds": {"left": 50, "top": 60, "width": 70, "height": 80},
                    },
                },
            ]
        }

        plan = build_chart_plan(
            config,
            classic=False,
            temp_enabled=True,
            ogtt=False,
            temp_correlation=False,
            rng_lcol=30,
        )

        self.assertEqual([chart["source"] for chart in plan], ["A", "A", "B"])
        self.assertEqual(get_bounds(plan[0]), (1.0, 2.0, 3.0, 4.0))
        self.assertEqual(get_bounds(plan[1]), (10.0, 20.0, 30.0, 40.0))
        self.assertEqual(plan[1]["ignore_series"], [2])
        self.assertEqual(plan[1]["extra_series"], [{"enabled": True, "target_id": "Diff35"}])
        self.assertNotIn("duplicate", plan[1])
        self.assertEqual(get_bounds(plan[2]), (50.0, 60.0, 70.0, 80.0))
        self.assertEqual(plan[2]["ignore_series"], [1, 3])
        self.assertNotIn("duplicate", plan[2])

    def test_build_chart_plan_applies_ogtt_and_temp_correlation(self):
        config = {
            "charts": [
                {
                    "source": "Diff1550-Diff1050",
                    "title": "Diff1550-Diff1050",
                    "temp": 0,
                    "info": False,
                    "bounds": {"left": 1, "top": 2, "width": 3, "height": 4},
                },
                {
                    "source": "1609",
                    "title": "1609nm单环吸光度",
                    "temp": 0,
                    "info": False,
                    "bounds": {"left": 5, "top": 6, "width": 7, "height": 8},
                    "temp_correlation": {
                        "corrected": {
                            "source": "Diff1550-Diff1050-temp",
                            "title": "温度校正后的波长差分",
                            "temp": 0,
                            "bounds": {"left": 9, "top": 10, "width": 11, "height": 12},
                        },
                        "original": {
                            "bounds": {"left": 13, "top": 14, "width": 15, "height": 16},
                        },
                    },
                },
            ],
        }

        plan = build_chart_plan(
            config,
            classic=False,
            temp_enabled=True,
            ogtt=True,
            temp_correlation=True,
            rng_lcol=30,
        )

        self.assertEqual(plan[0]["temp"], "32")
        self.assertEqual(plan[0]["secondary_axis"]["source"], "glucose")
        self.assertEqual(plan[0]["title"], "Diff1550-Diff1050 vs.血糖值")
        self.assertEqual(plan[1]["source"], "Diff1550-Diff1050-temp")
        self.assertEqual(plan[1]["title"], "温度校正后的波长差分")
        self.assertEqual(get_bounds(plan[1]), (9.0, 10.0, 11.0, 12.0))
        self.assertEqual(plan[2]["source"], "1609")
        self.assertEqual(plan[2]["title"], "1609nm单环吸光度")
        self.assertEqual(get_bounds(plan[2]), (13.0, 14.0, 15.0, 16.0))

    def test_validate_requires_absolute_chart_and_duplicate_bounds(self):
        config = {
            "charts": [
                {
                    "source": "A",
                    "title": "A",
                    "bounds": {"left": 1, "top": 2, "width": 3, "height": 4},
                    "duplicate": {
                        "enabled": True,
                        "bounds": {"left": 5, "top": 6, "width": 7, "height": 8},
                    },
                }
            ]
        }
        validate_chart_config(config)

        missing_chart_bounds = {"charts": [{"source": "A", "title": "A"}]}
        with self.assertRaises(ValueError):
            validate_chart_config(missing_chart_bounds)

        missing_duplicate_bounds = {
            "charts": [
                {
                    "source": "A",
                    "title": "A",
                    "bounds": {"left": 1, "top": 2, "width": 3, "height": 4},
                    "duplicate": {"enabled": True},
                }
            ]
        }
        with self.assertRaises(ValueError):
            validate_chart_config(missing_duplicate_bounds)

    def test_default_config_file_loads(self):
        config = load_chart_config()
        self.assertEqual(config["version"], 1)
        for classic, key in ((False, "charts"), (True, "classic_charts")):
            plan = build_chart_plan(config, classic, True, False, False, 30)
            self.assertEqual(
                [chart["source"] for chart in plan],
                [chart["source"] for chart in config[key]],
            )

    def test_secondary_axis_config_is_normalized_and_prefers_new_format(self):
        config = {
            "charts": [
                {
                    "source": "Diff45",
                    "title": "Diff45",
                    "temp": 33,
                    "bounds": {"left": 1, "top": 2, "width": 3, "height": 4},
                    "secondary_axis": {
                        "enabled": True,
                        "source": "temperature",
                        "match": {
                            "type": "header",
                            "value": "测头相对扶手高度(cm)",
                            "fallback_column": 33,
                        },
                    },
                },
                {
                    "source": "Diff1050",
                    "title": "Diff1050",
                    "temp": ["Diff1050", "Diff1219"],
                    "bounds": {"left": 1, "top": 2, "width": 3, "height": 4},
                },
            ]
        }

        plan = build_chart_plan(
            config,
            classic=False,
            temp_enabled=True,
            ogtt=False,
            temp_correlation=False,
            rng_lcol=30,
        )

        self.assertEqual(plan[0]["secondary_axis"]["source"], "temperature")
        self.assertEqual(plan[0]["secondary_axis"]["match"]["type"], "header")
        self.assertEqual(plan[0]["secondary_axis"]["match"]["fallback_column"], 33)
        self.assertEqual(plan[1]["secondary_axis"]["source"], "spectral")
        self.assertEqual(plan[1]["secondary_axis"]["series"], ["Diff1050", "Diff1219"])

    def test_resolve_header_column_matches_header_name_and_uses_fallback(self):
        headers = ["时间", "TC1实际温度", "测头相对扶手高度(cm)"]

        self.assertEqual(resolve_header_column(headers, " tc1实际温度 ", None), 2)
        self.assertEqual(resolve_header_column(headers, "missing", 33), 33)
        self.assertIsNone(resolve_header_column(headers, "missing", None))

    def test_resolve_secondary_axis_value_for_runtime_charting(self):
        headers = ["时间", "TC1实际温度", "测头相对扶手高度(cm)"]

        self.assertEqual(
            resolve_secondary_axis_value(
                {
                    "enabled": True,
                    "source": "temperature",
                    "match": {"type": "header", "value": "TC1实际温度", "fallback_column": 12},
                },
                headers,
                rng_lcol=30,
            ),
            "2",
        )
        self.assertEqual(
            resolve_secondary_axis_value(
                {
                    "enabled": True,
                    "source": "temperature",
                    "match": {"type": "header", "value": "missing", "fallback_column": 12},
                },
                headers,
                rng_lcol=30,
            ),
            "12",
        )
        self.assertEqual(
            resolve_secondary_axis_value({"enabled": True, "source": "glucose"}, headers, rng_lcol=30),
            "32",
        )
        self.assertEqual(
            resolve_secondary_axis_value(
                {"enabled": True, "source": "spectral", "series": ["Diff1050", "Diff1219"]},
                headers,
                rng_lcol=30,
            ),
            ["Diff1050", "Diff1219"],
        )
        self.assertEqual(
            resolve_secondary_axis_value({"enabled": False, "source": "none"}, headers, rng_lcol=30),
            "0",
        )


if __name__ == "__main__":
    unittest.main()
