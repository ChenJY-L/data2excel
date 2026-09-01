import copy
import json
import sys
from pathlib import Path


BOUNDS_KEYS = ("left", "top", "width", "height")


def default_config_path():
    """
    默认配置路径。

    打包成 onefile 后 __file__ 指向临时解包目录，配置文件可能随发布包放在
    exe 同级，因此按 源码同级 → exe同级src → exe同级 的顺序查找。
    """
    exe_dir = Path(sys.executable).parent
    candidates = (
        Path(__file__).with_name("chart_config.json"),
        exe_dir / "src" / "chart_config.json",
        exe_dir / "chart_config.json",
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return candidates[0]


def load_chart_config(config_path=None):
    path = Path(config_path) if config_path else default_config_path()
    with path.open("r", encoding="utf-8") as fp:
        config = json.load(fp)
    validate_chart_config(config)
    return config


def validate_chart_config(config):
    charts = config.get("charts")
    if not isinstance(charts, list) or not charts:
        raise ValueError("chart_config.json 必须包含非空 charts 列表")

    _validate_chart_list(charts, "charts")

    classic_charts = config.get("classic_charts")
    if classic_charts is not None:
        if not isinstance(classic_charts, list):
            raise ValueError("classic_charts 必须是列表")
        _validate_chart_list(classic_charts, "classic_charts")


def build_chart_plan(config, classic, temp_enabled, ogtt, temp_correlation, rng_lcol):
    source_charts = config.get("classic_charts") if classic else None
    if not source_charts:
        source_charts = config.get("charts", [])

    plan = [_normalize_chart(chart) for chart in source_charts]

    if not temp_enabled:
        for chart in plan:
            chart["temp"] = "0"
            chart["secondary_axis"] = _disabled_secondary_axis()

    if ogtt:
        chart = _find_chart(plan, "Diff1550-Diff1050")
        if chart is not None:
            chart["temp"] = str(rng_lcol + 2)
            chart["secondary_axis"] = {"enabled": True, "source": "glucose"}
            if not chart["title"].endswith(" vs.血糖值"):
                chart["title"] = chart["title"] + " vs.血糖值"

    if temp_correlation:
        chart_index = _find_chart_index(plan, "1609")
        if chart_index is not None:
            chart = plan[chart_index]
            temp_config = chart.get("temp_correlation", {})

            corrected_chart = copy.deepcopy(chart)
            corrected_chart.pop("temp_correlation", None)
            corrected_override = temp_config.get("corrected", {})
            corrected_chart.update(corrected_override)
            corrected_chart["source"] = str(corrected_override.get("source", "Diff1550-Diff1050-temp"))
            corrected_chart["key"] = str(corrected_override.get("key") or corrected_chart["source"])
            corrected_chart["title"] = str(corrected_override.get("title", "温度校正后的波长差分"))
            corrected_chart["temp"] = _normalize_temp(corrected_override.get("temp", "0"))
            corrected_chart.pop("duplicate", None)

            original_chart = copy.deepcopy(chart)
            original_chart.pop("temp_correlation", None)
            original_chart.update(temp_config.get("original", {}))
            original_chart = _normalize_chart(original_chart)

            plan[chart_index] = _normalize_chart(corrected_chart)
            plan.append(original_chart)

    return _expand_duplicate_charts(plan)


def get_bounds(item):
    bounds = item.get("bounds", {})
    return tuple(float(bounds[key]) for key in BOUNDS_KEYS)


def get_annotation_bounds(config, key):
    annotations = config.get("annotations", {})
    bounds = annotations.get(key)
    if not isinstance(bounds, dict):
        return None
    return tuple(float(bounds[name]) for name in BOUNDS_KEYS)


def resolve_header_column(headers, header_name, fallback_column=None):
    target = _normalize_header(header_name)
    if target == "":
        return int(fallback_column) if fallback_column is not None else None

    for index, header in enumerate(headers or [], start=1):
        if _normalize_header(header) == target:
            return index

    return int(fallback_column) if fallback_column is not None else None


def resolve_secondary_axis_value(secondary_axis, headers, rng_lcol):
    if not isinstance(secondary_axis, dict) or not secondary_axis.get("enabled"):
        return "0"

    source = str(secondary_axis.get("source", "none")).strip().lower()
    if source in ("", "none"):
        return "0"

    if source == "glucose":
        return str(rng_lcol + 2)

    if source == "spectral":
        return [str(item) for item in secondary_axis.get("series", [])]

    if source in ("temperature", "temp"):
        match = secondary_axis.get("match", {})
        match_type = str(match.get("type", "column")).strip().lower()
        if match_type == "header":
            column = resolve_header_column(
                headers,
                match.get("value", ""),
                match.get("fallback_column"),
            )
            return str(column) if column is not None else "0"
        return str(int(match.get("value", 0)))

    return "0"


def _validate_chart_list(charts, list_name):
    for index, chart in enumerate(charts):
        path = f"{list_name}[{index}]"
        if not isinstance(chart, dict):
            raise ValueError(f"{path} 必须是对象")
        if not chart.get("source"):
            raise ValueError(f"{path} 缺少 source")
        if not chart.get("title"):
            raise ValueError(f"{path} 缺少 title")
        _validate_bounds(chart.get("bounds"), f"{path}.bounds")

        duplicate = chart.get("duplicate")
        if isinstance(duplicate, dict) and duplicate.get("enabled"):
            _validate_bounds(duplicate.get("bounds"), f"{path}.duplicate.bounds")


def _validate_bounds(bounds, path):
    if not isinstance(bounds, dict):
        raise ValueError(f"{path} 必须是对象")
    for key in BOUNDS_KEYS:
        if key not in bounds:
            raise ValueError(f"{path} 缺少 {key}")
        float(bounds[key])


def _normalize_chart(chart):
    normalized = copy.deepcopy(chart)
    normalized["key"] = str(normalized.get("key") or normalized["source"])
    normalized["source"] = str(normalized["source"])
    normalized["title"] = str(normalized["title"])
    normalized["temp"] = _normalize_temp(normalized.get("temp", "0"))
    normalized["secondary_axis"] = _normalize_secondary_axis(
        normalized.get("secondary_axis"),
        normalized["temp"],
    )
    normalized["info"] = bool(normalized.get("info", normalized.get("show_exp_info", False)))
    return normalized


def _expand_duplicate_charts(charts):
    expanded = []
    for chart in charts:
        duplicate = chart.get("duplicate")
        if not (isinstance(duplicate, dict) and duplicate.get("enabled")):
            item = copy.deepcopy(chart)
            item.pop("duplicate", None)
            expanded.append(item)
            continue

        if not duplicate.get("delete_original"):
            original_chart = copy.deepcopy(chart)
            original_chart.pop("duplicate", None)
            expanded.append(original_chart)

        duplicate_chart = copy.deepcopy(chart)
        duplicate_chart.pop("duplicate", None)
        duplicate_chart["bounds"] = copy.deepcopy(duplicate.get("bounds", duplicate_chart.get("bounds", {})))
        if "ignore_series" in duplicate:
            duplicate_chart["ignore_series"] = list(duplicate.get("ignore_series") or [])
        if "extra_series" in duplicate:
            duplicate_chart["extra_series"] = copy.deepcopy(duplicate.get("extra_series") or [])
        expanded.append(duplicate_chart)
    return expanded


def _normalize_temp(value):
    if isinstance(value, list):
        return [str(item) for item in value]
    if value is None:
        return "0"
    return str(value)


def _normalize_secondary_axis(axis, legacy_temp):
    if isinstance(axis, dict):
        if not axis.get("enabled", True):
            return _disabled_secondary_axis()

        source = str(axis.get("source", "temperature")).strip().lower()
        if source in ("", "none"):
            return _disabled_secondary_axis()

        if source in ("temperature", "temp"):
            match = axis.get("match") or {}
            match_type = str(match.get("type", "column")).strip().lower()
            if match_type == "header":
                normalized_match = {
                    "type": "header",
                    "value": str(match.get("value", "")).strip(),
                }
                if match.get("fallback_column") is not None:
                    normalized_match["fallback_column"] = int(match["fallback_column"])
                return {
                    "enabled": True,
                    "source": "temperature",
                    "match": normalized_match,
                }

            column = match.get("value", axis.get("column", legacy_temp))
            return {
                "enabled": True,
                "source": "temperature",
                "match": {"type": "column", "value": int(column)},
            }

        if source == "glucose":
            return {"enabled": True, "source": "glucose"}

        if source == "spectral":
            series = axis.get("series", legacy_temp if isinstance(legacy_temp, list) else [])
            return {
                "enabled": True,
                "source": "spectral",
                "series": [str(item) for item in series],
            }

    if isinstance(legacy_temp, list):
        return {
            "enabled": True,
            "source": "spectral",
            "series": [str(item) for item in legacy_temp],
        }

    try:
        column = int(legacy_temp)
    except (TypeError, ValueError):
        column = 0

    if column <= 0:
        return _disabled_secondary_axis()

    return {
        "enabled": True,
        "source": "temperature",
        "match": {"type": "column", "value": column},
        "_legacy_temp": True,
    }


def _disabled_secondary_axis():
    return {"enabled": False, "source": "none"}


def _normalize_header(value):
    if value is None:
        return ""
    return "".join(str(value).split()).casefold()


def _find_chart(plan, key):
    index = _find_chart_index(plan, key)
    return plan[index] if index is not None else None


def _find_chart_index(plan, key):
    for index, chart in enumerate(plan):
        if chart.get("key") == key or chart.get("source") == key:
            return index
    return None
