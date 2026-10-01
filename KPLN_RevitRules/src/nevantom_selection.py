"""Parse Nevatom KT selections and match them to a KPLN Revit family snapshot.

The module is deliberately independent of Revit and PDF libraries.  It consumes
plain page text exported from a manufacturer PDF and the JSON returned by the
KPLN Revit bridge, so parsing and matching can be tested offline before any
model edit is proposed.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


MM_PER_FOOT = 304.8

# Verified from the end-view overall dimensions in the supplied KT sheets.
# The mapping is intentionally scoped to KT with Zn/Zn panels.
KT_ZN_BODY_MM = {
    "1,8": (695.0, 435.0),
    "2,2": (695.0, 480.0),
    "3,0": (795.0, 530.0),
    "4,0": (895.0, 595.0),
    "9,8": (1070.0, 1025.0),
    "13,5": (1370.0, 1025.0),
}

_CONFUSABLES = str.maketrans(
    {
        "А": "A",
        "В": "B",
        "Е": "E",
        "К": "K",
        "М": "M",
        "Н": "H",
        "О": "O",
        "Р": "P",
        "С": "C",
        "Т": "T",
        "У": "Y",
        "Х": "X",
    }
)

_START_RE = re.compile(
    r"Установка\s+(?P<system>П\d+(?:\.\d+)?)\s+"
    r"\(ID установки\s+(?P<installation_id>\d+),\s*ID расчета\s+(?P<calculation_id>\d+)\)\s*"
    r"(?P<mark>.*?)\s+Серия\s+(?P<series>[A-ZА-Я]+)\s+Длина установки\s+"
    r"(?P<length>\d+(?:[.,]\d+)?)\s*мм",
    re.IGNORECASE | re.DOTALL,
)

_SECTION_TITLES = (
    "Гибкая вставка",
    "Воздушный клапан",
    "Фильтр",
    "Водяной нагреватель",
    "Шумоглушитель",
    "Вентилятор",
    "Водяной охладитель",
)
_SECTION_END = "|".join(re.escape(item) for item in _SECTION_TITLES)


def _number(value: str) -> float:
    return float(value.replace(",", "."))


def _one(pattern: str, text: str, flags: int = 0) -> Optional[re.Match[str]]:
    return re.search(pattern, text, flags)


def normalize_mark(value: str) -> str:
    """Return a stable mark key despite Cyrillic/Latin lookalikes and spacing."""

    value = unicodedata.normalize("NFKC", value).upper().translate(_CONFUSABLES)
    value = re.sub(r"^KT\s*", "", value.strip())
    return re.sub(r"[^A-Z0-9]+", "", value)


def normalize_system(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).upper().strip()
    return re.sub(r"\s+", "", value)


def _section(text: str, title: str) -> str:
    pattern = re.compile(
        rf"(?ms)^\s*\d+\.\s*{re.escape(title)}\s*$\n?"
        rf"(?P<body>.*?)(?=^\s*\d+\.\s*(?:{_SECTION_END})\s*$|\Z)"
    )
    match = pattern.search(text)
    return match.group("body") if match else ""


def _pair(text: str, label: str, unit: str) -> Optional[Dict[str, float]]:
    match = _one(
        rf"{label}\s+(?P<selected>\d+(?:[.,]\d+)?)\s*"
        rf"\((?P<reference>\d+(?:[.,]\d+)?)\)\s*{unit}",
        text,
        re.IGNORECASE,
    )
    if not match:
        return None
    return {
        "selected": _number(match.group("selected")),
        "reference": _number(match.group("reference")),
    }


def _single_number(text: str, label: str, unit: str) -> Optional[float]:
    match = _one(
        rf"{label}\s+(?P<value>\d+(?:[.,]\d+)?)\s*{unit}",
        text,
        re.IGNORECASE,
    )
    return _number(match.group("value")) if match else None


def _component_data(block: str) -> Dict[str, Any]:
    heater = _section(block, "Водяной нагреватель")
    fan = _section(block, "Вентилятор")
    cooler = _section(block, "Водяной охладитель")

    result: Dict[str, Any] = {}
    if heater:
        result["heater"] = {
            "power_kw": _pair(heater, r"Мощность", r"кВт"),
            "fluid_flow_m3h": _pair(heater, r"Объемный расход жидкости", r"м3/ч"),
            "fluid_pressure_loss_kpa": _pair(
                heater, r"Суммарные потери давления по", r"кПа"
            ),
            "connection": _connection(heater),
        }
    if fan:
        result["fan"] = {
            "airflow_m3h": _single_number(fan, r"Расход фактический", r"м3/ч"),
            "free_pressure_pa": _single_number(fan, r"Напор свободный", r"Па"),
            "motor_power_kw": _single_number(fan, r"Мощность двигателя", r"кВт"),
            "motor_current_a": _single_number(
                fan, r"Номинальный ток двигателя", r"A"
            ),
            "actual_speed_rpm": _single_number(
                fan, r"Обороты фактические", r"об/мин"
            ),
            "working_frequency_hz": _single_number(fan, r"Рабочая частота", r"Гц"),
        }
    if cooler:
        result["cooler"] = {
            "power_kw": _pair(cooler, r"Мощность", r"кВт"),
            "fluid_flow_m3h": _pair(cooler, r"Объемный расход жидкости", r"м3/ч"),
            "fluid_pressure_loss_kpa": _pair(
                cooler, r"Суммарные потери давления по", r"кПа"
            ),
            "connection": _connection(cooler),
        }
    return result


def _connection(text: str) -> Optional[str]:
    match = _one(
        r"Диаметр подключения\s*\(вход/выход\)\s+(?P<value>\S+\s*/\s*\S+)",
        text,
        re.IGNORECASE,
    )
    return re.sub(r"\s+", "", match.group("value")) if match else None


def _connector_dimensions(block: str) -> List[Tuple[float, float]]:
    matches = re.findall(
        r"В(?:Гп|П)-\s*(?:\r?\n\s*Наименование\s*)?"
        r"(?P<width>\d+)\s*[-xх*]\s*(?P<height>\d+)",
        block,
        re.IGNORECASE,
    )
    return [(_number(width), _number(height)) for width, height in matches]


def _parse_selection(block: str, start_page: int, end_page: int) -> Dict[str, Any]:
    match = _START_RE.search(block)
    if not match:
        raise ValueError("Selection block has no KT header")

    mark = re.sub(r"\s+", " ", match.group("mark")).strip()
    series = match.group("series").upper().translate(_CONFUSABLES)
    size_match = _one(r"Типоразмер\s+[A-ZА-Я]+_\s*(\d+[,.]\d+)", block)
    panel_match = _one(r"Тип панели\s+([^\r\n]+)", block)
    weight_match = _one(r"Вес\s+(\d+(?:[.,]\d+)?)\s*кг", block)
    quantity_match = _one(r"Количество\s+(\d+)\s*шт", block)
    side_match = _one(r"Сторона обслуживания\s+(Правая|Левая)", block, re.IGNORECASE)
    pressure_match = _one(r"Свободный напор\s+(\d+(?:[.,]\d+)?)\s*Па", block)
    airflow_match = _one(r"Производительность\s+(\d+(?:[.,]\d+)?)\s*м3/ч", block)
    temperature_match = _one(r"Температура\s+(-?\d+(?:[.,]\d+)?)\s*◦?C", block)
    connectors = _connector_dimensions(block)
    size = size_match.group(1).replace(".", ",") if size_match else None
    body = KT_ZN_BODY_MM.get(size) if series == "KT" and size else None

    result: Dict[str, Any] = {
        "system": match.group("system"),
        "installation_id": match.group("installation_id"),
        "calculation_id": match.group("calculation_id"),
        "mark": mark,
        "series": series,
        "size": size,
        "panel_type": re.sub(r"\s+", "", panel_match.group(1)) if panel_match else None,
        "overall_length_mm": _number(match.group("length")),
        "weight_kg": _number(weight_match.group(1)) if weight_match else None,
        "quantity": int(quantity_match.group(1)) if quantity_match else None,
        "service_side": side_match.group(1).capitalize() if side_match else None,
        "supply_airflow_m3h": _number(airflow_match.group(1)) if airflow_match else None,
        "free_pressure_pa": _number(pressure_match.group(1)) if pressure_match else None,
        "supply_temperature_c": _number(temperature_match.group(1)) if temperature_match else None,
        "source_pages": [start_page, end_page],
    }
    if body:
        result["body_width_mm"], result["body_height_mm"] = body
    if connectors:
        result["inlet_connector_mm"] = list(connectors[0])
        result["outlet_connector_mm"] = list(connectors[-1])
    result.update(_component_data(block))
    return result


def parse_selections(pages: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Parse every standard KT selection found in ordered page dictionaries."""

    starts = [index for index, page in enumerate(pages) if _START_RE.search(page["text"])]
    selections: List[Dict[str, Any]] = []
    for position, start in enumerate(starts):
        stop = starts[position + 1] if position + 1 < len(starts) else len(pages)
        block_pages = [
            page
            for page in pages[start:stop]
            if not (
                _one(r"Установка\s+SL_", page["text"], re.IGNORECASE)
                and _one(r"Система\s+П\d+", page["text"], re.IGNORECASE)
            )
        ]
        block = "\n".join(page["text"] for page in block_pages)
        selections.append(
            _parse_selection(block, block_pages[0]["page"], block_pages[-1]["page"])
        )
    return selections


def find_unmatched_offers(pages: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Identify non-KT offer pages that must not be forced onto the KT family."""

    result = []
    for page in pages:
        text = page["text"]
        offer = _one(r"Установка\s+(SL_\s*\d+[,.]\d+\s+[^\r\n]+)", text)
        system = _one(r"Система\s+(П\d+(?:\.\d+)?)", text)
        if offer and system:
            result.append(
                {
                    "page": page["page"],
                    "system": system.group(1),
                    "mark": re.sub(r"\s+", " ", offer.group(1)).strip(),
                    "reason": "series_not_supported_by_kt_family",
                }
            )
    return result


def validate_selection(selection: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Return blocking issues for the verified KT/Zn-Zn mapping scope."""

    issues: List[Dict[str, Any]] = []
    if selection.get("series") != "KT":
        issues.append({"code": "unsupported_series", "actual": selection.get("series")})
    if selection.get("panel_type", "").upper() != "ZN/ZN":
        issues.append(
            {"code": "unsupported_panel_type", "actual": selection.get("panel_type")}
        )
    required = (
        "system",
        "installation_id",
        "mark",
        "overall_length_mm",
        "body_width_mm",
        "body_height_mm",
        "weight_kg",
        "service_side",
        "inlet_connector_mm",
        "outlet_connector_mm",
        "heater",
        "fan",
        "cooler",
    )
    for field in required:
        if selection.get(field) in (None, "", []):
            issues.append({"code": "missing_required_field", "field": field})

    fan = selection.get("fan", {})
    for source_field, fan_field in (
        ("supply_airflow_m3h", "airflow_m3h"),
        ("free_pressure_pa", "free_pressure_pa"),
    ):
        source_value = selection.get(source_field)
        fan_value = fan.get(fan_field)
        if source_value is not None and fan_value is not None and source_value != fan_value:
            issues.append(
                {
                    "code": "summary_fan_mismatch",
                    "summary_field": source_field,
                    "summary_value": source_value,
                    "fan_field": fan_field,
                    "fan_value": fan_value,
                }
            )
    return issues


def _family_values(family: Dict[str, Any]) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, str]]:
    parameters = {str(item["id"]): item for item in family["parameters"]}
    ids = {item["name"]: str(item["id"]) for item in family["parameters"]}
    return parameters, ids


def match_family_types(
    selections: Sequence[Dict[str, Any]], family: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """Match selections to family types by installation ID, then system + mark."""

    _, ids = _family_values(family)
    code_id = ids.get("КП_О_Код изделия")
    mark_id = ids.get("КП_О_Марка")
    system_id = ids.get("КП_О_Имя Системы")
    by_code = {item["installation_id"]: item for item in selections}
    by_key = {
        (normalize_system(item["system"]), normalize_mark(item["mark"])): item
        for item in selections
    }
    matches: List[Dict[str, Any]] = []
    used = set()
    for family_type in family["types"]:
        values = {str(item["parameter_id"]): item["value"] for item in family_type["values"]}
        code = str(values.get(code_id) or "").strip() if code_id else ""
        mark = str(values.get(mark_id) or "").strip() if mark_id else ""
        system = str(values.get(system_id) or "").strip() if system_id else ""
        selection = by_code.get(code) if code else None
        method = "installation_id" if selection else None
        if not selection:
            selection = by_key.get((normalize_system(system), normalize_mark(mark)))
            method = "system_and_mark" if selection else None
        matches.append(
            {
                "family_type": family_type["name"],
                "selection": selection,
                "method": method,
                "values": values,
            }
        )
        if selection:
            used.add(selection["installation_id"])
    for selection in selections:
        if selection["installation_id"] not in used:
            matches.append(
                {
                    "family_type": None,
                    "selection": selection,
                    "method": None,
                    "values": {},
                }
            )
    return matches


def compare_family(
    selections: Sequence[Dict[str, Any]], family: Dict[str, Any], tolerance_mm: float = 0.5
) -> Dict[str, Any]:
    """Return a source-unit comparison; it never produces Revit write operations."""

    parameters, ids = _family_values(family)

    def value(values: Dict[str, Any], name: str) -> Any:
        parameter_id = ids.get(name)
        return values.get(parameter_id) if parameter_id else None

    def length_mm(values: Dict[str, Any], name: str) -> Optional[float]:
        raw = value(values, name)
        return None if raw is None else float(raw) * MM_PER_FOOT

    rows = []
    summary = {"matched": 0, "unmatched_family_types": 0, "unmatched_selections": 0}
    for item in match_family_types(selections, family):
        selection = item["selection"]
        family_type = item["family_type"]
        values = item["values"]
        if not family_type:
            summary["unmatched_selections"] += 1
            rows.append(item)
            continue
        if not selection:
            summary["unmatched_family_types"] += 1
            rows.append(item)
            continue
        summary["matched"] += 1
        checks = []

        def exact(parameter: str, source_field: str, expected: Any, actual: Any) -> None:
            checks.append(
                {
                    "parameter": parameter,
                    "source_field": source_field,
                    "expected": expected,
                    "actual": actual,
                    "status": "match" if str(expected) == str(actual) else "mismatch",
                }
            )

        def number(
            parameter: str, source_field: str, expected: Optional[float], actual: Optional[float]
        ) -> None:
            if expected is None or actual is None:
                status = "missing"
            else:
                status = "match" if abs(float(expected) - float(actual)) <= tolerance_mm else "mismatch"
            checks.append(
                {
                    "parameter": parameter,
                    "source_field": source_field,
                    "expected": expected,
                    "actual": None if actual is None else round(actual, 3),
                    "status": status,
                }
            )

        exact("КП_О_Имя Системы", "system", selection["system"], value(values, "КП_О_Имя Системы"))
        exact(
            "КП_О_Марка",
            "mark",
            normalize_mark(selection["mark"]),
            normalize_mark(str(value(values, "КП_О_Марка") or "")),
        )
        exact(
            "КП_О_Код изделия",
            "installation_id",
            selection["installation_id"],
            value(values, "КП_О_Код изделия") or "",
        )
        exact(
            "КП_О_Масса_Текст",
            "weight_kg",
            str(int(selection["weight_kg"])),
            value(values, "КП_О_Масса_Текст") or "",
        )
        number(
            "Установка_Длина",
            "overall_length_mm",
            selection.get("overall_length_mm"),
            length_mm(values, "Установка_Длина"),
        )
        number(
            "Установка_Ширина",
            "end_view_overall_width_mm",
            selection.get("body_width_mm"),
            length_mm(values, "Установка_Ширина"),
        )
        number(
            "Установка_Высота",
            "end_view_overall_height_mm",
            selection.get("body_height_mm"),
            length_mm(values, "Установка_Высота"),
        )
        inlet = selection.get("inlet_connector_mm") or [None, None]
        outlet = selection.get("outlet_connector_mm") or [None, None]
        number(
            "Секция_1_Соединитель_Ширина",
            "first_flexible_insert_width_mm",
            inlet[0],
            length_mm(values, "Секция_1_Соединитель_Ширина"),
        )
        number(
            "Секция_1_Соединитель_Высота",
            "first_flexible_insert_height_mm",
            inlet[1],
            length_mm(values, "Секция_1_Соединитель_Высота"),
        )
        number(
            "Секция_12_Соединитель_Ширина",
            "last_flexible_insert_width_mm",
            outlet[0],
            length_mm(values, "Секция_12_Соединитель_Ширина"),
        )
        number(
            "Секция_12_Соединитель_Высота",
            "last_flexible_insert_height_mm",
            outlet[1],
            length_mm(values, "Секция_12_Соединитель_Высота"),
        )
        right = 1 if selection.get("service_side") == "Правая" else 0
        exact("ЗО_Справа", "service_side", right, value(values, "ЗО_Справа"))

        missing_targets = []
        target_map = {
            "КП_И_Расход воздуха": selection.get("supply_airflow_m3h"),
            "КП_И_Свободный напор воздуха": selection.get("free_pressure_pa"),
            "КП_И_Тепловая мощность": selection.get("heater", {}).get("power_kw", {}).get("selected"),
            "КП_И_Холодильная мощность": selection.get("cooler", {}).get("power_kw", {}).get("selected"),
            "КП_И_Расход теплоносителя": selection.get("heater", {}).get("fluid_flow_m3h", {}).get("selected"),
            "КП_И_Расход холодоносителя": selection.get("cooler", {}).get("fluid_flow_m3h", {}).get("selected"),
            "КП_И_Потеря давления теплоносителя": selection.get("heater", {}).get("fluid_pressure_loss_kpa", {}).get("selected"),
            "КП_И_Потеря давления холодоносителя": selection.get("cooler", {}).get("fluid_pressure_loss_kpa", {}).get("selected"),
            "КП_И_Номинальная мощность": selection.get("fan", {}).get("motor_power_kw"),
            "КП_И_Ток": selection.get("fan", {}).get("motor_current_a"),
            "КП_И_Частота вращения двигателя": selection.get("fan", {}).get("actual_speed_rpm"),
        }
        for parameter, target in target_map.items():
            current = value(values, parameter)
            if target is not None and (current is None or current == 0):
                missing_targets.append(
                    {"parameter": parameter, "target": target, "status": "family_value_zero"}
                )

        rows.append(
            {
                "family_type": family_type,
                "selection_system": selection["system"],
                "selection_mark": selection["mark"],
                "match_method": item["method"],
                "source_pages": selection["source_pages"],
                "validation_issues": validate_selection(selection),
                "checks": checks,
                "missing_characteristics": missing_targets,
            }
        )

    status_counts: Dict[str, int] = {"match": 0, "mismatch": 0, "missing": 0}
    for row in rows:
        for check in row.get("checks", []):
            status_counts[check["status"]] = status_counts.get(check["status"], 0) + 1
    summary["identity_and_dimension_checks"] = status_counts
    summary["zero_characteristic_cells"] = sum(
        len(row.get("missing_characteristics", [])) for row in rows
    )
    summary["blocking_validation_issues"] = sum(
        len(row.get("validation_issues", [])) for row in rows
    )
    return {
        "schema_version": 1,
        "read_only": True,
        "family_document": family.get("document"),
        "family_revision": family.get("revision"),
        "summary": summary,
        "rows": rows,
    }
