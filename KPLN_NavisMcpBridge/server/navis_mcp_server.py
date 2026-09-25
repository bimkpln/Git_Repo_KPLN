"""
MCP-сервер поверх локального HTTP-моста KPLN_NavisMcpBridge (аддин внутри
Navisworks Manage 2020, порт 8765 на 127.0.0.1).

Запуск: python navis_mcp_server.py
(stdio-транспорт — подключается как обычный локальный MCP-сервер).

Требует, чтобы Navisworks 2020 был открыт, документ загружен, и аддин
"MCP Bridge" был один раз запущен с ленты (вкладка Add-ins) — тогда
HttpListener слушает http://127.0.0.1:8765/.

Тесты и результаты Clash Detective в COM API Navisworks 2020 не имеют
стабильного Guid — идентифицируются по имени (name), поэтому все тулы
принимают имя теста/результата, как оно показано в list_clash_tests /
get_clash_test_results.

РАЗДЕЛЕНИЕ ОТВЕТСТВЕННОСТИ (важно при доработке):
  * C#-аддин            — только сырые данные COM API (пути, статусы,
                          Distance, Pt1/Pt2, Bound, Item1Bound/Item2Bound).
                          Никакой инженерной логики.
  * этот MCP-сервер     — команды чтения/записи данных Navisworks и
                          нормализация этих данных: миллиметры, категории,
                          размеры, габариты, центры и простые метрики.
                          Никаких проектных порогов, вердиктов, обучения
                          или реестра правил.
  * навык (SKILL.md)    — ПОЛИТИКА: пороги проекта, порядок правил,
                          обучение на человеческой разметке и решение,
                          какие статусы предлагать.
"""

import difflib
import math
import re
from collections import Counter
from urllib.parse import quote

import httpx
from mcp.server.fastmcp import FastMCP

BRIDGE_URL = "http://127.0.0.1:8765"

# Сырые длины COM API Navisworks приходят в ФУТАХ (проверено: Tolerance
# 0.032808399 == ровно 0.01 м). Всё, что этот сервер отдаёт в полях
# метрик, уже переведено в миллиметры.
FT2MM = 304.8

mcp = FastMCP("KPLN_NavisMcpBridge")


def _client(timeout: float = 90.0) -> httpx.Client:
    return httpx.Client(base_url=BRIDGE_URL, timeout=timeout)


def _q(name: str) -> str:
    return quote(name, safe="")


def _bridge_error_hint(exc: Exception) -> str:
    if isinstance(exc, httpx.ConnectError):
        return (
            "Не удалось достучаться до Navisworks (127.0.0.1:8765). "
            "Убедитесь, что Navisworks Manage 2020 запущен, документ открыт, "
            "и на вкладке Add-ins один раз нажата кнопка 'MCP Bridge'."
        )
    if isinstance(exc, httpx.HTTPStatusError):
        response = exc.response
        body = response.text.strip()
        try:
            payload = response.json()
            if isinstance(payload, dict) and payload.get("error"):
                body = str(payload["error"])
        except ValueError:
            pass

        detail = body or response.reason_phrase
        return (
            f"Navisworks bridge вернул HTTP {response.status_code} "
            f"для {response.request.method} {response.request.url}: {detail}"
        )
    return str(exc)

def _status_bucket(status: str | None) -> str:
    text = str(status or "").upper()
    if "APPROVED" in text or "RESOLVED" in text or "REVIEWED" in text:
        return "Accepted"
    if "NEW" in text or "ACTIVE" in text:
        return "Open"
    return "Other"


def _find_clash_test(tests: list[dict], test_name: str) -> dict:
    for test in tests:
        if test.get("Name") == test_name:
            return test

    names = [test.get("Name") for test in tests if test.get("Name")]
    close = difflib.get_close_matches(test_name, names, n=5, cutoff=0.4)
    raise RuntimeError(
        f'ОТКАЗ: теста "{test_name}" нет в открытом документе.'
        + (
            "\nПохожие имена: " + ", ".join(f'"{name}"' for name in close)
            if close
            else "\nПолный список — list_clash_tests."
        )
    )


# ---------------------------------------------------------------------------
# Геометрия: считаем ЗДЕСЬ, а не в диалоге.
#
# Раньше эта арифметика каждый раз заново писалась одноразовым скриптом в
# сессии, а в контекст приходилось тянуть по 150 bounding box'ов — это и
# упиралось в лимит размера ответа инструмента, и давало разный результат
# от сессии к сессии. Теперь формулы зафиксированы в коде, а наружу идут
# готовые числа.
# ---------------------------------------------------------------------------


def _dims_mm(bound: dict) -> list[float]:
    """Габариты bounding box'а (dX, dY, dZ) в миллиметрах."""
    mn, mx = bound["Min"], bound["Max"]
    return [
        (mx["X"] - mn["X"]) * FT2MM,
        (mx["Y"] - mn["Y"]) * FT2MM,
        (mx["Z"] - mn["Z"]) * FT2MM,
    ]


def _center_mm(bound: dict) -> list[float]:
    """Центр bounding box'а (X, Y, Z) в миллиметрах."""
    mn, mx = bound["Min"], bound["Max"]
    return [
        (mx["X"] + mn["X"]) / 2.0 * FT2MM,
        (mx["Y"] + mn["Y"]) / 2.0 * FT2MM,
        (mx["Z"] + mn["Z"]) / 2.0 * FT2MM,
    ]


def _wall_relation(
    result: dict,
    wall_geometry: dict | None,
    mep_axis: list[float] | None,
) -> dict:
    """Raw geometry relation between a MEP axis and a wall in plan.

    The wall axis comes from mesh PCA in the add-in, not from an axis-aligned
    bounding box. This is essential for diagonal walls: their AABB can look
    almost square and contains no usable wall direction.
    """
    out = {
        "WallNormalAngle": None,
        "WallEndDistance": None,
        "WallPlanThickness": None,
        "WallPlanReliability": None,
        "WallPlanSource": None,
        "WallBroadFaceDistance": None,
        "WallEndFaceDistance": None,
    }
    if not wall_geometry:
        return out

    try:
        axis_x = float(wall_geometry["AxisX"])
        axis_y = float(wall_geometry["AxisY"])
        center_x = float(wall_geometry["CenterX"]) * FT2MM
        center_y = float(wall_geometry["CenterY"]) * FT2MM
        half_length = float(wall_geometry["HalfLength"]) * FT2MM
        half_thickness = float(wall_geometry["HalfThickness"]) * FT2MM
        reliability = float(wall_geometry["Reliability"])
    except (KeyError, TypeError, ValueError):
        return out

    out["WallPlanThickness"] = round(half_thickness * 2.0, 1)
    out["WallPlanReliability"] = round(reliability, 2)
    out["WallPlanSource"] = wall_geometry.get("Source")
    broad_face = wall_geometry.get("NearestBroadFace")
    end_face = wall_geometry.get("NearestEndFace")
    if broad_face is not None:
        out["WallBroadFaceDistance"] = round(float(broad_face) * FT2MM, 1)
    if end_face is not None:
        out["WallEndFaceDistance"] = round(float(end_face) * FT2MM, 1)

    if mep_axis:
        wall_normal = [-axis_y, axis_x, 0.0]
        out["WallNormalAngle"] = round(_angle_deg(mep_axis, wall_normal), 2)

    clash_bound = result.get("Bound")
    if clash_bound:
        clash_center = _center_mm(clash_bound)
        dx = clash_center[0] - center_x
        dy = clash_center[1] - center_y
        along = abs(dx * axis_x + dy * axis_y)
        out["WallEndDistance"] = round(max(0.0, half_length - along), 1)

    return out


def _axis(dims: list[float]) -> tuple[list[float] | None, float, float, float]:
    """Ось элемента по его AABB.

    У прямого цилиндра (труба, воздуховод) axis-aligned bounding box равен
    "длина вдоль оси ПЛЮС диаметр по каждой из трёх осей". Значит:
        диаметр  dia = min(dX, dY, dZ)
        вектор оси   = normalize(dX - dia, dY - dia, dZ - dia)
    Это работает и для ДИАГОНАЛЬНЫХ участков — снапить к X/Y/Z не нужно.

    Возвращает (единичный вектор оси | None, вытянутость, диаметр, длина).
    Вытянутость e = длина / диаметр: во сколько раз элемент длиннее своего
    поперечника. e < 2 — коротыш (отвод, врезка, диск), его собственное
    направление ненадёжно.
    """
    dia = min(dims)
    v = [max(x - dia, 0.0) for x in dims]
    n = math.sqrt(sum(x * x for x in v))
    if n <= 1e-9:
        return None, 0.0, dia, 0.0
    return [x / n for x in v], (n / dia if dia > 0 else 0.0), dia, n


def _angle_deg(a: list[float], b: list[float]) -> float:
    """Угол между осями, 0..90°. Модуль скалярного произведения — чтобы
    0° и 180° (встречное направление той же трассы) считались одним и тем
    же продольным случаем."""
    dot = abs(sum(x * y for x, y in zip(a, b)))
    return math.degrees(math.acos(max(-1.0, min(1.0, dot))))


def _kind(path: str | None) -> str:
    """Классификация стороны пары по пути в дереве модели.

    Три категории:
      "изоляция"    — слой изоляции трубы/воздуховода;
      "фитинг"      — соединительная деталь или арматура. Её изоляция по
                      проекту ЕСТЬ, но в NWC не экспортируется (прямая
                      оговорка регламента), поэтому геометрически фитинг
                      ведёт себя как изолированная труба;
      "голая труба" — тело трубы без изоляции: если чужая изоляция дошла
                      до него, свой слой изоляции уже пробит.
    """
    p = path or ""
    if "Изоляция" in p:
        return "изоляция"
    if any(k in p for k in ("Врезка", "Арматура", "Фитинг", "DN15", "/ LD /")):
        return "фитинг"
    return "голая труба"


# Более широкая классификация — нужна там, где в паре не трубы, а
# конструкции: стены/перекрытия АР-КР против воздуховодов ОВ2 и т.п.
# Kind1/Kind2 (трубная таксономия выше) для таких пар бессмысленны, поэтому
# они НЕ трогаются, а рядом добавляются Class1/Class2/PairClass.
_WALL_MARKERS = ("Базовая стена", "/ Стены /", "Перегород")
_SLAB_MARKERS = ("Перекрыти", "/ Полы /", "Кровля")
# Отделочные слои: штукатурка по сетке, облицовка, утеплитель.
# Сервер только распознаёт категорию; решение о допуске принимает skill.
# Проверяется раньше стены: путь такого слоя тоже содержит "Базовая стена".
_STRUCT = ("стена", "перекрытие", "отделка")
_FINISH_MARKERS = ("штукатурк", "по сетке", "Отделк", "отделк", "Утеплител",
                   "утеплител", "Облицовк", "облицовк")


def _class(path: str | None) -> str:
    """Категория элемента по его пути в дереве модели.

    Главный признак — КАТЕГОРИЯ REVIT, а это предпоследний-перед-уровнем
    сегмент пути (…/ <категория> / <уровень> / <файл>.nwc): "Воздуховоды",
    "Соединительные детали воздуховодов", "Арматура воздуховодов",
    "Материалы изоляции воздуховодов", "Стены", "Перекрытия".

    Именно категория, а НЕ поиск слов по всей строке: имя типа сплошь и
    рядом содержит слова-ловушки — прямой воздуховод лежит в типе
    "ASML_ОВ_Сталь δ=0.8_ГОСТ 14918-80_Врезки", и поиск "Врезк" по всему
    пути записывал 48 прямых участков из 62 в фасонину.
    """
    p = path or ""
    if any(k in p for k in _FINISH_MARKERS):
        return "отделка"

    segs = [x for x in p.split(" / ") if x.strip()]
    cat = segs[-3] if len(segs) >= 3 else ""
    if cat:
        if "изоляции воздуховод" in cat:
            return "изоляция воздуховода"
        if "изоляции труб" in cat:
            return "изоляция трубы"
        if cat.startswith("Соединительные детали"):
            return "фасонина"
        if cat.startswith("Арматура") or "Воздухораспределител" in cat:
            return "клапан"
        if "Воздуховод" in cat:
            return "воздуховод"
        if cat.startswith("Труб") or "Трубопровод" in cat:
            return "труба"
        if cat in ("Стены", "Несущие стены") or "Перегород" in cat:
            return "стена"
        if cat in ("Перекрытия", "Полы", "Крыши", "Потолки"):
            return "перекрытие"
        if "лотк" in cat.lower() or "Короб" in cat:
            return "лоток"

    # Запасной путь для деревьев, не похожих на выгрузку из Revit.
    if "Изоляция воздуховода" in p:
        return "изоляция воздуховода"
    if "Изоляция" in p:
        return "изоляция трубы"
    if "Воздуховод" in p:
        return "воздуховод"
    if any(k in p for k in _WALL_MARKERS):
        return "стена"
    if any(k in p for k in _SLAB_MARKERS):
        return "перекрытие"
    if any(k in p for k in ("Врезка", "Арматура", "Фитинг", "DN15", "/ LD /")):
        return "фитинг"
    if "Труб" in p:
        return "труба"
    return "прочее"


def _thickness_mm(path: str | None) -> float | None:
    """Толщина конструкции из ИМЕНИ ТИПА, а не из геометрии.

    У стены bounding box НЕ даёт толщину: диагональная в плане стена
    имеет большой габарит и по X, и по Y (реальный случай: перегородка
    80 мм, а min(dX,dY,dZ) = 2425 мм). Зато в имени типа толщина есть
    явно — "к-1.2_01_ВН_ГБ200_200", "к-2.1_01_ВН_ПГП80_80": последняя
    группа после "_" и есть толщина в мм.
    """
    for seg in (path or "").split(" / "):
        tail = seg.rsplit("_", 1)[-1].strip()
        if tail.isdigit():
            v = int(tail)
            if 20 <= v <= 1000:
                return float(v)
    return None


# Разбор НОМИНАЛЬНОГО сечения из строки свойства "Размер".
#
# Мерить сечение по bounding box'у нельзя: у круглого противопожарного
# клапана ø200 коробка даёт 253 мм (корпус), у отвода R=1.5 — вдвое больше
# диаметра, а у диагонального в плане воздуховода раздувается по обеим
# горизонтальным осям. В свойстве "Размер" лежит то, что нужно:
#   "ø200 мм-ø200 мм"      -> круглый, обе грани 200
#   "300x200-300x200"      -> прямоугольный, грани 300 и 200
# Строка вида "начало-конец" описывает переход; берём НАИБОЛЬШУЮ грань по
# обеим сторонам — отверстие должно пропустить самое широкое место.
_ROUND_RE = re.compile(r"[øØ⌀]|\bd\s*\d", re.IGNORECASE)
_MULT_RE = re.compile(r"[x\u0445\u0425\u00d7X]")
_NUM_RE = re.compile(r"\d+(?:[.,]\d+)?")


def _parse_size(text: str | None) -> tuple[float | None, float | None]:
    """Строка размера -> (меньшая грань, большая грань) в мм."""
    if not text:
        return None, None
    edges: list[tuple[float, float]] = []
    for part in re.split(r"\s*[-\u2013\u2014]\s*", str(text)):
        nums = [float(x.replace(",", ".")) for x in _NUM_RE.findall(part)]
        nums = [n for n in nums if n > 0]
        if not nums:
            continue
        if _ROUND_RE.search(part):
            edges.append((nums[0], nums[0]))
        elif _MULT_RE.search(part) and len(nums) >= 2:
            edges.append((min(nums[0], nums[1]), max(nums[0], nums[1])))
        else:
            edges.append((nums[0], nums[0]))
    if not edges:
        return None, None
    return max(e[0] for e in edges), max(e[1] for e in edges)


def _metrics(result: dict) -> dict:
    """Готовые геометрические метрики одной пары.

    Это нормализация данных Navisworks, а не инженерное решение:
    здесь нет проектных порогов, статусов Approved/Active или обучения.
    """
    p1 = result.get("Item1Path")
    p2 = result.get("Item2Path")
    k1, k2 = _kind(p1), _kind(p2)
    c1, c2 = _class(p1), _class(p2)
    out: dict = {
        "Kind1": k1,
        "Kind2": k2,
        "PairKind": "+".join(sorted([k1, k2])),
        "Class1": c1,
        "Class2": c2,
        "PairClass": "+".join(sorted([c1, c2])),
        "Thick1": _thickness_mm(p1) if c1 in _STRUCT else None,
        "Thick2": _thickness_mm(p2) if c2 in _STRUCT else None,
        "Group": result.get("Group"),
        "Size1": result.get("Item1Size"),
        "Size2": result.get("Item2Size"),
        "DistMm": round(abs(result.get("Distance") or 0.0) * FT2MM, 1),
        "Angle": None,
        "PerpRel": None,
        "Dia1": None,
        "Dia2": None,
        "Len1": None,
        "Len2": None,
        "Elong1": None,
        "Elong2": None,
        "WallNormalAngle": None,
        "WallEndDistance": None,
        "WallPlanThickness": None,
        "WallPlanReliability": None,
        "WallPlanSource": None,
        "WallBroadFaceDistance": None,
        "WallEndFaceDistance": None,
        "GeomOk": False,
    }

    b1, b2 = result.get("Item1Bound"), result.get("Item2Bound")
    if not b1 or not b2:
        # Аддин не смог собрать габариты элемента (нет геометрии /
        # не резолвится ModelItem). Судить по геометрии нельзя —
        # такие случаи навык должен показывать пользователю отдельно.
        out["GeomNote"] = "нет Item1Bound/Item2Bound (запрошено без include_item_bounds?)"
        return out

    d1, d2 = _dims_mm(b1), _dims_mm(b2)
    v1, e1, dia1, l1 = _axis(d1)
    v2, e2, dia2, l2 = _axis(d2)

    out["Dia1"] = round(dia1, 1)
    out["Dia2"] = round(dia2, 1)
    out["Len1"] = round(l1, 1)
    out["Len2"] = round(l2, 1)
    out["Elong1"] = round(e1, 2)
    out["Elong2"] = round(e2, 2)

    if v1 and v2:
        out["Angle"] = round(_angle_deg(v1, v2), 2)

    # PerpRel — "насколько мимо оси трубы проходит второй элемент":
    # от вектора между центрами берём составляющую, ПЕРПЕНДИКУЛЯРНУЮ оси
    # более длинного элемента, и делим на больший из двух диаметров.
    #   ~0    — второй элемент идёт точно через ось первого (протыкает
    #           саму трубу, а не мнёт слой изоляции);
    #   >=0.2 — задевает по касательной.
    if l1 >= l2:
        bl, bs, vl = b1, b2, v1
    else:
        bl, bs, vl = b2, b1, v2
    if vl:
        cl, cs = _center_mm(bl), _center_mm(bs)
        delta = [cs[i] - cl[i] for i in range(3)]
        proj = sum(delta[i] * vl[i] for i in range(3))
        perp = math.sqrt(sum((delta[i] - proj * vl[i]) ** 2 for i in range(3)))
        dmax = max(dia1, dia2)
        out["PerpRel"] = round(perp / dmax, 3) if dmax > 0 else None
    else:
        out["PerpRel"] = None

    # "union-of-N" означает, что габарит собран объединением потомков —
    # при большом N это не один цилиндр, и ось/диаметр ненадёжны.
    s1 = result.get("Item1BoundSource") or ""
    s2 = result.get("Item2BoundSource") or ""
    if s1.startswith("union") or s2.startswith("union"):
        out["GeomNote"] = f"ббокс объединён по потомкам ({s1} / {s2}) — ось ненадёжна"

    # Признаки для пар "конструкция + инженерный элемент": нужно ли вообще
    # отверстие и какого размера.
    #   MepMinEdge — самая узкая грань габарита инженерного элемента. Для
    #     воздуховода это его сечение (у круглого — диаметр), то есть меньшая
    #     грань отверстия, которое пришлось бы вырезать.
    #   Through   — прошёл ли элемент конструкцию НАСКВОЗЬ: глубина захода
    #     (DistMm) не меньше толщины конструкции. Касание насквозь не идёт,
    #     и отверстие для него не нужно.
    struct_side = 1 if c1 in _STRUCT else (2 if c2 in _STRUCT else 0)
    if struct_side:
        mep_d = d2 if struct_side == 1 else d1
        mep_axis = v2 if struct_side == 1 else v1
        thick = out["Thick1"] if struct_side == 1 else out["Thick2"]
        out["MepMinEdge"] = round(min(mep_d), 1)
        out["StructThick"] = thick
        # Сечение по свойству "Размер" — приоритетнее габарита.
        raw_size = result.get("Item2Size") if struct_side == 1 else result.get("Item1Size")
        s_min, s_max = _parse_size(raw_size)
        out["SectionMin"] = s_min
        out["SectionMax"] = s_max
        out["SectionSource"] = "свойство Размер" if s_max is not None else "габарит (свойства нет)"
        if s_max is None:
            # Запасной вариант: узкая грань габарита. Это НИЖНЯЯ оценка —
            # вторая грань сечения по габариту не восстанавливается.
            out["SectionMin"] = out["MepMinEdge"]
            out["SectionMax"] = out["MepMinEdge"]
        out["Through"] = (
            round(out["DistMm"], 1) >= round(thick * 0.95, 1) if thick else None
        )

        wall_side = 1 if c1 == "стена" else (2 if c2 == "стена" else 0)
        if wall_side:
            wall_geometry = result.get(
                "Item1PlanGeometry" if wall_side == 1 else "Item2PlanGeometry"
            )
            out.update(_wall_relation(result, wall_geometry, mep_axis))
    else:
        out["MepMinEdge"] = None
        out["StructThick"] = None
        out["Through"] = None
        out["SectionMin"] = None
        out["SectionMax"] = None
        out["SectionSource"] = None

    if struct_side:
        out["GeomNote"] = (
            "в паре есть конструкция: общий Angle по двум AABB не использовать. "
            "Для стены смотреть WallNormalAngle и WallEndDistance, рассчитанные "
            "по mesh-PCA стены; для размера — SectionMax/MepMinEdge."
        )

    if struct_side and (c1 == "стена" or c2 == "стена"):
        out["GeomOk"] = (
            out["WallNormalAngle"] is not None
            and out["WallEndDistance"] is not None
            and (out["WallPlanReliability"] or 0.0) >= 2.0
        )
    else:
        out["GeomOk"] = out["Angle"] is not None and out["PerpRel"] is not None
    return out



@mcp.tool()
def get_bridge_health() -> dict:
    """Проверить, отвечает ли локальный Navisworks HTTP-мост.

    Это только диагностика транспорта и открытого Clash Detective: сервер не
    применяет правила, пороги проекта и не принимает инженерные решения.
    """
    health = {
        "bridge_url": BRIDGE_URL,
        "ok": False,
        "tests_available": False,
        "test_count": 0,
        "result_count_total": 0,
        "test_status_counts": {},
        "error": None,
        "hint": None,
    }
    try:
        with _client(timeout=10.0) as c:
            r = c.get("/clash/tests")
            r.raise_for_status()
            tests = r.json()
    except Exception as exc:
        health["error"] = str(exc)
        health["hint"] = _bridge_error_hint(exc)
        return health

    health["ok"] = True
    health["tests_available"] = isinstance(tests, list)
    health["test_count"] = len(tests) if isinstance(tests, list) else 0
    health["result_count_total"] = sum(int(test.get("ResultCount") or 0) for test in tests)
    health["test_status_counts"] = dict(Counter(str(test.get("Status") or "") for test in tests))
    return health


@mcp.tool()
def get_clash_test_overview(test_name: str) -> dict:
    """Краткая карточка теста Clash Detective без чтения всех результатов.

    Возвращает только данные Navisworks: статус теста, количество результатов,
    тип проверки и настроенный tolerance. Порогов проекта и вердиктов здесь нет.
    """
    try:
        with _client() as c:
            r = c.get("/clash/tests")
            r.raise_for_status()
            tests = r.json()
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc

    test = _find_clash_test(tests, test_name)
    tolerance = test.get("Tolerance")
    tolerance_mm = round(float(tolerance) * FT2MM, 1) if tolerance is not None else None
    return {
        "Name": test.get("Name"),
        "Status": test.get("Status"),
        "ResultCount": test.get("ResultCount"),
        "Tolerance": tolerance,
        "ToleranceMm": tolerance_mm,
        "ToleranceType": test.get("ToleranceType"),
        "TestType": test.get("TestType"),
    }


@mcp.tool()
def get_clash_status_summary(test_name: str) -> dict:
    """Сводка текущих статусов результатов без тяжёлых путей и геометрии.

    Команда нужна для быстрой проверки после пакетных изменений и для оценки
    объёма работы. Она не классифицирует коллизии и не предлагает статусы.
    """
    try:
        with _client() as c:
            tests_response = c.get("/clash/tests")
            tests_response.raise_for_status()
            test = _find_clash_test(tests_response.json(), test_name)

            r = c.get(
                f"/clash/tests/{_q(test_name)}/results",
                params={"includePaths": "false"},
            )
            r.raise_for_status()
            results = r.json()
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc

    status_counts = Counter(str(item.get("Status") or "") for item in results)
    bucket_counts = Counter(_status_bucket(item.get("Status")) for item in results)
    return {
        "test_name": test.get("Name"),
        "test_status": test.get("Status"),
        "result_count_from_test": test.get("ResultCount"),
        "rows_read": len(results),
        "status_counts": dict(status_counts),
        "status_bucket_counts": dict(bucket_counts),
    }


@mcp.tool()
def list_clash_tests() -> list[dict]:
    """Список тестов Clash Detective в текущем открытом документе Navisworks:
    имя, статус теста (NEW/OLD/PARTIAL/OK), количество результатов."""
    try:
        with _client() as c:
            r = c.get("/clash/tests")
            r.raise_for_status()
            return r.json()
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc


@mcp.tool()
def get_clash_test_results(
    test_name: str,
    include_paths: bool = True,
    status: str | None = None,
    include_geometry: bool = True,
    limit: int | None = None,
    offset: int = 0,
    include_item_bounds: bool = False,
    include_sizes: bool = False,
    include_worksets: bool = False,
    include_metrics: bool = False,
    keep_raw_bounds: bool = False,
) -> list[dict]:
    """Результаты конкретного теста Clash Detective (по имени теста, как
    в list_clash_tests): пары столкнувшихся элементов (путь в дереве модели),
    статус результата (NEW/ACTIVE/APPROVED/RESOLVED/REVIEWED/UNMERGED),
    дистанция, кто утвердил, комментарии, геометрия (Pt1/Pt2/Bound —
    точки контакта и bounding box зоны пересечения, единицы — ФУТЫ).

    include_metrics=True — ГЛАВНЫЙ режим для разбора допусков. Сервер сам
    запрашивает у моста габариты элементов и добавляет к каждому результату
    готовые числа (уже в миллиметрах и градусах):
        Kind1/Kind2/PairKind — изоляция | фитинг | голая труба (по пути);
        Angle    — угол между осями элементов, 0..90° (0 — продольное,
                   90 — поперечное пересечение);
        PerpRel  — промах мимо оси более длинного элемента, в долях
                   большего диаметра (0 — точно через ось, >=0.2 — по
                   касательной);
        Dia1/Dia2, Len1/Len2, Elong1/Elong2 — диаметр, длина и вытянутость
                   (длина/диаметр) каждого элемента;
        Group    — имя группы коллизий Navisworks (или null);
        Size1/Size2 — строка свойства "Размер" как есть;
        SectionMin/SectionMax — грани НОМИНАЛЬНОГО сечения из этой строки
                   ("ø200 мм-ø200 мм" -> 200/200; "300x200-300x200" ->
                   200/300). SectionSource говорит, откуда взято: свойство
                   или, если свойства нет, узкая грань габарита;
        MepMinEdge — узкая грань габарита элемента (запасной вариант);
        Through/StructThick — прошёл ли насквозь и толщина конструкции;
        WallNormalAngle — угол оси инженерного элемента к нормали стены:
                   0° — нормальный проход через широкую грань, 90° —
                   продольное движение в плоскости стены;
        WallEndDistance — расстояние центра коллизии до внешнего торца стены;
        WallBroadFaceDistance/WallEndFaceDistance — расстояния до ближайшей
                   широкой и торцевой/откосной mesh-грани рядом с коллизией;
                   WallPlanThickness/WallPlanReliability/WallPlanSource
                   описывают mesh-PCA, из которой восстановлена ось стены;
        DistMm   — Distance в миллиметрах;
        GeomOk   — удалось ли посчитать; GeomNote — почему нет / почему
                   ненадёжно (ббокс "union-of-N").
    Пороги и вердикты сервер НЕ применяет — это дело навыка
    navisworks-clash-tolerances.

    keep_raw_bounds=True оставит в ответе сырые Item1Bound/Item2Bound.
    По умолчанию при include_metrics=True они выбрасываются: 150 боксов —
    это как раз то, что упирается в лимит размера ответа инструмента,
    а вся информация из них уже свёрнута в метрики.

    include_paths=False пропускает резолв пути элементов (Item1Path/
    Item2Path будут null) — на тестах с тысячами результатов резолв путей
    может быть медленным/таймаутить. Для include_metrics пути нужны
    (по ним определяется изоляция/фитинг), поэтому он их включает сам.

    status — необязательный фильтр по статусу (New/Active/Approved/
    Resolved/Reviewed). На тестах с тысячами результатов почти всегда
    реально нужны только "открытые" (например, Active) — фильтрация на
    стороне аддина на порядок сокращает время ответа.

    include_geometry=False убирает поля Pt1/Pt2/Bound из каждого результата
    перед возвратом. Обрезка происходит здесь, на стороне MCP-сервера,
    ПОСЛЕ получения полного ответа от моста. На тестах с тысячами
    результатов даже с status="Active" полный ответ может упираться в лимит
    размера ответа инструмента (~1 МБ) — include_geometry=False обычно
    решает это без потери status/Distance.

    limit/offset — нарезка результата на стороне MCP-сервера (сам мост
    пагинацию не поддерживает, поэтому весь список сначала забирается
    целиком, а затем обрезается перед отправкой обратно).

    include_item_bounds=True добавляет Item1Bound/Item2Bound — bounding
    box КАЖДОГО ЭЛЕМЕНТА пары ЦЕЛИКОМ (не зоны пересечения, как Bound),
    в футах. Нужен, если хочется считать геометрию самостоятельно;
    для обычного разбора берите include_metrics=True.

    include_worksets=True добавляет Item1Workset/Item2Workset — значение
    свойства "Рабочий набор" с вкладки "Объект". Мост ищет его у самого
    clash-узла и выше по AncestorsAndSelf, потому что геометрия часто лежит
    ниже Revit-элемента с параметрами.

    MCP-сервер не знает, обучен ли тест и какие проектные параметры нужны.
    Он только возвращает данные. Решение о применимости правил, запрос
    проектного порога и классификация результатов живут в skill."""
    need_bounds = include_item_bounds or include_metrics
    need_paths = include_paths or include_metrics
    need_sizes = include_sizes or include_metrics
    try:
        with _client() as c:
            # ВНИМАНИЕ: это query-строка МОСТА, а не параметры проекта из
            # аргумента params — имена не путать, иначе одно затрёт другое.
            query = {"includePaths": "true" if need_paths else "false"}
            if status:
                query["status"] = status
            if need_bounds:
                query["includeItemBounds"] = "true"
            if need_sizes:
                query["includeSizes"] = "true"
            if include_worksets:
                query["includeWorksets"] = "true"
            r = c.get(f"/clash/tests/{_q(test_name)}/results", params=query)
            r.raise_for_status()
            results = r.json()
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc

    if include_metrics:
        for item in results:
            item.update(_metrics(item))

    if not keep_raw_bounds and not include_item_bounds:
        for item in results:
            item.pop("Item1Bound", None)
            item.pop("Item2Bound", None)

    if not include_geometry:
        for item in results:
            item.pop("Pt1", None)
            item.pop("Pt2", None)
            item.pop("Bound", None)

    if limit is not None:
        results = results[offset : offset + limit]
    elif offset:
        results = results[offset:]

    return results


@mcp.tool()
def run_clash_test(test_name: str) -> str:
    """Запустить/пересчитать тест Clash Detective по имени."""
    try:
        with _client() as c:
            r = c.post(f"/clash/tests/{_q(test_name)}/run")
            r.raise_for_status()
            return "Тест запущен"
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc


@mcp.tool()
def set_clash_result_status(
    test_name: str,
    result_name: str,
    status: str,
    comment: str = "",
) -> str:
    """Проставить статус результату коллизии.
    status: один из New, Active, Approved, Resolved, Reviewed.

    MCP-сервер не проверяет обученность теста и не принимает инженерное
    решение. Вызывать изменение статуса следует только после решения skill
    и явного пользовательского разрешения.
    ВНИМАНИЕ: параметр comment пока не реализован на стороне аддина
    (добавление комментария не удалось проверить без запуска в реальном
    Navisworks) — если передать непустой comment, запрос вернёт ошибку 501."""
    try:
        with _client() as c:
            r = c.post(
                f"/clash/results/{_q(test_name)}/{_q(result_name)}/status",
                json={"status": status, "comment": comment},
            )
            r.raise_for_status()
            return "Статус обновлён"
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc


@mcp.tool()
def set_clash_group_status(
    test_name: str,
    group_name: str,
    status: str,
) -> dict:
    """Проставить один статус всем дочерним результатам группы Clash Detective
    за один запрос к Navisworks.

    Инструмент не решает, можно ли утверждать группу целиком. Вызывать его
    следует только после проверки каждого дочернего результата применимым
    ruleset и явного разрешения пользователя на запись.
    """
    return batch_set_clash_group_status(test_name, [group_name], status)


@mcp.tool()
def batch_set_clash_group_status(
    test_name: str,
    group_names: list[str],
    status: str,
) -> dict:
    """Проставить один статус нескольким проверенным однородным группам.

    В отличие от batch_set_clash_result_status, мост получает один HTTP-запрос
    и за один проход по результатам теста обновляет все дочерние коллизии.
    Возвращает число обновлённых результатов по каждой группе; нулевое значение
    означает, что группа не найдена и требует отдельной проверки.
    """
    normalized = list(dict.fromkeys(name.strip() for name in group_names if name and name.strip()))
    if not normalized:
        raise ValueError("Не передано ни одного имени группы")

    try:
        with _client() as c:
            r = c.post(
                f"/clash/groups/{_q(test_name)}/status",
                json={"groupNames": normalized, "status": status, "comment": ""},
            )
            r.raise_for_status()
            payload = r.json()
            counts = payload.get("groups", {})
            return {
                "test_name": test_name,
                "status": status,
                "updated": payload.get("updated", sum(counts.values())),
                "groups": counts,
                "missing": [name for name in normalized if not counts.get(name)],
            }
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc


@mcp.tool()
def batch_set_clash_result_status(
    test_name: str,
    updates: list[dict],
) -> dict:
    """Пакетно проставить статус нескольким результатам ОДНОГО теста Clash
    Detective одним HTTP-запросом и одним проходом по результатам теста.

    updates — список объектов вида:
        [{"result_name": "<имя результата>", "status": "Approved"}]

    status в каждом элементе — один из New, Active, Approved, Resolved,
    Reviewed. Поле "comment" внутри элемента, если передано, игнорируется —
    как и в set_clash_result_status, комментарии сейчас не реализованы на
    стороне аддина и вызывают 501 при непустом значении.

    Мост сопоставляет все имена за один перебор test.results() и меняет
    найденные статусы внутри одного вызова главного потока Navisworks.
    Ненайденные элементы возвращаются в missing, ошибки отдельных записей —
    в failed; они не откатывают успешно применённые изменения.

    MCP-сервер не проверяет обученность теста и не выбирает статусы.
    Перед вызовом skill должен уже сформировать список изменений, а
    пользователь должен явно разрешить запись.
    """
    normalized: list[dict] = []
    failed: list[dict] = []
    seen: dict[str, str] = {}
    for item in updates:
        result_name = item.get("result_name")
        status = item.get("status")
        if not result_name or not status:
            failed.append(
                {"result_name": result_name, "error": "отсутствует result_name или status"}
            )
            continue
        previous = seen.get(result_name)
        if previous and previous.lower() != str(status).lower():
            failed.append(
                {"result_name": result_name, "error": "для результата переданы разные статусы"}
            )
            continue
        if previous:
            continue
        seen[result_name] = str(status)
        normalized.append({"resultName": result_name, "status": status})

    if not normalized:
        return {"test_name": test_name, "updated": [], "missing": [], "failed": failed}

    try:
        with _client(timeout=300.0) as c:
            r = c.post(
                f"/clash/results/{_q(test_name)}/status",
                json={"updates": normalized},
            )
            r.raise_for_status()
            payload = r.json()
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc

    failed.extend(
        {"result_name": name, "error": error}
        for name, error in (payload.get("failed") or {}).items()
    )
    return {
        "test_name": test_name,
        "updated": payload.get("updated", []),
        "missing": payload.get("missing", []),
        "failed": failed,
    }


@mcp.tool()
def export_clash_report(test_name: str, format: str = "html") -> str:
    """Экспортировать отчёт по тесту Clash Detective (html или xml).
    ПОКА НЕ РЕАЛИЗОВАНО на стороне аддина — точный механизм экспорта отчёта
    через COM API 2020 не определён статически, нужна проверка в реальном
    Navisworks (см. README и комментарий в ClashService.ExportReport)."""
    try:
        with _client() as c:
            r = c.get(f"/clash/tests/{_q(test_name)}/export", params={"format": format})
            r.raise_for_status()
            return r.json()["path"]
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc


if __name__ == "__main__":
    mcp.run(transport="stdio")
