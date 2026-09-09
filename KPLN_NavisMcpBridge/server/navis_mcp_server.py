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
  * этот MCP-сервер     — АРИФМЕТИКА: перевод футов в мм, векторы осей
                          элементов, угол между ними, вытянутость,
                          относительный промах мимо оси, классификация
                          стороны по пути (изоляция/фитинг/голая труба).
                          Никаких порогов и никаких вердиктов.
  * навык (SKILL.md)    — ПОЛИТИКА: пороги (5°, 0.15, e>=5), порядок
                          правил, исключения регламента, что делать
                          с "не проверено". Меняется без перезапуска.
"""

import difflib
import fnmatch
import json
import math
import re
from datetime import date
from pathlib import Path
from urllib.parse import quote

import httpx
from mcp.server.fastmcp import FastMCP

BRIDGE_URL = "http://127.0.0.1:8765"

# Сырые длины COM API Navisworks приходят в ФУТАХ (проверено: Tolerance
# 0.032808399 == ровно 0.01 м). Всё, что этот сервер отдаёт в полях
# метрик, уже переведено в миллиметры.
FT2MM = 304.8

mcp = FastMCP("KPLN_NavisMcpBridge")


def _client() -> httpx.Client:
    return httpx.Client(base_url=BRIDGE_URL, timeout=90.0)


def _q(name: str) -> str:
    return quote(name, safe="")


def _bridge_error_hint(exc: Exception) -> str:
    if isinstance(exc, httpx.ConnectError):
        return (
            "Не удалось достучаться до Navisworks (127.0.0.1:8765). "
            "Убедитесь, что Navisworks Manage 2020 запущен, документ открыт, "
            "и на вкладке Add-ins один раз нажата кнопка 'MCP Bridge'."
        )
    return str(exc)


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
# Отделочные слои: штукатурка по сетке, облицовка, утеплитель. По регламенту
# отверстия в них НЕ МОДЕЛИРУЮТСЯ ВООБЩЕ — пересечение с ними всегда допуск.
# Проверяется РАНЬШЕ стены: путь такого слоя тоже содержит "Базовая стена".
_STRUCT = ("стена", "перекрытие", "отделка")
# Элементы, у которых габарит bounding box'а НЕ равен сечению воздуховода:
# у отвода R=1.5 коробка вдвое больше диаметра, у противопожарного клапана
# в неё входят корпус и привод. Правило "сечение больше порога" к ним
# применять нельзя.
_BBOX_NOT_SECTION = ("фасонина", "клапан", "фитинг")
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
    """Готовые геометрические метрики одной пары. Только числа и
    классификация — никаких порогов и вердиктов (они в навыке)."""
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
    else:
        out["MepMinEdge"] = None
        out["StructThick"] = None
        out["Through"] = None
        out["SectionMin"] = None
        out["SectionMax"] = None
        out["SectionSource"] = None

    if struct_side:
        out["GeomNote"] = (
            "в паре есть конструкция: Dia/Angle/Elong посчитаны по её "
            "bounding box и СМЫСЛА НЕ ИМЕЮТ. Смотреть MepMinEdge (сечение "
            "инженерного элемента), StructThick, Through и DistMm."
        )

    out["GeomOk"] = out["Angle"] is not None and out["PerpRel"] is not None
    return out


# ---------------------------------------------------------------------------
# СЛОВАРЬ ПРОВЕРОК (clash_rules.json)
#
# Имена тестов Clash Detective у KPLN статичны, и именно имя — ключ к набору
# правил разбора. Словарь лежит РЯДОМ с этим файлом отдельным JSON и читается
# при КАЖДОМ вызове (не кэшируется) — значит новую обученную проверку можно
# добавить прямо во время работы, без перезапуска Claude Desktop.
#
# Смысл словаря — жёсткий отказ вместо угадывания:
#   * имени нет в открытом документе      -> отказ + список похожих имён (опечатка);
#   * имя есть, но нет в словаре          -> отказ давать вердикты и ставить
#                                            статусы; читать результаты можно
#                                            (иначе не на чем обучаться);
#   * status="assumed"                    -> набор правил предполагается по
#                                            аналогии, но НЕ проверен на
#                                            размеченных данных: тоже отказ,
#                                            пока пользователь не подтвердит.
# ---------------------------------------------------------------------------

RULES_PATH = Path(__file__).with_name("clash_rules.json")


def _load_registry() -> dict:
    try:
        with RULES_PATH.open(encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {"version": 1, "rulesets": {}, "tests": []}
    except Exception as exc:
        raise RuntimeError(
            f"Словарь проверок {RULES_PATH} повреждён и не читается: {exc}. "
            "Разбор коллизий остановлен — сначала почините файл."
        ) from exc


def _save_registry(reg: dict) -> None:
    reg["updated"] = date.today().isoformat()
    tmp = RULES_PATH.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(reg, f, ensure_ascii=False, indent=2)
    tmp.replace(RULES_PATH)


def _match_entry(test_name: str, reg: dict) -> dict | None:
    """Ищет строку словаря для имени теста. Точное совпадение приоритетнее
    шаблона; из шаблонов выигрывает самый длинный (самый специфичный)."""
    best = None
    for entry in reg.get("tests", []):
        pattern = entry.get("pattern", "")
        kind = entry.get("match", "exact")
        hit = (pattern == test_name) if kind == "exact" else fnmatch.fnmatch(test_name, pattern)
        if not hit:
            continue
        if kind == "exact":
            return entry
        if best is None or len(pattern) > len(best.get("pattern", "")):
            best = entry
    return best


def _known_patterns(reg: dict) -> list[str]:
    """Шаблоны имён, которые есть в словаре — для подсказки при опечатке.
    Берутся ИЗ ЛОКАЛЬНОГО ФАЙЛА: путь отказа не должен ходить в Navisworks,
    иначе "СТОП" на незнакомом тесте стоит столько же, сколько сам разбор."""
    return [e.get("pattern", "") for e in reg.get("tests", []) if e.get("pattern")]


def _require_ruleset(test_name: str, allow_untrained: bool = False) -> dict:
    """Возвращает набор правил для теста или ОТКАЗЫВАЕТ с объяснением."""
    reg = _load_registry()
    entry = _match_entry(test_name, reg)

    if entry is None:
        patterns = _known_patterns(reg)
        close = difflib.get_close_matches(test_name, patterns, n=3, cutoff=0.5)
        hint = "\nВ словаре есть: " + (
            ", ".join(f'"{n}"' for n in (close or patterns[:8])) or "(пусто)"
        )
        if allow_untrained:
            return {
                "test_name": test_name,
                "status": "untrained",
                "ruleset": None,
                "message": "Тест не описан в словаре — правил нет." + hint,
            }
        raise RuntimeError(
            f'СТОП. Теста "{test_name}" нет в словаре проверок '
            f"({RULES_PATH.name}) — правил для него нет."
            + hint
            + "\nРАБОТА ОСТАНОВЛЕНА: не читать результаты, не считать метрики, "
            "не ставить статусы, не угадывать правила по аналогии. Сказать об "
            "этом пользователю и ждать. Разбирать данные этого теста можно "
            "только если пользователь ЯВНО попросил обучить новую проверку "
            "(тогда get_clash_test_results с allow_untrained=True), а по итогам "
            "обучения занести её через register_clash_test."
        )

    status = entry.get("status", "untrained")
    rs_name = entry.get("ruleset")
    rs = (reg.get("rulesets") or {}).get(rs_name)
    out = {
        "test_name": test_name,
        "matched_pattern": entry.get("pattern"),
        "status": status,
        "ruleset_name": rs_name,
        "ruleset": rs,
        "note": entry.get("note"),
    }
    if status != "trained" and not allow_untrained:
        raise RuntimeError(
            f'СТОП. Тест "{test_name}" числится в словаре как "{status}", а не '
            f'"trained": набор правил "{rs_name}" предположен по аналогии и на '
            "размеченных данных этого теста НЕ проверялся.\nРАБОТА ОСТАНОВЛЕНА: "
            "спросить пользователя, действуют ли эти правила, и только после его "
            "подтверждения перевести строку словаря в trained через "
            "register_clash_test."
        )
    if rs is None:
        raise RuntimeError(
            f'Словарь повреждён: тест "{test_name}" ссылается на набор правил '
            f'"{rs_name}", которого нет в разделе rulesets.'
        )
    return out


def _require_params(info: dict, params: dict | None) -> dict:
    """Параметры проекта, без которых набор правил применять нельзя.

    Часть правил зависит не от геометрии, а от соглашения ПРОЕКТА — например,
    с какой грани отверстие вообще моделируется (в одном проекте 250 мм, в
    другом другое). Такие параметры перечислены в наборе правил как
    params_required, и пока пользователь не назвал их значения, инструмент
    ОТКАЗЫВАЕТ и печатает вопросы, которые надо ему задать. Значения нигде
    не запоминаются между вызовами намеренно: молча утащить в новый проект
    число из прошлого — худшая из возможных ошибок здесь.
    """
    rs = info.get("ruleset") or {}
    required = rs.get("params_required") or []
    if not required:
        return {}
    params = params or {}
    missing = [p for p in required if p.get("key") not in params]
    if missing:
        qs = "\n".join(
            f"  - {p.get('key')}: {p.get('question')}"
            + (f" (единицы: {p['unit']})" if p.get("unit") else "")
            + (f" [в разобранном примере было {p['example']}]" if p.get("example") is not None else "")
            for p in missing
        )
        raise RuntimeError(
            f'СТОП. Набор правил "{info.get("ruleset_name")}" зависит от '
            "соглашений ПРОЕКТА, а они не заданы.\nРАБОТА ОСТАНОВЛЕНА: сначала "
            "задать пользователю эти вопросы, получить числа и передать их в "
            "параметре params:\n" + qs
        )
    return params


@mcp.tool()
def list_clash_rulesets() -> dict:
    """Словарь проверок целиком: какие тесты разобраны, по каким наборам правил,
    в каком статусе обучения (trained / assumed / untrained) и на каких данных
    это проверялось.

    Смотреть ПЕРЕД разбором незнакомого теста и перед проходом по всему файлу:
    тесты, которых здесь нет, разбирать нельзя — по ним будет отказ."""
    return _load_registry()


@mcp.tool()
def get_clash_ruleset(test_name: str, allow_untrained: bool = False) -> dict:
    """Набор правил для конкретного теста Clash Detective по его имени.

    Если тест не описан в словаре (или описан как "assumed" — правила
    предполагаются, но не проверены), инструмент ОТКАЗЫВАЕТ с объяснением
    и, при опечатке в имени, со списком похожих имён из открытого документа.
    Это защита от разбора проверки по правилам, которым вас не учили.

    allow_untrained=True — вернуть описание без отказа (что известно, что нет).
    Использовать только чтобы ПОКАЗАТЬ пользователю состояние, а не чтобы
    обойти отказ и всё-таки проставить статусы."""
    return _require_ruleset(test_name, allow_untrained=allow_untrained)


@mcp.tool()
def register_clash_test(
    pattern: str,
    ruleset: str,
    status: str = "trained",
    note: str = "",
    match: str = "glob",
    ruleset_definition: dict | None = None,
    trained_on: str = "",
) -> dict:
    """Занести проверку в словарь — ЕДИНСТВЕННЫЙ способ добавить новую
    обученную проверку. Вызывать после того, как пользователь разметил отчёт
    и правила проверены на его разметке.

    pattern — имя теста или шаблон. Префикс со стадией/номером у KPLN меняется
    от модели к модели, поэтому обычно нужен шаблон вида
    "*ОВ1---Трубы (Самопересечение)", а не точное имя из одной модели.
    match — "glob" (по умолчанию) или "exact".

    ruleset — имя набора правил. Если такого набора ещё нет, ОБЯЗАТЕЛЬНО
    передать ruleset_definition — словарь с описанием: title, skill, section,
    thresholds (числовые пороги), pair_kinds_ok, pair_kinds_unknown.
    Сами правила (порядок проверок, что они значат) живут в навыке — здесь
    только пороги и ссылка на раздел навыка.

    status — "trained" (проверено на размеченных данных этого теста),
    "assumed" (предполагается по аналогии — сервер будет отказывать до
    подтверждения) или "untrained".

    trained_on — одна строка про данные, на которых проверено: модель, тест,
    сколько результатов, какая точность. Дописывается к набору правил.

    Повторный вызов с тем же pattern обновляет строку, а не плодит дубли."""
    if status not in ("trained", "assumed", "untrained"):
        raise RuntimeError('status должен быть "trained", "assumed" или "untrained"')
    if match not in ("glob", "exact"):
        raise RuntimeError('match должен быть "glob" или "exact"')

    reg = _load_registry()
    reg.setdefault("rulesets", {})
    reg.setdefault("tests", [])

    if ruleset not in reg["rulesets"]:
        if not ruleset_definition:
            raise RuntimeError(
                f'Набора правил "{ruleset}" в словаре нет. Передайте '
                "ruleset_definition с описанием (title, skill, section, thresholds, "
                "pair_kinds_ok, pair_kinds_unknown) или укажите существующий набор: "
                + (", ".join(reg["rulesets"]) or "(словарь пуст)")
            )
        reg["rulesets"][ruleset] = dict(ruleset_definition)
    elif ruleset_definition:
        reg["rulesets"][ruleset].update(ruleset_definition)

    if trained_on:
        reg["rulesets"][ruleset].setdefault("trained_on", []).append(trained_on)

    entry = {
        "match": match,
        "pattern": pattern,
        "ruleset": ruleset,
        "status": status,
        "note": note,
    }
    for i, old in enumerate(reg["tests"]):
        if old.get("pattern") == pattern:
            reg["tests"][i] = entry
            break
    else:
        reg["tests"].append(entry)

    _save_registry(reg)
    return {"saved": entry, "rulesets": list(reg["rulesets"]), "tests_total": len(reg["tests"])}


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
    include_metrics: bool = False,
    keep_raw_bounds: bool = False,
    allow_untrained: bool = False,
    params: dict | None = None,
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

    СТОП-ПРАВИЛО: если теста нет в словаре проверок со статусом
    "trained", инструмент ОТКАЗЫВАЕТ СРАЗУ — до любого обращения к
    Navisworks, не читая ни одного результата. Не тратить время и
    токены на разбор проверки, правилам которой не учили.

    allow_untrained=True — РЕЖИМ ОБУЧЕНИЯ: читать данные теста,
    которого нет в словаре. Ставить ТОЛЬКО когда пользователь явно
    попросил разобрать/разметить новую проверку. Статусы по ней всё
    равно не проставить, пока она не занесена через register_clash_test.

    params — значения проектных соглашений, которых требует набор правил
    (например {"opening_min_edge_mm": 250}). Если набор их требует, а они не
    переданы — тоже отказ со списком вопросов к пользователю."""
    if not allow_untrained:
        _require_params(_require_ruleset(test_name), params)

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
            r = c.get(f"/clash/tests/{_q(test_name)}/results", params=query)
            if r.status_code == 409:
                names = [t.get("Name") for t in c.get("/clash/tests").json() if t.get("Name")]
                close = difflib.get_close_matches(test_name, names, n=5, cutoff=0.4)
                raise RuntimeError(
                    f'ОТКАЗ: теста "{test_name}" нет в открытом документе.'
                    + (
                        "\nПохожие имена: " + ", ".join(f'"{n}"' for n in close)
                        if close
                        else "\nПолный список — list_clash_tests."
                    )
                )
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
    force_untrained: bool = False,
    params: dict | None = None,
) -> str:
    """Проставить статус результату коллизии.
    status: один из New, Active, Approved, Resolved, Reviewed.

    ОТКАЗЫВАЕТ, если тест не описан в словаре проверок как "trained":
    статусы нельзя расставлять по проверке, правилам которой вас не учили.
    force_untrained=True — только когда пользователь ЯВНО в этом разговоре
    попросил проставить статусы вручную, вопреки отсутствию правил.
    ВНИМАНИЕ: параметр comment пока не реализован на стороне аддина
    (добавление комментария не удалось проверить без запуска в реальном
    Navisworks) — если передать непустой comment, запрос вернёт ошибку 501."""
    if not force_untrained:
        _require_params(_require_ruleset(test_name), params)
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
def batch_set_clash_result_status(
    test_name: str,
    updates: list[dict],
    force_untrained: bool = False,
    params: dict | None = None,
) -> dict:
    """Пакетно проставить статус нескольким результатам ОДНОГО теста Clash
    Detective за один вызов инструмента — вместо цепочки вызовов
    set_clash_result_status по одному на каждый результат.

    updates — список объектов вида:
        [{"result_name": "Конфликт5", "status": "Approved"},
         {"result_name": "Конфликт7", "status": "Approved"}]

    status в каждом элементе — один из New, Active, Approved, Resolved,
    Reviewed. Поле "comment" внутри элемента, если передано, игнорируется —
    как и в set_clash_result_status, комментарии сейчас не реализованы на
    стороне аддина и вызывают 501 при непустом значении.

    ВНИМАНИЕ: сам мост KPLN_NavisMcpBridge пакетного эндпоинта не имеет — это
    обёртка на стороне MCP-сервера, которая внутри одного вызова инструмента
    последовательно шлёт мосту по одному HTTP-запросу на каждый result_name
    (тот же эндпоинт, что и set_clash_result_status). Экономится не сетевое
    время, а количество ходов туда-обратно на уровне диалога/инструмента.

    Пачки по ~40 штук — рабочий размер. Таймаут или разрыв соединения на
    большой пачке НЕ означает провал: операция довыполняется в фоне,
    надо подождать ~40 сек и проверить лёгким read-only вызовом, а не
    слепо повторять пачку.

    Ошибка на одном result_name не прерывает обработку остальных: такие
    элементы попадают в "failed" с текстом ошибки, а не откатывают весь
    пакет. Возвращает:
        {"test_name": ..., "updated": [...], "failed": [{"result_name":..., "error":...}]}

    ОТКАЗЫВАЕТ целиком (ничего не пишет), если тест не описан в словаре
    проверок как "trained" — см. get_clash_ruleset. force_untrained=True
    только по явной просьбе пользователя в этом разговоре.
    """
    if not force_untrained:
        _require_params(_require_ruleset(test_name), params)
    updated: list[str] = []
    failed: list[dict] = []
    try:
        with _client() as c:
            for item in updates:
                result_name = item.get("result_name")
                status = item.get("status")
                if not result_name or not status:
                    failed.append(
                        {
                            "result_name": result_name,
                            "error": "отсутствует result_name или status",
                        }
                    )
                    continue
                try:
                    r = c.post(
                        f"/clash/results/{_q(test_name)}/{_q(result_name)}/status",
                        json={"status": status, "comment": ""},
                    )
                    r.raise_for_status()
                    updated.append(result_name)
                except Exception as item_exc:
                    failed.append({"result_name": result_name, "error": str(item_exc)})
    except Exception as exc:
        raise RuntimeError(_bridge_error_hint(exc)) from exc

    return {"test_name": test_name, "updated": updated, "failed": failed}


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
