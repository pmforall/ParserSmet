#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reference.py — справочник укрупнённых физобъёмов и техники («Справочник новый.xlsx»).

Назначение
----------
Точный питоновский аналог браузерного модуля ``Reference`` из
``app/template.html`` (раздел «5A. СПРАВОЧНИК УКРУПНЕННЫХ ФИЗОБЪЁМОВ И ТЕХНИКИ»).
Поведение повторяется один в один: те же ключи, та же нормализация единиц
измерения, те же округления и тот же порядок сортировки строк отчётов.

Из справочника читаются три вкладки:

    «Расценки ФО»              — расценка → укрупнённый физобъём («Для ВПР»)
                                 и нормализованная единица измерения («ЕИ Примы»);
    «Множитель»                — единица «В смете» → множитель и ЕИ (нормализация);
    «Расценки ТЕХНИКА и ТРЗ»   — расценка машины/ТРЗ → код ресурса Primavera.

Отличия от браузерного модуля (только технические)
--------------------------------------------------
* Строка листа — обычный список, **индекс 0 = столбец A** (в браузере строки
  собираются как 1-индексные разреженные массивы, где индекс 1 = столбец A).
  Поэтому все обращения ``row[1..6]`` эталона сдвинуты на единицу:
  ``code = row[0]``, ``name = row[1]`` и т. д.
* ``load()`` повторяет форму массива строк браузера: перед строками листа
  добавляется пустая строка-заглушка (в браузере индекс 0 массива пуст). Из-за
  этого правило «первая строка = шапка» (``i == 0``) на реальной шапке не
  срабатывает — ровно как в эталоне, где шапка лежит по индексу 1. Благодаря
  этому число записей и поле ``row`` совпадают с браузером до единицы.
* Вместо индекса по «всем цифрам» (``foByDigits``) для поиска расценки ФО
  используется индекс «без букв» (``foByNoLetters``): ключ без букв сохраняет
  ведущий дефис, поэтому «ФЕР-01-01-001-01» не совпадает с «М01-01-001-01».
  Индекс ``foByDigits`` сохранён в результате для совместимости, но в поиске
  не участвует.

Отчёт по физобъёмам (``build_fo_report``)
-----------------------------------------
Каждая позиция может нести два набора показателей: ``values`` (обычно уровень
строки «ВСЕГО ПО СМЕТЕ» вкладки Source) и ``values_base`` (базисный уровень).
Если ``values_base`` нет, он считается равным ``values`` (тогда
``cost_base == cost``). Отчёт возвращает агрегаты по группам (с полями
``cost_base`` и ``unit_cost_base``), статистику (с ``cost_base``,
``unmatched_cost``, ``unmatched_all``, ``unmatched_skipped``, ``excluded_cost``),
а также перечни позиций ``unmatched`` (расценка не найдена в справочнике; только
категории :data:`UNMATCHED_CATS` — «Работы» и «Перевозка грузов») и
``excluded_positions`` (найдена, но группа «Не идет в физику»). Позиции прочих
категорий без расценки в перечень не выгружаются, но учтены в ``unmatched_all``
и ``unmatched_skipped``.

Запуск самопроверки
-------------------
    python parser/reference.py [путь_к_справочнику.xlsx] [путь_к_smetas.json]

По умолчанию: ``/tmp/spr.xlsx`` и ``<репозиторий>/data/smetas.json``.

Зависимости: openpyxl (поставляется со средой DSH).
"""

from __future__ import annotations

import json
import math
import os
import re
import sys
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:  # openpyxl нужен только функции load(); разбор/отчёты работают и без него
    import openpyxl
except ImportError:  # pragma: no cover — среда без openpyxl
    openpyxl = None  # type: ignore[assignment]


# ---------------------------------------------------------------------------
# Константы (имена — как в браузерном эталоне)
# ---------------------------------------------------------------------------

SHEET_FO = "Расценки ФО"
SHEET_MULT = "Множитель"
SHEET_TECH = "Расценки ТЕХНИКА и ТРЗ"
NO_FO = "Вне справочника ФО"
EXCLUDED = "Не идет в физику"
# Группа-строка «не распределено»: разница между суммой позиций сметы и её итогом
# «ВСЕГО ПО СМЕТЕ» (внутреннее расхождение исходных книг-корректировок).
RESIDUAL_GROUP = "Не распределено (расхождение позиций с итогом сметы)"
# Лист «ФО не сопоставлено» формируется только по этим категориям затрат
# (коды позиций Source: ``works`` — Работы, ``transport`` — Перевозка грузов).
UNMATCHED_CATS = ["works", "transport"]
UNMATCHED_CAT_LABELS = ["Работы", "Перевозка грузов"]


def is_unmatched_cat(cat: Any) -> bool:
    """True, если категория позиции входит в перечень листа «ФО не сопоставлено»."""
    value = "" if cat is None else str(cat).replace("\u00a0", " ").strip()
    return value in UNMATCHED_CATS

#: Латинские буквы-двойники → кириллические (как в эталоне).
LAT2CYR = {"M": "М", "T": "Т", "K": "К", "X": "Х", "A": "А", "B": "В",
           "E": "Е", "H": "Н", "O": "О", "P": "Р", "C": "С", "Y": "У"}

#: Число, стоящее в начале единицы измерения: «100 м2 покрытия» → 100 + «м2 покрытия».
_LEAD_NUM_RE = re.compile(r"^([0-9]+(?:[.,][0-9]+)?)\s*(.+)$")

#: Служебные единицы, для которых множителя в справочнике нет.
_BASE_UNITS_RE = re.compile(
    r"^(ШТ|М|М2|М3|Т|КГ|КОМП|КОМПЛ|СТЫК|СНИМОК|УЧАСТОК|КОНЕЦ|КОНЦОВ|КАНАТ|М2ПОКРЫТИЯ)$"
)

#: Коды машин/ТРЗ в текстах позиций и ресурсов (в эталоне — /gi).
_TOKEN_RE = re.compile(
    r"(ТСЭМ-[0-9]{3,6}|[0-9]{2}\.[0-9]{2}\.[0-9]{2}-[0-9]{2,4}|[0-9]{6})",
    re.IGNORECASE,
)

#: Number.EPSILON из JavaScript — участвует в округлении эталона.
_EPSILON = 2.220446049250313e-16

_LAT_RE = re.compile(r"[MTKXA BEHOPCY]")


# ---------------------------------------------------------------------------
# Мелкие утилиты (дословный перенос функций эталона)
# ---------------------------------------------------------------------------

def _js_string(v: Any) -> str:
    """Приближение ``String(v)`` из JavaScript (числа без лишнего «.0»)."""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float):
        if math.isnan(v):
            return "NaN"
        if math.isinf(v):
            return "Infinity" if v > 0 else "-Infinity"
        if v.is_integer() and abs(v) < 1e21:
            return str(int(v))
        return repr(v)
    return str(v)


def txt(v: Any) -> str:
    """``txt(v)`` эталона: None → "", иначе строка без неразрывных пробелов по краям."""
    if v is None:
        return ""
    return _js_string(v).replace("\u00a0", " ").strip()


def nz(v: Any) -> str:
    """``nz(v)`` эталона: None → "", иначе строка без обрезки краевых пробелов."""
    if v is None:
        return ""
    return _js_string(v).strip()


def num(v: Any) -> float:
    """``num(v)`` эталона: число из числа/строки, иначе 0.

    Строка чистится от неразрывных пробелов и всех пробельных символов, первая
    запятая заменяется точкой; далее разбирается ведущий числовой фрагмент
    (аналог ``parseFloat``, который не падает на «12 шт.»).
    """
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        f = float(v)
        return f if math.isfinite(f) else 0.0
    if isinstance(v, str):
        s = v.replace("\u00a0", "")
        s = re.sub(r"\s", "", s)
        s = s.replace(",", ".", 1)
        m = re.match(r"^[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", s)
        return float(m.group(0)) if m else 0.0
    return 0.0


def _js_round(x: float) -> float:
    """``Math.round`` из JavaScript (половинка округляется в сторону +∞)."""
    return math.floor(x + 0.5)


def r2(v: float) -> float:
    """Округление до 2 знаков, как ``Math.round((v+Number.EPSILON)*100)/100``."""
    return _js_round((v + _EPSILON) * 100) / 100


def r4(v: float) -> float:
    """Округление до 4 знаков, как ``Math.round((v+Number.EPSILON)*10000)/10000``."""
    return _js_round((v + _EPSILON) * 10000) / 10000


def lo(v: Any) -> str:
    """«Свободный» ключ единицы измерения: верхний регистр, без точек и пробелов.

    «²» → «2», «³» → «3», затем латинские буквы-двойники заменяются
    кириллическими (одним проходом, как регулярка эталона).
    """
    s = _js_string("" if not v else v).upper()
    s = re.sub(r"[.\s\ufeff]", "", s)
    s = s.replace("²", "2").replace("³", "3")
    return _LAT_RE.sub(lambda m: LAT2CYR.get(m.group(0), m.group(0)), s)


def fo_key(code: Any) -> str:
    """Ключ расценки: верхний регистр, без пробелов, без префикса вида «(МО8)»."""
    s = txt(code).upper()
    s = re.sub(r"\s+", "", s)
    return re.sub(r"^\([^)]*\)", "", s)


def fo_key_no_letters(code: Any) -> str:
    """Ключ расценки без букв: «м08-01-067-07» → «08-01-067-07».

    Буквенный префикс оставляет ведущий дефис, поэтому «ФЕР-01-01-001-01»
    превращается в «-01-01-001-01» и не совпадает с «01-01-001-01».
    """
    return re.sub(r"[^0-9.\-]", "", fo_key(code))


def digits(v: Any) -> str:
    """Все цифры значения (аналог ``String(v||"").replace(/\\D/g,"")``)."""
    return re.sub(r"[^0-9]", "", _js_string("" if not v else v))


def _js_truthy(v: Any) -> bool:
    """Истинность значения по правилам JavaScript (пустой объект/массив — истина)."""
    if v is None or v is False:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0 and not (isinstance(v, float) and math.isnan(v))
    if isinstance(v, str):
        return v != ""
    return True


def _js_or(a: Any, b: Any) -> Any:
    """Оператор ``a || b`` из JavaScript (важно для ``{}``: пустой объект истинен)."""
    return a if _js_truthy(a) else b


def _field(vals: Any, name: str) -> Any:
    """Поле набора показателей (аналог ``vals.name`` в JS; не-словарь → None)."""
    return vals.get(name) if isinstance(vals, dict) else None


def _cell(row: Sequence[Any], i: int) -> Any:
    """Значение столбца ``i`` строки листа (индекс 0 = столбец A)."""
    return row[i] if i < len(row) else None


# ---------------------------------------------------------------------------
# Столбцы листов справочника: определение по шапке
# ---------------------------------------------------------------------------
# Порядок столбцов в книгах справочника меняется от выпуска к выпуску: в новой версии
# «Справочник новый.xlsx» на листе «Расценки ФО» столбцы «Тип ресурса» и «Код ресурса
# Primavera» поменялись местами (а в шапке «RSRC» другая структура). Поэтому столбцы
# определяются по НАЗВАНИЮ в первой непустой строке листа; прежний фиксированный порядок
# остаётся запасным вариантом — для книг, где шапки нет вовсе.

# Синонимы названий столбцов (сравнение — по нормализованному тексту шапки)
COLS_FO = {
    "code": ("расценка", "шифр", "код расценки"),
    "name": ("наименование", "название"),
    "primavera": ("код ресурса primavera", "код ресурса", "primavera"),
    "type": ("тип ресурса", "тип"),
    "group": ("для впр", "укрупненный физобъем", "укрупнённый физобъем"),
    "unit": ("еи примы", "еи", "единица измерения"),
}
COLS_TECH = {
    "code": ("расценка", "шифр", "код"),
    "name": ("наименование", "название"),
    "primavera": ("код ресурса primavera", "код ресурса", "primavera"),
    "type": ("тип ресурса", "тип"),
}
COLS_MULT = {
    "unit_in": ("в смете", "единица измерения в смете"),
    "k": ("множитель",),
    "unit_out": ("еи", "единица измерения"),
}
# Запасной вариант — порядок столбцов в книгах прежних выпусков
DEFAULT_FO = {"code": 0, "name": 1, "primavera": 2, "type": 3, "group": 4, "unit": 5}
DEFAULT_TECH = {"code": 0, "name": 1, "primavera": 2, "type": 3}
DEFAULT_MULT = {"unit_in": 0, "k": 1, "unit_out": 2}


def norm_header(v: Any) -> str:
    """Нормализованный текст шапки: нижний регистр, без лишних пробелов и знаков."""
    t = re.sub(r"\s+", " ", str(v if v is not None else "")).strip().lower()
    t = t.replace("ё", "е")
    return t.strip(" .:;,")


def header_cols(row: Optional[Sequence[Any]], spec: Dict[str, Tuple[str, ...]],
                default: Dict[str, int]) -> Dict[str, int]:
    """Номера столбцов по названиям шапки: ``{поле: индекс}``.

    Для каждого поля берётся столбец, название которого совпадает с одним из синонимов
    (сначала точное совпадение, затем — начало названия). Если шапка пустая или поле
    не найдено, подставляется индекс из ``default`` (порядок столбцов прежних выпусков).
    """
    heads = [norm_header(_cell(row or [], i)) for i in range(len(row or []))]
    out: Dict[str, int] = {}
    for field, aliases in spec.items():
        found = None
        for alias in aliases:                       # точное совпадение — приоритет
            for i, h in enumerate(heads):
                if h and h == norm_header(alias):
                    found = i
                    break
            if found is not None:
                break
        if found is None:
            for alias in aliases:                   # затем — начало названия
                for i, h in enumerate(heads):
                    if h and h.startswith(norm_header(alias)):
                        found = i
                        break
                if found is not None:
                    break
        out[field] = default[field] if found is None else found
    return out


def looks_like_header(row: Optional[Sequence[Any]], spec: Dict[str, Tuple[str, ...]]) -> bool:
    """Первая строка листа — шапка? (в ней встречаются названия столбцов, а не данные)."""
    if not row:
        return False
    heads = [norm_header(v) for v in row if norm_header(v)]
    if not heads:
        return False
    hits = 0
    for aliases in spec.values():
        for alias in aliases:
            a = norm_header(alias)
            if any(h == a or h.startswith(a) for h in heads):
                hits += 1
                break
    return hits >= 2


# ---------------------------------------------------------------------------
# Разбор книги справочника
# ---------------------------------------------------------------------------

def parse(rows_fo: Optional[Iterable[Sequence[Any]]] = None,
          rows_mult: Optional[Iterable[Sequence[Any]]] = None,
          rows_tech: Optional[Iterable[Sequence[Any]]] = None,
          names: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Разбирает строки трёх вкладок справочника и строит индексы поиска.

    ``rows_*`` — списки строк, где строка это ``list``/``tuple`` со значениями
    столбцов, индекс 0 = столбец A.

    Столбцы определяются **по названию в шапке листа** (:func:`header_cols`): порядок
    столбцов в книгах справочника меняется от выпуска к выпуску (в новой версии
    «Справочник новый.xlsx» на листе «Расценки ФО» столбцы «Тип ресурса» и «Код ресурса
    Primavera» поменялись местами). Поля: ``code`` (Расценка), ``name`` (Наименование),
    ``primavera`` (Код ресурса Primavera), ``type`` (Тип ресурса), ``group`` (Для ВПР),
    ``unit`` (ЕИ Примы); для «Множителя» — ``unit_in`` (В смете), ``k`` (множитель),
    ``unit_out`` (ЕИ). Если шапки нет, используется порядок столбцов прежних выпусков.

    Строка-шапка распознаётся по названиям столбцов и в данные не попадает; пустые
    коды/единицы пропускаются.

    Возвращает словарь с ключами ``fo``, ``mult``, ``tech``, ``foByKey``,
    ``foByNoLetters``, ``foByDigits``, ``multByKey``, ``techByKey``,
    ``techByDigits``, ``names``.
    """
    fo: List[Dict[str, Any]] = []
    mult: List[Dict[str, Any]] = []
    tech: List[Dict[str, Any]] = []

    rows_fo = list(rows_fo or [])
    rows_mult = list(rows_mult or [])
    rows_tech = list(rows_tech or [])
    # шапка может быть не первой строкой листа — берём первую непустую
    head_fo = next((r for r in rows_fo if any(txt(_cell(r, i)) for i in range(len(r or [])))), None)
    head_mult = next((r for r in rows_mult if any(txt(_cell(r, i)) for i in range(len(r or [])))), None)
    head_tech = next((r for r in rows_tech if any(txt(_cell(r, i)) for i in range(len(r or [])))), None)
    cols_fo = header_cols(head_fo, COLS_FO, DEFAULT_FO)
    cols_mult = header_cols(head_mult, COLS_MULT, DEFAULT_MULT)
    cols_tech = header_cols(head_tech, COLS_TECH, DEFAULT_TECH)
    hdr_fo = looks_like_header(head_fo, COLS_FO)
    hdr_mult = looks_like_header(head_mult, COLS_MULT)
    hdr_tech = looks_like_header(head_tech, COLS_TECH)

    for i, row in enumerate(rows_fo):
        if not row:
            continue
        code = txt(_cell(row, cols_fo["code"]))
        if not code:
            continue
        # строка-шапка: та же строка, что распознана как шапка (или первая строка с «Расценка»)
        if row is head_fo and (hdr_fo or re.search(r"расценка", code, re.IGNORECASE)):
            continue
        fo.append({
            "row": i,
            "code": code,
            "key": fo_key(code),
            "noletters": fo_key_no_letters(code),
            "digits": digits(code),
            "name": txt(_cell(row, cols_fo["name"])),
            "primavera": txt(_cell(row, cols_fo["primavera"])),
            "type": txt(_cell(row, cols_fo["type"])),
            "group": txt(_cell(row, cols_fo["group"])),
            "unit": txt(_cell(row, cols_fo["unit"])),
        })

    for i, row in enumerate(rows_mult):
        if not row:
            continue
        u = txt(_cell(row, cols_mult["unit_in"]))
        if not u:
            continue
        if row is head_mult and (hdr_mult or re.search(r"в смете", u, re.IGNORECASE)):
            continue
        mult.append({
            "row": i,
            "unit_in": u,
            "key": lo(u),
            "k": num(_cell(row, cols_mult["k"])),
            "unit_out": txt(_cell(row, cols_mult["unit_out"])),
        })

    for i, row in enumerate(rows_tech):
        if not row:
            continue
        code = txt(_cell(row, cols_tech["code"]))
        if not code:
            continue
        if row is head_tech and (hdr_tech or re.search(r"расценка", code, re.IGNORECASE)):
            continue
        tech.append({
            "row": i,
            "code": code,
            "key": re.sub(r"\s+", "", txt(code).upper()),
            "digits": digits(code),
            "name": txt(_cell(row, cols_tech["name"])),
            "primavera": txt(_cell(row, cols_tech["primavera"])),
            "type": txt(_cell(row, cols_tech["type"])),
        })

    fo_by_key: Dict[str, List[Dict[str, Any]]] = {}
    fo_by_noletters: Dict[str, List[Dict[str, Any]]] = {}
    fo_by_digits: Dict[str, List[Dict[str, Any]]] = {}
    for x in fo:
        fo_by_key.setdefault(x["key"], []).append(x)
        if x["noletters"]:
            fo_by_noletters.setdefault(x["noletters"], []).append(x)
        if x["digits"]:
            fo_by_digits.setdefault(x["digits"], []).append(x)

    mult_by_key: Dict[str, Dict[str, Any]] = {}
    for x in mult:
        if x["key"] not in mult_by_key:      # первая запись побеждает
            mult_by_key[x["key"]] = x

    tech_by_key: Dict[str, Dict[str, Any]] = {}
    tech_by_digits: Dict[str, Dict[str, Any]] = {}
    for x in tech:
        if x["key"] not in tech_by_key:
            tech_by_key[x["key"]] = x
        if x["digits"] and x["digits"] not in tech_by_digits:
            tech_by_digits[x["digits"]] = x

    return {
        "fo": fo,
        "mult": mult,
        "tech": tech,
        "foByKey": fo_by_key,
        "foByNoLetters": fo_by_noletters,
        "foByDigits": fo_by_digits,      # сохранён для совместимости, в поиске не участвует
        "multByKey": mult_by_key,
        "techByKey": tech_by_key,
        "techByDigits": tech_by_digits,
        "names": names or {},
        # распознанные по шапке столбцы листов: {поле: индекс} (0 = столбец A)
        "columns": {"fo": cols_fo, "mult": cols_mult, "tech": cols_tech},
    }


def _pick_sheet(names: Sequence[str], want: str) -> Optional[str]:
    """Лист по точному имени (без учёта регистра), иначе — по началу имени."""
    low = want.lower()
    for n in names:
        if str(n).strip().lower() == low:
            return n
    for n in names:
        if str(n).strip().lower().startswith(low):
            return n
    return None


def _sheet_rows(ws: Any) -> List[List[Any]]:
    """Строки листа openpyxl: строка-заглушка + строки листа (см. докстринг модуля)."""
    rows: List[List[Any]] = [[]]
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))
    return rows


def load(path: str) -> Optional[Dict[str, Any]]:
    """Читает книгу справочника через openpyxl (``data_only=True``).

    Листы ищутся по имени: точное совпадение без учёта регистра, иначе — по
    началу имени. Строки читаются с 1-й (шапка отбрасывается теми же правилами,
    что в эталоне), после чего вызывается :func:`parse`.

    В результат дополнительно кладётся ключ ``sheets``:
    ``{"fo": имя, "mult": имя, "tech": имя, "count": N}``, где ``count`` —
    сколько из трёх листов нашлось в книге.

    Возвращает ``None``, если в книге нет ни одного листа справочника
    (поведение браузерного ``loadWorkbook``).
    """
    if openpyxl is None:  # pragma: no cover
        raise RuntimeError("для load() требуется openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        names = list(wb.sheetnames)
        fo_name = _pick_sheet(names, SHEET_FO)
        mult_name = _pick_sheet(names, SHEET_MULT)
        tech_name = _pick_sheet(names, SHEET_TECH)
        if not fo_name and not mult_name and not tech_name:
            return None
        rows_of = lambda nm: _sheet_rows(wb[nm]) if nm else []  # noqa: E731
        ref = parse(rows_of(fo_name), rows_of(mult_name), rows_of(tech_name),
                    {"fo": fo_name, "mult": mult_name, "tech": tech_name})
        ref["sheets"] = {
            "fo": fo_name, "mult": mult_name, "tech": tech_name,
            "count": sum(1 for x in (fo_name, mult_name, tech_name) if x),
        }
        return ref
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Нормализация единицы измерения по вкладке «Множитель»
# ---------------------------------------------------------------------------

def unit_multiplier(ref: Optional[Dict[str, Any]], unit_text: Any) -> Optional[Dict[str, Any]]:
    """Множитель и нормализованная ЕИ для единицы измерения сметы.

    * точное совпадение по ``multByKey[lo(text)]`` → ``{"k", "unit", "src":"точное"}``;
    * иначе ведущее число отделяется от остатка: остаток ищется в справочнике
      (``k`` = число × множитель), либо остаток должен быть служебной единицей
      (``k`` = число), → ``src="по числу"``;
    * иначе ``None``.
    """
    if not ref:
        return None
    key = lo(unit_text)
    if not key:
        return None
    rec = ref["multByKey"].get(key)
    if rec is not None:
        return {"k": rec["k"], "unit": rec["unit_out"], "src": "точное"}
    # «100 м2 покрытия» → 100 × «м2»: отделяем ведущее число и служебные слова
    m = _LEAD_NUM_RE.match(_js_string(unit_text or "").strip())
    if m:
        k = num(m.group(1))
        tail = lo(m.group(2))
        rec = ref["multByKey"].get(tail)
        if rec is not None:
            return {"k": k * rec["k"], "unit": rec["unit_out"], "src": "по числу"}
        base = _BASE_UNITS_RE.match((tail.split(" ")[0] if tail.split(" ") else "") or "")
        if base:
            return {"k": k, "unit": base.group(1).replace("КОМПЛ", "компл.").lower(),
                    "src": "по числу"}
    return None


# ---------------------------------------------------------------------------
# Поиск расценок в справочнике
# ---------------------------------------------------------------------------

_CODE_NOTE_RE = re.compile(r"[.,;:]?\s*[кkКK]\s*[=:]")
_CODE_MULT_RE = re.compile(r"\s*[×x*]\s*\d+(?:[.,]\d+)?\s*$", re.I)


def norm_code(v: Any) -> str:
    """Шифр сметной расценки без служебных пометок коэффициентов (как в браузерном модуле).

    «13-03-004-26. к=2» → «13-03-004-26», «м08-01-067-08.к=1,5» → «м08-01-067-08»,
    «13-03-004-26 (к=2)» → «13-03-004-26», «13-03-004-26 ×2» → «13-03-004-26».
    """
    t = ("" if v is None else str(v)).replace("\u00a0", " ").strip()
    if not t:
        return ""
    m = _CODE_NOTE_RE.search(t)
    if m:
        t = t[:m.start()]
    t = _CODE_MULT_RE.sub("", t)
    return re.sub(r"[\s.,;:()\[\]]+$", "", t).strip()


def find_fo(ref: Optional[Dict[str, Any]], code: Any, unit: Any = "") -> Optional[Dict[str, Any]]:
    """Расценка ФО по коду сметы: точный ключ → без ведущей «М» → с «М» → без букв.

    Если кандидатов несколько, предпочитается запись с совпадающей (loose)
    единицей измерения, иначе берётся первая.
    """
    if not ref:
        return None
    k = fo_key(norm_code(code))
    cand: List[Dict[str, Any]] = []
    if k and k in ref["foByKey"]:
        cand.extend(ref["foByKey"][k])
    if k and k[0] == "М" and k[1:] in ref["foByKey"]:
        cand.extend(ref["foByKey"][k[1:]])
    if k and ("М" + k) in ref["foByKey"]:
        cand.extend(ref["foByKey"]["М" + k])
    nl = fo_key_no_letters(code)
    if not cand and nl and nl in ref["foByNoLetters"]:
        cand.extend(ref["foByNoLetters"][nl])
    if not cand:
        return None
    if len(cand) == 1:
        return cand[0]
    u = lo(unit)
    same = next((x for x in cand if u and lo(x["unit"]) == u), None)
    return same or cand[0]


def find_tech(ref: Optional[Dict[str, Any]], token: Any) -> Optional[Dict[str, Any]]:
    """Машина/ТРЗ по коду: ``techByKey`` → ``techByDigits`` → digits без ведущих нулей."""
    if not ref:
        return None
    k = re.sub(r"\s+", "", txt(token).upper())
    if k in ref["techByKey"]:
        return ref["techByKey"][k]
    d = digits(token)
    if d and d in ref["techByDigits"]:
        return ref["techByDigits"][d]
    d2 = re.sub(r"^0+", "", d)
    if d and d2 in ref["techByDigits"]:
        return ref["techByDigits"][d2]
    return None


def tokens_of(*texts: Any) -> List[str]:
    """Коды машин/ТРЗ, найденные в текстах, в порядке появления (повторы возможны)."""
    out: List[str] = []
    for text in texts:
        s = nz(text)
        if not s:
            continue
        for m in _TOKEN_RE.finditer(s):
            out.append(m.group(1))
    return out


# ---------------------------------------------------------------------------
# Отчёт по укрупнённым физобъёмам
# ---------------------------------------------------------------------------

def k_usp_of(s: Dict[str, Any]) -> float:
    """Коэффициент условий работ из отобранных строк Source."""
    k = float(s.get("source_work_coefficient") or 0.0)
    if k:
        return k
    tc = {x["code"]: x for x in (s.get("total_check") or [])}
    base, tot = tc.get("SOURCE_TOTAL_BASE"), tc.get("SOURCE_TOTAL")
    t = s.get("totals_source") or s.get("totals") or {}
    labor = sum(float(t.get(x, 0) or 0.0) for x in ("ozp", "emm", "nr", "sp"))
    if base and tot and labor > 0:
        kk = (float(tot["source"]) - float(base["source"])) / labor + 1.0
        if 0.5 < kk < 3.0:
            return round(kk, 6)
    return 0.0


def scenario_smetas(smetas: Optional[Iterable[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    """Сметы с показателями позиций на уровне строки «ВСЕГО ПО СМЕТЕ».

    Паритет с браузерной функцией ``scenarioSmetas()``: к позициям применяется
    коэффициент «Условия производства работ» (``ozp``, ``emm``, ``zpm``, ``nr``, ``sp``,
    трудозатраты — × k; итог и ПЗ пересчитываются через оплату труда), базисные значения
    сохраняются в ``values_base``. Именно поэтому стоимость по укрупнённым физобъёмам
    совпадает со суммой итогов смет: ``Σ позиций = «ВСЕГО ПО СМЕТЕ»``.
    """
    out: List[Dict[str, Any]] = []
    for s in smetas or []:
        k = k_usp_of(s) or 1.0
        items = []
        for it in s.get("items") or []:
            v = dict(_js_or(it.get("values"), {}))
            vv = dict(v)
            for key in ("ozp", "emm", "zpm", "nr", "sp", "tz_workers", "tz_mash"):
                vv[key] = r4(float(v.get(key, 0) or 0.0) * k)
            labor = sum(float(v.get(x, 0) or 0.0) for x in ("ozp", "emm", "nr", "sp"))
            vv["total"] = r2(float(v.get("total", 0) or 0.0) + (k - 1.0) * labor)
            vv["pz"] = r2(float(v.get("pz", 0) or 0.0)
                          + (k - 1.0) * (float(v.get("ozp", 0) or 0.0)
                                         + float(v.get("emm", 0) or 0.0)))
            items.append(dict(it, values=vv, values_base=v))
        out.append(dict(s, items=items, k_usp=k,
                        total=s.get("total"), total_base=s.get("total_base")))
    return out


def _group_items(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Разметка укрупнённых групп (см. :func:`smeta_parser.group_items`).

    Импорт отложенный: модуль разбора импортирует этот модуль, поэтому прямая
    зависимость дала бы цикл. Тесты и парсеры вызывают ``reference`` и как
    пакетный модуль, и как отдельный файл — поэтому поддерживаются оба вида импорта.
    """
    try:
        from .smeta_parser import group_items
    except ImportError:                     # отдельный файл без пакета
        from smeta_parser import group_items  # type: ignore
    return group_items(items)


def _group_key(smeta: Dict[str, Any], it: Dict[str, Any], index: int) -> str:
    """Ключ укрупнённой группы строки сметы: смета + «№ строки сметы».

    У строки без ведущей работы (группа из одной строки) ключ — её собственный номер.
    Тот же ключ использует браузерный модуль (``Reference.buildFoReport``).
    """
    no = it.get("group_no")
    if no is None or no == "":
        no = it.get("num") or ""
    return str(smeta.get("number") or "") + "#" + (str(no) if no != "" else "r" + str(index))


def build_fo_report(ref: Optional[Dict[str, Any]], smetas: Optional[Iterable[Dict[str, Any]]]) -> Dict[str, Any]:
    """Отчёт по укрупнённым физобъёмам: агрегат по группам «Для ВПР» + ЕИ.

    Каждая позиция сметы даёт объём ``qty × множитель(ЕИ)`` и стоимость
    ``values.total`` (уровень строки «ВСЕГО ПО СМЕТЕ» вкладки Source). Если у
    позиции есть ``values_base``, параллельно считается стоимость базисного
    уровня (``cost_base``); без него ``cost_base == cost``. Группы — это
    ``group + "|" + unit`` записи справочника (для позиций вне справочника
    группа — :data:`NO_FO`).

    Позиции без расценки в справочнике попадают в перечень ``unmatched``,
    найденные, но с группой :data:`EXCLUDED` — в ``excluded_positions``.
    В ``unmatched`` включаются только категории :data:`UNMATCHED_CATS` («Работы»,
    «Перевозка грузов»): позиции прочих категорий учитываются в статистике
    ``unmatched_all`` / ``unmatched_skipped``, но в перечень и лист Excel не
    выгружаются.

    УКРУПНЁННЫЕ ГРУППЫ «РАБОТА + ЕЁ МАТЕРИАЛЫ».
    Материал может не иметь собственной расценки в справочнике: тогда его
    укрупнённый физобъём берётся не из справочника, а ОТ РАБОТЫ, под которой он
    стоит в смете (поле ``group_no`` — «№ строки сметы», см.
    :func:`smeta_parser.group_items`). Стоимость и трудозатраты такого материала
    суммируются в укрупнённую группу его работы, а объём группы остаётся объёмом
    самой работы (в ЕИ справочника). Такие позиции попадают в
    ``absorbed_positions`` и в перечень ``unmatched`` НЕ включаются.
    """
    by_group: Dict[str, Dict[str, Any]] = {}
    order: List[str] = []
    by_smeta: List[Dict[str, Any]] = []
    unmatched: List[Dict[str, Any]] = []
    excluded_positions: List[Dict[str, Any]] = []
    absorbed_positions: List[Dict[str, Any]] = []
    smetas = list(smetas or [])
    stats: Dict[str, Any] = {
        "positions": 0, "matched": 0, "unmatched": 0, "excluded": 0,
        "unit_known": 0, "unit_unknown": 0,
        "volume": 0.0, "cost": 0.0, "cost_base": 0.0, "pz": 0.0, "nr": 0.0, "sp": 0.0,
        "emm": 0.0, "ozp": 0.0,
        "unmatched_cost": 0.0, "excluded_cost": 0.0,
        "unmatched_all": 0, "unmatched_all_cost": 0.0,
        "unmatched_skipped": 0, "unmatched_skipped_cost": 0.0,
        "residual": 0, "residual_cost": 0.0, "residual_cost_base": 0.0,
        # материалы и оборудование, вошедшие в укрупнённый объём своей работы
        "absorbed": 0, "absorbed_cost": 0.0, "absorbed_groups": 0,
        # трудозатраты по укрупнённым физобъёмам (чел.-ч): строителей, машинистов и всего
        "tz_workers": 0.0, "tz_mash": 0.0, "tz_total": 0.0,
    }

    for s in smetas or []:
        st = {"number": s.get("number"), "name": s.get("name"),
              "positions": 0, "matched": 0, "unmatched": 0, "volume": 0.0,
              "cost": 0.0, "cost_base": 0.0, "pz": 0.0, "nr": 0.0, "sp": 0.0,
              "absorbed": 0}
        items = s.get("items") or []
        # Группы «работа + её материалы» проставляет разбор книги (smeta_parser.group_items).
        # Если позиции пришли без них (вызов отчёта на «сырых» данных), размечаем здесь же —
        # правило одно и то же, поэтому результат совпадает с разобранной книгой.
        if items and any("group_no" not in it for it in items):
            items = _group_items([dict(x) for x in items])
        # --- первый проход: укрупнённая группа (работа) → её расценка в справочнике ---
        item_um: List[Optional[Dict[str, Any]]] = []
        item_rec: List[Optional[Dict[str, Any]]] = []
        lead_rec: Dict[str, Dict[str, Any]] = {}
        for idx, it in enumerate(items):
            um = unit_multiplier(ref, it.get("unit"))
            unit = um["unit"] if um else (it.get("unit") or "")
            rec = find_fo(ref, it.get("code") or "", unit)
            item_um.append(um)
            item_rec.append(rec)
            gk = _group_key(s, it, idx)
            if rec and gk not in lead_rec:
                lead_rec[gk] = {"rec": rec, "unit": unit, "um": um}
        for idx, it in enumerate(items):
            stats["positions"] += 1
            st["positions"] += 1
            um = item_um[idx]
            unit = um["unit"] if um else (it.get("unit") or "")
            k = um["k"] if um else 1
            if um:
                stats["unit_known"] += 1
            else:
                stats["unit_unknown"] += 1
            rec = item_rec[idx]
            gk = _group_key(s, it, idx)
            # материал/оборудование без собственной расценки входит в объём своей работы
            absorbed = None
            if not rec and gk in lead_rec:
                absorbed = lead_rec[gk]
                rec = absorbed["rec"]
            qty = it.get("qty")
            qty = 0 if qty is None else qty
            # объём группы = объём работы: у вошедшего материала своё количество и ЕИ,
            # в физический объём укрупнённой группы они не суммируются
            volume = 0.0 if absorbed else r4(qty * k)
            values = _js_or(it.get("values"), {})
            values_base = _js_or(it.get("values_base"), values)
            # суммы накапливаются без промежуточного округления: округление до копеек
            # выполняется один раз при выводе — тогда «Всего с НР и СП» и стоимость по
            # укрупнённым физобъёмам совпадают до копейки (как в браузерном модуле)
            cost_raw = _field(values, "total") or 0
            cost_base_raw = _field(values_base, "total") or 0
            cost = r2(cost_raw)
            cost_base = r2(cost_base_raw)
            excluded = bool(rec and lo(rec["group"]) == lo(EXCLUDED) and not absorbed)
            group = (rec["group"] or "") if rec else NO_FO
            gkey = group + "|" + ((rec["unit"] or unit) if rec else unit)
            if gkey not in by_group:
                by_group[gkey] = {
                    "group": group, "unit": (rec["unit"] or unit) if rec else unit,
                    "positions": 0, "volume": 0.0, "cost": 0.0, "cost_base": 0.0,
                    "pz": 0.0, "nr": 0.0, "sp": 0.0, "emm": 0.0, "ozp": 0.0,
                    "in_reference": bool(rec), "excluded": excluded, "smetas": {}, "codes": {},
                    "fo_codes": {}, "tz_workers": 0.0, "tz_mash": 0.0,
                    "absorbed": 0, "absorbed_cost": 0.0, "absorbed_units": {},
                }
                order.append(gkey)
            g = by_group[gkey]
            g["positions"] += 1
            g["volume"] = r4(g["volume"] + volume)
            g["cost"] += cost_raw
            g["cost_base"] += cost_base_raw
            g["pz"] += _field(values, "pz") or 0
            g["nr"] += _field(values, "nr") or 0
            g["sp"] += _field(values, "sp") or 0
            g["emm"] += _field(values, "emm") or 0
            g["ozp"] += _field(values, "ozp") or 0
            # трудозатраты группы: ТЗ строителей и машинистов (чел.-ч), без промежуточного округления
            g["tz_workers"] += _field(values, "tz_workers") or 0
            g["tz_mash"] += _field(values, "tz_mash") or 0
            if absorbed:
                g["absorbed"] += 1
                g["absorbed_cost"] += cost_raw
                g["absorbed_units"][it.get("unit") or ""] = 1
                stats["absorbed"] += 1
                st["absorbed"] += 1
                stats["absorbed_cost"] += cost_raw
            # код укрупненного ФО — «Код ресурса Primavera» расценки справочника
            # (вкладка «RSRC» книги справочника относит коды CEWFV.* к «Укрупненным ФО»)
            if rec and rec.get("primavera"):
                g["fo_codes"][rec["primavera"]] = 1
            g["smetas"][_js_string(s.get("number")) if s.get("number") is not None else "undefined"] = 1
            if it.get("code"):
                g["codes"][it["code"]] = 1
            if rec and not absorbed:
                stats["matched"] += 1
                st["matched"] += 1
            if excluded:
                stats["excluded"] += 1
            stats["volume"] = r4(stats["volume"] + volume)
            stats["cost"] += cost_raw
            stats["cost_base"] += cost_base_raw
            stats["pz"] += _field(values, "pz") or 0
            stats["nr"] += _field(values, "nr") or 0
            stats["sp"] += _field(values, "sp") or 0
            stats["emm"] += _field(values, "emm") or 0
            stats["ozp"] += _field(values, "ozp") or 0
            stats["tz_workers"] += _field(values, "tz_workers") or 0
            stats["tz_mash"] += _field(values, "tz_mash") or 0
            stats["tz_total"] += (_field(values, "tz_workers") or 0) + (_field(values, "tz_mash") or 0)
            st["volume"] = r4(st["volume"] + volume)
            st["cost"] += cost_raw
            st["cost_base"] += cost_base_raw
            st["pz"] += _field(values, "pz") or 0
            st["nr"] += _field(values, "nr") or 0
            st["sp"] += _field(values, "sp") or 0
            row = {
                "smeta": s.get("number"), "num": it.get("num") or "",
                "code": it.get("code") or "(без шифра)", "name": it.get("name") or "",
                "basis": it.get("basis") or "", "unit": it.get("unit") or "",
                "unit_norm": unit, "k": k, "qty": qty, "volume": volume,
                "cost": cost, "cost_base": cost_base, "pz": r2(_field(values, "pz") or 0),
                "nr": r2(_field(values, "nr") or 0), "sp": r2(_field(values, "sp") or 0),
                "category": it.get("cat") or "",
                # укрупнённая группа строки сметы: работа, под которой стоит материал
                "group_no": it.get("group_no") or "",
                "group_code": it.get("group_code") or "",
                "group_name": it.get("group_name") or "",
                "absorbed": bool(absorbed),
                "absorbed_by": ((it.get("group_no") or "") +
                                ((" · " + it["group_code"]) if it.get("group_code") else "")) if absorbed else "",
                "group": (rec["group"] or "") if rec else "",
                "position_kind": ("материал в объёме работы" if absorbed else
                                  ("не идет в физику" if rec else "нет расценки в справочнике ФО")),
            }
            if absorbed:
                # материал/оборудование не попал в перечень «ФО не сопоставлено»: его место —
                # в укрупнённом объёме работы, а не отдельной строкой отчёта
                stats["unmatched_all"] += 1
                stats["unmatched_all_cost"] += cost_raw
                absorbed_positions.append(row)
            elif not rec:
                stats["unmatched_all"] += 1
                stats["unmatched_all_cost"] += cost_raw
                if is_unmatched_cat(it.get("cat")):
                    stats["unmatched"] += 1
                    st["unmatched"] += 1
                    stats["unmatched_cost"] += cost_raw
                    unmatched.append(row)
                else:
                    stats["unmatched_skipped"] += 1
                    stats["unmatched_skipped_cost"] += cost_raw
            elif excluded:
                stats["excluded_cost"] += cost_raw
                excluded_positions.append(row)
        by_smeta.append(st)

    # Стоимость по укрупнённым физобъёмам должна соответствовать «ВСЕГО ПО СМЕТЕ».
    # В книгах-корректировках сумма позиций не равна итогу сметы (внутреннее расхождение
    # исходного файла) — разницу показываем отдельной строкой «не распределено», чтобы
    # итог отчёта совпадал с «ВСЕГО ПО СМЕТЕ» и расхождение было видно явно.
    for s in smetas:
        model_total = num(s.get("total"))
        model_base = num(s.get("total_base"))
        if s.get("total") is None and s.get("total_base") is None:
            continue
        st = None
        for x in by_smeta:
            if x["number"] == s.get("number"):
                st = x
                break
        if st is None:
            continue
        res_raw = model_total - st["cost"]
        res_base_raw = model_base - st["cost_base"]
        # Расхождение меньше половины рубля — округление по позициям: отдельной строки
        # «не распределено» не заводим, разницу распределяет корректировка округления ниже.
        if abs(res_raw) > 0.5 or abs(res_base_raw) > 0.5:
            gkey = RESIDUAL_GROUP + "|—"
            if gkey not in by_group:
                by_group[gkey] = {
                    "group": RESIDUAL_GROUP, "unit": "", "positions": 0, "volume": 0.0,
                    "cost": 0.0, "cost_base": 0.0, "pz": 0.0, "nr": 0.0, "sp": 0.0,
                    "emm": 0.0, "ozp": 0.0,
                    "in_reference": False, "excluded": False, "residual": True,
                    "smetas": {}, "codes": {}, "fo_codes": {},
                    "tz_workers": 0.0, "tz_mash": 0.0,
                }
                order.append(gkey)
            g = by_group[gkey]
            g["cost"] += res_raw
            g["cost_base"] += res_base_raw
            g["smetas"][_js_string(s.get("number")) if s.get("number") is not None else "undefined"] = 1
            stats["residual"] = (stats.get("residual") or 0) + 1
            stats["residual_cost"] = (stats.get("residual_cost") or 0.0) + res_raw
            stats["residual_cost_base"] = (stats.get("residual_cost_base") or 0.0) + res_base_raw
        # итог по смете в отчёте = итогу сметы («ВСЕГО ПО СМЕТЕ»), в том числе когда
        # расхождение объясняется округлением по позициям
        stats["cost"] += res_raw
        stats["cost_base"] += res_base_raw
        st["cost"] += res_raw
        st["cost_base"] += res_base_raw
        st["residual"] = (st.get("residual") or 0.0) + res_raw

    rows: List[Dict[str, Any]] = []
    for gk in order:
        g = by_group[gk]
        # трудозатраты считаем по округлённым значениям: тогда «всего ч/ч» = «строители» +
        # «машинисты» в каждой строке, а колонки сходятся с итоговой строкой листа точно
        tzw = r4(g["tz_workers"])
        tzm = r4(g["tz_mash"])
        tzt = r4(tzw + tzm)
        fo_codes = sorted(g.get("fo_codes", {}).keys())
        rows.append({
            "group": g["group"], "unit": g["unit"], "positions": g["positions"],
            "volume": r4(g["volume"]), "cost": r2(g["cost"]), "cost_base": r2(g["cost_base"]),
            "unit_cost": (r2(g["cost"] / g["volume"]) if g["volume"] else None),
            "unit_cost_base": (r2(g["cost_base"] / g["volume"]) if g["volume"] else None),
            # «Код укрупненного ФО» — код ресурса Primavera (вкладка «Расценки ФО»);
            # если у одной группы справочника их несколько — перечисляются через запятую
            "fo_code": ", ".join(fo_codes),
            # трудозатраты: всего = ТЗ строителей + ТЗ машинистов, на единицу = всего / объём
            "tz_workers": tzw, "tz_mash": tzm, "tz_total": tzt,
            "tz_unit": (r4(tzt / g["volume"]) if g["volume"] else None),
            "pz": r2(g["pz"]), "nr": r2(g["nr"]), "sp": r2(g["sp"]),
            "emm": r2(g["emm"]), "ozp": r2(g["ozp"]),
            "in_reference": g["in_reference"], "excluded": g["excluded"],
            "residual": bool(g.get("residual")),
            # материалы и оборудование, вошедшие в укрупнённый объём этой группы
            "absorbed": g.get("absorbed") or 0,
            "absorbed_cost": r2(g.get("absorbed_cost") or 0),
            "absorbed_units": sorted(u for u in (g.get("absorbed_units") or {}) if u),
            "smetas": sorted(g["smetas"].keys()), "codes": list(g["codes"].keys())[:20],
        })
    rows.sort(key=lambda r: (-r["cost"], -r["volume"]))
    # Перечни позиций: стоимость убыв., затем код по возрастанию (обычное сравнение
    # строк, без локали — по ТЗ). Браузер во вторичной сортировке использует
    # String(code).localeCompare(code,"ru"); на реальном справочнике и сметах оба
    # правила дают одинаковый порядок (групп равной стоимости с разными кодами
    # нет) — расхождение возможно только при равной стоимости у кодов из разных
    # алфавитов/регистров.
    unmatched.sort(key=lambda r: (-r["cost"], str(r["code"])))
    excluded_positions.sort(key=lambda r: (-r["cost"], str(r["code"])))
    # округление итоговых сумм — один раз, после накопления
    for key in ("cost", "cost_base", "pz", "nr", "sp", "emm", "ozp", "unmatched_cost",
                "excluded_cost", "unmatched_all_cost", "unmatched_skipped_cost",
                "residual_cost", "residual_cost_base"):
        stats[key] = r2(stats[key])
    for st in by_smeta:
        st["cost"] = r2(st["cost"])
        st["cost_base"] = r2(st["cost_base"])
        st["pz"] = r2(st["pz"])
        st["nr"] = r2(st["nr"])
        st["sp"] = r2(st["sp"])
    # Корректировка округления: сумма округлённых значений групп должна в точности
    # равняться итогу (иначе в Excel столбец не сходится с итоговой строкой на копейки).
    # Разница ≤ половины последнего знака × число групп относится на самую крупную группу.
    if rows:
        def _fix(key: str, total: float, rnd: Any, extra: Optional[Any] = None) -> None:
            total_sum = rnd(sum(r[key] if math.isfinite(r[key]) else 0 for r in rows))
            d = rnd(total - total_sum)
            if d == 0 or not math.isfinite(d):
                return
            bi = 0
            for i, r in enumerate(rows):
                if abs(r[key]) > abs(rows[bi][key]):
                    bi = i
            rows[bi][key] = rnd(rows[bi][key] + d)
            if extra is not None:
                extra(rows[bi])

        _fix("cost", stats["cost"], r2,
             lambda r: r.update(unit_cost=r2(r["cost"] / r["volume"])) if r["volume"] else None)
        _fix("cost_base", stats["cost_base"], r2,
             lambda r: r.update(unit_cost_base=r2(r["cost_base"] / r["volume"])) if r["volume"] else None)
    # Итоги по трудозатратам — суммы округлённых значений строк: колонки листа сходятся
    # с итоговой строкой точно, и «всего ч/ч» = «строители» + «машинисты» в каждой строке.
    stats["tz_workers"] = r4(sum(r["tz_workers"] for r in rows))
    stats["tz_mash"] = r4(sum(r["tz_mash"] for r in rows))
    stats["tz_total"] = r4(sum(r["tz_total"] for r in rows))
    stats["rows"] = len(rows)
    # сколько укрупнённых групп справочника получили материалы своих работ
    stats["absorbed_groups"] = sum(1 for r in rows if (r.get("absorbed") or 0) > 0)
    return {"rows": rows, "by_smeta": by_smeta, "stats": stats, "unmatched": unmatched,
            "absorbed_positions": absorbed_positions,
            "excluded_positions": excluded_positions,
            "no_fo_group": NO_FO, "excluded_group": EXCLUDED,
            "unmatched_cats": list(UNMATCHED_CATS),
            "unmatched_cat_labels": list(UNMATCHED_CAT_LABELS)}


# ---------------------------------------------------------------------------
# Отчёт по укрупнённому справочнику машин и механизмов
# ---------------------------------------------------------------------------

def build_tech_report(ref: Optional[Dict[str, Any]],
                      smetas: Optional[Iterable[Dict[str, Any]]]) -> Dict[str, Any]:
    """Отчёт по технике: упоминания кодов машин/ТРЗ в позициях и ресурсах смет.

    Один и тот же код в одной строке (позиции или ресурсе) учитывается один раз;
    позиции дают ``qty`` позиции и стоимость ``values.emm``, ресурсы — ``qty``
    и ``cost`` ресурса.
    """
    by_key: Dict[str, Dict[str, Any]] = {}
    order: List[str] = []
    by_prim: Dict[str, Dict[str, Any]] = {}
    smetas = list(smetas or [])
    stats: Dict[str, Any] = {
        "rows_scanned": 0, "mentions": 0, "machines": 0, "primavera": 0,
        "cost": 0.0, "qty": 0.0, "smetas": 0,
    }

    def add(rec: Dict[str, Any], where: Dict[str, Any], qty: Any, cost: Any, kind: str) -> None:
        key = rec["key"] + "|" + (rec["primavera"] or "")
        if key not in by_key:
            by_key[key] = {
                "code": rec["code"], "name": rec["name"], "primavera": rec["primavera"],
                "type": rec["type"], "mentions": 0, "qty": 0.0, "cost": 0.0,
                "items": 0, "resources": 0, "smetas": {}, "samples": [],
            }
            order.append(key)
        m = by_key[key]
        m["mentions"] += 1
        m["qty"] = r4(m["qty"] + (qty or 0))
        m["cost"] = r2(m["cost"] + (cost or 0))
        if kind == "resource":
            m["resources"] += 1
        else:
            m["items"] += 1
        m["smetas"][_js_string(where["smeta"]) if where["smeta"] is not None else "undefined"] = 1
        if len(m["samples"]) < 5 and where["text"]:
            m["samples"].append(where["text"])

    for s in smetas or []:
        for it in s.get("items") or []:
            stats["rows_scanned"] += 1
            found: Dict[str, Any] = {}
            for t in tokens_of(it.get("code"), it.get("basis"), it.get("name")):
                rec = find_tech(ref, t)
                if rec and rec["key"] not in found:
                    found[rec["key"]] = rec
                    where = {
                        "smeta": s.get("number"),
                        "text": (("№" + _js_string(it["num"]) + " ") if it.get("num") else "")
                                + (it.get("name") or "")[:60],
                    }
                    add(rec, where, it.get("qty") or 0, _field(_js_or(it.get("values"), {}), "emm") or 0, "item")
        for rc in s.get("resources") or []:
            stats["rows_scanned"] += 1
            found = {}
            for t in tokens_of(rc.get("code"), rc.get("name")):
                rec = find_tech(ref, t)
                if rec and rec["key"] not in found:
                    found[rec["key"]] = rec
                    where = {"smeta": s.get("number"), "text": (rc.get("name") or "")[:60]}
                    add(rec, where, rc.get("qty") or 0, rc.get("cost") or 0, "resource")

    rows: List[Dict[str, Any]] = []
    for k in order:
        m = by_key[k]
        rows.append({
            "code": m["code"], "name": m["name"], "primavera": m["primavera"], "type": m["type"],
            "mentions": m["mentions"], "qty": r4(m["qty"]), "cost": r2(m["cost"]),
            "items": m["items"], "resources": m["resources"],
            "smetas": sorted(m["smetas"].keys()), "samples": m["samples"],
        })
    rows.sort(key=lambda r: (-r["cost"], -r["mentions"], str(r["code"])))

    for r in rows:
        p = r["primavera"] or "(без кода Primavera)"
        if p not in by_prim:
            by_prim[p] = {"primavera": p, "machines": 0, "mentions": 0, "qty": 0.0,
                          "cost": 0.0, "names": []}
        b = by_prim[p]
        b["machines"] += 1
        b["mentions"] += r["mentions"]
        b["qty"] = r4(b["qty"] + r["qty"])
        b["cost"] = r2(b["cost"] + r["cost"])
        if len(b["names"]) < 3 and r["name"]:
            b["names"].append(r["name"])

    prim = sorted(by_prim.values(), key=lambda b: (-b["cost"], -b["mentions"]))
    stats["mentions"] = sum(r["mentions"] for r in rows)
    stats["machines"] = len(rows)
    stats["primavera"] = len(prim)
    stats["cost"] = r2(sum(r["cost"] for r in rows))
    stats["qty"] = r4(sum(r["qty"] for r in rows))
    stats["smetas"] = len(smetas)
    return {"rows": rows, "by_primavera": prim, "stats": stats}


# ---------------------------------------------------------------------------
# Самопроверка
# ---------------------------------------------------------------------------

def _print_fo_row(r: Dict[str, Any]) -> None:
    print("    %-42s %-16s поз. %-4d объём %-14s стоимость %-14s руб/ед. %-12s %s%s"
          % (r["group"][:42], r["unit"][:16], r["positions"], r["volume"], r["cost"],
             r["unit_cost"], "в справочнике" if r["in_reference"] else "ВНЕ справочника",
             ", НЕ в физике" if r["excluded"] else ""))


def _print_fo_position(r: Dict[str, Any]) -> None:
    print("    №%-10s %-30s %-16s ЕИ %-10s→%-10s k=%-7s qty=%-10s объём %-12s cost %-12s базис %-12s [%s]"
          % (str(r["num"])[:10], r["name"][:30], str(r["code"])[:16], r["unit"][:10],
             str(r["unit_norm"])[:10], r["k"], r["qty"], r["volume"], r["cost"], r["cost_base"],
             r["position_kind"]))


def _print_tech_row(r: Dict[str, Any]) -> None:
    print("    %-14s %-34s %-18s упом. %-3d qty %-10s cost %-12s (поз. %d, рес. %d)"
          % (r["code"][:14], r["name"][:34], (r["primavera"] or "—")[:18], r["mentions"],
             r["qty"], r["cost"], r["items"], r["resources"]))


def self_check(xlsx_path: str, smetas_path: str) -> int:
    """Печатает контрольные сведения о справочнике и отчётах. Возвращает код выхода."""
    print("=" * 100)
    print("САМОПРОВЕРКА reference.py — справочник:", xlsx_path)
    print("=" * 100)
    ref = load(xlsx_path)
    if ref is None:
        print("ОШИБКА: в книге нет ни одного листа справочника.")
        return 2
    sh = ref["sheets"]
    print("Листы: ФО=%r, Множитель=%r, Техника=%r (найдено %d)"
          % (sh["fo"], sh["mult"], sh["tech"], sh["count"]))
    print("Записей: ФО=%d, множителей=%d, техники=%d"
          % (len(ref["fo"]), len(ref["mult"]), len(ref["tech"])))
    print("Индексы: foByKey=%d, foByNoLetters=%d, foByDigits=%d, multByKey=%d, "
          "techByKey=%d, techByDigits=%d"
          % (len(ref["foByKey"]), len(ref["foByNoLetters"]), len(ref["foByDigits"]),
             len(ref["multByKey"]), len(ref["techByKey"]), len(ref["techByDigits"])))

    print("\nПримеры unit_multiplier:")
    for u in ("100 м2", "1000 м", "ШТ", "нечто"):
        print("    %-10s -> %s" % (u, unit_multiplier(ref, u)))

    print("\nПримеры find_fo (защита от ложных совпадений по цифрам):")
    for code, unit in (("ФЕР-01-01-001-01", "100 м3"), ("м08-01-067-07", "100 м2"),
                       ("м08-01-067-08", "100 м2"), ("01-01-001-01", "100 м2")):
        rec = find_fo(ref, code, unit)
        print("    find_fo(%-20r, %-9r) -> %s"
              % (code, unit, "None" if rec is None else "%s / key=%s / ЕИ=%s"
                 % (rec["code"], rec["key"], rec["unit"])))
    print("    примечание: кода «м08-01-067-07» в справочнике нет (есть «м08-01-067-08»),")
    print("    поэтому None для него — верный результат; «ФЕР-01-01-001-01» больше не")
    print("    совпадает по цифрам с «м01-01-001-01» (индекс без букв вместо индекса по цифрам).")
    print("\nНормализация шифров («13-03-004-26. к=2» → «13-03-004-26»):")
    for raw in ("13-03-004-26. к=2", "м08-01-067-08.к=1,5", "13-03-004-26 (к=2)",
                "13-03-004-26 ×2", "13-03-004-26", "(МО8)08-01-007-01"):
        rec = find_fo(ref, raw, "м2")
        print("    %-24r -> %-16r %s" % (raw, norm_code(raw),
              "нет в справочнике" if rec is None else "key=%s" % rec["key"]))

    try:
        with open(smetas_path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except OSError as exc:
        print("\nСметы не прочитаны (%s): %s" % (smetas_path, exc))
        return 0
    smetas = data.get("smetas") if isinstance(data, dict) else data
    smetas = smetas or []
    print("\nСметы: %s (смет: %d)" % (smetas_path, len(smetas)))

    fo = build_fo_report(ref, smetas)
    st = fo["stats"]
    print("\n--- Отчёт по укрупнённым физобъёмам ---")
    print("Статистика: positions=%d matched=%d unmatched=%d excluded=%d unit_known=%d "
          "unit_unknown=%d rows=%d" % (st["positions"], st["matched"], st["unmatched"],
                                       st["excluded"], st["unit_known"], st["unit_unknown"],
                                       st["rows"]))
    print("Суммы: volume=%s cost=%s cost_base=%s pz=%s emm=%s ozp=%s"
          % (st["volume"], st["cost"], st["cost_base"], st["pz"], st["emm"], st["ozp"]))
    print("Перечни позиций: unmatched=%d на %s руб. (базис %s), excluded_positions=%d на %s руб."
          % (len(fo["unmatched"]), st["unmatched_cost"],
             r2(sum(x["cost_base"] for x in fo["unmatched"])),
             len(fo["excluded_positions"]), st["excluded_cost"]))
    print("Вне справочника всего: %d на %s руб.; в перечень (%s) включено %d на %s руб.; "
          "прочие категории: %d на %s руб."
          % (st["unmatched_all"], st["unmatched_all_cost"],
             " / ".join(UNMATCHED_CAT_LABELS), st["unmatched"], st["unmatched_cost"],
             st["unmatched_skipped"], st["unmatched_skipped_cost"]))
    print("Проверка unit_known == positions: %s (%d == %d)"
          % ("OK" if st["unit_known"] == st["positions"] else "РАСХОЖДЕНИЕ",
             st["unit_known"], st["positions"]))
    print("По сметам:")
    for b in fo["by_smeta"]:
        print("    №%-32s поз. %-4d в справочнике %-4d вне справочника %-4d объём %-14s "
              "стоимость %-14s базис %s"
              % (str(b["number"])[:32], b["positions"], b["matched"], b["unmatched"],
                 b["volume"], b["cost"], b["cost_base"]))
    print("Первые 5 строк отчёта ФО (сортировка: стоимость убыв., затем объём убыв.):")
    for r in fo["rows"][:5]:
        _print_fo_row(r)
    print("Первые 3 позиции «нет расценки в справочнике ФО» (unmatched):")
    for r in fo["unmatched"][:3]:
        _print_fo_position(r)
    print("Первые 3 позиции «не идет в физику» (excluded_positions):")
    for r in fo["excluded_positions"][:3]:
        _print_fo_position(r)

    tech = build_tech_report(ref, smetas)
    ts = tech["stats"]
    print("\n--- Отчёт по технике ---")
    print("Статистика: rows_scanned=%d mentions=%d machines=%d primavera=%d cost=%s qty=%s smetas=%d"
          % (ts["rows_scanned"], ts["mentions"], ts["machines"], ts["primavera"],
             ts["cost"], ts["qty"], ts["smetas"]))
    print("Первые 5 строк отчёта по технике (сортировка: cost убыв., mentions убыв., code возр.):")
    for r in tech["rows"][:5]:
        _print_tech_row(r)
    if not tech["rows"]:
        print("    (совпадений нет — кодов машин/ТРЗ в этих сметах не найдено)")
    return 0


def _default_smetas_path() -> str:
    """Путь к data/smetas.json рядом с репозиторием."""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(root, "data", "smetas.json")


if __name__ == "__main__":
    # самопроверка: python parser/reference.py /tmp/spr.xlsx <smetas.json>
    _xlsx = sys.argv[1] if len(sys.argv) > 1 else "/tmp/spr.xlsx"
    _smetas = sys.argv[2] if len(sys.argv) > 2 else _default_smetas_path()
    sys.exit(self_check(_xlsx, _smetas))
