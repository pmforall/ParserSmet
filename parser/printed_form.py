#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
printed_form.py — чтение печатной формы локальной сметы (листы, содержащие «11 граф»).

Назначение
----------
Печатная форма — официальный документ сметы. По требованию технического задания
проверка корректности собранных данных выполняется именно по ней: суммарные данные
разбора вкладки «Source» должны совпадать с печатной формой.

Что читается с листа «…11 граф…»
--------------------------------
1. Шапка (строки 1–30):
   * шифр документа, «ЛОКАЛЬНАЯ СМЕТА № …», наименование, основание, объект, НБ;
   * «Сметная стоимость», в том числе строительных / монтажных работ, оборудования,
     прочих; возврат; «Нормативная трудоемкость»; «Сметная заработная плата» (тыс. → руб.).
2. Итоговый блок (после позиций). Печатная форма содержит несколько блоков итогов —
   по подразделам, разделам и смете целиком, поэтому значение показателя выбирается
   как наибольшее из одноимённых строк (итог сметы всегда не меньше итога её части).
   Различаются уровни цен:
     «Всего по смете»        — базисный уровень (НР и СП без коэффициента условий работ);
     «ВСЕГО ПО СМЕТЕ»        — с коэффициентом «Условия производства работ» (напр. 1,15);
     «ВСЕГО, СТОИМОСТЬ … РАБОТ» и «… с учетом "Условия производства работ"».
3. Итоги разделов и подразделов: «Итого по разделу (без НР и СП) <наименование>».
4. Позиции формы: № поз., шифр, наименование, ед. изм., количество, общая стоимость —
   для контроля количества и суммы позиций.

Модуль используется и парсером (smeta_parser.py), и независимой приёмкой
(verify_parser.py): обе программы читают печатную форму одним и тем же способом,
но сравнивают её с данными, полученными разными путями.
"""

from __future__ import annotations

import math
import re
from typing import Any, Dict, List, Optional, Tuple

MARK = "11 граф"          # признак листа печатной формы «ЛС 11 граф (АО Атомпроект)»
MARK17 = "17 граф"        # признак листа «Смета 17 граф с оборуд.»
MARKS = (MARK, MARK17)    # поддерживаемые печатные формы
MARK_LABEL = "11/17 граф"
HEAD_LAST_ROW = 30        # строки 1..30 — шапка формы
TOL = 0.02

# показатели, которые сверяются с разбором «Source»: ключ → подпись
COMPARE_KEYS: List[Tuple[str, str]] = [
    # «Прямые затраты» = стоимость работ без НР и СП: ВСЕГО − НР − СП = МатОб + ОЗП + ЭММ
    ("pz", "Прямые затраты (стоимость без НР и СП)"),
    ("total", "Всего с НР и СП"),
    ("ozp", "Основная ЗП рабочих"),
    ("emm", "Эксплуатация машин"),
    ("zpm", "ЗП машинистов"),
    ("material", "Стоимость материалов (всего)"),
    ("transport", "Перевозка грузов"),
    ("nr", "Накладные расходы"),
    ("sp", "Сметная прибыль"),
    ("equipment", "Стоимость оборудования (всего)"),
    # разбивка стоимости СМР: печатается формой «17 граф» (и строкой «Всего, стоимость … работ»
    # формы «11 граф»); если строки нет — показатель считается отсутствующим в форме
    ("smr_build", "Строительные работы с НР и СП"),
    ("smr_install", "Монтажные работы с НР и СП"),
    ("other", "Прочие работы с НР и СП"),
    ("tz_mash", "Трудозатраты машинистов"),
    ("tz_workers", "Трудозатраты строителей"),
]


def _next_row_labor(ws, r: int) -> Tuple[float, float]:
    """ЗПМ и трудозатраты машинистов со строки-продолжения (без своей подписи в столбце A).

    В форме ЗПМ и ТЗМ печатаются строкой ниже итога, но эта строка есть не всегда: если
    следующая строка имеет собственную подпись («Прямые затраты», «Итого по разделу…»,
    «Подраздел …»), брать из неё значения нельзя — иначе в итог попадут чужие суммы.
    """
    if _s(ws.cell(r, 1).value):
        return 0.0, 0.0
    return _f(ws.cell(r, 9).value), _f(ws.cell(r, 11).value)


def header_last_row(ws, max_row: Optional[int] = None) -> int:
    """Последняя строка шапки формы — строка с нумерацией столбцов 1…11.

    У разных выпусков АтомСметы шапка занимает 26 или 29 строк (в книгах ЛСР-02-… таблица
    начинается раньше), поэтому границу определяем по строке с номерами столбцов, а не
    по фиксированному HEAD_LAST_ROW. Если такой строки нет — прежнее значение.
    """
    n = max_row if max_row is not None else ws.max_row
    for r in range(1, min(n, 90) + 1):
        ok = True
        for c in range(1, 12):
            v = _f(ws.cell(r, c).value)
            if v != c:
                ok = False
                break
        if ok:
            return r
    return HEAD_LAST_ROW


def paren_value(v: Any) -> Optional[float]:
    """Значение из скобок — базисный уровень цены в части выпусков формы («(6793,19)»)."""
    if isinstance(v, str):
        t = v.strip()
        if t.startswith("(") and t.endswith(")"):
            return _f(t[1:-1])
    return None


def _s(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _f(v: Any) -> float:
    if isinstance(v, bool) or v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.replace("\xa0", "").replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0
    return 0.0


def _cell(row, idx: int):
    """Значение ячейки по номеру столбца (с нуля): строка может быть и списком, и словарём."""
    if row is None:
        return None
    try:
        return row[idx]
    except (TypeError, KeyError, IndexError):
        return None


def looks_like_printed(rows) -> bool:
    """Лист является локальной сметой по содержанию.

    Часть книг АтомСмета называет лист по шифру сметы («ЛСР-02-14.3-101С»), без слова
    «граф». Такой лист опознаётся по заголовку «ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ»; ниже должна
    быть либо строка заголовков граф, либо строки позиций (номер во втором столбце,
    наименование в четвёртом).
    """
    title = False
    for row in rows[:400]:
        for idx in range(8):
            text = str(_cell(row, idx) if _cell(row, idx) is not None else "").lower()
            if not text:
                continue
            if ("локальный сметный расчет" in text or "локальный сметный расчёт" in text
                    or "локальная смета" in text):
                title = True
                break
        if title:
            break
    if not title:
        return False
    return _has_column_header(rows) or _count_positions(rows) > 3


HEADER_WORDS = ("обоснование", "единиц", "наименование", "кол-во", "количество")


def _has_column_header(rows) -> bool:
    """Строка заголовков граф: «№ п/п … Обоснование … Единица измерения …»."""
    for row in rows[:400]:
        hits = 0
        for idx in range(10):
            text = str(_cell(row, idx) if _cell(row, idx) is not None else "").lower()
            if text and any(word in text for word in HEADER_WORDS):
                hits += 1
        if hits >= 2:
            return True
    return False


def _count_positions(rows) -> int:
    """Число строк-позиций (проверяем не больше пяти)."""
    count = 0
    for row in rows:
        if count > 4:
            break
        num = str(_cell(row, 1) if _cell(row, 1) is not None else "").strip()
        name = str(_cell(row, 3) if _cell(row, 3) is not None else "").strip()
        if num and name and num.replace(",", ".").replace(".", "", 1).isdigit():
            count += 1
    return count


SERVICE_SHEETS = ("source", "sourceobsm", "smtres", "etalonres", "srcka",
                  "srcpoprs", "calctmp", "unionsheet", "лист")


def sheets_with_mark(wb, rows_of=None) -> List[str]:
    """Листы печатной формы: «граф» в названии, иначе — распознанные по содержанию.

    rows_of(name) отдаёт строки листа; без него работает только признак названия.
    """
    names = [name for name in wb.sheetnames if "граф" in name.lower()]
    if rows_of is None:
        return names
    for name in wb.sheetnames:
        if name in names or name.lower().startswith(SERVICE_SHEETS):
            continue
        rows = rows_of(name)
        if rows and looks_like_printed(rows):
            names.append(name)
    return names


LABOR_KEYS = ("tz_workers", "tz_mash")
# Разбивка стоимости СМР по видам работ. Программа определяет вид работ позиции по «виду
# затрат» и группе ресурсов вкладки «Source» и сходится с показателями узлов 214/215/216/217
# до копейки; печатная форма делит те же суммы по собственному правилу (по видам работ сметы).
# Расхождение здесь — разница методик, а не ошибка данных: строка сверки получает статус
# «kind_diff» и выводится предупреждением (PRINTED_KIND_DIFF), решение принимает пользователь.
KIND_KEYS = ("smr_build", "smr_install", "other")
LABOR_DECIMALS = 4
ROUND_HALF = 0.005       # половина младшего разряда «грубого» (2-знакового) представления


def decimals(v: float) -> int:
    """Сколько знаков после запятой реально несёт значение (грубость представления).

    Нужна, чтобы отличить точное значение от напечатанного округлённым: в печатной форме
    один и тот же показатель встречается дважды — в итоговой строке с полной точностью
    (761,9711 / 1793,688813) и в справочной строке «Трудозатраты …» с двумя знаками
    (761,97 / 1793,69).
    """
    for n in range(0, LABOR_DECIMALS + 1):
        scaled = v * (10 ** n)
        if abs(scaled - round(scaled)) <= 1e-6:
            return n
    return 9             # точность выше учётной (например, 70,242565) — считаем самой полной
_EPS_JS = 2.220446049250313e-16          # Number.EPSILON в JavaScript


def js_round(x: float) -> float:
    """Math.round: половина вверх (как в браузерном модуле)."""
    return math.floor(x + 0.5)


def round2(v: float) -> float:
    return js_round((v + _EPS_JS) * 100) / 100


def round4(v: float) -> float:
    return js_round((v + _EPS_JS) * 10000) / 10000


def _is_position_num(a: str) -> bool:
    """Номер позиции печатной формы: «13», «13,1» (поправка), но не «02-24.0.1И3»."""
    t = a.strip()
    if t.isdigit():
        return True
    head, sep, tail = t.replace(".", ",").partition(",")
    return bool(sep) and head.isdigit() and tail.isdigit() and len(tail) <= 2


# ---------------------------------------------------------------------------
# Печатная форма «Смета 17 граф с оборуд.» (АтомСмета, шаблон с оборудованием)
# ---------------------------------------------------------------------------
# Раскладка столбцов этой формы (строка нумерации 1…17):
#   1 № | 2 обоснование | 3 наименование | 4 ед. изм. | 5 кол-во |
#   6 стоимость единицы, всего | 7 ОЗП | 8 ЭММ | 9 ЗПМ (на единицу) |
#   10 оборудование (общая стоимость) | 11 всего (общая стоимость) |
#   12 ОЗП | 13 ЭММ | 14 ЗПМ (общая стоимость) | 15 Т/з на ед. | 16 Т/з всего | 17 масса
FO_COL = {"num": 1, "basis": 2, "name": 3, "unit": 4, "qty": 5, "unit_cost": 6, "unit_ozp": 7,
          "unit_emm": 8, "unit_zpm": 9, "equipment": 10, "total": 11, "ozp": 12, "emm": 13,
          "zpm": 14, "tz_unit": 15, "tz_workers": 16, "mass": 17}


def header_last_row_17(ws, max_row: Optional[int] = None) -> int:
    """Строка нумерации столбцов 1…17 — граница шапки формы «17 граф»."""
    n = max_row if max_row is not None else ws.max_row
    for r in range(1, min(n, 90) + 1):
        if all(_f(ws.cell(r, c).value) == c for c in range(1, 18)):
            return r
    return HEAD_LAST_ROW


def _is_position_num17(a: str) -> bool:
    t = (a or "").strip()
    if t.isdigit():
        return True
    head, sep, tail = t.replace(".", ",").partition(",")
    return bool(sep) and head.isdigit() and tail.isdigit() and len(tail) <= 2


def parse17(ws) -> Dict[str, Any]:
    """Разобрать лист «Смета 17 граф с оборуд.».

    Итоги берутся из итогового блока в конце формы («ИТОГО по смете, в т.ч.»,
    «Строительные работы с НР и СП», «Монтажные работы с НР и СП», «Оборудование»,
    «Прочие с НР и СП», «Накладные расходы», «Сметная прибыль», «Прямые затраты»,
    «Стоимость материальных ресурсов (без оборудования)», «ЭММ (без ЗПМ)»,
    «Сметная заработная плата (ФОТ)», «Нормативная трудоемкость»), а шапка —
    из верхнего блока («Сметная стоимость», в том числе строительных / монтажных работ,
    оборудования, прочих; средства на оплату труда; сметная трудоемкость;
    трудозатраты механизаторов). Верхний блок печатается в тыс. руб.
    """
    out: Dict[str, Any] = {
        "sheet": ws.title, "doc_code": "", "number": "", "name": "", "osnovanie": "",
        "object": "", "norm_base": "",
        "head": {}, "totals": {}, "totals_with_k": {}, "rows": {},
        "sections": [], "n_positions": 0, "sum_position_total": 0.0,
        "positions": [], "index": {}, "positions_sum": {},
    }
    max_row = ws.max_row
    head_last = header_last_row_17(ws, max_row)
    out["head_last_row"] = head_last

    def put_head(key: str, value: float, digits: int = 4, scale: bool = True) -> None:
        if not value:
            return
        out["head"][key] = round(value * (1000.0 if scale else 1.0), digits)

    # ---------- шапка ----------
    for r in range(1, head_last + 1):
        a = _s(ws.cell(r, 1).value)
        d = _s(ws.cell(r, 4).value)
        if a:
            low = a.lower()
            if "сметный расчет" in low or "смета" in low and "№" in a and not out["number"]:
                m = re.search(r"№\s*(.+)$", a)
                if m:
                    out["number"] = m.group(1).strip()
            elif low.startswith("наименование стройки") or "проектная документация" in low:
                out["doc_code"] = a
        if d:
            low = d.lower()
            val = _f(ws.cell(r, 7).value)
            if low.startswith("основание"):
                out["osnovanie"] = _s(ws.cell(r, 7).value)
            elif low.startswith("сметная стоимость"):
                put_head("smeta_cost", val, 2)
            elif low.startswith("строительных работ"):
                put_head("cost_build", val, 2)
            elif low.startswith("монтажных работ"):
                put_head("cost_install", val, 2)
            elif low.startswith("оборудования"):
                put_head("cost_equipment", val, 2)
            elif low.startswith("прочих работ"):
                put_head("cost_other", val, 2)
            elif low.startswith("средства на оплату труда"):
                put_head("smeta_wage", val, 2)
            elif low.startswith("сметная трудоемкость"):
                put_head("norm_labor", val, 4, scale=False)
            elif low.startswith("трудозатраты механизаторов"):
                put_head("tz_mash", val, 4, scale=False)
    # наименование работ и затрат — строка перед блоком «Основание»/«Сметная стоимость»
    for r in range(1, head_last + 1):
        a = _s(ws.cell(r, 1).value)
        if a and not _s(ws.cell(r, 4).value) and r > 1:
            prev = _s(ws.cell(r - 1, 1).value)
            if prev and ("№ оср" in prev.lower() or prev.lower().startswith("№")):
                out["name"] = a
                break

    # ---------- итоговый блок и позиции ----------
    cand: Dict[str, List[Tuple[float, int]]] = {}

    def add(key: str, value: float, row: int) -> None:
        if value:
            cand.setdefault(key, []).append((value, row))

    cur_section = None
    for r in range(head_last + 1, max_row + 1):
        a = _s(ws.cell(r, 1).value)
        c = _s(ws.cell(r, 3).value)
        val = _f(ws.cell(r, 10).value)          # деньги итогового блока — столбец 10
        total = _f(ws.cell(r, FO_COL["total"]).value)
        if a:
            low = a.lower()
            if low.startswith("итого по смете, в т.ч"):
                add("total", val, r)
            elif low.startswith("накладные расходы") and "в том числе" not in low:
                add("nr", val, r)
            elif low.startswith("сметная прибыль") and "в том числе" not in low:
                add("sp", val, r)
            elif low.startswith("прямые затраты"):
                add("pz", val, r)
            elif low.startswith("стоимость материальных ресурсов"):
                add("material", val, r)
            elif low.startswith("стоимость оборудования") or low == "оборудование":
                add("equipment", val, r)
            elif low.startswith("эмм (без зпм)"):
                add("emm_wo_zpm", val, r)
            elif low.startswith("сметная заработная плата"):
                add("fot", val, r)
            elif low.startswith("нормативная трудоемкость"):
                add("norm_labor", val, r)
            elif low.startswith("строительные работы с нр и сп"):
                add("smr_build", val, r)
            elif low.startswith("монтажные работы с нр и сп"):
                add("smr_install", val, r)
            elif low.startswith("прочие с нр и сп"):
                add("other", val, r)
            elif low.startswith("всего, стоимость") and "работ" in low:
                m = re.match(r"всего,\s*стоимость\s+(.+?)\s+работ", low)
                kind = (m.group(1) if m else "").strip()
                if "общестроительн" in kind or "строительн" == kind:
                    add("smr_build", val, r)
                elif "монтажн" in kind:
                    add("smr_install", val, r)
            elif low.startswith("всего, стоимость оборудования"):
                add("equipment", val, r)
            elif low.startswith("трудозатраты машинистов"):
                add("tz_mash", val, r)
            elif low.startswith("трудозатраты строителей"):
                add("tz_workers", val, r)
        if c.lower().startswith("раздел"):
            cur_section = c.strip()
        # позиции: номер в столбце 1, обоснование — в столбце 2
        b = _s(ws.cell(r, FO_COL["basis"]).value)
        if _is_position_num17(a) and b:
            out["n_positions"] += 1
            out["sum_position_total"] += total
            for k, col in (("ozp", FO_COL["ozp"]), ("emm", FO_COL["emm"]), ("zpm", FO_COL["zpm"]),
                           ("equipment", FO_COL["equipment"]), ("total", FO_COL["total"]),
                           ("tz_workers", FO_COL["tz_workers"])):
                out["positions_sum"][k] = out["positions_sum"].get(k, 0.0) + _f(ws.cell(r, col).value)
            if len(out["positions"]) < 5000:
                out["positions"].append({
                    "num": a, "basis": b, "name": _s(ws.cell(r, FO_COL["name"]).value),
                    "unit": _s(ws.cell(r, FO_COL["unit"]).value),
                    "qty": _f(ws.cell(r, FO_COL["qty"]).value),
                    "unit_cost": _f(ws.cell(r, FO_COL["unit_cost"]).value),
                    "equipment": round2(_f(ws.cell(r, FO_COL["equipment"]).value)),
                    "total": round2(total),
                    "ozp": round2(_f(ws.cell(r, FO_COL["ozp"]).value)),
                    "emm": round2(_f(ws.cell(r, FO_COL["emm"]).value)),
                    "zpm": round2(_f(ws.cell(r, FO_COL["zpm"]).value)),
                    "tz_workers": round4(_f(ws.cell(r, FO_COL["tz_workers"]).value)),
                    "section": cur_section or "",
                })

    def pick_max(key: str) -> Optional[float]:
        if key not in cand:
            return None
        best = max(cand[key], key=lambda x: abs(x[0]))
        out["rows"][key] = best[1]
        return best[0]

    for key in ("pz", "total", "nr", "sp", "material", "equipment", "smr_build", "smr_install",
                "other", "ozp", "emm", "zpm", "tz_workers", "tz_mash", "return_mat"):
        v = pick_max(key)
        if v is not None:
            out["totals"][key] = round4(v) if key in LABOR_KEYS else round2(v)
    # Итоговый блок «17 граф» печатает ФОТ (ОЗП + ЗПМ) и ЭММ без ЗПМ, а показатели
    # технического задания — ОЗП, ЭММ (с ЗПМ) и ЗПМ по отдельности. Значения берём из сумм
    # по позициям формы (столбцы 12/13/14 — они совпадают с показателями узлов до копейки),
    # а ФОТ, «нормативную трудоёмкость» и «трудозатраты механизаторов» — из отпечатанных строк.
    psum = {k: round4(v) if k == "tz_workers" else round2(v)
            for k, v in (out.get("positions_sum") or {}).items()}
    zpm_pos = psum.get("zpm")
    if zpm_pos is not None:
        out["totals"]["zpm"] = round2(zpm_pos)
    if "ozp" not in out["totals"] and psum.get("ozp") is not None:
        out["totals"]["ozp"] = round2(psum["ozp"])
    if psum.get("emm") is not None:
        out["totals"]["emm"] = round2(psum["emm"])          # ЭММ с ЗПМ (показатель ТЗ)
    emm17 = pick_max("emm_wo_zpm")
    if emm17 is not None:
        out["totals"]["emm_wo_zpm"] = round2(emm17)          # напечатано «ЭММ (без ЗПМ)»
    fot = pick_max("fot")
    if fot is not None:
        out["head"]["fot"] = round2(fot)
        if zpm_pos is not None and "ozp" not in out["totals"]:
            out["totals"]["ozp"] = round2(fot - zpm_pos)
    # «Сметная трудоемкость» в шапке напечатана с полной точностью (172 603,320649),
    # а в итоговом блоке — округлённой (172 603): берём точное значение, если оно есть
    norm_block = pick_max("norm_labor")
    if not out["head"].get("norm_labor") and norm_block:
        out["head"]["norm_labor"] = round4(norm_block)
    if "tz_mash" not in out["totals"] and out["head"].get("tz_mash"):
        out["totals"]["tz_mash"] = round4(out["head"]["tz_mash"])
    norm_labor = out["head"].get("norm_labor")
    if norm_labor and "tz_workers" not in out["totals"] and "tz_mash" in out["totals"]:
        out["totals"]["tz_workers"] = round4(norm_labor - out["totals"]["tz_mash"])
    if psum.get("equipment"):
        out["positions_sum"]["equipment"] = round2(psum["equipment"])
    out["sum_position_total"] = round2(out["sum_position_total"])
    return out


def parse(ws) -> Dict[str, Any]:
    """Разобрать лист печатной формы (11 граф или 17 граф)."""
    title = (ws.title or "").lower()
    if MARK17 in title and MARK not in title:
        return parse17(ws)
    out: Dict[str, Any] = {
        "sheet": ws.title, "doc_code": "", "number": "", "name": "", "osnovanie": "",
        "object": "", "norm_base": "",
        "head": {}, "totals": {}, "totals_with_k": {}, "rows": {},
        "sections": [], "n_positions": 0, "sum_position_total": 0.0,
        "positions": [], "index": {},
    }
    cand: Dict[str, List[Tuple[float, int]]] = {}
    cand_idx: Dict[str, List[Tuple[float, int]]] = {}     # итоги с учётом индексов пересчёта

    def add(key: str, value: float, row: int, idx: bool = False) -> None:
        if value == 0.0:
            return
        store = cand_idx if idx else cand
        store.setdefault(key, []).append((value, row))

    def pick(key: str, store: Optional[Dict[str, List[Tuple[float, int]]]] = None) -> Optional[float]:
        """Наибольшее значение показателя (итог сметы не меньше итога её части).

        Среди значений, совпадающих с наибольшим с точностью до округления (полкопейки /
        0,005 чел.-ч), выбирается самое точное: иначе при округлении вверх справочная строка
        формы (1793,69) вытеснила бы точный итог (1793,6888) и сверка потеряла бы знаки.
        """
        src = store if store is not None else cand
        if key not in src:
            return None
        best = max(src[key], key=lambda x: abs(x[0]))
        same = [x for x in src[key] if abs(abs(x[0]) - abs(best[0])) <= ROUND_HALF + 1e-9]
        value, row = max(same, key=lambda x: (decimals(x[0]), abs(x[0])))
        if store is None:
            out["rows"][key] = row
        return value

    max_row = ws.max_row
    head_last = header_last_row(ws, max_row)
    out["head_last_row"] = head_last
    ls_row = 0

    # ---------- шапка ----------
    # В части выпусков формы рядом с текущей ценой печатается базисная — в скобках:
    # «Сметная стоимость: 69426.41 (6793,19) тыс.руб.». Для сверки с вкладкой «Source»
    # (базисный уровень цен 2001 г.) берём значение из скобок, текущее сохраняем отдельно.
    def head_set(key: str, cell_i: Any, cell_j: Any, digits: int) -> None:
        base = paren_value(cell_j)
        if base is None:
            out["head"][key] = round(_f(cell_i) * 1000, digits)
        else:
            out["head"][key] = round(base * 1000, digits)
            out["head"][key + "_index"] = round(_f(cell_i) * 1000, digits)

    for r in range(1, head_last + 1):
        a = _s(ws.cell(r, 1).value)
        c = _s(ws.cell(r, 3).value)
        i = _f(ws.cell(r, 9).value)
        if not a:
            continue
        low = a.lower()
        if low.startswith("локальная смета"):
            ls_row = r
            m = re.search(r"№\s*(.+)$", a)
            out["number"] = (m.group(1).strip() if m else "")
            out["name"] = _s(ws.cell(r + 1, 1).value)
        elif low.startswith("объект"):
            out["object"] = c
        elif low.startswith("основание"):
            out["osnovanie"] = c
        elif low.startswith("сметная стоимость"):
            head_set("smeta_cost", ws.cell(r, 9).value, ws.cell(r, 10).value, 2)
        elif low.startswith("строительных работ"):
            head_set("cost_build", ws.cell(r, 9).value, ws.cell(r, 10).value, 2)
        elif low.startswith("монтажных работ"):
            head_set("cost_install", ws.cell(r, 9).value, ws.cell(r, 10).value, 2)
        elif low.startswith("оборудования"):
            head_set("cost_equipment", ws.cell(r, 9).value, ws.cell(r, 10).value, 2)
        elif low.startswith("прочие"):
            head_set("cost_other", ws.cell(r, 9).value, ws.cell(r, 10).value, 2)
        elif low.startswith("возврат"):
            head_set("return_mat", ws.cell(r, 9).value, ws.cell(r, 10).value, 2)
        elif low.startswith("нормативная трудоемкость"):
            head_set("norm_labor", ws.cell(r, 9).value, ws.cell(r, 10).value, 4)
        elif low.startswith("сметная заработная плата"):
            head_set("smeta_wage", ws.cell(r, 9).value, ws.cell(r, 10).value, 2)
        elif low.startswith("составлена в"):
            m = re.search(r"НБ:\s*(.+)$", a)
            if m:
                out["norm_base"] = m.group(1).strip()
    if ls_row > 1:
        prev = _s(ws.cell(ls_row - 1, 1).value)
        if prev and not prev.lower().startswith(("объект", "стройка", "номер")):
            out["doc_code"] = prev

    # ---------- итоговые блоки и позиции ----------
    # В книгах с пересчётом базовых цен в текущие после строки «ВСЕГО ПО СМЕТЕ» печатается
    # блок «Итоги с учётом индексов пересчёта …» с теми же строками в текущем уровне цен.
    # Для сверки с «Source» (базисный уровень) используем первый блок, а итоги с индексами
    # складываем отдельно — они выводятся в отчёте как справочные.
    in_index = False
    cur_razdel = ""          # текущий раздел формы (для привязки подразделов)
    for r in range(head_last + 1, max_row + 1):
        a = _s(ws.cell(r, 1).value)
        g = _f(ws.cell(r, 7).value)
        h = _f(ws.cell(r, 8).value)
        i = _f(ws.cell(r, 9).value)
        k = _f(ws.cell(r, 11).value)
        if a:
            low = a.lower()
            if low.startswith("итоги с учётом индекс"):
                in_index = True
                out["index"]["row"] = r
                out["index"]["title"] = a
                continue
            if a.startswith("Итого по смете") or low.startswith("итого по смете (без нр и сп)") \
                    or low.startswith("прямые затраты"):
                add("pz", g, r, in_index)
                if h:
                    add("ozp", h, r, in_index)
                if i:
                    add("emm", i, r, in_index)
                if k:
                    add("tz_workers", k, r, in_index)
                nxt_i, nxt_k = _next_row_labor(ws, r + 1)
                if nxt_i:
                    add("zpm", nxt_i, r + 1, in_index)
                if nxt_k:
                    add("tz_mash", nxt_k, r + 1, in_index)
            elif a == "Всего по смете":
                add("total", g, r, in_index)
                if h:
                    add("ozp", h, r, in_index)
                if i:
                    add("emm", i, r, in_index)
                if k:
                    add("tz_workers", k, r, in_index)
                nxt_i, nxt_k = _next_row_labor(ws, r + 1)
                if nxt_i:
                    add("zpm", nxt_i, r + 1, in_index)
                if nxt_k:
                    add("tz_mash", nxt_k, r + 1, in_index)
            elif a == "ВСЕГО ПО СМЕТЕ":
                add("total_with_k", g, r, in_index)
            elif low.startswith("всего, накладные расходы"):
                add("nr", g, r, in_index)
            elif low.startswith("всего, сметная прибыль"):
                add("sp", g, r, in_index)
            elif low.startswith("накладные расходы"):
                add("nr_part", g, r, in_index)
            elif low.startswith("сметная прибыль"):
                add("sp_part", g, r, in_index)
            elif low.startswith("в т.ч. стоимость материалов"):
                add("material", g, r, in_index)
            elif low.startswith("перевозка грузов"):
                add("transport", g, r, in_index)
            elif "стоимость оборудования" in low:
                add("equipment", g, r, in_index)
            elif low == "оборудование":
                add("equipment", g, r, in_index)
            elif low.startswith("трудозатраты строителей"):
                add("tz_workers", g, r, in_index)
            elif low.startswith("трудозатраты машинистов"):
                add("tz_mash", g, r, in_index)
            elif low.startswith("возврат материалов"):
                add("return_mat", g, r, in_index)
            elif low.startswith("итого смр"):
                add("smr_total_k", g, r, in_index)
            elif low.startswith("всего, стоимость") and "работ" in low:
                m = re.match(r"всего,\s*стоимость\s+(.+?)\s+работ(.*)$", low)
                kind = (m.group(1) if m else "прочих").strip()
                key = "smr_" + _kind_key(kind) + ("_k" if "с учетом" in low else "")
                add(key, g, r, in_index)
            elif low.startswith("стоимость") and "работ" in low:
                m = re.match(r"стоимость\s+(.+?)\s+работ", low)
                if m:
                    add("pz_" + _kind_key(m.group(1)), g, r)
            elif low.startswith("раздел "):
                cur_razdel = a[7:].strip()      # заголовок блока «Раздел …»
            elif low.startswith("подраздел "):
                pass                            # заголовок блока «Подраздел …» (не итог)
            elif low.startswith("итого по разделу") or low.startswith("итого по подразделу"):
                kind = "раздел" if low.startswith("итого по разделу") else "подраздел"
                nm = re.sub(r"^итого по (разделу|подразделу)\s*\(без нр и сп\)\s*", "", a, flags=re.I)
                # в части выпусков строка «Итого по подразделу» печатается без чисел, а суммы —
                # на следующей строке «Прямые затраты»; ЗПМ и трудозатраты машинистов — ещё ниже
                num_row = r + 1
                if not (g or h or i or k) and _s(ws.cell(num_row, 1).value).lower().startswith("прямые затраты"):
                    g = _f(ws.cell(num_row, 7).value)
                    h = _f(ws.cell(num_row, 8).value)
                    i = _f(ws.cell(num_row, 9).value)
                    k = _f(ws.cell(num_row, 11).value)
                    num_row += 1
                nxt_i, nxt_k = _next_row_labor(ws, num_row)
                entry = {
                    "kind": kind, "name": nm.strip(), "pz": round2(g),
                    "ozp": round2(h), "emm": round2(i), "tz_workers": round4(k),
                    "zpm": round4(nxt_i), "tz_mash": round4(nxt_k), "row": r}
                # подраздел печатается внутри своего раздела — запоминаем раздел для
                # однозначного сопоставления с узлами (в смете бывает несколько «Трубопроводов»)
                if kind == "подраздел":
                    if cur_razdel:
                        entry["parent"] = cur_razdel
                else:
                    cur_razdel = nm.strip()
                out["sections"].append(entry)
            elif low.startswith("коэфф") and g > 1.0:
                out["head"]["k_usp"] = round(g, 6)

        # позиции формы: номер в столбце A, шифр — в столбце B. Номер может быть
        # составным («13,1», «34,2») — так печатаются поправки к позиции: они входят
        # в итог формы отдельными строками, поэтому входят и в перечень позиций.
        b = _s(ws.cell(r, 2).value)
        if _is_position_num(a) and b and not b.isdigit():
            out["n_positions"] += 1
            out["sum_position_total"] += g
            if len(out["positions"]) < 5000:
                out["positions"].append({
                    "num": a, "basis": b, "name": _s(ws.cell(r, 3).value),
                    "qty": _f(ws.cell(r, 4).value), "total": round(g, 2),
                    "unit_cost": _f(ws.cell(r, 5).value), "ozp": round(h, 2),
                    "emm": round(i, 2), "zpm": round(_f(ws.cell(r, 24).value), 2),
                    "nr": round(_f(ws.cell(r, 27).value), 2), "sp": round(_f(ws.cell(r, 28).value), 2),
                    "tz_workers": round4(k)})

    for key in ("pz", "total", "nr", "sp", "material", "transport", "equipment",
                "tz_workers", "tz_mash", "ozp", "emm", "zpm", "return_mat", "smr_total_k"):
        v = pick(key)
        if v is not None:
            # трудозатраты — в чел.-ч, они хранятся с точностью 4 знака; деньги — копейки
            out["totals"][key] = round4(v) if key in LABOR_KEYS else round2(v)
    v = pick("total_with_k")
    if v is not None:
        out["totals_with_k"]["total"] = round(v, 2)
    for key in list(cand):
        if key.startswith("smr_") or key.startswith("pz_"):
            v = pick(key)
            if v is not None:
                target = out["totals_with_k"] if key.endswith("_k") else out["totals"]
                target[key] = round(v, 2)
    # итоги с учётом индексов пересчёта (текущий уровень цен) — справочно
    for key in list(cand_idx):
        v = pick(key, cand_idx)
        if v is not None:
            out["index"].setdefault("totals", {})[key] = round4(v) if key in LABOR_KEYS else round2(v)
    # НР и СП: если агрегированных строк «Всего, …» нет — берём сумму частей
    for key in ("nr", "sp"):
        if key not in out["totals"] and (key + "_part") in cand:
            total = sum(v for v, _ in cand[key + "_part"])
            out["totals"][key] = round(total, 2)
            out["rows"][key] = cand[key + "_part"][0][1]
    # --- уточнение итогов разделов/подразделов: НР, СП, материалы внутри блока ---
    marks = sorted([(sec["row"], i) for i, sec in enumerate(out["sections"])
                    if sec.get("pz") is not None] +
                   [(int(out["rows"].get("pz", 10 ** 9)), -1)], key=lambda x: x[0])
    for pos, (row, idx) in enumerate(marks):
        if idx < 0:
            break
        stop = marks[pos + 1][0] if pos + 1 < len(marks) else max_row + 1
        sec = out["sections"][idx]
        nr = sp = mat = 0.0
        for r in range(row + 1, min(stop, max_row + 1)):
            a = _s(ws.cell(r, 1).value)
            g = _f(ws.cell(r, 7).value)
            if not a:
                continue
            low = a.lower()
            if low.startswith("накладные расходы") and "%" in a:
                nr += g
            elif low.startswith("сметная прибыль") and "%" in a:
                sp += g
            elif low.startswith("в т.ч. стоимость материалов"):
                mat += g
            elif low.startswith("всего, стоимость материалов без начислений"):
                mat += g
        sec["nr"] = round(nr, 2)
        sec["sp"] = round(sp, 2)
        sec["material"] = round(mat, 2)
        sec["total"] = round((sec.get("pz") or 0.0) + nr + sp, 2)

    out["sum_position_total"] = round(out["sum_position_total"], 2)
    return out


def _kind_key(text: str) -> str:
    t = text.lower()
    if "общестроит" in t:
        return "build"
    if "монтаж" in t:
        return "install"
    if "ремонт" in t and "реставр" in t:
        return "restore"
    if "ремонт" in t:
        return "repair"
    if "металл" in t:
        return "metal"
    if "сантех" in t:
        return "plumb"
    if "буровзрыв" in t:
        return "drill"
    if "горнопроход" in t:
        return "mine"
    if "пусконалад" in t:
        return "start"
    if "проч" in t:
        return "other"
    if "оборуд" in t:
        return "equip"
    return re.sub(r"\W+", "_", t)[:24] or "other"


def source_missing(source_totals: Dict[str, float]) -> bool:
    """Все значимые показатели «Source» пусты — значит, в книге нет кэша формул
    (файл пересохранён сторонней программой без пересчёта)."""
    keys = [k for k, _ in COMPARE_KEYS if k not in ("transport", "other")]
    return all(abs(float(source_totals.get(k, 0.0) or 0.0)) < 1e-9 for k in keys)


def compare(source_totals: Dict[str, float], printed: Dict[str, float], tol: float = TOL
            ) -> List[Dict[str, Any]]:
    """Сравнение показателей: разбор «Source» против печатной формы.

    Статусы строк: «ok» — совпадает, «absent» — показателя нет в форме, «source_missing» —
    в «Source» нет значений, «kind_diff» — расходится разбивка СМР по видам работ
    (см. KIND_KEYS), «mismatch» — расхождение данных.
    """
    rows: List[Dict[str, Any]] = []
    no_source = source_missing(source_totals)
    for key, label in COMPARE_KEYS:
        src = float(source_totals.get(key, 0.0) or 0.0)
        prn = printed.get(key)
        if prn is None:
            rows.append({"key": key, "label": label, "source": round(src, 4),
                         "printed": None, "delta": None, "status": "absent"})
            continue
        delta = round(prn - src, 4)
        if no_source:
            status = "source_missing"
        elif abs(delta) <= max(tol, abs(prn) * 1e-9):
            status = "ok"
        elif key in KIND_KEYS:
            status = "kind_diff"      # разница методик деления СМР, см. KIND_KEYS
        else:
            status = "mismatch"
        rows.append({"key": key, "label": label, "source": round(src, 4),
                     "printed": round(prn, 4), "delta": delta, "status": status})
    return rows


def check_printed_self(printed: Dict[str, Any], tol: float = TOL) -> List[Dict[str, Any]]:
    """Внутренняя согласованность печатной формы (независимо от «Source»).

    Учитываются особенности формы: шапка печатается в тыс. руб. с двумя знаками
    (округление до 5 руб.), а коэффициент «Условия производства работ» применяется
    программой к НР и СП каждой позиции с округлением до целых рублей, поэтому итог
    с коэффициентом отличается от линейного пересчёта на доли рубля на позицию.
    """
    problems: List[Dict[str, Any]] = []
    head = printed.get("head", {})
    tot_k = printed.get("totals_with_k", {}).get("total")
    n_pos = max(1, int(printed.get("n_positions") or 1))
    if head.get("smeta_cost") and tot_k:
        d = head["smeta_cost"] - tot_k
        if abs(d) > max(5.0, abs(tot_k) * 1e-9):     # 0,01 тыс. руб. = 10 руб.
            problems.append({"code": "HEAD_COST", "severity": "warn", "delta": round(d, 2),
                             "message": "«Сметная стоимость» в шапке ≠ строке «ВСЕГО ПО СМЕТЕ»"})
    base = printed.get("totals", {})
    if base.get("total") and tot_k and tot_k + tol < base["total"]:
        problems.append({"code": "K_LT_BASE", "severity": "warn", "delta": round(tot_k - base["total"], 2),
                         "message": "«ВСЕГО ПО СМЕТЕ» меньше «Всего по смете» (базисный уровень)"})
    k = head.get("k_usp")
    if k and base.get("total") and tot_k:
        # коэффициент «Условия производства работ» применяется к ОЗП, ЭММ, НР и СП
        # (ЗПМ входит в ЭММ и отдельно не масштабируется)
        labor = (base.get("ozp", 0.0) + base.get("emm", 0.0)
                 + base.get("nr", 0.0) + base.get("sp", 0.0))
        expect = base["total"] + (k - 1.0) * labor
        d = tot_k - expect
        limit = max(5.0, n_pos * 1.0)      # округление НР и СП по каждой позиции
        if abs(d) > limit:
            problems.append({"code": "K_APPLY", "severity": "warn", "delta": round(d, 2),
                             "message": "Итог с коэффициентом условий работ %s не сходится "
                                        "с базисными данными (расхождение %.2f руб.)" % (k, d)})
        else:
            problems.append({"code": "K_APPLY_ROUND", "severity": "info", "delta": round(d, 2),
                             "message": "Итог с коэффициентом условий работ %s (округление НР и СП "
                                        "по позициям: %.2f руб.)" % (k, d)})
    if head.get("norm_labor") and base.get("tz_workers") is not None and base.get("tz_mash") is not None:
        kk = k or 1.0
        expect = (base.get("tz_workers", 0.0) + base.get("tz_mash", 0.0)) * kk
        d = head["norm_labor"] - expect
        # «Нормативная трудоемкость» печатается в тыс. чел.-ч с двумя знаками,
        # поэтому допуск — половина младшего разряда (5 чел.-ч)
        if abs(d) > max(5.0, abs(expect) * 1e-4):
            problems.append({"code": "NORM_LABOR", "severity": "warn", "delta": round(d, 4),
                             "message": "«Нормативная трудоемкость» шапки ≠ трудозатратам формы × %s" % kk})
    if head.get("smeta_wage") and base.get("ozp") is not None and base.get("zpm") is not None:
        kk = k or 1.0
        expect = (base.get("ozp", 0.0) + base.get("zpm", 0.0)) * kk
        d = head["smeta_wage"] - expect
        if abs(d) > max(5.0, abs(expect) * 1e-4):
            problems.append({"code": "SMETA_WAGE", "severity": "warn", "delta": round(d, 2),
                             "message": "«Сметная заработная плата» шапки ≠ (ОЗП + ЗПМ) × %s" % kk})
    return problems
