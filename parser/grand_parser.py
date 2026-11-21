#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
grand_parser.py — чтение смет, выгруженных из ПК «ГРАНД-Смета» (печатная форма).

Назначение
----------
Книги ГРАНД-Сметы не содержат служебной вкладки «Source», на которую опирается
основной парсер (smeta_parser.py). Единственный машиночитаемый слой такой книги —
печатная форма «Локальный сметный расчет (смета)»: шапка документа, таблица позиций
с блоками ресурсов и сводными строками (ОТ / ЭМ / ОТм / М / ЗТ / ЗТм / ФОТ / НР / СП),
итоги разделов и итоговый блок «Итоги по смете:».

Модуль распознаёт эту форму и собирает данные во внутреннюю модель, полностью
совместимую с разбором вкладки «Source» АтомСметы: тот же набор показателей
(pz, mat_ob, material, equipment, transport, other, ozp, emm, zpm, nr, sp,
tz_workers, tz_mash, total, kind_pz), та же структура WorkbookParse
(header / nodes / items / resources / summary_rows / conditions), поэтому
отчёты строятся общими модулями без изменений.

Раскладка печатной формы ГРАНД-Сметы (листы «…по Методике 2020 (БИМ)…»)
----------------------------------------------------------------------
A  — № п/п позиции / «Раздел N. …»;   B — шифр (обоснование);
C  — наименование;                    F — единица измерения;
G  — количество на единицу;           H — коэффициент;
I  — всего с учётом коэффициентов;    J — расценка (стоимость) на единицу;
K  — коэффициент к расценке;          L — стоимость в базисном уровне, руб.;
M  — индекс пересчёта в текущий уровень;
N  — стоимость в текущем уровне, руб.

Строки внутри позиции:
  * ресурс:            B=шифр ресурса, C=наименование (1-4-0/1-4-… — ЗП рабочих,
                       91.… — эксплуатация машин, остальные — материалы);
  * сводка позиции:    B=1..4, C=ОТ / ЭМ / в т.ч. ОТм / М (L — базис, N — текущ.);
  * трудозатраты:      C=ЗТ / ЗТм (I — чел.-ч / маш.-ч с коэффициентом УПР);
  * C=ФОТ              (L — оплата труда рабочих + машинистов);
  * C=Итого по расценке (L — прямые затраты позиции);
  * B=Пр/…, C=НР…/СП…  (L — накладные расходы / сметная прибыль);
  * C=Всего по позиции (L — позиция с НР и СП, базис; N — то же, текущий уровень).

Итоговые блоки «Итоги по разделу N …» и «Итоги по смете:» дают показатели узлов
и сметы целиком; они же используются для внутренней сверки (код проверки GRAND_*).

Запуск (независимая проверка):
    python grand_parser.py <папка_или_файл> [-o grand.json]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from openpyxl import load_workbook
except Exception as exc:  # pragma: no cover
    sys.stderr.write("Не найден openpyxl: %s\n" % exc)
    raise

EPS = 0.01          # допуск сверки денег, руб.
EPS_TZ = 0.001      # допуск сверки трудозатрат, чел.-ч

# признаки книги ГРАНД-Сметы (печатная форма вместо вкладки Source)
GRAND_MARKS = ("ГРАНД-Смета", "ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ")

RE_SECTION = re.compile(r"^Раздел\s*(\d+)\s*[.:]\s*(.*)$", re.I)
RE_SEC_TOTALS = re.compile(r"^Итоги\s+по\s+разделу\s+(\d+)\b(.*)$", re.I)
RE_SEC_VSEGO = re.compile(r"^Всего\s+по\s+разделу\s+(\d+)\b(.*)$", re.I)
RE_NUM = re.compile(r"^\d+(?:[.,]\d+)?$")
RE_POS_A = re.compile(r"^\d+(?:[.,]\d+)?\s*(О?|\.\d+)?$")
RE_RES_B = re.compile(r"^[\dа-яA-Za-z_.\-/()]")
RE_ZP_RES = re.compile(r"^1-\d+-\d+")          # ресурс «Затраты труда рабочих»
RE_MASH_RES = re.compile(r"^91\.")             # машины и механизмы
# шифр сметной нормы: «ФЕР08-…», «ТЕРм р-35-01-…», «ГЭСН2020 6-01-001-01», «ФССЦпг …»
RE_NORM_CODE = re.compile(
    r"^(?:фер|тер|гэсм?|фссц|тссц|рес)[а-яё0-9]*(?:\s*[-–]?\s*\d{2,4})?", re.I)


def is_norm_code(code: Any) -> bool:
    """True, если шифр строки — смета норма (расценка), а не код материала/машины."""
    t = s(code).strip()
    return bool(t) and bool(RE_NORM_CODE.match(t))


def s(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


def f(v: Any) -> float:
    """Число из ячейки; запятая как десятичный разделитель тоже принимается."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace("\u00a0", " ").replace(" ", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def money(v: float) -> float:
    return round(v + 0.0, 2)


def tz(v: float) -> float:
    return round(v + 0.0, 4)


def looks_like_grand(path: str) -> bool:
    """True, если книга — выгрузка ГРАНД-Сметы (нет «Source», есть печатная форма)."""
    wb = load_workbook(path, read_only=True)
    try:
        names = [n.strip().lower() for n in wb.sheetnames]
        if "source" in names:
            return False
        for name in wb.sheetnames[:3]:
            ws = wb[name]
            hits = 0
            for row in ws.iter_rows(min_row=1, max_row=60, min_col=1, max_col=14,
                                    values_only=True):
                for v in row:
                    t = s(v)
                    if any(mk in t for mk in GRAND_MARKS):
                        hits += 1
                if hits >= 2:
                    break
            if hits >= 2:
                return True
        return False
    finally:
        wb.close()


class _Block:
    """Накопитель показателей одного узла (позиции / подраздела / раздела / сметы)."""

    KEYS = ("pz", "mat_ob", "material", "equipment", "transport", "other",
            "ozp", "emm", "zpm", "nr", "sp", "tz_workers", "tz_mash", "total")

    def __init__(self) -> None:
        self.v: Dict[str, float] = {k: 0.0 for k in self.KEYS}
        self.cur: Dict[str, float] = {k: 0.0 for k in self.KEYS}   # текущий уровень цен
        self.work_kind: Optional[str] = None
        self.k_upr: Optional[float] = None      # коэффициент условий производства работ (H)
        self.pct_nr: Optional[float] = None     # норматив НР, % (строка Пр/…НР…)
        self.pct_sp: Optional[float] = None     # норматив СП, %

    def totals(self) -> Dict[str, float]:
        out = {k: money(self.v[k]) for k in self.KEYS if k not in ("tz_workers", "tz_mash")}
        out["tz_workers"] = tz(self.v["tz_workers"])
        out["tz_mash"] = tz(self.v["tz_mash"])
        out["kind_pz"] = {}
        return out


def _label(c_val: Any, b_val: Any) -> str:
    t = s(c_val)
    if not t:
        return ""
    return t.strip().rstrip(":").lower()


def parse_book(path: str) -> Dict[str, Any]:
    """Разобрать книгу ГРАНД-Сметы → словарь, совместимый с WorkbookParse."""
    file_name = os.path.basename(path)
    wb = load_workbook(path, data_only=True)
    try:
        sheet_names = list(wb.sheetnames)
        ws = wb[sheet_names[0]]
        rows: List[List[Any]] = []
        ncols = 15
        for r in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ncols):
            rows.append([c.value for c in r])
    finally:
        wb.close()

    header: Dict[str, Any] = {"number": "", "name": "", "osnovanie": "", "object": "",
                              "norm_base": "", "made_by": "", "checked_by": "",
                              "tr": "", "corrections": "", "method": "", "price_level": ""}
    head_totals: Dict[str, float] = {}     # шапка: сметная стоимость и прочее (текущий уровень)
    head_labor: Dict[str, float] = {}      # правая колонка шапки (средства на оплату труда и пр.)

    table_start = None                     # строка с «№ п/п» (1-based)
    footer_start = None                    # строка «Итоги по смете:»
    for i, row in enumerate(rows, 1):
        a, b, c = s(row[0]), s(row[1]), s(row[2])
        if a.startswith("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ"):
            m = re.search(r"№\s*([\w./-]+)", a)
            header["number"] = m.group(1) if m else ""
        elif a == "(наименование стройки)" or a == "(наименование объекта капитального строительс"[:40]:
            pass
        if b == "Основание" and not header["osnovanie"]:
            header["osnovanie"] = s(row[1]) and s(row[2 - 1])  # значение в B
            header["osnovanie"] = s(row[1])
        if a == "Составлен" and not header.get("method"):
            header["method"] = s(row[1])
        if a.startswith("Сметная стоимость"):
            head_totals["total_current"] = f(row[2]) * 1000.0
            head_totals["total_current_paren"] = f(row[3]) * 1000.0
        if a == "Составлен(а) в текущем (базисном) уровне цен ":
            header["price_level"] = s(row[3]) or s(row[2])
        if b in ("строительных работ", "монтажных работ", "оборудования", "прочих затрат"):
            key = {"строительных работ": "build", "монтажных работ": "install",
                   "оборудования": "equipment", "прочих затрат": "other"}[b]
            head_totals["head_" + key] = f(row[2]) * 1000.0
            head_totals["head_" + key + "_paren"] = f(row[3]) * 1000.0
        if g := _label(c, b):
            pass
        if s(row[6]).startswith("Средства на оплату труда рабочих"):
            head_labor["fot_current"] = f(row[11]) * 1000.0
        if s(row[6]).startswith("Нормативные затраты труда рабочих"):
            head_labor["tz_workers_norm"] = f(row[11])
        if s(row[6]).startswith("Нормативные затраты труда машинистов"):
            head_labor["tz_mash_norm"] = f(row[11])
        if a == "№ п/п" and s(row[1]).startswith("Обоснование"):
            table_start = i
        if c == "Итоги по смете:":
            footer_start = i
        if a == "Составил:" and not header.get("made_by"):
            header["made_by"] = s(row[7])
        if a == "Проверил:" and not header.get("checked_by"):
            header["checked_by"] = s(row[7])

    # наименование сметы и объекта — строки перед маркером «(наименование …)»
    for i in range(1, min(table_start or 35, len(rows))):
        a = s(rows[i - 1][0])
        nxt = s(rows[i][0]) if i < len(rows) else ""
        if nxt.startswith("(наименование работ и затрат)") and a and not header["name"]:
            header["name"] = a
        if nxt.startswith("(наименование объекта") and a and not header["object"]:
            header["object"] = a
    header["name"] = header["name"] or header["object"]

    if table_start is None:
        raise ValueError("книга «%s»: не найдена шапка таблицы («№ п/п» / «Обоснование»)" % file_name)
    if footer_start is None:
        footer_start = len(rows) + 1

    # ------------------------------------------------------------------ тело
    nodes: List[Dict[str, Any]] = []          # разделы (kind=4) в порядке следования
    items: List[Dict[str, Any]] = []
    resources: List[Dict[str, Any]] = []

    pos: Optional[Dict[str, Any]] = None      # текущая позиция
    cur_section: Optional[Dict[str, Any]] = None
    section_rows: List[Tuple[int, Dict[str, Any]]] = []

    def close_position() -> None:
        """Завершить текущую позицию: довести итоги и выложить её строки."""
        nonlocal pos
        if pos is None:
            return
        p, pos = pos, None
        blk: _Block = p["blk"]
        it = p["item"]
        if not blk.v["pz"]:
            blk.v["pz"] = blk.v["ozp"] + blk.v["emm"] + blk.v["mat_ob"]
        # материалы позиции: если печатная форма не напечатала сводную строку «М»,
        # берём сумму строковых ресурсов позиции (ПЗ тогда сходится с ОТ+ЭМ+М)
        if not blk.v["material"] and p["resources"]:
            m_sum = money(sum(rc["cost"] for rc in p["resources"]))
            if m_sum:
                blk.v["material"] = m_sum
                blk.v["mat_ob"] = max(blk.v["mat_ob"], m_sum)
                blk.cur.setdefault("material", money(
                    sum(rc.get("cost_current", 0.0) for rc in p["resources"])))
        if not blk.v["total"]:
            blk.v["total"] = blk.v["pz"] + blk.v["nr"] + blk.v["sp"]
        vals = {k: money(blk.v[k]) for k in _Block.KEYS}
        vals["tz_workers"] = tz(blk.v["tz_workers"])
        vals["tz_mash"] = tz(blk.v["tz_mash"])
        vals["total_current"] = money(blk.cur.get("total", 0.0))
        vals["fot"] = money(blk.v.get("fot", 0.0))
        vals["k_upr"] = blk.k_upr or 1.0
        it["values"] = vals
        it["node"] = cur_section["id"] if cur_section else 1
        it["nr_pct"] = blk.pct_nr or 0.0
        it["sp_pct"] = blk.pct_sp or 0.0
        items.append(it)
        for rc in p["resources"]:
            rc["ls"] = None
            resources.append(rc)

    pending_tags: List[Tuple[int, str]] = []
    sec_no = 0
    i = table_start + 1
    while i < footer_start:
        row = rows[i - 1]
        r = i
        a, b, c = s(row[0]), s(row[1]), s(row[2])
        lab = _label(c, b)
        plain = lab.replace(" ", "")

        m = RE_SECTION.match(a)
        if m:
            close_position()
            sec_no += 1
            cur_section = {"id": sec_no + 1, "kind": 4, "row": r,
                           "code": "Раздел %s" % m.group(1), "name": m.group(2).strip()}
            nodes.append(cur_section)
            section_rows.append((r, cur_section))
            i += 1
            continue

        # служебные строки позиции
        if pos is not None:
            blk = pos["blk"]
            handled = True
            if b and RE_NUM.match(b) and lab in ("от", "эм", "в т.ч. отм", "м"):
                base, curv = f(row[11]), f(row[13])
                if lab == "от":
                    blk.v["ozp"] += base; blk.cur["ozp"] = blk.cur.get("ozp", 0.0) + curv
                elif lab == "эм":
                    blk.v["emm"] += base; blk.cur["emm"] = blk.cur.get("emm", 0.0) + curv
                elif lab.startswith("в т.ч"):
                    blk.v["zpm"] += base; blk.cur["zpm"] = blk.cur.get("zpm", 0.0) + curv
                else:
                    blk.v["material"] += base; blk.v["mat_ob"] += base
                    blk.cur["material"] = blk.cur.get("material", 0.0) + curv
                # K — коэффициент условий производства работ, применённый к строке
                if f(row[10]):
                    blk.k_upr = f(row[10])
            elif plain == "зт":
                blk.v["tz_workers"] += f(row[8])
            elif plain == "зтм":
                blk.v["tz_mash"] += f(row[8])
            elif plain == "фот":
                blk.v["fot"] = blk.v.get("fot", 0.0) + f(row[11])
                blk.cur["fot"] = blk.cur.get("fot", 0.0) + f(row[13])
            elif plain == "итогопорасценке":
                blk.v["pz"] += f(row[11]); blk.cur["pz"] = blk.cur.get("pz", 0.0) + f(row[13])
            elif b.startswith("Пр/") and (plain.startswith("нр") or "накладн" in plain):
                blk.v["nr"] += f(row[11]); blk.cur["nr"] = blk.cur.get("nr", 0.0) + f(row[13])
                blk.pct_nr = f(row[6]) or blk.pct_nr
            elif b.startswith("Пр/") and (plain.startswith("сп") or "прибыл" in plain):
                blk.v["sp"] += f(row[11]); blk.cur["sp"] = blk.cur.get("sp", 0.0) + f(row[13])
                blk.pct_sp = f(row[6]) or blk.pct_sp
            elif plain == "всегопопозиции":
                blk.v["total"] += f(row[11]); blk.cur["total"] = blk.cur.get("total", 0.0) + f(row[13])
                close_position()
            else:
                handled = False
            if handled:
                i += 1
                continue

        # новая позиция: номер в A (целое число или «N О» — оборудование)
        if RE_POS_A.match(a) and (b or c):
            close_position()
            code, name, unit = b, c, s(row[5])
            qty, price = f(row[8]), f(row[9])
            cat, wk = _classify(code, name, [(r0, t0) for (r0, t0) in pending_tags],
                                cur_section["name"] if cur_section else "")
            blk = _Block()
            blk.work_kind = wk
            pos = {"item": {"num": a.replace("\u00a0", " ").strip(), "row": r,
                            "code": code, "name": name, "basis": code, "unit": unit,
                            "qty": round(qty, 6), "cat": cat, "vid": "", "vid_code": 0,
                            "group": "", "work_kind": wk, "nr_pct": 0.0, "sp_pct": 0.0,
                            "price_variant": 0, "notes": [], "row_end": r},
                   "blk": blk, "resources": []}
            pending_tags = []
            i += 1
            continue

        # пометка вида работ/ресурсов над позицией: «(Электротехнические установки…)»
        if not a and not b and c.startswith("(") and pos is None:
            pending_tags.append((r, c))
            i += 1
            continue

        # заголовок подраздела в столбце A без номера раздела (например «ЭЛЕКТРОАППАРАТЫ»)
        if a and not b and not c and pos is None and not RE_SECTION.match(a):
            i += 1
            continue

        # условия производства работ / примечания к позиции (B=шифр условия, C=текст)
        if b and c and pos is None:
            pending_tags.append((r, c))
            i += 1
            continue

        if pos is not None and b and c:
            # строка ресурса внутри позиции либо её примечание
            has_val = any(f(row[k]) for k in (8, 9, 11, 13))
            if not has_val and not RE_RES_B.match(b):
                pos["item"]["notes"].append(c)
                pos["item"]["row_end"] = r
                i += 1
                continue
            base_cost, cur_cost = f(row[11]), f(row[13])
            rc = {"row": r, "parent": pos["item"]["num"], "code": b, "name": c,
                  "unit": s(row[5]), "qty": round(f(row[8]), 6),
                  "price": money(f(row[9])), "cost": money(base_cost),
                  "cost_current": money(cur_cost), "vid_code": 0, "vid": "", "group": "",
                  "kind": 0, "ls": None}
            if RE_ZP_RES.match(b):
                rc["cat"], rc["vid"], rc["group"] = "zp", "Затраты труда рабочих", "Работы"
                rc["vid_code"] = 1101
            elif RE_MASH_RES.match(b):
                rc["cat"], rc["vid"], rc["group"] = "machine", "Эксплуатация машин", "Машины и механизмы"
                rc["vid_code"] = 1200
            elif b.startswith("ТЦ_"):
                rc["cat"], rc["vid"], rc["group"] = "equipment" if pos["item"]["cat"] == "equipment" else "material", \
                    "Оборудование" if pos["item"]["cat"] == "equipment" else "Материалы", ""
                rc["vid_code"] = 116 if pos["item"]["cat"] == "equipment" else 1100
            else:
                rc["cat"], rc["vid"], rc["group"] = "material", "Материалы", ""
                rc["vid_code"] = 1100
            pos["resources"].append(rc)
            pos["item"]["row_end"] = r
            i += 1
            continue

        i += 1

    close_position()

    # --------------------------------------------------- печатные блоки итогов
    section_blocks: Dict[int, Dict[str, Any]] = {}
    i = table_start + 1
    while i < footer_start:
        c = s(rows[i - 1][2]).strip()
        m = RE_SEC_TOTALS.match(c)
        if m:
            num = int(m.group(1))
            blk = {"row": i, "v": {}, "cur": {}}
            _read_totals_block(rows, i - 1, blk["v"], blk["cur"])
            section_blocks[num] = blk
        i += 1

    fv: Dict[str, float] = {}
    fc: Dict[str, float] = {}
    _read_totals_block(rows, footer_start - 1, fv, fc)
    # строки «Всего по разделу N … | L=…» — печатный итог раздела с НР и СП
    for i in range(table_start, footer_start):
        m = RE_SEC_VSEGO.match(s(rows[i - 1][2]).strip())
        if m:
            num = int(m.group(1))
            if num in section_blocks:
                section_blocks[num]["v"]["total"] = f(rows[i - 1][11])
                section_blocks[num]["printed_total_row"] = i

    # ------------------------------------------------------------ сборка узлов
    ls_id = 1
    result_nodes: List[Dict[str, Any]] = [{
        "id": ls_id, "kind": 1, "row": footer_start, "code": header.get("number", ""),
        "name": header.get("name", ""), "totals": {}, "printed": dict(fv),
    }]
    by_sec_no: Dict[int, Dict[str, Any]] = {}
    for nd in nodes:
        m = re.search(r"\d+", nd["code"])
        num = int(m.group()) if m else None
        pr = dict(section_blocks.get(num, {}).get("v") or {})
        n2 = {"id": nd["id"], "kind": 4, "row": nd["row"], "code": nd["code"],
              "name": nd["name"], "totals": {}, "printed": pr, "parent": ls_id}
        result_nodes.append(n2)
        if num is not None:
            by_sec_no[num] = n2

    sec_rows_sorted = sorted((n["row"], n["id"]) for n in result_nodes if n["kind"] == 4)

    def section_for(row: int) -> int:
        cur = ls_id
        for srow, sid in sec_rows_sorted:
            if srow <= row:
                cur = sid
        return cur

    for it in items:
        it["node"] = section_for(it["row"])
    for rc in resources:
        rc["node"] = section_for(rc["row"])

    # суммы узлов по позициям
    agg: Dict[int, Dict[str, float]] = {n["id"]: {k: 0.0 for k in _Block.KEYS} for n in result_nodes}
    kind_pz: Dict[int, Dict[str, float]] = {n["id"]: {} for n in result_nodes}
    fot_sum: Dict[int, float] = {n["id"]: 0.0 for n in result_nodes}
    for it in items:
        tgt = agg[it["node"]]
        for k in _Block.KEYS:
            tgt[k] += it["values"].get(k, 0.0)
        fot_sum[it["node"]] += it["values"].get("fot", 0.0)
        wk = it.get("work_kind") or "other"
        kind_pz[it["node"]][wk] = kind_pz[it["node"]].get(wk, 0.0) + it["values"].get("pz", 0.0)
    for n in result_nodes:
        if n["kind"] == 4:
            for k in _Block.KEYS:
                agg[ls_id][k] += agg[n["id"]][k]
            fot_sum[ls_id] += fot_sum[n["id"]]
            for k, v in kind_pz[n["id"]].items():
                kind_pz[ls_id][k] = kind_pz[ls_id].get(k, 0.0) + v

    def fill_totals(n: Dict[str, Any]) -> None:
        a = agg[n["id"]]
        pr = n.get("printed") or {}
        t = {k: money(a[k]) for k in _Block.KEYS}
        t["tz_workers"] = tz(a["tz_workers"])
        t["tz_mash"] = tz(a["tz_mash"])
        # печатные итоги приоритетнее сумм позиций: они гарантируют паритет с формой
        for key in ("pz", "ozp", "emm", "zpm", "material", "nr", "sp", "equipment",
                    "build", "install", "other", "tz_workers", "tz_mash", "total"):
            if key in pr:
                t[key] = tz(pr[key]) if key.startswith("tz") else money(pr[key])
        if "material" in pr:
            t["mat_ob"] = money(t["material"] + t.get("equipment", 0.0) + t["transport"])
        if not t["pz"]:
            t["pz"] = money(t["ozp"] + t["emm"] + t["mat_ob"])
        if not t["total"]:
            t["total"] = money(t["pz"] + t["nr"] + t["sp"])
        t.setdefault("smr_build", money(pr.get("build", 0.0)))
        t.setdefault("smr_install", money(pr.get("install", 0.0)))
        t["kind_pz"] = {k: money(v) for k, v in kind_pz[n["id"]].items()}
        n["totals"] = t

    for n in result_nodes:
        fill_totals(n)

    ls = result_nodes[0]
    total_base = ls["totals"]["total"]
    total_cur = fv.get("vsego_po_smete") or total_base
    labor = (ls["totals"]["ozp"] + ls["totals"]["emm"] + ls["totals"]["nr"] + ls["totals"]["sp"])
    # коэффициент условий производства работ: медиана отношения фактических НР позиции
    # к напечатанной базе (в БИМ-форме НР уже идут с индексом, поэтому k выводим из
    # отношения ΣНР футера к ΣНР позиций только при явном расхождении) — иначе 1.
    k_upr = _median([it["values"].get("k_upr") for it in items]) or 1.0
    conditions: List[Dict[str, Any]] = []
    if abs(k_upr - 1.0) > 1e-6:
        conditions.append({"key": "КСт_УПР", "label": "Условия производства работ",
                           "title": "Условия производства работ (коэффициент строк затрат)",
                           "on": True, "value": k_upr, "num": "", "kind": 0})
    calc_k = money(total_base + (k_upr - 1.0) * labor)

    summary_rows = [
        {"row": footer_start, "node": ls_id, "key": "ВсегоПоСмете",
         "display": "Всего по смете", "label": "Всего по смете (базисный уровень)",
         "value": money(total_base)},
        {"row": footer_start, "node": ls_id, "key": "ВСЕГОПОСМЕТЕ",
         "display": "ВСЕГО ПО СМЕТЕ", "label": "ВСЕГО по смете",
         "value": money(total_cur)},
    ]

    wp = {
        "file": file_name,
        "path": path,
        "format": "grand",
        "sheet": sheet_names[0],
        "sheets": sheet_names,
        "header": header,
        "nodes": result_nodes,
        "items": items,
        "resources": resources,
        "notes": [],
        "conditions": conditions,
        "doc_summary": {},
        "summary_rows": summary_rows,
        "has_sections": bool(nodes),
        "open_row": {ls_id: table_start}, "close_row": {ls_id: footer_start},
        "node_indicators": {},
        "n_spec_rows": len(resources),
        "price_levels": [], "price_variants": {}, "main_variant": None,
        "price_selection": {},
        "printed": [],
        "grand": {
            "head_totals": head_totals,
            "head_labor": head_labor,
            "footer": fv,
            "footer_current": fc,
            "vat": fc.get("nds"),
            "total_base": money(total_base),
            "total_current": money(total_cur),
            "work_coefficient": k_upr,
            "sections_printed": {n: b["v"] for n, b in section_blocks.items()},
        },
    }
    return wp


def _median(vals: List[Optional[float]]) -> Optional[float]:
    xs = sorted(float(v) for v in vals if v)
    if not xs:
        return None
    return xs[len(xs) // 2]


def _classify(code: str, name: str, pending: List[Tuple[int, str]],
              section_name: str = "") -> Tuple[str, str]:
    """Категория и вид работ позиции по шифру, пометкам и разделу печатной формы."""
    low = ("%s %s" % (code, name)).lower()
    tags = " ".join(s(p[1]).lower() for p in pending)
    sec = s(section_name).lower()
    if is_norm_code(code):
        # нормативная расценка — это работа; перевозка ФССЦпг — отдельная категория
        if re.search(r"перевозк", low) or code.lower().startswith("фссцпг"):
            return "transport", "build"
        if "монтажн" in sec or "монтажн" in tags:
            return "works", "install"
        return "works", "build"
    # без шифра нормы: оборудование / материалы по цене поставщика («ТЦ_…», ФССЦ, прайсы)
    if "оборудован" in tags or "оборудован" in sec:
        return "equipment", "equipment"
    if re.search(r"перевозк", low):
        return "transport", "build"
    if "монтажн" in tags or "монтажн" in sec:
        return "material", "install"
    return "material", "other"


def row_kind_of(it: Dict[str, Any]) -> str:
    """Тип строки для укрупнённых групп: работа (норма), машина или материал."""
    code = s(it.get("code"))
    if is_norm_code(code):
        return "work"
    if RE_MASH_RES.match(code):
        return "machine"
    return "material"


def group_items(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Группы «работа + её материалы»: номер и шифр нормы ведущей строки — всем строкам.

    Тот же алгоритм, что в smeta_parser.group_items (паритет браузерного разбора и
    автономного парсера): новая группа начинается с позиции-работы (шифр сметной
    нормы); машины и материалы продолжают группу предыдущей работы.
    """
    cur: Optional[Dict[str, Any]] = None
    for it in items:
        kind = row_kind_of(it)
        if kind == "work" or cur is None:
            cur = {"it": it, "starts": True, "cur": it, "kind": kind, "lead": it}
            it["group_no"] = s(it.get("num"))
            it["group_code"] = s(it.get("code")) if kind == "work" else ""
            it["group_name"] = s(it.get("name"))
            it["group_lead"] = True
            it["group_size"] = 1
        else:
            it["group_no"] = cur["lead"].get("group_no", "")
            it["group_code"] = cur["lead"].get("group_code", "")
            it["group_name"] = cur["lead"].get("group_name", "")
            it["group_lead"] = False
            cur["lead"]["group_size"] = int(cur["lead"].get("group_size", 1)) + 1
    return items


def _read_totals_block(rows: List[List[Any]], start_idx: int, out: Dict[str, float],
                       cur: Dict[str, float], stop_at_vsego: bool = True) -> None:
    """Прочитать блок «Итоги по …»: строки C=подпись, L=базис, N=текущий.

    start_idx — 0-based индекс строки-маркера в списке rows. Подстроки вида
    «оплата труда», «эксплуатация машин», «материалы», «НР», «СП» внутри блоков
    «Строительные работы» / «Монтажные работы» сохраняются с суффиксом режима.
    """
    mode = ""   # '', 'build', 'install'
    i = start_idx + 1
    n = len(rows)
    while i < n:
        c = s(rows[i][2]).strip()
        if not c:
            i += 1
            continue
        low = c.lower().rstrip(":")
        plain = low.replace(" ", "")
        if plain.startswith("итогипосмете"):
            break
        if RE_SEC_TOTALS.match(low) or RE_SEC_VSEGO.match(low):
            break
        value = f(rows[i][11])
        value_cur = f(rows[i][13])
        key = None
        if plain.startswith("всегопрямыезатраты"):
            key = "pz"
        elif plain.startswith("оплататрударабочих"):
            key = "ozp"
        elif plain.startswith("эксплуатациямашин"):
            key = "emm"
        elif plain.startswith("втомчислеоплататруд"):
            key = "zpm"
        elif plain.startswith("материалы"):
            key = "material"
        elif plain.startswith("накладныерасходы"):
            key = "nr"
        elif plain.startswith("сметнаяприбыль"):
            key = "sp"
        elif plain.startswith("строительныеработы"):
            mode = "build"
            key = "build"
        elif plain.startswith("монтажныеработы"):
            mode = "install"
            key = "install"
        elif plain.startswith("оборудование") and "отсутствующее" not in plain:
            mode = ""
            key = "equipment"
        elif plain.startswith("прочиезатраты"):
            mode = ""
            key = "other"
        elif plain.startswith("всегофот"):
            key = "fot"
        elif plain.startswith("всегонатра"):
            key = "vsego_nr_sp"       # «Всего накладные расходы (справочно)»/«Всего сметная прибыль» см. ниже
            if "прибыл" in plain:
                key = "vsego_sp"
            else:
                key = "vsego_nr"
        elif plain.startswith("всего"):
            if "посмете" in plain:
                if plain.startswith("всегопо"):     # «Всего по смете» не бывает — есть «Всего» и «ВСЕГО по смете»
                    pass
                key = "vsego_po_smete" if plain.isupper() or "всегопосмете" == plain else None
                if plain == "всего":
                    key = "total"
                elif plain.startswith("всегопосмете"):
                    key = "vsego_po_smete"
            else:
                key = "total"
        elif plain.startswith("ндс"):
            key = "nds"
        elif plain.startswith("затратытрударабочих"):
            key = "tz_workers"
            value = f(rows[i][8])     # I — чел.-ч
            value_cur = value
        elif plain.startswith("затратытрудмаши"):
            key = "tz_mash"
            value = f(rows[i][8])
            value_cur = value
        elif plain.startswith("справочно"):
            mode = ""
        if key is None:
            i += 1
            continue
        full = key + ("_" + mode if mode and key in ("ozp", "emm", "zpm", "material", "nr", "sp") else "")
        out[full] = value
        cur[full] = value_cur
        if key == "vsego_po_smete" and plain.startswith("всего"):
            # «ВСЕГО по смете» — итог с НДС; базовую строку «Всего» не перекрываем
            pass
        i += 1


def parse_file(path: str) -> Dict[str, Any]:
    return parse_book(path)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Парсер печатной формы ГРАНД-Сметы")
    ap.add_argument("input", help="папка или файл *.xlsx")
    ap.add_argument("-o", "--out", default="", help="файл результата (JSON)")
    args = ap.parse_args(argv)
    paths: List[str] = []
    if os.path.isdir(args.input):
        for root, _d, files in os.walk(args.input):
            for fn in sorted(files):
                if fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                    paths.append(os.path.join(root, fn))
    else:
        paths = [args.input]
    books = []
    for p in paths:
        try:
            wp = parse_book(p)
        except Exception as exc:
            print("%s: НЕ РАЗОБРАНО — %s" % (os.path.basename(p), exc))
            continue
        ls = wp["nodes"][0]
        t = ls["totals"]
        print("%-46s поз.%4d разд.%2d ПЗ=%12.2f ОТ=%9.2f ЭМ=%9.2f М=%13.2f НР=%9.2f СП=%9.2f Всего=%12.2f ЗТ=%10.4f ЗТм=%8.4f" % (
            wp["file"][:46], len(wp["items"]), sum(1 for n in wp["nodes"] if n["kind"] == 4),
            t["pz"], t["ozp"], t["emm"], t["material"], t["nr"], t["sp"], t["total"],
            t["tz_workers"], t["tz_mash"]))
        books.append({k: v for k, v in wp.items() if k != "resources"} | {"n_resources": len(wp["resources"])})
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(books, fh, ensure_ascii=False, separators=(",", ":"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
