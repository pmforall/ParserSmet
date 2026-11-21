#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_multi_ls.py — генератор синтетической сметы АтомСмета, в которой ОДИН файл
содержит НЕСКОЛЬКО локальных смет (ЛС).

Зачем: в реальных папках встречаются книги, где на листе «Source» описано
несколько ЛС (несколько узлов вида «Смета (ЛС)» = kind 1). Такие книги должны
разбираться в несколько строк отчёта — по одной на каждую ЛС, каждая со своей
печатной формой «11 граф».

Файл формируется программно (обычные значения, без формул), чтобы оба независимых
разбора — питоновский (parser/smeta_parser.py) и браузерный (app/template.html)
— читали одни и те же числа.

Структура листа «Source» (нумерация столбцов 1-based, как в openpyxl):

    A(1)  тип записи: 1 шапка, 3/4/5 открытие узла, 17 позиция,
          50 показатель узла, 51 итоговая запись узла
    C(3)  вид узла: 1 ЛС, 3 документ, 4 раздел
    D(4)  идентификатор узла (для A=51/52) либо ссылка на строку итога (A=3/4/5)
    E(5)  номер позиции
    F(6)  шифр / код
    G(7)  наименование
    K(11) код показателя (для A=50), W(23) — значение показателя
    O(15) ПЗ, P(16) мат.ресурсы, Q(17) ЭММ, R(18) ЗПМ, S(19) ОЗП,
    U(21) ТЗ строителей, V(22) ТЗ машинистов, X(24) НР, Y(25) СП
"""
from __future__ import annotations

import os
from typing import Any, Dict, List

from openpyxl import Workbook

# --- столбцы листа «Source» -------------------------------------------------
C_TYPE, C_KIND, C_REF, C_NUM, C_CODE, C_NAME = 1, 3, 4, 5, 6, 7
C_UNIT, C_QTY, C_OSN, C_IND, C_INDVAL = 8, 9, 10, 11, 23
C_PZ, C_MAT, C_EMM, C_ZPM, C_OZP = 15, 16, 17, 18, 19
C_TZS, C_TZM, C_NR, C_SP, C_TOTAL = 21, 22, 24, 25, 195
C_BASIS, C_DV, C_GROUP, C_KINDF, C_VID, C_VIDNAME = 62, 126, 137, 140, 141, 142
C_DISP, C_M, C_SHORT = 8, 13, 14          # отображаемое имя (H), признак (M), краткое имя (N)
C_PARENT = 3                              # номер позиции-владельца для строки ресурса

MARK = "11 граф"
K_USP = 1.15

# показатели узлов (код -> значение-ключ)
IND_CODES = [(201, "pz"), (202, "mat_ob"), (203, "emm"), (204, "zpm"), (205, "ozp"),
             (207, "tz_workers"), (208, "tz_mash"), (210, "nr"), (211, "sp"), (224, "total"),
             (226, "mat_ob"), (216, "equipment"), (233, "transport"), (217, "other")]


class Rows:
    """Набор строк листа с отложенными ссылками на строки итогов."""

    def __init__(self) -> None:
        self.rows: List[Dict[int, Any]] = []
        self.refs: List[tuple] = []          # (номер строки, столбец, id узла)
        self.footer: Dict[int, int] = {}     # id узла -> номер строки итога (A=51)

    def add(self, cells: Dict[int, Any]) -> int:
        self.rows.append(dict(cells))
        return len(self.rows)

    def ref(self, cells: Dict[int, Any], node: int) -> int:
        r = self.add(cells)
        self.refs.append((r, C_REF, node))
        return r

    def footer_at(self, cells: Dict[int, Any], node: int, kind: int) -> int:
        r = self.add(cells)
        self.footer[node] = r
        return r

    def resolve(self) -> List[Dict[int, Any]]:
        for r, col, node in self.refs:
            self.rows[r - 1][col] = self.footer.get(node)
        return self.rows


# --- описание данных --------------------------------------------------------

# основания позиций с кодами машин/механизмов — для отчёта по справочнику «ТЕХНИКА и ТРЗ»
BASIS_BY_NUM = {3: "ФЕРм 91.21.22-589, приказ Минстроя России"}

# шифры с пометками коэффициентов, как их пишет АтомСмета: должны нормализоваться
# при разборе («13-03-004-26. к=2» → «13-03-004-26»)
CODE_BY_NUM = {2: "13-03-004-26. к=2", 11: "м08-01-067-08.к=1,5", 13: "13-03-004-26 ×2"}
CODE_EXPECT = {2: "13-03-004-26", 11: "м08-01-067-08", 13: "13-03-004-26"}


def _item(num: int, name: str, qty: float, vals: Dict[str, float]) -> Dict[int, Any]:
    return {
        C_TYPE: 17, C_NUM: num,
        C_CODE: CODE_BY_NUM.get(num, "ФЕР-01-01-001-%02d" % num), C_NAME: name,
        C_UNIT: "100 м3", C_QTY: qty,
        C_BASIS: BASIS_BY_NUM.get(num, "ФЕР 01-01-001-%02d" % num), C_DV: "100 м3",
        C_GROUP: "Работы", C_KINDF: 0, C_VID: 9999, C_VIDNAME: "Работы строительные",
        C_PZ: vals["pz"], C_MAT: vals["mat_ob"], C_EMM: vals["emm"], C_ZPM: vals["zpm"],
        C_OZP: vals["ozp"], C_TZS: vals["tz_workers"], C_TZM: vals["tz_mash"],
        C_NR: vals["nr"], C_SP: vals["sp"], C_TOTAL: vals["total"],
    }


def _v(pz: float, mat: float, emm: float, zpm: float, ozp: float,
       tzs: float, tzm: float, nr: float, sp: float) -> Dict[str, float]:
    return {"pz": pz, "mat_ob": mat, "emm": emm, "zpm": zpm, "ozp": ozp,
            "tz_workers": tzs, "tz_mash": tzm, "nr": nr, "sp": sp, "total": pz + nr + sp}


def _sum(a: Dict[str, float], b: Dict[str, float]) -> Dict[str, float]:
    return {k: round(a.get(k, 0.0) + b.get(k, 0.0), 4) for k in
            ("pz", "mat_ob", "emm", "zpm", "ozp", "tz_workers", "tz_mash", "nr", "sp", "total")}


# позиции ЛС № 1: три позиции в разделе + одна на уровне сметы
SEC1 = [_v(1000, 600, 100, 20, 300, 2.0, 0.20, 240, 150),
        _v(500, 200, 50, 10, 250, 1.5, 0.10, 200, 120),
        _v(500, 100, 50, 10, 350, 2.0, 0.15, 280, 170)]
SEC1_T = _sum(_sum(SEC1[0], SEC1[1]), SEC1[2])
LS1_FREE = _v(500, 200, 50, 10, 250, 1.5, 0.10, 200, 120)
LS1_T = _sum(SEC1_T, LS1_FREE)

# позиции ЛС № 2: две позиции в разделе + одна на уровне сметы
SEC2 = [_v(300, 100, 20, 5, 180, 1.0, 0.05, 144, 90),
        _v(200, 50, 30, 5, 120, 0.5, 0.02, 96, 60)]
SEC2_T = _sum(SEC2[0], SEC2[1])
LS2_FREE = _v(300, 100, 20, 5, 180, 1.0, 0.05, 144, 90)
LS2_T = _sum(SEC2_T, LS2_FREE)
DOC_T = _sum(LS1_T, LS2_T)

# позиции ЛС в порядке листа «Source»: (номер, значения, количество)
LS1_ITEMS = [(1, SEC1[0], 1.0), (2, SEC1[1], 1.5), (3, SEC1[2], 2.0), (4, LS1_FREE, 2.0)]
LS2_ITEMS = [(11, SEC2[0], 1.0), (12, SEC2[1], 1.5), (13, LS2_FREE, 2.0)]

LS1 = {"id": 12, "code": "01-01-01-01", "sec_id": 21,
       "name": "ЛС № 1. Монолитные конструкции", "sec_name": "Раздел 1. Стены"}
LS2 = {"id": 13, "code": "01-01-01-02", "sec_id": 22,
       "name": "ЛС № 2. Гидроизоляция", "sec_name": "Раздел 2. Изоляция"}
DOC = {"id": 20, "code": "RCNS.TEST.0001",
       "name": "Дополнительная спецификация (синтетическая книга с двумя ЛС)"}

# ресурсы (машины и механизмы) по сметам — с кодами из справочника «ТЕХНИКА и ТРЗ»
RESOURCES = {
    "01-01-01-01": [{"parent": 4, "code": "010311", "qty": 8.0, "cost": 12000.0,
                     "name": "Тракторы на гусеничном ходу при работе на других видах строительства"}],
    "01-01-01-02": [{"parent": 13, "code": "ТСЭМ-110054", "qty": 3.5, "cost": 4500.0,
                     "name": "Машины бурильные самоходные"}],
}

HDR = {C_TYPE: 1, C_CODE: DOC["code"], C_NAME: DOC["name"], C_OSN: "Приказ № 1 от 01.01.2024",
       C_UNIT: "ФЕР-2001", C_INDVAL: "Объект: учебный корпус", 28: "Инженер",
       29: "Иванов И.И.", 30: "Проверил", 31: "Петров П.П.", 78: "ТЕР-2001", 83: "нет"}


def _source_rows() -> List[Dict[int, Any]]:
    r = Rows()
    r.add(HDR)
    r.ref({C_TYPE: 3}, DOC["id"])                       # структура документа
    r.add({C_TYPE: 52, C_KIND: 3, C_REF: DOC["id"], C_NAME: DOC["name"]})  # начало документа

    def one_ls(ls: Dict[str, Any], sec_t: Dict[str, float], ls_t: Dict[str, float],
               items: List[tuple], n_sec: int) -> None:
        r.add({C_TYPE: 52, C_KIND: 1, C_REF: ls["id"], C_NAME: ls["name"]})   # начало ЛС
        r.ref({C_TYPE: 4}, ls["sec_id"])                                       # открытие раздела
        r.add({C_TYPE: 52, C_KIND: 4, C_REF: ls["sec_id"], C_NAME: ls["sec_name"]})
        for num, vals, qty in items[:n_sec]:
            r.add(_item(num, "Работа %d в разделе" % num, qty, vals))
        r.footer_at({C_TYPE: 51, C_KIND: 4, C_REF: ls["sec_id"], C_CODE: "Раздел",
                     C_NAME: ls["sec_name"], C_PZ: sec_t["pz"], C_MAT: sec_t["mat_ob"],
                     C_EMM: sec_t["emm"], C_ZPM: sec_t["zpm"], C_OZP: sec_t["ozp"],
                     C_TZS: sec_t["tz_workers"], C_TZM: sec_t["tz_mash"], C_NR: sec_t["nr"],
                     C_SP: sec_t["sp"]}, ls["sec_id"], 4)
        _indicators(r, ls["sec_id"], sec_t)
        r.ref({C_TYPE: 3}, ls["id"])                                          # открытие самой ЛС
        for num, vals, qty in items[n_sec:]:
            r.add(_item(num, "Работа вне разделов", qty, vals))
        for res in RESOURCES.get(ls["code"], []):
            _resource(r, res["parent"], res["code"], res["name"], res["qty"], res["cost"])
        _summary_rows(r, ls_t, round(ls_t["total"] + (K_USP - 1) *
                                     (ls_t["ozp"] + ls_t["emm"] + ls_t["nr"] + ls_t["sp"]), 2))
        r.footer_at({C_TYPE: 51, C_KIND: 1, C_REF: ls["id"], C_CODE: ls["code"],
                     C_NAME: ls["name"], C_PZ: ls_t["pz"], C_MAT: ls_t["mat_ob"],
                     C_EMM: ls_t["emm"], C_ZPM: ls_t["zpm"], C_OZP: ls_t["ozp"],
                     C_TZS: ls_t["tz_workers"], C_TZM: ls_t["tz_mash"], C_NR: ls_t["nr"],
                     C_SP: ls_t["sp"]}, ls["id"], 1)
        _indicators(r, ls["id"], ls_t)

    one_ls(LS1, SEC1_T, LS1_T, LS1_ITEMS, len(SEC1))
    one_ls(LS2, SEC2_T, LS2_T, LS2_ITEMS, len(SEC2))

    r.footer_at({C_TYPE: 51, C_KIND: 3, C_REF: DOC["id"], C_CODE: DOC["code"],
                 C_NAME: DOC["name"], C_PZ: DOC_T["pz"], C_MAT: DOC_T["mat_ob"],
                 C_EMM: DOC_T["emm"], C_ZPM: DOC_T["zpm"], C_OZP: DOC_T["ozp"],
                 C_TZS: DOC_T["tz_workers"], C_TZM: DOC_T["tz_mash"], C_NR: DOC_T["nr"],
                 C_SP: DOC_T["sp"]}, DOC["id"], 3)
    _indicators(r, DOC["id"], DOC_T)
    return r.resolve()


def _resource(r: Rows, parent: int, code: str, name: str, qty: float, cost: float) -> None:
    """Строка ресурса (A=18): машина/механизм с кодом из справочника."""
    r.add({C_TYPE: 18, C_KIND: 0, C_REF: 2, C_NUM: None, C_CODE: code, C_NAME: name,
           C_UNIT: "маш.-ч", C_DV: "маш.-ч", C_QTY: qty, C_PZ: cost, C_TOTAL: cost,
           C_PARENT: parent, C_GROUP: "Строительные машины", C_KINDF: 2,
           C_VID: 3001, C_VIDNAME: "Эксплуатация машин"})


def _summary_rows(r: Rows, t: Dict[str, float], total_k: float) -> None:
    """Сводные строки вкладки «Source» (D=2): базисный итог и «ВСЕГО ПО СМЕТЕ»."""
    r.add({C_TYPE: 50, C_KIND: 0, C_REF: 2, C_NUM: 0, C_IND: 212, C_NAME: "ВСЕГО_ПЗ",
           C_DISP: "ВСЕГО ПО СМЕТЕ (без стесненности с учетом НР и СП)",
           C_SHORT: "Всего по Смете", C_M: 3, C_CODE: t["total"], C_INDVAL: t["total"]})
    r.add({C_TYPE: 50, C_KIND: 0, C_REF: 2, C_NUM: 213, C_IND: 212,
           C_NAME: "КСт_Всего ПО СМЕТЕ _КСт", C_DISP: "ВСЕГО ПО СМЕТЕ", C_M: 1,
           C_CODE: total_k, C_INDVAL: total_k})


def _indicators(r: Rows, node: int, t: Dict[str, float]) -> None:
    for code, key in IND_CODES:
        r.add({C_TYPE: 50, C_KIND: 0, C_REF: 1, C_NUM: code, C_IND: code,
               C_NAME: key, C_INDVAL: t.get(key, 0.0)})


# --- печатная форма «11 граф» ----------------------------------------------

def _printed(ws, ls: Dict[str, Any], t: Dict[str, float], sec_t: Dict[str, float],
             sec_name: str, items: List[tuple]) -> None:
    """Заполнить лист печатной формы так, как его читает разбор (см. printed_form.py)."""
    k = K_USP
    total_k = round(t["total"] + (k - 1) * (t["ozp"] + t["emm"] + t["nr"] + t["sp"]), 2)
    norm_labor = round((t["tz_workers"] + t["tz_mash"]) * k, 4)
    smeta_wage = round((t["ozp"] + t["zpm"]) * k, 2)

    ws.cell(row=4, column=1).value = ls["code"]                  # шифр (doc_code)
    ws.cell(row=5, column=1).value = "ЛОКАЛЬНАЯ СМЕТА № %s" % ls["code"]
    ws.cell(row=6, column=1).value = ls["name"]
    ws.cell(row=8, column=1).value = "Основание:"
    ws.cell(row=8, column=3).value = "Приказ № 1 от 01.01.2024"
    ws.cell(row=10, column=1).value = "Сметная стоимость"
    ws.cell(row=10, column=9).value = round(total_k / 1000.0, 6)  # тыс. руб.
    ws.cell(row=12, column=1).value = "Строительных работ"
    ws.cell(row=12, column=9).value = round(total_k / 1000.0, 6)
    ws.cell(row=18, column=1).value = "Нормативная трудоемкость"
    ws.cell(row=18, column=9).value = round(norm_labor / 1000.0, 6)
    ws.cell(row=19, column=1).value = "Сметная заработная плата"
    ws.cell(row=19, column=9).value = round(smeta_wage / 1000.0, 6)
    ws.cell(row=25, column=1).value = "Составлена в базисных ценах 01.01.2000 по НБ: ФЕР-2001"

    # блок раздела: итог раздела и НР/СП/материалы внутри блока
    ws.cell(row=40, column=1).value = "Итого по разделу (без НР и СП) %s" % sec_name
    ws.cell(row=40, column=7).value = sec_t["pz"]
    ws.cell(row=40, column=8).value = sec_t["ozp"]
    ws.cell(row=40, column=9).value = sec_t["emm"]
    ws.cell(row=40, column=11).value = sec_t["tz_workers"]
    ws.cell(row=41, column=1).value = "Накладные расходы 80 %"
    ws.cell(row=41, column=7).value = sec_t["nr"]
    ws.cell(row=42, column=1).value = "Сметная прибыль 60 %"
    ws.cell(row=42, column=7).value = sec_t["sp"]
    ws.cell(row=43, column=1).value = "в т.ч. Стоимость материалов"
    ws.cell(row=43, column=7).value = sec_t["mat_ob"]

    # позиции формы — все позиции сметы в порядке листа «Source»
    row = 50
    for num, vals, qty in items:
        ws.cell(row=row, column=1).value = str(num)
        ws.cell(row=row, column=2).value = "ФЕР 01-01-001-%02d" % num
        ws.cell(row=row, column=3).value = "Работа %s" % num
        ws.cell(row=row, column=4).value = qty
        ws.cell(row=row, column=5).value = vals["pz"]
        ws.cell(row=row, column=7).value = vals["pz"]
        ws.cell(row=row, column=8).value = vals["ozp"]
        ws.cell(row=row, column=9).value = vals["emm"]
        ws.cell(row=row, column=11).value = vals["tz_workers"]
        ws.cell(row=row, column=24).value = vals["zpm"]
        ws.cell(row=row, column=27).value = vals["nr"]
        ws.cell(row=row, column=28).value = vals["sp"]
        row += 1

    ws.cell(row=60, column=1).value = "Прямые затраты"
    ws.cell(row=60, column=7).value = t["pz"]
    ws.cell(row=60, column=8).value = t["ozp"]
    ws.cell(row=60, column=9).value = t["emm"]
    ws.cell(row=60, column=11).value = t["tz_workers"]
    ws.cell(row=61, column=9).value = t["zpm"]
    ws.cell(row=61, column=11).value = t["tz_mash"]
    ws.cell(row=62, column=1).value = "Всего, Накладные расходы"
    ws.cell(row=62, column=7).value = t["nr"]
    ws.cell(row=63, column=1).value = "Всего, Сметная прибыль"
    ws.cell(row=63, column=7).value = t["sp"]
    ws.cell(row=64, column=1).value = "в т.ч. Стоимость материалов"
    ws.cell(row=64, column=7).value = t["mat_ob"]
    ws.cell(row=65, column=1).value = "Трудозатраты строителей"
    ws.cell(row=65, column=7).value = t["tz_workers"]
    ws.cell(row=66, column=1).value = "Трудозатраты машинистов"
    ws.cell(row=66, column=7).value = t["tz_mash"]
    ws.cell(row=67, column=1).value = "Всего по смете"
    ws.cell(row=67, column=7).value = t["total"]
    ws.cell(row=67, column=8).value = t["ozp"]
    ws.cell(row=67, column=9).value = t["emm"]
    ws.cell(row=67, column=11).value = t["tz_workers"]
    ws.cell(row=68, column=9).value = t["zpm"]
    ws.cell(row=68, column=11).value = t["tz_mash"]
    ws.cell(row=69, column=1).value = "ВСЕГО ПО СМЕТЕ"
    ws.cell(row=69, column=7).value = total_k
    ws.cell(row=70, column=1).value = "Коэфф. согласно пр. Минстроя РФ № 421/пр"
    ws.cell(row=70, column=7).value = k


def build(path: str, printed_count: int = 2) -> str:
    """Создать книгу с двумя ЛС. printed_count — сколько листов «11 граф» добавить."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Source"
    for i, cells in enumerate(_source_rows(), start=1):
        if cells[C_TYPE] not in (0, 1):
            ws.cell(row=i, column=2).value = cells.get(2, 1)
        for col, val in cells.items():
            if val is not None:
                ws.cell(row=i, column=col).value = val

    sheets = [(LS1, LS1_T, SEC1_T, LS1["sec_name"], LS1_ITEMS),
              (LS2, LS2_T, SEC2_T, LS2["sec_name"], LS2_ITEMS)]
    for n, (ls, t, sec_t, sec_name, items) in enumerate(sheets[:max(1, printed_count)]):
        pws = wb.create_sheet("%s %s (%s)" % (MARK, ls["code"], ls["name"][:28]))
        _printed(pws, ls, t, sec_t, sec_name, items)

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


EXPECTED = {
    "smetas": 2,
    "numbers": [LS1["code"], LS2["code"]],
    "items": [len(LS1_ITEMS), len(LS2_ITEMS)],
    "codes": {str(k): v for k, v in CODE_EXPECT.items()},
    "totals": {LS1["code"]: LS1_T, LS2["code"]: LS2_T},
    "doc_total": DOC_T,
}


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "out", "multi_ls", "Две_сметы.xlsx")
    print(build(out))
