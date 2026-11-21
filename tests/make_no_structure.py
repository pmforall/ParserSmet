#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_no_structure.py — генератор синтетической книги АтомСмета, в которой на листе
«Source» есть заголовок и позиции, но НЕТ записей структуры (A=3/4/5 и A=51/52).

Зачем: сводная строка сметы строится по узлам структуры. Если в книге их нет (модель
сохранена без структуры, выгрузка повреждена и т. п.), раньше книга не давала ни одной
сметы — «сводные данные» пропадали вместе с позициями. Теперь разбор собирает одну смету
по заголовку и позициям и помечает файл предупреждением NO_STRUCTURE.

Структура листа «Source» (нумерация столбцов 1-based, как в openpyxl):
    A(1) тип записи: 1 — шапка, 17 — позиция, 50 — показатели (здесь не нужны)
    E(5) номер позиции, F(6) шифр, G(7) наименование, H(8)/DV(126) единицы, I(9) количество
    O(15) ПЗ, P(16) материалы, Q(17) ЭММ, R(18) ЗПМ, S(19) ОЗП, U(21)/V(22) трудозатраты,
    X(24) НР, Y(25) СП, GM(195) всего с НР и СП
"""
from __future__ import annotations

import os
from typing import Any, Dict, List

from openpyxl import Workbook

C_TYPE, C_KIND, C_NUM, C_CODE, C_NAME = 1, 3, 5, 6, 7
C_UNIT, C_QTY, C_OSN = 8, 9, 10
C_PZ, C_MAT, C_EMM, C_ZPM, C_OZP = 15, 16, 17, 18, 19
C_TZS, C_TZM, C_NR, C_SP, C_TOTAL = 21, 22, 24, 25, 195
C_BASIS, C_DV, C_GROUP, C_KINDF, C_VID, C_VIDNAME = 62, 126, 137, 140, 141, 142

HEADER = {"code": "01-02-03-04", "name": "ЛС без структуры (синтетическая книга)",
          "osnovanie": "том 1 инв.№ Т-0001", "object": "Объект синтетический",
          "norm_base": "ФЕР-2001 (редакция 2020 г.)"}

# (номер позиции, шифр, наименование, количество, значения)
ITEMS: List[Dict[str, Any]] = [
    {"num": 1, "code": "ФЕР-01-01-001-01", "name": "Разработка грунта", "qty": 10.0,
     "pz": 1000.0, "mat": 0.0, "emm": 200.0, "zpm": 40.0, "ozp": 800.0,
     "tzs": 4.0, "tzm": 0.4, "nr": 748.0, "sp": 496.0},
    {"num": 2, "code": "ФЕР-01-01-001-02", "name": "Устройство основания", "qty": 5.5,
     "pz": 2500.0, "mat": 1800.0, "emm": 100.0, "zpm": 20.0, "ozp": 600.0,
     "tzs": 3.0, "tzm": 0.2, "nr": 533.0, "sp": 354.0},
    {"num": 3, "code": "13-03-004-26. к=2", "name": "Окраска металлических поверхностей",
     "qty": 2.25, "pz": 700.0, "mat": 100.0, "emm": 50.0, "zpm": 10.0, "ozp": 550.0,
     "tzs": 2.5, "tzm": 0.1, "nr": 440.0, "sp": 292.0},
]


def totals() -> Dict[str, float]:
    """Итоги сметы: сумма позиций (то, что должен собрать разбор без структуры)."""
    keys = ("pz", "mat", "emm", "zpm", "ozp", "tzs", "tzm", "nr", "sp")
    acc = {k: round(sum(float(it[k]) for it in ITEMS), 4) for k in keys}
    acc["total"] = round(acc["pz"] + acc["nr"] + acc["sp"], 2)
    return acc


EXPECTED = {
    "smetas": 1,
    "number": HEADER["code"],
    "items": len(ITEMS),
    "totals": totals(),
    "warning": "NO_STRUCTURE",
}


def _source_rows() -> List[Dict[int, Any]]:
    rows: List[Dict[int, Any]] = []
    rows.append({C_TYPE: 0, 2: "АтомСмета"})
    rows.append({C_TYPE: 1, C_CODE: HEADER["code"], C_NAME: HEADER["name"],
                 C_OSN: HEADER["osnovanie"], 21: HEADER["norm_base"], 23: HEADER["object"]})
    for it in ITEMS:
        rows.append({
            C_TYPE: 17, C_NUM: it["num"], C_CODE: it["code"], C_NAME: it["name"],
            C_UNIT: "100 м3", C_QTY: it["qty"], C_OSN: "ФЕР 01-01-001", C_DV: "100 м3",
            C_GROUP: "Работы", C_KINDF: 0, C_VID: 9999, C_VIDNAME: "Работы строительные",
            C_PZ: it["pz"], C_MAT: it["mat"], C_EMM: it["emm"], C_ZPM: it["zpm"],
            C_OZP: it["ozp"], C_TZS: it["tzs"], C_TZM: it["tzm"],
            C_NR: it["nr"], C_SP: it["sp"], C_TOTAL: it["pz"] + it["nr"] + it["sp"],
        })
    return rows


def build(path: str) -> str:
    """Создать книгу без записей структуры (A=51/52) — только шапка и позиции."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Source"
    for i, cells in enumerate(_source_rows(), start=1):
        if cells[C_TYPE] not in (0, 1):
            ws.cell(row=i, column=2).value = cells.get(2, 1)
        for col, val in cells.items():
            if val is not None:
                ws.cell(row=i, column=col).value = val
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    print(build(os.path.join(os.path.dirname(os.path.abspath(__file__)), "out",
                             "Без_структуры.xlsx")))
