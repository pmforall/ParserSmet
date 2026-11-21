#!/usr/bin/env python3
"""Сопоставление позиций, границы ЛС и паритет браузерного алгоритма с Python."""
from copy import deepcopy
import json
from pathlib import Path
import re
import subprocess
import sys
import unittest

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "parser"))
import printed_positions as pp


def item(row, num="1", code="08-07-002-01", name="Работа", total=100):
    return dict(row=row, num=num, code=code, name=name,
                values=dict(total=total, ozp=20, emm=10, nr=30, sp=10))


def smeta(items, number="LS-1"):
    return dict(number=number, name=number, file="input.xlsx", items=items, checks=[],
                source_total=dict(value=180), source_work_coefficient=1.13)


def printed(positions, number="", sheet="ЛС 11 граф", name=""):
    return dict(sheet=sheet, number=number, name=name,
                position_rows=[dict(row=i, sheet=sheet, num=str(n), basis=b, name=t)
                               for i, (n, b, t) in enumerate(positions, 31)])


def selection(smetas, sheets):
    result = deepcopy(smetas)
    pp.apply(result, sheets, [p["sheet"] for p in sheets])
    return result


class PrintedPositionTests(unittest.TestCase):
    def run_both(self, smetas, sheets):
        expected = selection(smetas, sheets)
        html = (ROOT / "app/template.html").read_text()
        module = re.search(r"const PrintedPositions = \(function\(\)\{.*?\n\}\)\(\);", html, re.S).group()
        js = module + "\nconst smetas=" + json.dumps(smetas, ensure_ascii=False) + ";\n"
        js += "const printed=" + json.dumps(sheets, ensure_ascii=False) + ";\n"
        js += "PrintedPositions.apply(smetas,printed,printed.map(p=>p.sheet));process.stdout.write(JSON.stringify(smetas));"
        actual = json.loads(subprocess.run(["node"], input=js, text=True, capture_output=True, check=True).stdout)
        self.assertEqual([s["position_selection"] for s in expected], [s["position_selection"] for s in actual])
        # Формат чисел в тексте различается, коды/ссылки/отклонения должны совпадать.
        for left, right in zip(expected, actual):
            fields = ("ref", "code", "severity", "delta")
            self.assertEqual([{k: c[k] for k in fields} for c in left["checks"]],
                             [{k: c[k] for k in fields} for c in right["checks"]])
        return expected

    def test_subset_source_values_and_coefficient(self):
        source = [smeta([item(4), item(8, "2")])]
        sheets = [printed([("1", "ФЕР 08-07-002-01", "Работа")])]
        result = self.run_both(source, sheets)[0]
        self.assertEqual(result["position_selection"]["matches"][0]["source_row"], 4)
        self.assertEqual(result["position_selection"]["excluded_rows"], [8])
        self.assertEqual(result["position_selection"]["total"], 109.1)
        self.assertEqual(result["items"], source[0]["items"])
        self.assertEqual(result["source_total"], source[0]["source_total"])
        self.assertIn("PRINTED_SELECTION_TOTAL", [c["code"] for c in result["checks"]])

    def test_code_conflict_does_not_match_by_number_or_name(self):
        result = self.run_both([smeta([item(4)])], [printed([("1", "ФЕР 46-02-005-04", "Работа")])])
        self.assertEqual(result[0]["position_selection"]["matches"], [])

    def test_corrections_numeric_codes_and_name_without_code(self):
        items = [item(4, "13,1", "01.7.16.02-0001"), item(5, "14", "", "Оборудование по прайсу")]
        rows = [("13.1", "ФССЦ-2001, 01.7.16.02-0001, приказ № 876/пр", "Работа"),
                ("14", 'К/п АО «Прогресс» от 28.06.2021', "Оборудование по прайсу")]
        result = self.run_both([smeta(items)], [printed(rows)])
        self.assertEqual(len(result[0]["position_selection"]["matches"]), 2)

    def test_mounting_and_transport_prefix_and_printed_unit(self):
        short = item(7, "3", "", "Труба Корсис АРМ 800 SN12")
        short["unit"] = "м"
        result = self.run_both([smeta([item(4, "1", "м08-01-062-02"),
                                      item(5, "2", "т03-02-01-030"), short])],
                              [printed([("1", "ФЕРм 08-01-062-02", "Работа"),
                                        ("2", "ФССЦ 03-02-01-030", "Работа"),
                                        ("3", "К/П от 01.04.2021", "Труба Корсис АРМ 800 SN12, м")])])
        self.assertEqual(len(result[0]["position_selection"]["matches"]), 3)

    def test_multiple_ls_reordered_sheets(self):
        source = [smeta([item(4)], "LS-1"), smeta([item(9)], "LS-2")]
        sheets = [printed([("1", "08-07-002-01", "Работа")], number="LS-2", sheet="17 ГРАФ LS-2")]
        result = self.run_both(source, sheets)
        self.assertEqual([len(s["position_selection"]["matches"]) for s in result], [0, 1])

    def test_unknown_ls_uses_identity_without_order_fallback(self):
        source = [smeta([item(4)], "LS-1"), smeta([item(9, code="46-02-005-04")], "LS-2")]
        sheets = [printed([("1", "46-02-005-04", "Работа")], number="OTHER")]
        result = self.run_both(source, sheets)
        self.assertEqual([len(s["position_selection"]["matches"]) for s in result], [0, 1])

    def test_identical_position_in_different_ls_is_ambiguous(self):
        result = self.run_both([smeta([item(4)]), smeta([item(9)], "LS-2")],
                               [printed([("1", "08-07-002-01", "Работа")])])
        self.assertEqual([s["position_selection"]["ambiguous_rows"] for s in result], [[4], [9]])
        self.assertTrue(all(not s["position_selection"]["matches"] for s in result))

    def test_repeated_number_with_other_code(self):
        result = self.run_both([smeta([item(4), item(5, code="46-02-005-04")])],
                              [printed([("1", "46-02-005-04", "Работа")])])
        self.assertEqual(result[0]["position_selection"]["matches"][0]["source_row"], 5)

    def test_duplicate_rows_and_copies_of_form(self):
        source = [smeta([item(4), item(5)])]
        one = printed([("1", "08-07-002-01", "Работа")])
        copy = printed([("1", "08-07-002-01", "Работа")], sheet="Копия 11 граф")
        result = self.run_both(source, [one, copy])
        self.assertEqual(len(result[0]["position_selection"]["matches"]), 1)
        two = printed([("1", "08-07-002-01", "Работа")] * 2)
        result = self.run_both(source, [two])
        self.assertEqual(len(result[0]["position_selection"]["matches"]), 2)

    def test_missing_form(self):
        result = self.run_both([smeta([item(4)])], [])
        self.assertEqual(result[0]["position_selection"]["matches"], [])
        self.assertIn("POSITIONS_NO_GRAPH", [c["code"] for c in result[0]["checks"]])

    def test_membership_reader_has_no_5000_row_limit_and_skips_column_numbers(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Печатная форма ГРАФЫ"
        ws.append([1, 2, 3])
        for n in range(1, 5003):
            ws.append([n, "01.7.16.02-0001", "Материал"])
        positions = pp.read(ws)
        self.assertEqual(len(positions), 5002)
        self.assertEqual(positions[-1]["row"], 5003)
        rows = [None] + [[None] + list(r) for r in ws.values]
        html = (ROOT / "app/template.html").read_text()
        module = re.search(r"const PrintedPositions = \(function\(\)\{.*?\n\}\)\(\);", html, re.S).group()
        js = module + "\nprocess.stdout.write(JSON.stringify(PrintedPositions.read(" + json.dumps(rows) + "," + json.dumps(ws.title) + ")));"
        actual = json.loads(subprocess.run(["node"], input=js, text=True, capture_output=True, check=True).stdout)
        self.assertEqual(actual, positions)
        wb.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
