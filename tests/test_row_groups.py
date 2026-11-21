#!/usr/bin/env python3
"""Укрупнённые группы строк сметы («работа + её материалы») и их учёт в отчёте по ФО.

Проверяется:
  * определение шифра сметной нормы (ФЕР/ТЕР/ГЭСН) и его отличие от кодов материалов
    и машин («14.4.01.20-0012», «91.14.01-011», прайс-лист «Расчёт 1231Р»);
  * разбиение позиций на группы по порядку следования строк вкладки «Source»:
    номер и шифр нормы ведущей строки проставляются всем строкам группы;
  * материалы, у которых нет расценки в справочнике, входят в укрупнённый объём
    своей работы — их стоимость учитывается в группе справочника, а объём группы
    остаётся объёмом работы;
  * паритет браузерного модуля и автономного парсера (одни и те же поля, те же итоги).
"""
import json
from pathlib import Path
import re
import subprocess
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "parser"))
import reference as ref
import smeta_parser as sp


# ---------------------------------------------------------------------------
# вспомогательные построители
# ---------------------------------------------------------------------------
def item(row, num, code="", name=None, unit="100 м2", qty=1.0, total=100.0, cat="works",
         basis=""):
    return dict(row=row, num=num, code=code, name=name or ("Позиция " + str(num)),
                basis=basis, unit=unit, qty=qty, cat=cat,
                values=dict(total=total, pz=total * 0.6, nr=total * 0.25, sp=total * 0.15,
                            ozp=total * 0.2, emm=total * 0.1, tz_workers=1.0, tz_mash=0.5))


def smeta(items, number="ЛС-1"):
    return dict(number=number, name=number, file="input.xlsx", items=items,
                total=sum(i["values"]["total"] for i in items),
                total_base=sum(i["values"]["total"] for i in items))


def fo_row(code, group, unit, primavera="CEWFV.R-1", name="Работа"):
    return [code, name, None, primavera, group, unit]


def mult_row(unit_in, k, unit_out):
    return [unit_in, k, unit_out]


def make_ref(fo_rows=None, mult_rows=None):
    """Справочник из двух вкладок: «Расценки ФО» и «Множитель»."""
    fo_rows = fo_rows if fo_rows is not None else [
        fo_row("м08-01-062-02", "Электротехническое оборудование, шт.", "шт"),
        fo_row("м08-02-147-01", "Кабель силовой, м", "м"),
        fo_row("13-03-004-26", "Окраска, м2", "м2"),
        fo_row("т03-21-01-030", "Перевозка", "т"),
    ]
    mult_rows = mult_rows if mult_rows is not None else [
        mult_row("100 м2", 100, "м2"), mult_row("м2", 1, "м2"),
        mult_row("шт", 1, "шт"), mult_row("м", 1, "м"), mult_row("т", 1, "т"),
        mult_row("кг", 0.001, "т"),
    ]
    head_fo = ["Расценка", "Наименование", "Тип ресурса", "Код ресурса Primavera",
               "Для ВПР", "ЕИ Примы"]
    head_mult = ["В смете", "множитель", "ЕИ"]
    return ref.parse([head_fo] + fo_rows, [head_mult] + mult_rows, [])


def js_group(items):
    """Разбиение тех же позиций браузерным модулем (SourceParser.groupItems)."""
    html = (ROOT / "app/template.html").read_text()
    mod = re.search(r"const SourceParser = \(function\(\)\{.*?\n\}\)\(\);", html, re.S).group()
    # вырезаем определения шифра нормы и группировки, чтобы не тянуть весь модуль
    helpers = "\n".join([
        re.search(r"  const NORM_RE = .*?;\n", mod, re.S).group(),
        re.search(r"  function isNormCode\(code\)\{.*?\n  \}\n", mod, re.S).group(),
        re.search(r"  function rowKindOf\(it\)\{.*?\n  \}\n", mod, re.S).group(),
        re.search(r"  function groupItems\(items\)\{.*?\n  \}\n", mod, re.S).group(),
    ])
    js = ("function str(v){ return v===null||v===undefined? '' : String(v).trim(); }\n"
          + helpers
          + "\nconst items=" + json.dumps(items, ensure_ascii=False) + ";\n"
          + "groupItems(items);"
          + "process.stdout.write(JSON.stringify(items));")
    out = subprocess.run(["node"], input=js, text=True, capture_output=True, check=True)
    return json.loads(out.stdout)


GROUP_FIELDS = ("group_no", "group_code", "group_name", "group_lead", "group_size")


def pick(items):
    return [{k: it.get(k) for k in GROUP_FIELDS} for it in items]


# ---------------------------------------------------------------------------
# шифр сметной нормы
# ---------------------------------------------------------------------------
class NormCodeTests(unittest.TestCase):
    def test_fer_ter_gesn_codes_are_norms(self):
        for code in ("06-01-001-01", "м08-01-062-02", "м08-02-147-04",
                     "11-01-011-03", "м39-02-012-01", "01-01-002-07"):
            self.assertTrue(sp.is_norm_code(code), code)

    def test_material_machine_and_price_codes_are_not_norms(self):
        # коды материалов ГЭСН-сборников, коды машин и прайс-листы нормами не являются
        for code in ("14.4.01.20-0012", "01.7.06.03-0003", "14.5.09.07-0030",
                     "91.14.01-011", "62.1.02.22-0092", "Расчёт 1231Р",
                     "21.1.06.10-0194", "08.4.03.03-0009", "", None):
            self.assertFalse(sp.is_norm_code(code), code)

    def test_js_and_python_agree_on_norm_detection(self):
        codes = ["06-01-001-01", "м08-01-062-02", "14.4.01.20-0012", "91.14.01-011",
                 "Расчёт 1231Р", "т03-21-01-030", "м39-02-012-01", "62.1.02.22-0092"]
        py = [sp.is_norm_code(c) for c in codes]
        html = (ROOT / "app/template.html").read_text()
        mod = re.search(r"const SourceParser = \(function\(\)\{.*?\n\}\)\(\);", html, re.S).group()
        helpers = (re.search(r"  const NORM_RE = .*?;\n", mod, re.S).group()
                   + re.search(r"  function isNormCode\(code\)\{.*?\n  \}\n", mod, re.S).group())
        js = ("function str(v){ return v===null||v===undefined? '' : String(v).trim(); }\n"
              + helpers + "\nprocess.stdout.write(JSON.stringify("
              + json.dumps(codes, ensure_ascii=False) + ".map(isNormCode)));")
        out = subprocess.run(["node"], input=js, text=True, capture_output=True, check=True)
        self.assertEqual(json.loads(out.stdout), py)


# ---------------------------------------------------------------------------
# группировка по порядку следования
# ---------------------------------------------------------------------------
class GroupingTests(unittest.TestCase):
    def test_materials_follow_their_work(self):
        items = [item(10, "1", "06-01-001-01", "Устройство подготовки"),
                 item(11, "2", "Расчёт 1231Р", "Бетон В35", cat="material"),
                 item(12, "3", "91.14.01-011", "Автобетоносмесители", unit="маш.-ч."),
                 item(13, "4", "11-01-011-03", "Устройство стяжек")]
        sp.group_items(items)
        g = pick(items)
        self.assertEqual([x["group_no"] for x in g], ["1", "1", "1", "4"])
        self.assertEqual([x["group_code"] for x in g],
                         ["06-01-001-01", "06-01-001-01", "06-01-001-01", "11-01-011-03"])
        self.assertEqual([x["group_lead"] for x in g], [True, False, False, True])
        self.assertEqual([x["group_size"] for x in g], [3, 3, 3, 1])

    def test_transport_correction_rows_join_the_same_group(self):
        """Строки «13,1»/«13,2» перевозки «исключен» идут подряд — та же группа."""
        items = [item(10, "13", "06-03-004-12", "Армирование"),
                 item(11, "13,1", "т03-21-01-030", "Перевозка (исключен)", cat="transport"),
                 item(12, "13,2", "т03-21-01-125", "Перевозка", cat="transport"),
                 item(13, "14", "08.4.03.03-0004", "Сталь арматурная", cat="material")]
        sp.group_items(items)
        g = pick(items)
        self.assertEqual([x["group_no"] for x in g], ["13", "13", "13", "13"])
        self.assertEqual([x["group_lead"] for x in g], [True, False, False, False])

    def test_rows_are_grouped_by_source_row_order_not_input_order(self):
        items = [item(30, "3", "11-01-011-04", "Третья работа"),
                 item(10, "1", "06-01-001-01", "Первая работа"),
                 item(11, "2", "", "Материал", cat="material"),
                 item(20, "2", "06-01-001-16", "Вторая работа")]
        sp.group_items(items)
        by_num = {it["row"]: (it["group_no"], it["group_lead"]) for it in items}
        self.assertEqual(by_num[10], ("1", True))
        self.assertEqual(by_num[11], ("1", False))   # материал идёт сразу за строкой 10
        self.assertEqual(by_num[20], ("2", True))
        self.assertEqual(by_num[30], ("3", True))

    def test_smeta_without_norms_keeps_every_row_its_own_group(self):
        items = [item(10, "1", "", "Материал 1", cat="material"),
                 item(11, "2", "", "Материал 2", cat="material")]
        sp.group_items(items)
        self.assertEqual([it["group_lead"] for it in items], [True, True])
        self.assertEqual([it["group_no"] for it in items], ["1", "2"])
        self.assertEqual([it["group_size"] for it in items], [1, 1])

    def test_leading_row_without_norm_still_opens_a_group(self):
        """Оборудование без норматива (КТПСН) — ведущая строка своей группы."""
        items = [item(10, "1", "", "КТПСН 6/0,4 кВ", unit="компл.", cat="equipment"),
                 item(11, "2", "", "КТПСН 6/0,4 кВ (второй)", unit="компл.", cat="equipment"),
                 item(12, "3", "м08-01-062-02", "Трансформатор силовой")]
        sp.group_items(items)
        g = pick(items)
        self.assertEqual([x["group_no"] for x in g], ["1", "1", "3"])
        self.assertEqual([x["group_lead"] for x in g], [True, False, True])

    def test_python_and_browser_produce_the_same_groups(self):
        items = [item(10, "1", "06-01-001-01", "Работа"),
                 item(11, "2", "Расчёт 1231Р", "Материал", cat="material"),
                 item(12, "3", "91.14.01-011", "Машина", unit="маш.-ч."),
                 item(13, "13,1", "т03-21-01-030", "Перевозка", cat="transport"),
                 item(20, "4", "11-01-011-03", "Работа 2"),
                 item(21, "", "", "Материал без номера", cat="material")]
        py = sp.group_items([dict(x) for x in items])
        js = js_group([dict(x) for x in items])
        self.assertEqual(pick(js), pick(py))


# ---------------------------------------------------------------------------
# учёт материалов в укрупнённом объёме работы
# ---------------------------------------------------------------------------
class AbsorbedMaterialsTests(unittest.TestCase):
    def report(self, items):
        r = make_ref()
        return ref.build_fo_report(r, [smeta(items)])

    def test_material_without_reference_joins_its_work_volume(self):
        items = [item(10, "1", "13-03-004-26", "Окраска", unit="100 м2", qty=2.0, total=1000.0),
                 item(11, "2", "14.4.01.20-0012", "Грунтовка цинкнаполненная",
                      unit="т", qty=0.5, total=500.0, cat="material")]
        rep = self.report(items)
        self.assertEqual(rep["stats"]["unmatched"], 0)
        self.assertEqual(rep["stats"]["absorbed"], 1)
        self.assertEqual(rep["stats"]["absorbed_cost"], 500.0)
        paint = [r for r in rep["rows"] if r["group"] == "Окраска, м2"][0]
        # стоимость работы и её материала — в одной группе справочника
        self.assertEqual(paint["cost"], 1500.0)
        self.assertEqual(paint["positions"], 2)
        self.assertEqual(paint["absorbed"], 1)
        self.assertEqual(paint["absorbed_cost"], 500.0)
        # объём группы — объём самой работы: 2 × 100 м2 = 200 м2
        self.assertEqual(paint["volume"], 200.0)
        # материала нет ни в перечне «не сопоставлено», ни в отдельной строке «вне справочника»
        self.assertFalse([r for r in rep["rows"] if "Вне справочника" in r["group"]])
        self.assertEqual([p.get("absorbed") for p in rep["absorbed_positions"]], [True])

    def test_material_without_reference_and_without_work_is_still_reported(self):
        """Оборудование без норматива и без расценки остаётся отдельной строкой."""
        items = [item(10, "1", "", "КТПСН 6/0,4 кВ", unit="компл.", qty=1.0,
                      total=700.0, cat="equipment")]
        rep = self.report(items)
        self.assertEqual(rep["stats"]["absorbed"], 0)
        self.assertEqual([r["group"] for r in rep["rows"]], [ref.NO_FO])
        self.assertEqual(rep["stats"]["unmatched"], 0)          # категория не в UNMATCHED_CATS
        self.assertEqual(rep["stats"]["unmatched_skipped"], 1)

    def test_work_cost_is_not_changed_by_grouping(self):
        """Итог отчёта по-прежнему равен сумме стоимостей позиций (объединение не теряет деньги)."""
        items = [item(10, "1", "13-03-004-26", "Окраска", total=1000.0),
                 item(11, "2", "14.4.01.20-0012", "Грунтовка", cat="material", total=500.0),
                 item(12, "3", "м08-02-147-01", "Кабель", unit="м", total=2000.0),
                 item(13, "4", "91.14.01-011", "Машина", unit="маш.-ч.", total=300.0)]
        rep = self.report(items)
        self.assertEqual(rep["stats"]["cost"], 3800.0)
        self.assertEqual(round(sum(r["cost"] for r in rep["rows"]), 2), 3800.0)
        self.assertEqual(rep["stats"]["absorbed"], 2)

    def test_material_found_in_reference_keeps_its_own_group(self):
        """Материал, который ЕСТЬ в справочнике, остаётся в своей группе справочника.

        Объединение касается только строк без расценки: найденная в справочнике
        группа важнее положения строки в смете.
        """
        r = make_ref(fo_rows=[
            fo_row("13-03-004-26", "Окраска, м2", "м2"),
            fo_row("14.4.01.20-0012", "Грунтовка, м2", "м2"),
            fo_row("т03-21-01-030", "Перевозка", "т"),
        ])
        items = [item(10, "1", "13-03-004-26", "Окраска", total=1000.0),
                 item(11, "2", "14.4.01.20-0012", "Грунтовка", total=500.0)]
        rep = ref.build_fo_report(r, [smeta(items)])
        self.assertEqual(rep["stats"]["absorbed"], 0)
        groups = {row["group"]: row["cost"] for row in rep["rows"]}
        self.assertEqual(groups.get("Окраска, м2"), 1000.0)
        self.assertEqual(groups.get("Грунтовка, м2"), 500.0)
        self.assertEqual(rep["stats"]["cost"], 1500.0)

    def test_absorbed_material_does_not_create_extra_excluded_row(self):
        """Материал без расценки под работой не заводит строку «Не идет в физику»."""
        r = make_ref(fo_rows=[
            fo_row("13-03-004-26", "Окраска, м2", "м2"),
            fo_row("т03-21-01-030", "Перевозка", "т"),
        ])
        items = [item(10, "1", "13-03-004-26", "Окраска", total=1000.0),
                 item(11, "2", "14.4.01.20-0012", "Грунтовка", unit="т",
                      total=500.0, cat="material"),
                 item(12, "2,1", "т03-21-01-030", "Перевозка", unit="т",
                      total=-100.0, cat="transport")]
        rep = ref.build_fo_report(r, [smeta(items)])
        self.assertEqual(rep["stats"]["absorbed"], 1)
        self.assertEqual([row["group"] for row in rep["rows"]], ["Окраска, м2", "Перевозка"])
        self.assertEqual(rep["stats"]["cost"], 1400.0)

if __name__ == "__main__":
    unittest.main(verbosity=2)
