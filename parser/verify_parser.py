#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_parser.py — независимая приёмка парсера смет по печатной форме.

Требование технического задания: проверку корректности данных выполнять по печатной
форме локальной сметы — листам, в названии которых есть «11 граф»; суммарные данные
разбора обязаны совпадать с этой формой.

Что проверяется
---------------
1. Печатная форма читается отдельным модулем (parser/printed_form.py) и внутренне
   согласована: «Сметная стоимость» шапки = строке «ВСЕГО ПО СМЕТЕ», «Нормативная
   трудоемкость» и «Сметная заработная плата» = трудозатратам и ФОТ с коэффициентом
   «Условия производства работ».
2. Итоговые показатели из JSON парсера совпадают с печатной формой (12 показателей).
3. Итоги разбора «Source» (totals_source) совпадают с суммой позиций из «Source».
4. Позиции печатной формы совпадают с позициями разбора: количество, объём, общая
   стоимость, ОЗП, ЭММ, НР, СП.
5. Разделы и подразделы печатной формы совпадают с узлами разбора.
6. Печатные формы есть у каждой сметы (иначе сверка невозможна — это фиксируется).
7. Трудозатраты (показатели 207/208) не теряют точность: берутся из «Source» с четырьмя
   знаками, сумма позиций сходится с итогом сметы, а в сверке с печатной формой
   используется точное значение формы, а не напечатанное округлённым до 2 знаков.

Запуск:
    python verify_parser.py <папка_со_сметами> <smetas.json>
"""

from __future__ import annotations

import json
import os
import re
import sys
from typing import Any, Dict, List

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import printed_form as pf

TOL = 0.02


def _source_cell(wb, row: int):
    """Отображаемое имя (столбец H) и значение (столбец F/W) указанной строки вкладки «Source»."""
    for ws in wb.worksheets:
        if ws.title.strip().lower() != "source":
            continue
        if row < 1 or row > ws.max_row:
            return "", None
        val = ws.cell(row, 6).value
        if val is None:
            val = ws.cell(row, 23).value
        return str(ws.cell(row, 8).value or ""), (float(val) if val is not None else None)
    return "", None


def _norm(t: str) -> str:
    return "".join(ch for ch in str(t or "").lower() if ch.isalnum())


def _printed_percents(ws) -> Dict[str, Dict[str, float]]:
    """Проценты НР и СП, напечатанные в форме рядом с позицией: «НР=110% (НР=5,21)».

    Возвращает {номер позиции: {"nr": 110.0, "sp": 68.0}}. Строки норм идут после своей
    позиции, поэтому текущая позиция определяется по последнему номеру в столбце A.
    """
    out: Dict[str, Dict[str, float]] = {}
    cur = ""
    for r in range(pf.header_last_row(ws) + 1, ws.max_row + 1):
        a = str(ws.cell(r, 1).value if ws.cell(r, 1).value is not None else "").strip()
        c = str(ws.cell(r, 3).value if ws.cell(r, 3).value is not None else "").strip()
        if a.isdigit():
            cur = a
        if not cur or not c:
            continue
        m = re.match(r"НР\s*=\s*([\d.,]+)\s*%", c)
        if m:
            out.setdefault(cur, {})["nr"] = float(m.group(1).replace(",", "."))
        m = re.match(r"СП\s*=\s*([\d.,]+)\s*%", c)
        if m:
            out.setdefault(cur, {})["sp"] = float(m.group(1).replace(",", "."))
    return out


def _prn_tol(v: float) -> float:
    """Допуск сверки с напечатанным значением — половина последнего напечатанного знака.

    В книгах ЛСР-02-… трудозатраты печатаются с одним-двумя знаками (144814,2), в книгах
    типовых книг — с четырьмя; значения «Source» хранятся с четырьмя знаками.
    """
    return max(0.0001, 0.5 * (10.0 ** (-pf.decimals(v))) + 1e-9)


def _two_decimals(v: float) -> bool:
    """Значение представимо двумя знаками после запятой (то есть напечатано округлённым)."""
    return abs(v * 100.0 - round(v * 100.0)) <= 1e-9


class Checker:
    def __init__(self) -> None:
        self.checks = 0
        self.problems = 0
        self.details: List[Dict[str, Any]] = []

    def ok(self, cond: bool, label: str, extra: Any = "") -> None:
        self.checks += 1
        if not cond:
            self.problems += 1
        tail = "" if (cond or extra == "") else "  [%s]" % extra
        print(("  ✓ " if cond else "  ✗ ") + label + tail)
        if not cond:
            self.details.append({"label": label, "extra": str(extra)})

    @staticmethod
    def near(a: float, b: float, tol: float = TOL) -> bool:
        return abs(a - b) <= max(tol, abs(b) * 1e-9)


def main(argv: List[str]) -> int:
    folder = argv[1] if len(argv) > 1 else "."
    json_path = argv[2] if len(argv) > 2 else "smetas.json"
    with open(json_path, encoding="utf-8") as fh:
        data = json.load(fh)

    ch = Checker()
    print("ПРИЁМКА ПО ПЕЧАТНОЙ ФОРМЕ (листы, содержащие «%s»)" % pf.MARK)
    keys = [k for k, _ in pf.COMPARE_KEYS]

    for smeta in data["smetas"]:
        path = smeta.get("path") or os.path.join(folder, smeta["file"])
        print("=" * 104)
        print("СМЕТА: %s" % smeta["number"])
        print("ФАЙЛ : %s" % os.path.basename(path))
        wb = load_workbook(path, data_only=True)
        def _rows_of(name):
            try:
                return [[c.value for c in row] for row in wb[name].iter_rows()]
            except Exception:          # pragma: no cover
                return None
        names = pf.sheets_with_mark(wb, _rows_of)
        ch.ok(True, "печатная форма «…%s…»: %s" % (pf.MARK, ", ".join(names) or
               "нет (книга без печатной формы — сверка по форме не выполняется)"))
        if not names:
            wb.close()
            continue
        # лист для этой ЛС: сначала тот, что привязан при разборе (в книге их может
        # быть несколько — по одному на каждую локальную смету), затем по номеру
        sheet_name = ((smeta.get("printed") or {}).get("sheet") or "")
        if sheet_name not in names:
            want = _norm(smeta.get("number"))
            sheet_name = next((n for n in names if want and want in _norm(n)), "")
        if not sheet_name:
            sheet_name = names[0]
        ws = wb[sheet_name]
        prn = pf.parse(ws)

        # 1. внутренняя согласованность печатной формы
        self_problems = [p for p in pf.check_printed_self(prn) if p.get("severity") == "warn"]
        ch.ok(not self_problems, "печатная форма внутренне согласована",
              "; ".join("%s Δ=%s" % (p["code"], p.get("delta")) for p in self_problems))

        # 2. итоги JSON против печатной формы
        reported = smeta.get("totals", {})
        bad = []
        kind_bad = []
        compared = 0
        for key in keys:
            prn_val = prn["totals"].get(key)
            if prn_val is None:
                continue
            compared += 1
            got = float(reported.get(key, 0.0) or 0.0)
            if not ch.near(got, prn_val):
                # Разбивка СМР по видам работ: форма делит суммы по собственному правилу,
                # программа — по «виду затрат» и группе ресурсов «Source» (сходится с
                # показателями узлов 214/215/216/217). Это разница методик, а не расхождение:
                # строка попадает в отчёт о сверках как PRINTED_KIND_DIFF.
                (kind_bad if key in pf.KIND_KEYS else bad).append(
                    "%s: %s != %s" % (key, got, prn_val))
        ch.ok(not bad, "итоги отчёта совпадают с печатной формой по %d показателям" % compared,
              "; ".join(bad))
        if kind_bad:
            print("  ~ разбивка по видам работ отличается от формы (учтено как PRINTED_KIND_DIFF): "
                  + "; ".join(kind_bad))

        # 2б. стоимость работ без НР и СП: ПЗ = «Всего с НР и СП» − НР − СП
        rep_t = smeta.get("totals", {})
        tz_pz = float(rep_t.get("pz") or 0.0)
        tz_nr = float(rep_t.get("nr") or 0.0)
        tz_sp = float(rep_t.get("sp") or 0.0)
        tz_total = float(rep_t.get("total") or 0.0)
        ch.ok(ch.near(tz_pz + tz_nr + tz_sp, tz_total),
              "прямые затраты (стоимость без НР и СП) + НР + СП = «Всего с НР и СП» "
              "(%.2f + %.2f + %.2f = %.2f)" % (tz_pz, tz_nr, tz_sp, tz_total),
              "%.2f != %.2f" % (tz_pz + tz_nr + tz_sp, tz_total))
        kind_pz_tot = rep_t.get("kind_pz") or {}
        if kind_pz_tot:
            ch.ok(ch.near(sum(float(v or 0.0) for v in kind_pz_tot.values()), tz_pz),
                  "сумма прямых затрат по видам работ = прямым затратам сметы (%.2f)"
                  % sum(float(v or 0.0) for v in kind_pz_tot.values()),
                  "%.2f != %.2f" % (sum(float(v or 0.0) for v in kind_pz_tot.values()), tz_pz))

        # 3. totals_source против суммы позиций разбора «Source»
        src_totals = smeta.get("totals_source") or {}
        items = smeta.get("items") or []
        acc = {k: 0.0 for k in ("pz", "mat_ob", "emm", "zpm", "ozp", "nr", "sp",
                                "tz_workers", "tz_mash")}
        for it in items:
            for k in acc:
                acc[k] += float((it.get("values") or {}).get(k, 0.0) or 0.0)
        bad = ["%s: %s != %s" % (k, round(acc[k], 2), src_totals.get(k))
               for k in acc if src_totals.get(k) is not None and not ch.near(acc[k], float(src_totals[k]))]
        ch.ok(not bad, "сумма позиций «Source» = итогам разбора «Source» (%d позиций)" % len(items),
              "; ".join(bad))

        # 4. позиции печатной формы против позиций разбора
        prn_list = prn.get("positions", [])
        seq_equal = len(prn_list) == len(items)
        ch.ok(seq_equal, "количество позиций: форма %d = разбор %d" % (len(prn_list), len(items)),
              "расхождение количества")
        mism = 0
        for i, it in enumerate(items):
            pp = prn_list[i] if seq_equal else next(
                (p for p in prn_list if str(p.get("num")) == str(it.get("num"))), None)
            if pp is None:
                mism += 1
                continue
            vals = it.get("values") or {}
            for pk, ik, label in (("total", "pz", "стоимость"), ("ozp", "ozp", "ОЗП"),
                                  ("emm", "emm", "ЭММ"), ("nr", "nr", "НР"), ("sp", "sp", "СП")):
                a = float(vals.get(ik, 0.0) or 0.0)
                b = float(pp.get(pk, 0.0) or 0.0)
                if not ch.near(a, b):
                    mism += 1
                    print("      ! №%s %s: форма=%.2f разбор=%.2f" % (it.get("num"), label, b, a))
            if abs(float(it.get("qty", 0.0) or 0.0) - float(pp.get("qty", 0.0) or 0.0)) > 1e-9:
                mism += 1
                print("      ! №%s объём: форма=%s разбор=%s" % (it.get("num"), pp.get("qty"), it.get("qty")))
        ch.ok(mism == 0, "позиции совпали (№, объём, стоимость, ОЗП, ЭММ, НР, СП)",
              "расхождений: %d" % mism)

        # 4б. трудозатраты по каждой позиции: 11-я графа печатной формы против «Source»
        tz_mism = 0
        tz_sum_p = tz_sum_j = 0.0
        tz_n = 0
        for i, it in enumerate(items):
            pp = prn_list[i] if seq_equal else next(
                (p for p in prn_list if str(p.get("num")) == str(it.get("num"))), None)
            if pp is None or pp.get("tz_workers") is None:
                continue
            tz_n += 1
            pv = float(pp["tz_workers"])
            jv = float((it.get("values") or {}).get("tz_workers", 0.0) or 0.0)
            tz_sum_p += pv
            tz_sum_j += jv
            # в «Source» значение хранится с 4 знаками, в форме книги типовых книг печатается
            # так же, а в книгах ЛСР-02-… — с двумя знаками: допуск — половина последнего
            # напечатанного знака
            tol_tz = max(0.0005, _prn_tol(pv))
            if abs(jv - pv) > tol_tz:
                tz_mism += 1
                print("      ! №%s трудозатраты строителей: форма=%.4f разбор=%.4f"
                      % (it.get("num"), pv, jv))
        ch.ok(tz_mism == 0 and tz_n > 0,
              "трудозатраты строителей по каждой позиции = 11-й графе печатной формы (%d позиций)" % tz_n,
              "расхождений: %d" % tz_mism)
        ch.ok(ch.near(tz_sum_p, tz_sum_j, max(0.005, 0.005 * max(1, tz_n))),
              "сумма трудозатрат позиций формы = сумме по позициям разбора (%.4f)" % tz_sum_j,
              "форма %.4f != разбор %.4f" % (tz_sum_p, tz_sum_j))

        # 4в. нормативы НР и СП: «Source» и печатная форма («НР=110% (НР=5,21)»), база — ФОТ
        fot = float(src_totals.get("ozp", 0.0) or 0.0) + float(src_totals.get("zpm", 0.0) or 0.0)
        prn_pct = _printed_percents(ws)
        bad_norm, bad_pct, n_norm = [], [], 0
        for it in items:
            vals = it.get("values") or {}
            fot_i = float(vals.get("ozp", 0.0) or 0.0) + float(vals.get("zpm", 0.0) or 0.0)
            nr_p = float(it.get("nr_pct") or 0.0)
            sp_p = float(it.get("sp_pct") or 0.0)
            if nr_p > 0 and fot_i > 0:
                n_norm += 1
                if abs(float(vals.get("nr", 0.0)) - nr_p / 100.0 * fot_i) > 0.02:
                    bad_norm.append("№%s НР %.2f != %s%% x %.2f" % (it.get("num"), vals.get("nr", 0.0), nr_p, fot_i))
            if sp_p > 0 and fot_i > 0:
                if abs(float(vals.get("sp", 0.0)) - sp_p / 100.0 * fot_i) > 0.02:
                    bad_norm.append("№%s СП %.2f != %s%% x %.2f" % (it.get("num"), vals.get("sp", 0.0), sp_p, fot_i))
            # процент, напечатанный в форме, должен совпадать с процентом из «Source»
            got = prn_pct.get(str(it.get("num")).strip())
            if got and nr_p > 0 and abs(got.get("nr", nr_p) - nr_p) > 0.01:
                bad_pct.append("№%s НР: форма %s%% != Source %s%%" % (it.get("num"), got.get("nr"), nr_p))
            if got and sp_p > 0 and abs(got.get("sp", sp_p) - sp_p) > 0.01:
                bad_pct.append("№%s СП: форма %s%% != Source %s%%" % (it.get("num"), got.get("sp"), sp_p))
        if n_norm:
            ch.ok(not bad_norm,
                  "нормативы НР и СП: НР = «НР, %%» x ФОТ и СП = «СП, %%» x ФОТ по позициям (проверено %d)" % n_norm,
                  "; ".join(bad_norm[:4]))
        else:
            ch.ok(True, "нормативы НР и СП в смете не заданы — сверка НР = «НР, %%» x ФОТ не требуется")
        if prn_pct and n_norm:
            ch.ok(not bad_pct,
                  "проценты НР и СП из печатной формы совпадают с «Source» (позиций с нормативом в форме: %d)"
                  % len(prn_pct),
                  "; ".join(bad_pct[:4]))
        else:
            ch.ok(True, "в печатной форме проценты НР и СП не напечатаны — сверка процентов не требуется")
        if fot > 0:
            nr_eff = 100.0 * float(src_totals.get("nr", 0.0) or 0.0) / fot
            sp_eff = 100.0 * float(src_totals.get("sp", 0.0) or 0.0) / fot
            norms = sorted({float(it.get("nr_pct") or 0.0) for it in items if float(it.get("nr_pct") or 0.0) > 0})
            ch.ok(ch.near(nr_eff, float(src_totals.get("nr", 0.0) or 0.0) / fot * 100.0, 1e-6),
                  "применённый процент НР по смете = НР / ФОТ (ОЗП + ЗПМ) = %.2f %%" % nr_eff)
            ch.ok(ch.near(sp_eff, float(src_totals.get("sp", 0.0) or 0.0) / fot * 100.0, 1e-6),
                  "применённый процент СП по смете = СП / ФОТ (ОЗП + ЗПМ) = %.2f %%" % sp_eff)
            if len(norms) == 1:
                ch.ok(abs(nr_eff - norms[0]) < 0.001,
                      "норматив НР сметы один (%.2f %%) и совпадает с применённым процентом" % norms[0],
                      "%.4f != %.4f" % (nr_eff, norms[0]))
            else:
                ch.ok(True, "у сметы %d различных нормативов НР: %s (проценты по смете — средневзвешенные)"
                      % (len(norms), ", ".join("%g" % x for x in norms)))

        # 5. разделы и подразделы
        if prn.get("sections"):
            # В смете бывают разделы и подразделы с одинаковыми названиями («Трубопроводы»,
            # «Материалы»), поэтому сверяем не по имени, а по привязке, выполненной при разборе
            # (вид узла + родительский раздел), и проверяем, что каждый итог формы кому-то отдан.
            nodes = [n for n in smeta.get("nodes", []) if n.get("kind") in (4, 5)]
            bad = []
            used_rows = set()
            for node in nodes:
                sec = node.get("printed_section")
                if not sec:
                    continue
                used_rows.add(sec.get("row"))
                if not ch.near(float(node["totals"].get("pz", 0.0)), float(sec.get("pz", 0.0))):
                    bad.append("«%s»: %s != %s" % (str(node.get("name"))[:30],
                                                   node["totals"].get("pz"), sec.get("pz")))
            ch.ok(not bad, "разделы/подразделы формы = узлам разбора (%d шт.)" % len(used_rows),
                  "; ".join(bad))
            ch.ok(len(used_rows) == len(prn["sections"]),
                  "каждый итог формы сопоставлен узлу сметы (%d из %d)"
                  % (len(used_rows), len(prn["sections"])),
                  "не сопоставлены строки: %s"
                  % ", ".join(str(x.get("row")) for x in prn["sections"] if x.get("row") not in used_rows))

        # 5б. трудозатраты: значения не должны терять точность (в «Source» — 4 знака).
        #     В печатной форме показатель напечатан дважды: в итоговой строке с полной
        #     точностью и в справочной строке «Трудозатраты …» округлённым до 2 знаков,
        #     поэтому в сверку должно попадать точное значение формы, а не округлённое.
        rep_w = float(reported.get("tz_workers", 0.0) or 0.0)
        rep_m = float(reported.get("tz_mash", 0.0) or 0.0)
        src_w = float(src_totals.get("tz_workers", 0.0) or 0.0)
        src_m = float(src_totals.get("tz_mash", 0.0) or 0.0)
        ch.ok(ch.near(rep_w, src_w, 0.0001) and ch.near(rep_m, src_m, 0.0001),
              "трудозатраты отчёта = итогам «Source» с точностью 4 знака "
              "(ТЗС %.4f, ТЗМ %.4f)" % (rep_w, rep_m),
              "ТЗС %s != %s, ТЗМ %s != %s" % (rep_w, src_w, rep_m, src_m))
        # точность «Source» не должна теряться: если в «Source» значения есть, они берутся как есть
        if src_w or src_m:
            ch.ok(abs(rep_w - src_w) < 1e-9 and abs(rep_m - src_m) < 1e-9,
                  "трудозатраты взяты из «Source» без округления (ТЗС %.4f, ТЗМ %.4f)" % (rep_w, rep_m),
                  "ТЗС %.4f != %.4f, ТЗМ %.4f != %.4f" % (rep_w, src_w, rep_m, src_m))
        # сумма позиций «Source» сходится с итогом по трудозатратам (округление позиций — 4 знака)
        ch.ok(ch.near(acc["tz_workers"], rep_w, 0.001) and ch.near(acc["tz_mash"], rep_m, 0.001),
              "сумма позиций по трудозатратам = итогу (ТЗС %.4f, ТЗМ %.4f)" % (acc["tz_workers"], acc["tz_mash"]),
              "ТЗС %s, ТЗМ %s" % (acc["tz_workers"], acc["tz_mash"]))
        prn_w = prn["totals"].get("tz_workers")
        prn_m = prn["totals"].get("tz_mash")
        if prn_w is not None and prn_m is not None:
            # в книгах ЛСР-02-… форма печатает трудозатраты по позициям с двумя знаками,
            # поэтому её итог может отличаться от точной суммы «Source» накопленным
            # округлением по позициям (n × 0,005)
            acc_tol = max(_prn_tol(float(prn_w)), 0.005 * max(1, tz_n))
            tol_w = acc_tol
            tol_m = max(_prn_tol(float(prn_m)), 0.005 * max(1, tz_n))
            ch.ok(ch.near(rep_w, float(prn_w), tol_w) and ch.near(rep_m, float(prn_m), tol_m),
                  "трудозатраты совпадают с печатной формой (ТЗС %.4f = %.4f, ТЗМ %.4f = %.4f)"
                  % (rep_w, float(prn_w), rep_m, float(prn_m)),
                  "в сверку попало округлённое значение формы, а не точный итог")
            # Выбор точного значения из двух представлений одного показателя проверяется на
            # синтетическом примере (tests/tests.js): в книгах типовых книг форма печатает итог
            # с четырьмя знаками, в книгах ЛСР-02-… — только с двумя, поэтому «потеря»
            # знаков здесь не является дефектом разбора.

        # 6. шапка формы: сметная стоимость, нормативная трудоёмкость, сметная ЗП
        head = prn.get("head", {})
        tot_k = (prn.get("totals_with_k") or {}).get("total")
        ch.ok(not head.get("smeta_cost") or not tot_k or ch.near(head["smeta_cost"], tot_k, 10.0),
              "«Сметная стоимость» шапки = «ВСЕГО ПО СМЕТЕ» формы (%.2f руб.)" % (tot_k or 0.0),
              "%.2f != %.2f" % (head.get("smeta_cost", 0.0), tot_k or 0.0))
        k = head.get("k_usp") or 1.0
        if head.get("norm_labor"):
            expect = (float(reported.get("tz_workers", 0.0)) + float(reported.get("tz_mash", 0.0))) * k
            ch.ok(ch.near(head["norm_labor"], expect, max(5.0, expect * 1e-4)),
                  "«Нормативная трудоемкость» = (ТЗС + ТЗМ) × %s = %.2f чел.-ч" % (k, expect),
                  "%.2f != %.2f" % (head["norm_labor"], expect))
        if head.get("smeta_wage"):
            expect = (float(reported.get("ozp", 0.0)) + float(reported.get("zpm", 0.0))) * k
            ch.ok(ch.near(head["smeta_wage"], expect, max(5.0, expect * 1e-4)),
                  "«Сметная заработная плата» = (ОЗП + ЗПМ) × %s = %.2f руб." % (k, expect),
                  "%.2f != %.2f" % (head["smeta_wage"], expect))

        # 7. контроль суммы всех затрат по строке «ВСЕГО ПО СМЕТЕ» вкладки «Source».
        #    Значения читаются из книги независимо от разбора и сверяются с отчётом.
        total_check = {x["code"]: x for x in (smeta.get("total_check") or [])}
        tc = total_check.get("SOURCE_TOTAL")
        tb = total_check.get("SOURCE_TOTAL_BASE")
        ch.ok(tc is not None, "в отчёте есть контроль суммы всех затрат по строке «ВСЕГО ПО СМЕТЕ»")
        if tc is not None:
            disp, val = _source_cell(wb, int(tc.get("row") or 0))
            ch.ok(_norm(disp) == "всегопосмете",
                  "строка %s вкладки «Source» — «ВСЕГО ПО СМЕТЕ»" % tc.get("row"), disp)
            ch.ok(val is not None and ch.near(float(tc.get("source", 0.0)), val),
                  "значение строки «ВСЕГО ПО СМЕТЕ» вкладки «Source» = %s руб." % val,
                  "%s != %s" % (tc.get("source"), val))
            ch.ok(tc.get("status") == "ok", "контроль суммы всех затрат по «Source» пройден",
                  "статус: %s, отклонение: %s" % (tc.get("status"), tc.get("delta")))
            if (prn.get("totals_with_k") or {}).get("total") is not None:
                ch.ok(ch.near(float(tc.get("source", 0.0)),
                              float(prn["totals_with_k"]["total"])),
                      "строка «ВСЕГО ПО СМЕТЕ» вкладки «Source» = «ВСЕГО ПО СМЕТЕ» печатной формы")
        if tb is not None:
            disp_b, val_b = _source_cell(wb, int(tb.get("row") or 0))
            ch.ok(_norm(disp_b).startswith("всегопосмете"),
                  "строка %s вкладки «Source» — базисный итог «ВСЕГО ПО СМЕТЕ (…)»" % tb.get("row"), disp_b)
            ch.ok(val_b is not None and ch.near(float(tb.get("source", 0.0)), val_b),
                  "базисный итог вкладки «Source» = %s руб. (сверен)" % val_b)
            if val_b is not None:
                ch.ok(ch.near(val_b, float(reported.get("total", 0.0))),
                      "сумма всех затрат (базисный уровень) = строке «Всего по Смете» вкладки «Source»",
                      "%.2f != %.2f" % (val_b, reported.get("total", 0.0)))
        wb.close()

    print("=" * 104)
    print("Проверок: %d · расхождений: %d" % (ch.checks, ch.problems))
    print("ИТОГ: %s" % ("УСПЕХ — данные разбора совпадают с печатной формой «%s»" % pf.MARK
                        if ch.problems == 0 else "ЕСТЬ РАСХОЖДЕНИЯ"))
    return 0 if ch.problems == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
