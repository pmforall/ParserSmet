#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_multi_ls.py — приёмка разбора книг, в которых НЕСКОЛЬКО локальных смет (ЛС).

Проверяется:
1. Одна книга с двумя ЛС разбирается в две сметы отчёта, каждая — со своими
   итогами, своими позициями и своей печатной формой «11 граф».
2. Числа разбора «Source» совпадают с печатной формой (0 расхождений), ошибок
   контроля нет.
3. Независимая приёмка parser/verify_parser.py на такой книге проходит без замечаний.
4. Если листов печатной формы меньше, чем ЛС, фиксируется предупреждение
   PRINTED_COUNT, а «лишняя» ЛС получает NO_PRINTED (без ложных расхождений).

Запуск: python tests/test_multi_ls.py
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from typing import Any, Dict, List

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "tests", "out", "multi_ls")
sys.path.insert(0, os.path.join(ROOT, "parser"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import make_multi_ls as gen            # noqa: E402
import smeta_parser as sp              # noqa: E402

TOL = 0.02


class Checker:
    def __init__(self) -> None:
        self.checks = 0
        self.problems = 0
        self.details: List[str] = []

    def ok(self, cond: bool, label: str, extra: Any = "") -> None:
        self.checks += 1
        if not cond:
            self.problems += 1
            self.details.append(label + (("  [%s]" % extra) if extra != "" else ""))
        print(("  ✓ " if cond else "  ✗ ") + label + (("  [%s]" % extra) if (extra != "" and not cond) else ""))

    @staticmethod
    def near(a: float, b: float, tol: float = TOL) -> bool:
        return abs(a - b) <= max(tol, abs(b) * 1e-9)


def codes(checks: List[Dict[str, Any]]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for c in checks:
        out[c["code"]] = out.get(c["code"], 0) + 1
    return out


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    ch = Checker()
    print("КОНТРОЛЬ КНИГ С НЕСКОЛЬКИМИ ЛОКАЛЬНЫМИ СМЕТАМИ (ЛС)")

    two = gen.build(os.path.join(OUT, "Две_сметы.xlsx"), printed_count=2)
    one = gen.build(os.path.join(OUT, "Две_сметы_одна_форма.xlsx"), printed_count=1)
    exp = gen.EXPECTED

    # --- 1. Один файл -> две сметы ---------------------------------------
    print("\n1. Книга с двумя ЛС: разбор «Source»")
    data = sp.build_dataset([two])
    ch.ok(len(data["smetas"]) == exp["smetas"],
          "файл разобран в %d сметы (ЛС)" % len(data["smetas"]), len(data["smetas"]))
    nums = [s["number"] for s in data["smetas"]]
    ch.ok(nums == exp["numbers"], "шифры ЛС разобраны по порядку: " + ", ".join(nums))
    ch.ok(all(s["n_smetas_in_file"] == 2 for s in data["smetas"]),
          "каждая смета знает, что в файле их 2 (n_smetas_in_file)")
    ch.ok(data["meta"]["files"][0]["smetas"] == 2, "в метаданных файла указано 2 сметы")
    ch.ok(len(data["meta"]["files"][0]["printed_sheets"]) == 2,
          "найдены оба листа печатной формы «%s»" % gen.MARK)

    # --- 2. Итоги и позиции каждой ЛС ------------------------------------
    print("\n2. Итоги и позиции каждой ЛС (без смешивания)")
    for i, s in enumerate(data["smetas"]):
        want = exp["totals"][s["number"]]
        diffs = [k for k in ("pz", "ozp", "emm", "zpm", "nr", "sp", "total", "mat_ob")
                 if not ch.near(s["totals"].get(k, 0.0), want.get(k, 0.0))]
        ch.ok(not diffs, "ЛС %s: итоги соответствуют ожидаемым" % s["number"], ", ".join(diffs))
        ch.ok(s["n_items"] == exp["items"][i],
              "ЛС %s: позиций %d" % (s["number"], s["n_items"]), s["n_items"])
        ch.ok(all(it["ls"] == s["node_id"] for it in s["items"]),
              "ЛС %s: все позиции принадлежат своей смете (узел %s)" % (s["number"], s["node_id"]))
        # сумма позиций сметы равна её итогу по «Всего с НР и СП»
        ssum = round(sum(it["values"].get("total", 0.0) for it in s["items"]), 2)
        ch.ok(ch.near(ssum, s["totals"]["total"]),
              "ЛС %s: сумма позиций (%s) = итогу сметы (%s)" % (s["number"], ssum, s["totals"]["total"]))
        for it in s["items"]:
            want = (exp.get("codes") or {}).get(str(it["num"]))
            if want is None:
                continue
            ch.ok(it["code"] == want,
                  "позиция %s: шифр с пометкой коэффициента приведён к «%s» (в файле %r)"
                  % (it["num"], want, it.get("code_raw")))
        with_notes = [it for it in s["items"]
                      if it.get("code_raw") and it["code_raw"] != it["code"]]
        ch.ok(all(("к=" in it["code_raw"]) or ("×" in it["code_raw"]) or ("*" in it["code_raw"])
                  for it in with_notes),
              "исходные шифры с пометками сохранены в code_raw: %s"
              % ", ".join(repr(it["code_raw"]) for it in with_notes))
    ch.ok(exp["totals"][exp["numbers"][0]]["total"] != exp["totals"][exp["numbers"][1]]["total"],
          "итоги ЛС различаются — смешивание было бы заметно")

    # --- 3. Печатная форма у каждой ЛС -----------------------------------
    print("\n3. Сверка с печатной формой «%s» по каждой ЛС" % gen.MARK)
    for s in data["smetas"]:
        prn = s.get("printed") or {}
        ch.ok(bool(prn), "ЛС %s: печатная форма сопоставлена (%s)" % (s["number"], prn.get("sheet", "—")))
        ch.ok(prn.get("number") == s["number"],
              "ЛС %s: в печатной форме тот же номер" % s["number"], prn.get("number"))
        mis = [c for c in s["cross_check"] if c["status"] == "mismatch"]
        ch.ok(not mis, "ЛС %s: расхождений «Source» ↔ печатная форма нет" % s["number"],
              ", ".join("%s Δ=%s" % (c["key"], c["delta"]) for c in mis))
        keys = {c["key"] for c in s["cross_check"] if c["status"] == "ok"}
        ch.ok({"pz", "total", "ozp", "emm", "nr", "sp", "material"} <= keys,
              "ЛС %s: сверено показателей «совпадает» — %d" % (s["number"], len(keys)))
    sheets = {s["printed"]["sheet"] for s in data["smetas"]}
    ch.ok(len(sheets) == 2, "каждой ЛС сопоставлен свой лист печатной формы")
    by_code = codes(data["checks"])
    ch.ok(by_code.get("PRINTED_MISMATCH", 0) == 0, "ложных расхождений с печатной формой нет")
    ch.ok(by_code.get("PRINTED_COUNT", 0) == 0, "число печатных форм совпадает с числом ЛС")
    ch.ok(all(c["severity"] != "error" for c in data["checks"]),
          "ошибок контроля данных нет", json.dumps(by_code, ensure_ascii=False))

    # --- 3б. контроль суммы всех затрат по строке «ВСЕГО ПО СМЕТЕ» ---------
    print("\n3б. Контроль суммы всех затрат по строке «ВСЕГО ПО СМЕТЕ» вкладки «Source»")
    for s in data["smetas"]:
        tc = {x["code"]: x for x in (s.get("total_check") or [])}
        ch.ok(s.get("source_total") is not None,
              "ЛС %s: найдена строка «ВСЕГО ПО СМЕТЕ» вкладки «Source»" % s["number"])
        tot = tc.get("SOURCE_TOTAL")
        base = tc.get("SOURCE_TOTAL_BASE")
        ch.ok(tot is not None and base is not None,
              "ЛС %s: проверены и итог с коэффициентом, и базисный итог" % s["number"])
        if tot:
            ch.ok(ch.near(tot["source"], exp["totals"][s["number"]]["total"] +
                          (gen.K_USP - 1) * (exp["totals"][s["number"]]["ozp"] +
                                             exp["totals"][s["number"]]["emm"] +
                                             exp["totals"][s["number"]]["nr"] +
                                             exp["totals"][s["number"]]["sp"])),
                  "ЛС %s: «ВСЕГО ПО СМЕТЕ» = базисный итог + (k−1)·(ОЗП+ЭММ+НР+СП) = %.2f"
                  % (s["number"], tot["source"]))
            ch.ok(tot["status"] == "ok", "ЛС %s: расхождения с итогом «Source» нет (Δ=%s, k=%s)"
                  % (s["number"], tot.get("delta"), tot.get("k")))
        if base:
            ch.ok(ch.near(base["source"], s["totals_source"]["total"]),
                  "ЛС %s: базисный итог «Source» = сумме всех затрат (%.2f)"
                  % (s["number"], base["source"]))
    ch.ok(all(not [x for x in (s.get("total_check") or []) if x["status"] == "mismatch"]
              for s in data["smetas"]), "ни одна ЛС не помечена расхождением с итогом «Source»")

    # --- 4. Независимая приёмка (verify_parser.py) ------------------------
    print("\n4. Независимая приёмка разбора (parser/verify_parser.py)")
    jpath = os.path.join(OUT, "smetas_multi.json")
    with open(jpath, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)
    proc = subprocess.run([sys.executable, os.path.join(ROOT, "parser", "verify_parser.py"), OUT, jpath],
                          capture_output=True, text=True)
    tail = [ln for ln in proc.stdout.strip().splitlines() if ln.strip()][-3:]
    ch.ok(proc.returncode == 0, "приёмка по печатной форме: " + (tail[-1] if tail else "нет вывода"),
          " | ".join(tail))

    # --- 5. Печатных форм меньше, чем ЛС ---------------------------------
    print("\n5. Книга с двумя ЛС, но одним листом печатной формы")
    d2 = sp.build_dataset([one])
    ch.ok(len(d2["smetas"]) == 2, "смет по-прежнему две", len(d2["smetas"]))
    cc = codes(d2["checks"])
    ch.ok(cc.get("PRINTED_COUNT", 0) >= 1,
          "зафиксировано предупреждение PRINTED_COUNT (листов меньше, чем ЛС)")
    ch.ok(cc.get("NO_PRINTED", 0) >= 1, "«лишняя» ЛС помечена NO_PRINTED")
    ch.ok(all(c["severity"] != "error" for c in d2["checks"]),
          "ошибок нет — отсутствие формы не выдаётся за расхождение")
    with_form = [s for s in d2["smetas"] if s.get("printed")]
    without = [s for s in d2["smetas"] if not s.get("printed")]
    ch.ok(len(with_form) == 1 and len(without) == 1,
          "печатная форма привязана ровно к одной ЛС")
    if with_form and without:
        ch.ok(ch.near(with_form[0]["totals"]["total"], exp["totals"][with_form[0]["number"]]["total"]),
              "у ЛС с формой итог по печатной форме (%s)" % with_form[0]["totals"]["total"])
        ch.ok(ch.near(without[0]["totals"]["total"], exp["totals"][without[0]["number"]]["total"]),
              "у ЛС без формы итог взят из «Source» (%s)" % without[0]["totals"]["total"])

    print("=" * 78)
    print("Проверок: %d · провалов: %d" % (ch.checks, ch.problems))
    print("ИТОГ: %s" % ("УСПЕХ — книги с несколькими ЛС разбираются корректно"
                        if ch.problems == 0 else "ЕСТЬ ЗАМЕЧАНИЯ"))
    if ch.details:
        print("\nЗамечания:")
        for d in ch.details:
            print("  · " + d)
    return 0 if ch.problems == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
