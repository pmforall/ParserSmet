#!/usr/bin/env python3
"""Приёмка веб-сборки: настоящий Chromium, локальные XLSX, скачанный Excel.

Запуск: .venv/bin/python tests/test_browser_contract.py [--baseline-html FILE]
Браузер берётся из SMETA_CHROMIUM либо установленного Playwright.
Результаты, тестовые книги и отчёты остаются в tests/out/browser.
"""
from __future__ import annotations

import argparse
import functools
import json
import os
from pathlib import Path
import sys
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/out/browser"
sys.path[:0] = [str(ROOT / "parser"), str(ROOT / "tests")]
import smeta_parser as sp
import make_multi_ls
import reference as ref
sys.path.insert(0, str(ROOT / "tools"))
from package_web import validate


def check(condition, message):
    if not condition:
        raise AssertionError(message)
    print("PASS: " + message, flush=True)


def marker_fixture(path, base=175, total=180, graph=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "Source"
    rows = [
        {1: 1, 6: "MARKER-TEST", 7: "Проверка отбора", 10: "Тестовое основание"},
        {1: 75, 2: "Базовый", 3: "2000", 14: 1},
        {1: 3, 2: 1, 4: 13},
        {1: 17, 2: 1, 5: 1, 6: "kept-a", 7: "Выбрано A", 8: "шт", 9: 1,
         15: 100, 16: 100, 27: 1, 195: 100},
        {1: 17, 2: 0, 5: 2, 6: "excluded", 7: "Исключено", 9: 1,
         15: 900000, 16: 900000, 195: 900000},
        {1: 18, 2: 1, 3: 1, 6: "resource-kept", 7: "Выбранный ресурс", 9: 2, 15: 5, 195: 10},
        {1: 18, 2: 0, 3: 1, 6: "resource-excluded", 7: "Исключённый ресурс", 9: 999},
        {1: 17, 2: "1", 5: 3, 6: "kept-b", 7: "Выбрано B", 8: "шт", 9: 1,
         15: 50, 16: 50, 27: 1, 195: 50},
        {1: 17, 5: 4, 6: "empty-marker", 7: "Без отметки", 15: 800000, 195: 800000},
        {1: 70, 2: 0, 5: "КСт_ОС", 7: 2},
        {1: 70, 2: 1, 5: "КСт_ОС", 7: 1},
        {1: 50, 2: 0, 4: 2, 11: 212, 7: "fake", 8: "ВСЕГО ПО СМЕТЕ", 6: 900000},
        {1: 51, 2: 1, 3: 1, 4: 100, 6: "MARKER-TEST", 7: "Проверка отбора",
         15: 150, 16: 150, 195: 150},
        {1: 50, 2: 1, 4: 1, 11: 224, 6: 150, 23: 150},
        {1: 50, 2: 0, 4: 1, 11: 224, 6: 900000},
        {1: 50, 2: 1, 4: 2, 11: 212, 7: "ВСЕГО_ПЗ",
         8: "ВСЕГО ПО СМЕТЕ (базисный уровень)", 14: "Всего по Смете", 6: base},
        {1: 50, 2: 1, 4: 2, 11: 212, 7: "КСт_Всего", 8: "ВСЕГО ПО СМЕТЕ", 6: total},
        {1: 50, 2: 0, 4: 2, 11: 212, 7: "fake", 8: "ВСЕГО ПО СМЕТЕ", 6: 999999},
        {1: 51, 2: 0, 3: 1, 4: 200, 6: "excluded-node", 195: 999999},
        {1: 50, 2: 1, 4: 1, 11: 224, 6: 999999},
    ]
    for r, cells in enumerate(rows, 1):
        for c, value in cells.items():
            ws.cell(r, c, value)
    if graph:
        printed = wb.create_sheet("Печатная форма (ГРАФЫ)")
        printed.append([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11])
        printed.append([1, "kept-a", "Выбрано A", 999, 999, None, 999999])
        printed.append([2, "excluded", "Исключено", 999, 999, None, 900000])
        printed.append([3, "kept-b", "Выбрано B", 999, 999, None, 999999])
    wb.save(path)
    wb.close()
    return path


def load(page, files):
    page.locator("#fXlsx").set_input_files([str(p) for p in files])
    page.wait_for_function("""() => /^(Загружено из папки:|Не удалось)/.test(
        document.getElementById('loadInfo').textContent)""", timeout=180000)
    info = page.locator("#loadInfo").inner_text()
    check(info.startswith("Загружено из папки:"), info)
    page.wait_for_function("() => UI.state.computed.length > 0", timeout=30000)
    return page.evaluate("SmetaData.get()")


def download(page, name):
    with page.expect_download(timeout=180000) as event:
        page.locator("#bExcel").click()
    path = OUT / name
    event.value.save_as(path)
    return path


def report_contract(path):
    wb = load_workbook(path, read_only=False, data_only=True)
    result = {}
    for ws in wb:
        result[ws.title] = {
            "headers": [c.value for c in ws[1]] if ws.title != "Параметры" else [ws.cell(1, 1).value],
            "freeze": ws.freeze_panes,
            "filter": bool(ws.auto_filter.ref),
            "formats": [ws.cell(2, c).number_format for c in range(1, ws.max_column + 1)],
        }
    wb.close()
    return result


def check_report(path, expected_total, expected_items=None):
    wb = load_workbook(path, read_only=False, data_only=True)
    for name in ("Сметные расчёты", "Позиции", "Контроль", "Параметры", "Сверка итога (Source)"):
        check(name in wb.sheetnames, "лист Excel: " + name)
    ws = wb["Сметные расчёты"]
    headers = [c.value for c in ws[1]]
    idx = headers.index("Всего с НР и СП, руб") + 1
    check(abs(ws.cell(ws.max_row, idx).value - expected_total) < 0.011,
          "итог скачанного Excel равен отмеченным итогам Source")
    check(ws.cell(2, idx).data_type == "n", "денежные ячейки Excel числовые")
    if expected_items is not None:
        check(wb["Позиции"].max_row - 1 == expected_items, "Excel содержит только выбранные позиции")
        control = [tuple(row) for row in wb["Контроль"].values]
        check(any("SOURCE_TOTAL_MISMATCH" in str(row) for row in control),
              "расхождение Source и выбранных позиций есть на листе «Контроль»")
    wb.close()


def source_view(s):
    return {k: s.get(k) for k in ("number", "name", "totals", "items", "resources",
                                  "total", "total_base", "source_work_coefficient")}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--baseline-html", type=Path)
    ap.add_argument("--input", type=Path, default=ROOT / "Исходные данные")
    args = ap.parse_args()
    OUT.mkdir(parents=True, exist_ok=True)
    validate()
    check(True, "проверены состав сборки, пустые данные и хеш скрипта CSP")
    marker = marker_fixture(OUT / "markers.xlsx")
    zero = marker_fixture(OUT / "zero.xlsx", 0, 0)
    cents = marker_fixture(OUT / "cents.xlsx", 150, 150.01)
    no_graph = marker_fixture(OUT / "no-graph.xlsx", graph=False)
    subset = marker_fixture(OUT / "subset.xlsx")
    wb = load_workbook(subset)
    wb["Печатная форма (ГРАФЫ)"].delete_rows(4)
    wb.save(subset)
    wb.close()
    multi = Path(make_multi_ls.build(str(OUT / "multi.xlsx")))
    changed_print = OUT / "multi-changed-print.xlsx"
    wb = load_workbook(multi)
    for ws in wb:
        if ws.title != "Source":
            for row in ws:
                for c in row:
                    if isinstance(c.value, (int, float)):
                        c.value *= 7
    wb.save(changed_print)
    wb.close()
    alien = OUT / "not-an-estimate.xlsx"
    wb = Workbook()
    wb.save(alien)
    wb.close()
    for path in (ROOT / "dist").iterdir():
        check(path.name in ("index.html", "Парсер_смет.html"), "в dist только файлы приложения: " + path.name)
    html = (ROOT / "dist/index.html").read_text()
    check("connect-src 'none'" in html, "CSP запрещает сетевую отправку данных приложением")

    handler = functools.partial(SimpleHTTPRequestHandler, directory=str(ROOT / "dist"))
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        with sync_playwright() as p:
            launch = {"headless": True, "args": ["--no-sandbox"]}
            if os.environ.get("SMETA_CHROMIUM"):
                launch["executable_path"] = os.environ["SMETA_CHROMIUM"]
            elif not Path(p.chromium.executable_path).is_file():
                installed = sorted((Path.home() / ".cache/ms-playwright").glob("chromium-*/chrome-linux*/chrome"))
                if installed:
                    launch["executable_path"] = str(installed[-1])
            browser = p.chromium.launch(**launch)
            page = browser.new_page(accept_downloads=True)
            page.set_default_timeout(30000)
            errors, requests = [], []
            page.on("pageerror", lambda e: errors.append(str(e)))
            page.on("request", lambda req: requests.append((req.method, req.url)))
            page.goto("http://127.0.0.1:%d/" % server.server_port)
            check(page.evaluate("SmetaData.get().smetas.length") == 0, "пустой запуск в браузере")
            seed = page.evaluate("JSON.parse(document.getElementById('smeta-data').textContent)")
            check(seed["smetas"] == [] and seed["checks"] == [] and "files" not in seed["meta"],
                  "в HTML нет скрытых тестовых смет и их реквизитов")

            data = load(page, [marker])
            s = data["smetas"][0]
            check(len(data["smetas"]) == 1 and s["number"] == "MARKER-TEST", "реквизиты читаются из служебной строки")
            check([it["code"] for it in s["items"]] == ["kept-a", "kept-b"], "позиции: только B = 1")
            check([r["code"] for r in s["resources"]] == ["resource-kept"], "ресурсы: только B = 1")
            check(s["totals"]["total"] == 150 and len(s["nodes"]) == 1, "узлы и показатели: только B = 1")
            check(s["source_work_coefficient"] == 1, "неотмеченные условия работ исключены")
            check(s["total"] == 180 and s["total_base"] == 175, "итоги Source имеют приоритет")
            check(s["source_total_calc"] == 150, "контроль считает выбранные позиции отдельно")
            check(page.evaluate("UI.state.computed[0].res.v.total") == 180, "сводка показывает 180, позиции дают 150")
            marker_report = download(page, "markers-report.xlsx")
            check_report(marker_report, 180, 2)
            contract = report_contract(marker_report)
            expected_contract = json.loads((ROOT / "tests/excel_contract.json").read_text())
            check(contract == expected_contract, "формат Excel соответствует исходной версии")
            selected = s["position_selection"]
            check(len(selected["matches"]) == 2, "лист с «граф» в произвольном названии задаёт состав позиций")
            wb = load_workbook(marker_report, data_only=True)
            ph = [c.value for c in wb["Позиции"][1]]
            ti = ph.index("Всего с НР и СП, руб") + 1
            check([wb["Позиции"].cell(r, ti).value for r in (2, 3)] == [100, 50],
                  "стоимость позиций берётся из Source, числа печатной формы не подменяют её")
            wb.close()
            page.evaluate("UI.state.sc.kConv=2; UI.state.dirty={table:1,kpi:1}; UI.render()")
            page.wait_for_function("UI.state.computed[0].res.v.total === 360")
            check(True, "общий коэффициент пересчитывает итог Source")
            page.evaluate("UI.state.sc.kConv=1")
            load(page, [zero])
            check(page.evaluate("UI.state.computed[0].res.v.total") == 0, "нулевой отмеченный итог Source сохраняется")
            data = load(page, [cents])
            check(any(c["code"] == "SOURCE_TOTAL_MISMATCH" for c in data["checks"]), "расхождение в одну копейку видно в контроле")

            data = load(page, [subset])
            check_report(download(page, "subset-report.xlsx"), 180, 1)
            check(len(data["smetas"][0]["items"]) == 2, "фильтр выгрузки сохраняет исходный состав Source")
            check(any(c["code"] == "POSITION_NOT_IN_PRINTED" and "Source!" in c["ref"] for c in data["checks"]),
                  "исключённая позиция со строкой Source видна в контроле")
            check(any(c["code"] == "PRINTED_SELECTION_TOTAL" and c["delta"] == 80 for c in data["checks"]),
                  "контроль показывает разницу итога Source и выгруженных позиций")
            data = load(page, [no_graph])
            check_report(download(page, "no-graph-report.xlsx"), 180, 0)
            check(any(c["code"] == "POSITIONS_NO_GRAPH" for c in data["checks"]),
                  "без листа с «граф» выгружается пустой лист позиций с предупреждением")

            original = load(page, [multi])
            modified = load(page, [changed_print])
            check([source_view(s) for s in original["smetas"]] == [source_view(s) for s in modified["smetas"]],
                  "изменение печатной формы не меняет данные отчёта из Source")
            check(any(c["code"] == "PRINTED_MISMATCH" for c in modified["checks"]), "печатная форма используется для сверки")
            check(all(ref.k_usp_of(s) == s["source_work_coefficient"] for s in modified["smetas"]),
                  "Python-отчёт справочника берёт коэффициент только из Source")
            check(len(original["smetas"]) == 2, "две локальные сметы одной книги разделяются")

            reference_path = ROOT / "Исходные данные/Справочник новый.xlsx"
            if reference_path.is_file():
                load(page, [zero, reference_path])
                check(page.evaluate("Math.abs(UI.refReports().fo.stats.cost) < 0.005"),
                      "отчёт по физобъёмам сохраняет нулевой итог Source")

            form17 = ROOT / "example/fixture-001.xlsx"
            if form17.is_file():
                data17 = load(page, [form17])
                matches17 = [m for s in data17["smetas"] for m in s["position_selection"]["matches"]]
                check(matches17 and all("17 граф" in m["printed_sheet"].lower() for m in matches17),
                      "реальная форма 17 граф: состав позиций сопоставлен с Source")
                total17 = sum(s["source_total"]["value"] if s.get("source_total") else s["total"]
                              for s in data17["smetas"])
                report17 = download(page, "17-graph-report.xlsx")
                check_report(report17, total17)
                wb = load_workbook(report17, data_only=True)
                check(wb["Позиции"].max_row - 1 == len(matches17),
                      f"Excel формы 17 граф содержит {len(matches17)} сопоставленных позиций")
                wb.close()

            paths = sorted(args.input.glob("*.xlsx"))
            data = load(page, paths + [alien])
            if args.input.resolve() == ROOT / "Исходные данные":
                check(len(data["smetas"]) == 16, "все исходные книги: 16 локальных смет")
                check(sum(len(s["items"]) for s in data["smetas"]) == 876, "все исходные книги: 876 выбранных позиций")
            check(any("not-an-estimate" in w for w in data["meta"]["warnings"]), "посторонняя книга не мешает загрузке")
            (OUT / "browser-dataset.json").write_text(json.dumps(data, ensure_ascii=False))
            total = page.evaluate("UI.state.computed.reduce((a,c)=>a+c.res.v.total,0)")
            expected_total = sum(s["source_total"]["value"] if s.get("source_total") else s["total"]
                                 for s in data["smetas"])
            check(abs(total - expected_total) < 0.011, "все сводные строки равны отмеченным итогам Source")
            report = download(page, "all-estimates-report.xlsx")
            check_report(report, expected_total)
            wb = load_workbook(report, data_only=True)
            selected_count = sum(len(s["position_selection"]["matches"]) for s in data["smetas"])
            check(wb["Позиции"].max_row - 1 == selected_count, "Excel содержит ровно сопоставленные позиции")
            if args.input.resolve() == ROOT / "Исходные данные":
                check(selected_count == 735, "в исходных книгах подтверждено 735 из 876 позиций Source")
                counts = {s["number"]: len(s["position_selection"]["matches"]) for s in data["smetas"]
                          if s["file"].startswith("320-")}
                check(sorted(counts.values()) == [0, 12], "добавочная ЛС без позиций печатной формы исключена")
            wb.close()
            check(all(method == "GET" and url.endswith("/") for method, url in requests),
                  "загрузка и экспорт не отправляют файлы или отчёты по сети")
            check(not errors, "нет ошибок JavaScript: " + "; ".join(errors))

            if args.baseline_html:
                page.goto(args.baseline_html.resolve().as_uri())
                load(page, [marker])
                before = report_contract(download(page, "baseline-report.xlsx"))
                check(contract == before, "формат Excel сохранён: листы, столбцы, закрепления, фильтры, форматы чисел")
            browser.close()
    finally:
        server.shutdown()
        server.server_close()

    check(True, "сверка браузерного и Python-разбора начата")
    py_data = sp.build_dataset([str(p) for p in paths], include_items=True)
    (OUT / "python-dataset.json").write_text(json.dumps(py_data, ensure_ascii=False))
    py_by_id = {s["id"]: s for s in py_data["smetas"]}
    check({s["id"] for s in data["smetas"]} == set(py_by_id), "Python/браузер: одинаковый набор смет")
    for js in data["smetas"]:
        py = py_by_id[js["id"]]
        check(js["number"] == py["number"] and len(js["items"]) == len(py["items"]),
              "Python/браузер: состав " + js["number"])
        check(js["position_selection"] == py["position_selection"],
              "Python/браузер: сопоставление с печатной формой " + js["number"])
        for key in ("total", "total_base", "source_total_calc", "source_work_coefficient"):
            check(abs(js[key] - py[key]) < 0.011, "Python/браузер: " + key + " " + js["number"])
    print("PASS: вся браузерная приёмка завершена", flush=True)


if __name__ == "__main__":
    main()
