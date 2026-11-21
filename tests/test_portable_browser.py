#!/usr/bin/env python3
"""Браузерная приёмка только на синтетических сметах, без закрытых книг."""
import functools
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import os
from pathlib import Path
import sys
import threading

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))
import test_browser_contract as browser_helpers


def main():
    out = ROOT / "tests/out/portable-browser"
    out.mkdir(parents=True, exist_ok=True)
    browser_helpers.OUT = out
    browser_helpers.validate()
    original = browser_helpers.marker_fixture(out / "synthetic.xlsx")
    subset = browser_helpers.marker_fixture(out / "synthetic-subset.xlsx")
    no_graph = browser_helpers.marker_fixture(out / "synthetic-no-graph.xlsx", graph=False)
    wb = load_workbook(subset)
    wb["Печатная форма (ГРАФЫ)"].delete_rows(4)
    wb.save(subset)
    wb.close()
    handler = functools.partial(SimpleHTTPRequestHandler, directory=str(ROOT / "dist"))
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    worker = threading.Thread(target=server.serve_forever, daemon=True)
    worker.start()
    errors, requests = [], []
    try:
        with sync_playwright() as p:
            options = {"headless": True, "args": ["--no-sandbox"]}
            if os.environ.get("SMETA_CHROMIUM"):
                options["executable_path"] = os.environ["SMETA_CHROMIUM"]
            elif not Path(p.chromium.executable_path).is_file():
                candidates = sorted((Path.home() / ".cache/ms-playwright").glob("chromium-*/chrome-linux*/chrome"))
                if candidates:
                    options["executable_path"] = str(candidates[-1])
            browser = p.chromium.launch(**options)
            page = browser.new_page(accept_downloads=True)
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.on("request", lambda req: requests.append((req.method, req.url)))
            url = "http://127.0.0.1:%d/" % server.server_port
            page.goto(url)
            assert page.evaluate("SmetaData.get().smetas.length") == 0
            for source, expected in ((original, 2), (subset, 1), (no_graph, 0)):
                data = browser_helpers.load(page, [source])
                smeta = data["smetas"][0]
                assert len(smeta["items"]) == 2
                assert len(smeta["position_selection"]["matches"]) == expected
                report = browser_helpers.download(page, "report-" + source.name)
                browser_helpers.check_report(report, 180, expected)
                wb = load_workbook(report, data_only=True)
                headers = [c.value for c in wb["Позиции"][1]]
                column = headers.index("Всего с НР и СП, руб") + 1
                amounts = [wb["Позиции"].cell(r, column).value for r in range(2, wb["Позиции"].max_row + 1)]
                assert amounts == [100, 50][:expected], amounts
                if expected == 1:
                    assert any(c["code"] == "PRINTED_SELECTION_TOTAL" and c["delta"] == 80 for c in data["checks"])
                if expected == 0:
                    assert any(c["code"] == "POSITIONS_NO_GRAPH" for c in data["checks"])
                wb.close()
            assert not errors, errors
            assert requests and all(method == "GET" and target == url for method, target in requests), requests
            browser.close()
    finally:
        server.shutdown()
        server.server_close()
        worker.join(timeout=5)
    print("PASS: синтетические XLSX → браузер → Excel; B = 1, состав по «граф», итог Source и контроль")
    print("PASS: сметы и отчёты не отправляются по сети")


if __name__ == "__main__":
    main()
