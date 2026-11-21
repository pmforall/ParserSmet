#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка веб-приложения без исходных смет и отчётов.

    python build.py

dist содержит HTML, код и описание столбцов из app/schema.json.
Для закрытых регрессионных стендов: --data FILE --with-data --out tests/out/fixture.html.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import re
from pathlib import Path

TOKEN = "/*__SMETA_DATA__*/"
ROOT = Path(__file__).resolve().parent
META_KEYS = ("app", "version", "plugin", "source_sheet", "printed_mark", "verification",
             "price_level", "indicators", "table_columns", "category_labels")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-d", "--data", type=Path, default=ROOT / "app/schema.json")
    ap.add_argument("-t", "--template", type=Path, default=ROOT / "app/template.html")
    ap.add_argument("-o", "--out", type=Path)
    ap.add_argument("--with-data", action="store_true",
                    help="включить сметы только в закрытый тестовый стенд вне dist")
    args = ap.parse_args(argv)
    outputs = [args.out] if args.out else [ROOT / "dist/index.html", ROOT / "dist/Парсер_смет.html"]
    if args.with_data and any(p.resolve().is_relative_to(ROOT / "dist") for p in outputs):
        ap.error("сметы нельзя включать в dist; укажите --out вне каталога сайта")
    data = json.loads(args.data.read_text(encoding="utf-8"))
    if args.with_data:
        data.pop("reference_report", None)
    else:
        data = {"meta": {k: data["meta"][k] for k in META_KEYS if k in data["meta"]},
                "smetas": [], "checks": []}
    html = args.template.read_text(encoding="utf-8")
    if html.count(TOKEN) != 1:
        ap.error("шаблон должен содержать одну метку %s" % TOKEN)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    for char, replacement in (("<", "\\u003c"), (">", "\\u003e"), ("&", "\\u0026")):
        payload = payload.replace(char, replacement)
    result = html.replace(TOKEN, payload)
    scripts = re.findall(r"<script(?![^>]*type=\"application/json\")[^>]*>(.*?)</script>", result, re.S)
    hashes = " ".join("'sha256-%s'" % base64.b64encode(hashlib.sha256(s.encode()).digest()).decode()
                      for s in scripts)
    policy = ("default-src 'none'; script-src " + hashes + "; style-src 'unsafe-inline'; "
              "img-src data: blob:; connect-src 'none'; object-src 'none'; "
              "base-uri 'none'; form-action 'none'")
    result = result.replace('<meta charset="utf-8">', '<meta charset="utf-8">\n'
                            '<meta http-equiv="Content-Security-Policy" content="' + policy + '">', 1)
    for output in outputs:
        output.parent.mkdir(parents=True, exist_ok=True)
        temporary = output.with_name(output.name + ".tmp")
        temporary.write_text(result, encoding="utf-8")
        temporary.replace(output)
        print("Сборка: %s (%.1f КБ), смет: %d" %
              (output, output.stat().st_size / 1024, len(data["smetas"])))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
