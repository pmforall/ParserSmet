#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_all.py — сборка сайта и браузерная приёмка с независимым Python-разбором.

    python run_all.py "папка_со_сметами"

Без аргумента используется папка «Исходные данные».
Книги и отчёты проверок сохраняются в tests/out, вне каталога сайта dist.
"""
from __future__ import annotations

import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable


def run(title, args, allow_fail=False):
    print("\n" + "=" * 78)
    print("▶ " + title)
    print("=" * 78, flush=True)
    rc = subprocess.call([PY] + args, cwd=ROOT)
    if rc != 0 and not allow_fail:
        print("\nШаг завершился с кодом %d — цикл остановлен." % rc)
        sys.exit(rc)
    return rc


def main() -> int:
    folder = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "Исходные данные")
    print("Папка со сметами: %s" % folder)

    run("Сборка сайта без встроенных смет", ["build.py"])
    run("Проверка состава веб-сборки", [os.path.join("tools", "package_web.py"), "--check"])
    run("Проверка исключения смет и архивов из сайта", [os.path.join("tests", "test_web_build.py")])
    run("Сопоставление позиций печатной формы и Source", [os.path.join("tests", "test_printed_positions.py")])
    run("Загрузка в Chromium, скачивание Excel и сверка с Python",
        [os.path.join("tests", "test_browser_contract.py"), "--input", os.path.abspath(folder)])
    print("\nГотово. Сайт: dist/index.html. Материалы приёмки: tests/out/browser/.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
