#!/usr/bin/env python3
"""Проверить и упаковать только веб-файлы, без смет и отчётов."""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
from pathlib import Path
import re
import sys
import tarfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from build import META_KEYS

PUBLIC_FILES = {"index.html", "Парсер_смет.html"}


def validate(directory=ROOT / "dist"):
    paths = list(directory.iterdir())
    if {p.name for p in paths} != PUBLIC_FILES or any(p.is_symlink() or not p.is_file() for p in paths):
        raise ValueError("dist должен содержать только index.html и Парсер_смет.html без символьных ссылок")
    schema = json.loads((ROOT / "app/schema.json").read_text())
    expected = {"meta": {k: schema["meta"][k] for k in META_KEYS if k in schema["meta"]},
                "smetas": [], "checks": []}
    contents = []
    for path in paths:
        html = path.read_text()
        blocks = re.findall(r"<script[^>]*>(.*?)</script>", html, re.S)
        if len(blocks) != 2 or json.loads(blocks[0]) != expected:
            raise ValueError("в HTML должны быть только схема столбцов и код, без данных смет")
        digest = base64.b64encode(hashlib.sha256(blocks[1].encode()).digest()).decode()
        if "'sha256-%s'" % digest not in html or "connect-src 'none'" not in html:
            raise ValueError("CSP не соответствует коду либо разрешает передачу данных по сети")
        if re.search(r'<script[^>]+\bsrc=|\bfetch\s*\(|XMLHttpRequest|importScripts|@import', html, re.I):
            raise ValueError("найдена внешняя зависимость или сетевой вызов")
        contents.append(path.read_bytes())
    if contents[0] != contents[1]:
        raise ValueError("основной и совместимый HTML отличаются")
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(paths)}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-o", "--out", type=Path, default=ROOT.parent / "releases/smeta-parser-web.tar.gz")
    ap.add_argument("--check", action="store_true", help="только проверить dist")
    args = ap.parse_args()
    if args.out.resolve().is_relative_to(ROOT / "dist"):
        ap.error("архив должен находиться вне каталога сайта")
    digests = validate()
    if args.check:
        print("OK: dist содержит только проверенные файлы приложения, без смет и отчётов")
        return
    args.out.parent.mkdir(parents=True, exist_ok=True)
    temporary = args.out.with_name(args.out.name + ".tmp")
    with tarfile.open(temporary, "w:gz") as archive:
        for name in sorted(digests):
            archive.add(ROOT / "dist" / name, arcname=name, recursive=False)
    temporary.replace(args.out)
    receipt = {"files": digests, "archive_sha256": hashlib.sha256(args.out.read_bytes()).hexdigest()}
    args.out.with_suffix(args.out.suffix + ".json").write_text(json.dumps(receipt, indent=2) + "\n")
    print("Пакет сайта: %s" % args.out)


if __name__ == "__main__":
    main()
