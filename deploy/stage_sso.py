#!/usr/bin/env python3
"""Разместить проверенные неизменяемые файлы; не включать службы и маршруты."""
import hashlib
import json
from pathlib import Path
import shutil
from prepare_sso import ROOT, TABEL_GATEWAY, GATEWAY_DIGEST, verify_gateway


def put(source, target):
    if target.exists():
        if target.is_symlink() or hashlib.sha256(target.read_bytes()).digest() != hashlib.sha256(source.read_bytes()).digest():
            raise ValueError("существующий файл отличается: " + str(target))
        return
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, target)
    target.chmod(0o444)


def main():
    manifest = verify_gateway()
    receipt = json.loads((ROOT / "deploy/prepared/receipt.json").read_text())
    gateway = Path("/opt/smeta-parser/gateway") / GATEWAY_DIGEST
    release = Path("/opt/smeta-parser/releases") / receipt["web_sha256"]
    for name in manifest["files"]:
        put(TABEL_GATEWAY / name, gateway / name)
    put(TABEL_GATEWAY / "manifest.json", gateway / "manifest.json")
    verify_gateway(gateway)
    for name in ("index.html", "Парсер_смет.html"):
        source = ROOT / "dist" / name
        if hashlib.sha256(source.read_bytes()).hexdigest() != receipt["web_sha256"]:
            raise ValueError("сборка изменилась после подготовки конфигурации")
        put(source, release / name)
    staged = {"schema": "eks.smeta.sso-stage/v1", "gateway": str(gateway), "release": str(release),
              "gateway_manifest_sha256": GATEWAY_DIGEST, "web_sha256": receipt["web_sha256"],
              "credentials_created": False, "activated": False}
    (ROOT / "deploy/prepared/stage-receipt.json").write_text(json.dumps(staged, indent=2) + "\n")
    print("Проверенные файлы размещены в %s; службы и маршруты не включены" % release)


if __name__ == "__main__":
    main()
