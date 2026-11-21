#!/usr/bin/env python3
"""Установить конфигурацию самостоятельного входа; запуск только после регистрации IdP."""
import json
import os
from pathlib import Path
import secrets
import shutil
import stat
import subprocess

ROOT = Path(__file__).resolve().parents[1]
PREPARED = ROOT / "deploy/prepared"


def install(source, destination, mode):
    content = source.read_bytes()
    if destination.exists():
        if destination.is_symlink() or destination.read_bytes() != content:
            raise ValueError("существующий файл отличается: " + str(destination))
    else:
        fd = os.open(destination, os.O_CREAT | os.O_EXCL | os.O_WRONLY, mode)
        with os.fdopen(fd, "wb") as stream:
            stream.write(content)
    destination.chmod(mode)


def main():
    if os.geteuid() != 0:
        raise SystemExit("установка выполняется локальным root")
    policy = json.loads((PREPARED / "client-registration.json").read_text())
    assert policy["clientId"] == "cii-smeta-parser" and policy["name"] == "Парсер смет"
    cfg = json.loads((PREPARED / "gateway.json").read_text())
    assert cfg["services"][0]["client_id"] == "cii-smeta-parser"
    assert "/tabel" not in json.dumps(cfg)
    folder = Path("/etc/smeta-parser")
    folder.mkdir(mode=0o700, exist_ok=True)
    if folder.is_symlink() or folder.stat().st_uid != 0:
        raise ValueError("некорректный каталог конфигурации")
    folder.chmod(0o700)
    key = folder / "store-key"
    if not key.exists():
        fd = os.open(key, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        with os.fdopen(fd, "wb") as stream:
            stream.write(secrets.token_bytes(32))
    info = key.lstat()
    if not stat.S_ISREG(info.st_mode) or stat.S_IMODE(info.st_mode) != 0o600 or info.st_size != 32 or info.st_uid != 0:
        raise ValueError("некорректный ключ хранилища")
    install(PREPARED / "gateway.json", folder / "gateway.json", 0o600)
    install(PREPARED / "client-registration.json", folder / "client-registration.json", 0o600)
    unit = PREPARED / "smeta-parser-gateway.service"
    assert "ConditionPathExists=/etc/smeta-parser/client-registration.ok" in unit.read_text()
    assert "cii-tabel-gateway" not in unit.read_text()
    install(unit, Path("/etc/systemd/system/smeta-parser-gateway.service"), 0o644)
    subprocess.run(["systemctl", "daemon-reload"], check=True)
    subprocess.run(["systemd-analyze", "verify", str(unit)], check=True)
    receipt = {"schema": "eks.smeta.standalone-installed/v1", "client_id": policy["clientId"],
        "display_name": policy["name"], "issuer": cfg["issuer"], "minimum_level": 1,
        "separate_store_key_created": True, "client_secret_installed": (folder / "client-secret").exists(),
        "idp_registration_confirmed": (folder / "client-registration.ok").exists(),
        "started": subprocess.run(["systemctl", "is-active", "--quiet", "smeta-parser-gateway.service"]).returncode == 0}
    (PREPARED / "standalone-installed.json").write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n")
    print("Самостоятельный шлюз установлен; запуск защищён подтверждением регистрации клиента в ЕКС")


if __name__ == "__main__":
    main()
