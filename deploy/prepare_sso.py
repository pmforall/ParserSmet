#!/usr/bin/env python3
"""Подготовка отдельного стандартного шлюза ЕКС по работающему «Табелю».

Создаёт только проверяемые файлы в deploy/prepared; ничего не активирует.
Ключи и пароли в подготовленные файлы не включаются.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from package_web import validate

GATEWAY_DIGEST = "273755939965b47b883b9c70f4ad328360ffe82239ccfdaa69e00036de2804b0"
TABEL_GATEWAY = Path("/opt/cii-tabel-gateway/releases") / GATEWAY_DIGEST
ISSUER = "https://id.cii-eks.ru/realms/cii"
ORIGIN = "https://box1.cii-eks.ru"
CLIENT = "cii-smeta-parser"
BASE = "/smeta-parser"
PORT = 18468


def verify_gateway(path=TABEL_GATEWAY):
    manifest_path = path / "manifest.json"
    raw = manifest_path.read_bytes()
    if hashlib.sha256(raw).hexdigest() != GATEWAY_DIGEST:
        raise ValueError("манифест шлюза не соответствует закреплённому релизу Табеля")
    manifest = json.loads(raw)
    for name, expected in manifest["files"].items():
        file = path / name
        if file.is_symlink() or not file.is_file() or hashlib.sha256(file.read_bytes()).hexdigest() != expected:
            raise ValueError("изменён файл закреплённого шлюза: " + name)
    if hashlib.sha256(Path(sys.executable).resolve().read_bytes()).hexdigest() != manifest["python_sha256"]:
        raise ValueError("версия Python отличается от проверенного шлюза")
    return manifest


def gateway_config(release, state=Path("/var/lib/smeta-parser-gateway"),
                   runtime=Path("/run/smeta-parser-gateway")):
    return {"issuer": ISSUER, "listen_host": "127.0.0.1", "listen_port": PORT,
            "store_path": str(state / "sessions.db"), "store_key_file": str(runtime / "store-key"),
            "revocation_seconds": 2, "max_session_seconds": 43200, "services": [{
                "name": "smeta-parser", "origin": ORIGIN, "client_id": CLIENT,
                "client_secret_file": str(runtime / "client-secret"), "minimum_level": 1,
                "base_path": BASE, "home": BASE + "/", "launch_path": BASE + "/auth/launch",
                "launch_return_prefixes": [BASE], "blocked_return_paths": [BASE + "/auth"],
                "upstream_auth": "oidc_bearer", "csrf_mode": "bff", "blocked_paths": [],
                "routes": [{"prefix": BASE, "static_files": {
                    BASE + "/": str(release / "index.html"),
                    BASE + "/index.html": str(release / "index.html"),
                    BASE + "/Парсер_смет.html": str(release / "Парсер_смет.html")}}]}]}


def nginx_location(html, *, host="box1.cii-eks.ru", port=PORT):
    if not re.fullmatch(r"[a-z0-9.:-]+", host) or not 1024 <= port <= 65535:
        raise ValueError("некорректный адрес собственного шлюза")
    policy = re.search(r'<meta http-equiv="Content-Security-Policy" content="([^"]+)">', html).group(1)
    policy += "; frame-ancestors 'none'"
    return f"""# Парсер смет: собственный шлюз, только зарегистрированные HTML-файлы.
location = {BASE} {{ absolute_redirect off; return 303 {BASE}/; }}
location ^~ {BASE}/ {{
    client_max_body_size 16k;
    client_body_timeout 5s;
    proxy_http_version 1.1;
    proxy_set_header Host {host};
    proxy_set_header X-Forwarded-Proto https;
    proxy_set_header X-Forwarded-For $remote_addr;
    proxy_set_header X-Forwarded-Host {host};
    proxy_set_header Connection "";
    proxy_connect_timeout 3s;
    proxy_read_timeout 30s;
    proxy_send_timeout 30s;
    proxy_intercept_errors on;
    error_page 502 503 504 =503 @smeta_parser_identity_unavailable;
    proxy_cache off;
    proxy_redirect off;
    proxy_hide_header Content-Security-Policy;
    proxy_hide_header Cache-Control;
    proxy_hide_header X-Content-Type-Options;
    proxy_hide_header X-Frame-Options;
    proxy_hide_header Referrer-Policy;
    add_header Content-Security-Policy "{policy}" always;
    add_header Cache-Control "no-store" always;
    add_header X-Content-Type-Options "nosniff" always;
    add_header X-Frame-Options "DENY" always;
    add_header Referrer-Policy "no-referrer" always;
    add_header Permissions-Policy "camera=(), microphone=(), geolocation=(), payment=(), usb=()" always;
    add_header Cross-Origin-Opener-Policy "same-origin" always;
    add_header Cross-Origin-Resource-Policy "same-origin" always;
    add_header Strict-Transport-Security "max-age=31536000" always;
    access_log off;
    error_log /dev/null crit;
    proxy_pass http://127.0.0.1:{port};
}}
location @smeta_parser_identity_unavailable {{
    charset utf-8;
    default_type text/plain;
    add_header Cache-Control "no-store" always;
    return 503 "Парсер смет. Сервис входа временно недоступен.\\n";
}}
"""


def client_registration():
    source = TABEL_GATEWAY / "source"
    spec = importlib.util.spec_from_file_location("realm_builder", source / "deploy/production/build_realm.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    pilot = json.loads((source / "config/pilot-realm.json").read_text())
    manifest = {"issuer": ISSUER, "clients": [{"id": CLIENT, "level": 1, "origin": ORIGIN,
        "callback": BASE + "/auth/callback", "launch": BASE + "/auth/launch", "logout": BASE + "/",
        "launch_prefixes": [BASE]}]}
    realm = module.build(pilot, manifest)
    client = next(c for c in realm["clients"] if c["clientId"] == CLIENT)
    client["name"] = "Парсер смет"
    client["attributes"]["cii.display-name"] = "Парсер смет"
    if any(k in client for k in ("secret", "id")):
        raise ValueError("регистрация должна содержать только публичную политику клиента")
    return client


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", type=Path, default=ROOT / "deploy/prepared")
    args = ap.parse_args()
    validate()
    manifest = verify_gateway()
    digest = hashlib.sha256((ROOT / "dist/index.html").read_bytes()).hexdigest()
    release = Path("/opt/smeta-parser/releases") / digest
    runtime = Path("/opt/smeta-parser/gateway") / GATEWAY_DIGEST
    args.out.mkdir(parents=True, exist_ok=True)
    config = gateway_config(release)
    (args.out / "gateway.json").write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n")
    (args.out / "client-registration.json").write_text(json.dumps(client_registration(), ensure_ascii=False, indent=2) + "\n")
    (args.out / "nginx-location.conf").write_text(nginx_location((ROOT / "dist/index.html").read_text()))
    unit = Path("/etc/systemd/system/cii-tabel-gateway.service").read_text()
    unit = unit.replace("Tabel native identity gateway", "Парсер смет — стандартный единый вход ЕКС")
    unit = unit.replace(" tabel-web.service", "")
    unit = unit.replace("Wants=network-online.target", "Wants=network-online.target\nConditionPathExists=/etc/smeta-parser/client-registration.ok")
    unit = unit.replace(str(TABEL_GATEWAY), str(runtime))
    unit = unit.replace("cii-tabel-gateway", "smeta-parser-gateway")
    # Configuration belongs to this service; no Tabel credential is copied.
    unit = unit.replace("/etc/smeta-parser-gateway", "/etc/smeta-parser")
    (args.out / "smeta-parser-gateway.service").write_text(unit)
    receipt = {"schema": "eks.smeta.sso-preparation/v1", "url": ORIGIN + BASE + "/",
        "reference": "tabel", "issuer": ISSUER, "client_id": CLIENT, "minimum_level": 1,
        "audience": "standard-authenticated-eks-users", "gateway_manifest_sha256": GATEWAY_DIGEST,
        "gateway_files_verified": len(manifest["files"]), "web_sha256": digest,
        "credentials_included": False, "activated": False}
    (args.out / "receipt.json").write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n")
    print("Подготовлено: %s; шлюз проверен (%d файлов); секреты не включены" % (args.out, len(manifest["files"])))


if __name__ == "__main__":
    main()
