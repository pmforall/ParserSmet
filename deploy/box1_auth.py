#!/usr/bin/env python3
"""Публикация через существующий стандартный вход Табеля на box1.

Nginx запрашивает только /tabel/auth/session. Native-токены остаются в штатном
шлюзе; у парсера нет серверной обработки книг, ключей и собственной базы сессий.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from package_web import validate

BASE = "/smeta-parser"
BRIDGE = "/tabel/smeta-parser/"
AUTH_CHECK = "/_smeta_parser_session"


def nginx_config(release, *, host="box1.cii-eks.ru", port=18462):
    if not re.fullmatch(r"[a-z0-9.:-]+", host) or not 1024 <= port <= 65535:
        raise ValueError("некорректный адрес штатного шлюза")
    release = Path(release).resolve()
    html = (release / "index.html").read_text()
    policy = re.search(r'<meta http-equiv="Content-Security-Policy" content="([^"]+)">', html).group(1)
    policy += "; frame-ancestors 'none'"
    gate = f"""
    auth_request {AUTH_CHECK};
    error_page 401 = @smeta_parser_login;
    error_page 500 = @smeta_parser_identity_unavailable;
    limit_except GET {{ deny all; }}
    autoindex off;
    charset utf-8;
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
"""
    return f"""# Парсер смет: выдача HTML только после проверки штатной native-сессии ЕКС.
location = {BASE} {{ absolute_redirect off; return 303 {BASE}/; }}
# Фиксированный адрес возврата лежит в зарегистрированном пространстве Табеля.
location = {BRIDGE} {{ absolute_redirect off; return 303 {BASE}/; }}
location = {BASE}/ {{{gate}
    alias {release}/;
    index index.html;
}}
location = {BASE}/index.html {{{gate}
    alias {release}/index.html;
}}
location = {BASE}/Парсер_смет.html {{{gate}
    alias {release}/Парсер_смет.html;
}}
location ^~ {BASE}/ {{ return 404; }}
location = {AUTH_CHECK} {{
    internal;
    proxy_pass http://127.0.0.1:{port}/tabel/auth/session;
    proxy_http_version 1.1;
    proxy_method GET;
    proxy_pass_request_body off;
    proxy_pass_request_headers off;
    proxy_set_header Host {host};
    proxy_set_header Cookie $http_cookie;
    proxy_set_header Accept application/json;
    proxy_set_header Content-Length "";
    proxy_set_header Connection "";
    proxy_connect_timeout 2s;
    proxy_read_timeout 15s;
    proxy_send_timeout 5s;
    proxy_cache off;
    proxy_intercept_errors off;
    access_log off;
    error_log /dev/null crit;
}}
location @smeta_parser_login {{
    absolute_redirect off;
    add_header Cache-Control "no-store" always;
    add_header Referrer-Policy "no-referrer" always;
    return 303 /tabel/auth/login?return_to=%2Ftabel%2Fsmeta-parser%2F;
}}
location @smeta_parser_identity_unavailable {{
    charset utf-8;
    default_type text/plain;
    add_header Cache-Control "no-store" always;
    return 503 "Сервис входа временно недоступен. Повторите позже.\\n";
}}
"""


def main():
    raise SystemExit("Связанный вход Табеля отменён пользователем. Используйте deploy/prepare_sso.py для самостоятельного входа парсера.")
    validate()
    digest = hashlib.sha256((ROOT / "dist/index.html").read_bytes()).hexdigest()
    release = Path("/opt/smeta-parser/releases") / digest
    for name in ("index.html", "Парсер_смет.html"):
        if (release / name).is_symlink() or hashlib.sha256((release / name).read_bytes()).hexdigest() != digest:
            raise ValueError("не совпадает размещённый HTML: " + name)
    cfg = json.loads(Path("/run/credentials/cii-tabel-gateway.service/gateway.json").read_text())
    service = next(s for s in cfg["services"] if s["name"] == "tabel")
    if (cfg["issuer"] != "https://id.cii-eks.ru/realms/cii" or cfg["listen_host"] != "127.0.0.1"
            or cfg["listen_port"] != 18462 or service["origin"] != "https://box1.cii-eks.ru"
            or service["client_id"] != "cii-tabel" or service["minimum_level"] != 1
            or service["base_path"] != "/tabel" or service["upstream_auth"] != "oidc_bearer"):
        raise ValueError("эффективная конфигурация Табеля изменилась")
    if any(BRIDGE.startswith(p.rstrip("/") + "/") for p in service.get("blocked_return_paths", [])):
        raise ValueError("возврат к парсеру запрещён политикой входа")
    out = ROOT / "deploy/prepared/box1-smeta-parser.conf"
    out.write_text(nginx_config(release))
    receipt = {"schema": "eks.smeta.box1-session-gate/v1", "url": "https://box1.cii-eks.ru/smeta-parser/",
               "auth_mode": "tabel-native-session-auth-request", "issuer": cfg["issuer"],
               "authentication_client": "cii-tabel", "minimum_level": 1,
               "credentials_copied": False, "release": str(release), "web_sha256": digest,
               "nginx_sha256": hashlib.sha256(out.read_bytes()).hexdigest(), "activated": False}
    out.with_suffix(".json").write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n")
    print("Подготовлен box1: проверка действующей сессии Табеля, три HTML-пути; ключи не копировались")


if __name__ == "__main__":
    main()
