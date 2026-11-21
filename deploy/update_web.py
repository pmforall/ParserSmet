#!/usr/bin/env python3
"""Обновить только HTML и CSP самостоятельного «Парсера смет», с откатом.

До запуска: build.py, prepare_sso.py, stage_sso.py и проверки текущей сборки.
Клиент ЕКС, ключи, сессии и службы других приложений не изменяются.
"""
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time
from urllib.parse import parse_qs, quote, urlsplit

from activate_standalone import request
from prepare_sso import BASE, CLIENT, ORIGIN, ROOT, gateway_config, nginx_location, verify_gateway, GATEWAY_DIGEST
from package_web import validate

PREPARED = ROOT / "deploy/prepared"
CONFIG = Path("/etc/smeta-parser/gateway.json")
SNIPPET = Path("/etc/nginx/snippets/smeta-parser.conf")
SERVICE = "smeta-parser-gateway.service"


def sha(data):
    return hashlib.sha256(data).hexdigest()


def replace(path, content, mode):
    temporary = path.with_name(path.name + ".smeta-update-tmp")
    fd = os.open(temporary, os.O_WRONLY | os.O_CREAT | os.O_EXCL, mode)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def ready():
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        try:
            status, _, data = request("http://127.0.0.1:18468" + BASE + "/auth/session",
                headers={"Host": "box1.cii-eks.ru", "Accept": "application/json"})
            if status == 401 and json.loads(data).get("login_url", "").startswith(BASE + "/auth/login"):
                return
        except (OSError, ValueError):
            pass
        time.sleep(0.2)
    raise RuntimeError("самостоятельный шлюз не готов")


def main():
    if os.geteuid() != 0:
        raise SystemExit("требуется локальный root на ekskz")
    hashes = validate()
    digest = hashes["index.html"]
    release = Path("/opt/smeta-parser/releases") / digest
    if set(p.name for p in release.iterdir()) != set(hashes):
        raise ValueError("неожиданный состав публичного релиза")
    for name, expected in hashes.items():
        file = release / name
        if file.is_symlink() or sha(file.read_bytes()) != expected:
            raise ValueError("не совпадает подготовленный HTML")
    proof = json.loads((PREPARED / "web-tests.json").read_text())
    if proof.get("web_sha256") != digest or not all(proof.get(k) is True for k in
            ("browser", "excel_format", "python_parity", "position_matching", "gateway", "nginx_browser")):
        raise ValueError("нет проверок текущей сборки")
    verify_gateway(Path("/opt/smeta-parser/gateway") / GATEWAY_DIGEST)
    if CONFIG.is_symlink() or SNIPPET.is_symlink():
        raise ValueError("некорректный файл конфигурации")
    old_config, old_snippet = CONFIG.read_bytes(), SNIPPET.read_bytes()
    old = json.loads(old_config)
    old_release = Path(old["services"][0]["routes"][0]["static_files"][BASE + "/"]).parent
    if old != gateway_config(old_release):
        raise ValueError("текущая политика шлюза отличается; автоматическое обновление остановлено")
    new_config = (PREPARED / "gateway.json").read_bytes()
    new_snippet = (PREPARED / "nginx-location.conf").read_bytes()
    if json.loads(new_config) != gateway_config(release) or new_snippet.decode() != nginx_location((release / "index.html").read_text()):
        raise ValueError("подготовленная конфигурация отличается от проверенной")
    subprocess.run(["nginx", "-t"], check=True)
    subprocess.run(["systemctl", "is-active", "--quiet", SERVICE], check=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup = ROOT.parent / "backups" / ("web-update-" + stamp)
    backup.mkdir(mode=0o700)
    for name, content in (("gateway.json", old_config), ("smeta-parser.conf", old_snippet)):
        file = backup / name
        file.write_bytes(content)
        file.chmod(0o600)
    pid_args = ["systemctl", "show", "cii-tabel-gateway.service", "--property=MainPID", "--value"]
    other_pid = subprocess.check_output(pid_args, text=True).strip()
    try:
        replace(CONFIG, new_config, 0o600)
        replace(SNIPPET, new_snippet, 0o644)
        subprocess.run(["nginx", "-t"], check=True)
        subprocess.run(["systemctl", "restart", SERVICE], check=True)
        ready()
        subprocess.run(["systemctl", "reload", "nginx"], check=True)
        # reload завершается до смены workers: ожидаем CSP именно нового HTML.
        deadline = time.monotonic() + 15
        while True:
            status, headers, _ = request(ORIGIN + BASE + "/", headers={"Sec-Fetch-Mode": "navigate", "Accept": "text/html"})
            csp = headers.get("Content-Security-Policy", "")
            if status == 303 and csp and csp in new_snippet.decode():
                break
            if time.monotonic() > deadline:
                raise RuntimeError("публичный маршрут не переключился на новую CSP")
            time.sleep(0.2)
        checks = []
        for path in (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"):
            status, headers, body = request(ORIGIN + quote(path), headers={"Sec-Fetch-Mode": "navigate", "Accept": "text/html"})
            target = urlsplit(headers.get("Location", ""))
            if status != 303 or target.path != BASE + "/auth/login" or b"smeta-data" in body:
                raise RuntimeError("не выполнена защита входа")
            checks.append(dict(path=path, status=status, login_path=target.path))
        status, headers, _ = request(ORIGIN + BASE + "/auth/login?return_to=%2Fsmeta-parser%2F")
        target = urlsplit(headers.get("Location", ""))
        query = parse_qs(target.query)
        if status != 302 or target.hostname != "id.cii-eks.ru" or query.get("client_id") != [CLIENT] or query.get("code_challenge_method") != ["S256"]:
            raise RuntimeError("не выполнена проверка собственного входа ЕКС")
        if subprocess.check_output(pid_args, text=True).strip() != other_pid:
            raise RuntimeError("изменился процесс другого приложения")
        receipt = dict(schema="eks.smeta.web-update/v1", url=ORIGIN + BASE + "/", activated=True,
                       web_sha256=digest, previous_release=str(old_release), release=str(release),
                       backup=str(backup), updated_at=stamp, tests=proof, public_checks=checks,
                       separate_eks_login=True, other_gateway_pid_unchanged=int(other_pid))
        (PREPARED / "web-update-live-receipt.json").write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n")
        print("Обновлено: " + receipt["url"] + "; SHA256 " + digest)
    except BaseException:
        replace(CONFIG, old_config, 0o600)
        replace(SNIPPET, old_snippet, 0o644)
        subprocess.run(["nginx", "-t"], check=True)
        subprocess.run(["systemctl", "restart", SERVICE], check=True)
        ready()
        subprocess.run(["systemctl", "reload", "nginx"], check=True)
        raise


if __name__ == "__main__":
    main()
