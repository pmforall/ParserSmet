#!/usr/bin/env python3
"""Включить только маршрут парсера в box1 с резервной копией и проверками."""
from __future__ import annotations

from datetime import datetime, timezone
import fcntl
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import time
import urllib.error
import urllib.request
from urllib.parse import quote, urlsplit

from box1_auth import ROOT, BASE, BRIDGE, AUTH_CHECK

SITE = Path("/etc/nginx/sites-available/box1.cii-eks.ru")
SNIPPET = Path("/etc/nginx/snippets/smeta-parser.conf")
PREPARED = ROOT / "deploy/prepared/box1-smeta-parser.conf"
BLOCK = "    # BEGIN SMETA PARSER APP\n    include /etc/nginx/snippets/smeta-parser.conf;\n    # END SMETA PARSER APP\n\n"


def sha(data):
    return hashlib.sha256(data).hexdigest()


def atomic(path, content, mode=0o644):
    fd, temp_name = tempfile.mkstemp(prefix=path.name + ".smeta-", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as output:
            output.write(content)
            output.flush()
            os.fsync(output.fileno())
            os.fchmod(output.fileno(), mode)
        os.replace(temp_name, path)
    finally:
        Path(temp_name).unlink(missing_ok=True)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def status(path):
    opener = urllib.request.build_opener(NoRedirect)
    request = urllib.request.Request("https://box1.cii-eks.ru" + quote(path, safe="/"),
        headers={"Sec-Fetch-Mode": "navigate", "Accept": "text/html"})
    try:
        response = opener.open(request, timeout=20)
    except urllib.error.HTTPError as error:
        response = error
    with response:
        body = response.read(512000)
        location = response.headers.get("Location", "")
        return response.code, location, body


def main():
    raise SystemExit("Активация через сессию Табеля отменена пользователем. Используйте самостоятельный шлюз парсера.")
    if os.geteuid() != 0:
        raise SystemExit("нужен локальный root для конфигурации nginx")
    with Path("/run/smeta-parser-deploy.lock").open("w") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        proof = (ROOT / "tests/box1-auth.log").read_text()
        if "Ran 5 tests" not in proof or not proof.rstrip().endswith("OK"):
            raise SystemExit("браузерная проверка nginx + штатного входа не прошла")
        settings = json.loads(PREPARED.with_suffix(".json").read_text())
        expected = PREPARED.read_bytes()
        if sha(expected) != settings["nginx_sha256"]:
            raise SystemExit("подготовленная конфигурация изменилась")
        subprocess.run(["nginx", "-t"], check=True)
        before_pid = subprocess.check_output(["systemctl", "show", "cii-tabel-gateway.service", "--property=MainPID", "--value"], text=True).strip()
        if not before_pid or before_pid == "0":
            raise SystemExit("штатный шлюз Табеля не запущен")
        original = SITE.read_bytes()
        if BLOCK.encode() in original:
            if not SNIPPET.is_file() or SNIPPET.read_bytes() != expected:
                raise SystemExit("существующий маршрут отличается; требуется проверка")
            print("Маршрут уже установлен; без изменений")
            return
        if SNIPPET.exists():
            raise SystemExit("файл назначения уже существует; он не будет перезаписан")
        marker = b"    # BEGIN TABEL APP\n"
        if original.count(marker) != 1:
            raise SystemExit("не найдена однозначная точка добавления маршрута")
        updated = original.replace(marker, BLOCK.encode() + marker, 1)
        if updated.replace(BLOCK.encode(), b"", 1) != original:
            raise SystemExit("не прошла проверка точного отката")
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        backup = ROOT.parent / "backups" / ("box1-activation-" + stamp)
        backup.mkdir(parents=True, mode=0o700)
        shutil.copy2(SITE, backup / "box1.cii-eks.ru")
        if sha((backup / "box1.cii-eks.ru").read_bytes()) != sha(original):
            raise SystemExit("резервная копия не совпадает")
        changed = False
        try:
            atomic(SNIPPET, expected)
            if SITE.read_bytes() != original:
                raise RuntimeError("конфигурация box1 параллельно изменилась")
            atomic(SITE, updated, SITE.stat().st_mode & 0o777)
            changed = True
            subprocess.run(["nginx", "-t"], check=True)
            subprocess.run(["systemctl", "reload", "nginx"], check=True)
            subprocess.run(["systemctl", "is-active", "--quiet", "nginx"], check=True)
            # Reload signals the master before new workers accept connections.
            # Wait for the new route before deciding whether activation failed.
            ready_at = time.monotonic() + 10
            while True:
                ready_code, ready_location, _ = status(BASE + "/")
                if ready_code == 303 and ready_location == "/tabel/auth/login?return_to=%2Ftabel%2Fsmeta-parser%2F":
                    break
                if time.monotonic() >= ready_at:
                    raise RuntimeError("маршрут не готов: HTTP %d, redirect path %s" %
                        (ready_code, urlsplit(ready_location).path))
                time.sleep(0.25)
            checks = []
            for path in (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"):
                code, location, body = status(path)
                if code != 303 or location != "/tabel/auth/login?return_to=%2Ftabel%2Fsmeta-parser%2F" or b"smeta-data" in body:
                    raise RuntimeError("публичная проверка входа: %s HTTP %d, redirect path %s" %
                        (path, code, urlsplit(location).path))
                checks.append({"path": path, "status": code, "login": "/tabel/auth/login"})
            for path in (AUTH_CHECK, BASE + "/data/smetas.json", BASE + "/tests/out/report.xlsx", BASE + "/Исходные данные/test.xlsx"):
                code, _, _ = status(path)
                if code != 404:
                    raise RuntimeError("не пройдена проверка закрытых файлов")
                checks.append({"path": path, "status": code})
            code, location, _ = status(BRIDGE)
            if code != 303 or location != BASE + "/":
                raise RuntimeError("не пройдена проверка возврата после входа")
            after_pid = subprocess.check_output(["systemctl", "show", "cii-tabel-gateway.service", "--property=MainPID", "--value"], text=True).strip()
            if before_pid != after_pid:
                raise RuntimeError("штатный шлюз Табеля изменился при включении")
            settings.update(activated=True, activated_at=stamp, backup=str(backup),
                box1_before_sha256=sha(original), box1_after_sha256=sha(updated),
                installed_snippet_sha256=sha(SNIPPET.read_bytes()), unchanged_tabel_gateway_pid=int(after_pid),
                public_checks=checks, rollback_bytes_verified=True)
            (ROOT / "deploy/prepared/box1-live-receipt.json").write_text(json.dumps(settings, ensure_ascii=False, indent=2) + "\n")
            print("Включено: https://box1.cii-eks.ru/smeta-parser/")
        except BaseException:
            if changed:
                if SITE.read_bytes() != updated:
                    raise RuntimeError("конфигурация изменилась параллельно; автоматический откат остановлен") from None
                atomic(SITE, original)
            if SNIPPET.is_file() and SNIPPET.read_bytes() == expected:
                SNIPPET.unlink()
            subprocess.run(["nginx", "-t"], check=True)
            subprocess.run(["systemctl", "reload", "nginx"], check=True)
            raise


if __name__ == "__main__":
    main()
