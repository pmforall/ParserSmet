#!/usr/bin/env python3
"""Настоящий nginx + штатный BFF + подписанный тестовый IdP + Chromium.

.venv/bin/python -B tests/test_box1_auth.py
Все аккаунты и ключи синтетические; корпоративный вход не вызывается.
"""
import asyncio
import dataclasses
import hashlib
import json
from pathlib import Path
import secrets
import socket
import subprocess
import sys
import unittest
from urllib.parse import parse_qs, unquote, urlencode, urlsplit

ROOT = Path(__file__).resolve().parents[1]
GATEWAY = Path("/opt/cii-tabel-gateway/releases/273755939965b47b883b9c70f4ad328360ffe82239ccfdaa69e00036de2804b0")
sys.path[:0] = [str(ROOT / "deploy"), str(GATEWAY / "source"), str(GATEWAY / "site-packages")]
import aiohttp
from gateway.app import GATEWAY as GATEWAY_KEY
from gateway.config import Route
from gateway.tests.test_gateway import GatewayIntegration
from prepare_sso import BASE, CLIENT, nginx_location

AUTH_CHECK = "/_smeta_parser_session"
BRIDGE = "/tabel/smeta-parser/"


class Box1AuthTests(GatewayIntegration):
    async def asyncSetUp(self):
        await super().asyncSetUp()
        self.level = 1
        files = tuple((path, str(ROOT / "dist/index.html")) for path in
                      (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"))
        self.service = dataclasses.replace(self.service, name="smeta-parser", client_id=CLIENT,
            minimum_level=1, base_path=BASE, home=BASE + "/", launch_path=BASE + "/auth/launch",
            launch_return_prefixes=(BASE,), blocked_return_paths=(BASE + "/auth",),
            routes=(Route(BASE, static_files=files),))
        gw = self.app[GATEWAY_KEY]
        gw.config = dataclasses.replace(gw.config, services=(self.service,))
        gw.oidc.config = gw.config
        directory = Path(self.directory.name)
        subprocess.run(["openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-days", "1",
            "-subj", "/CN=localhost", "-keyout", str(directory / "test.key"), "-out", str(directory / "test.crt")],
            check=True, capture_output=True)
        with socket.socket() as sock:
            sock.bind(("127.0.0.1", 0))
            port = sock.getsockname()[1]
        self.front = "https://127.0.0.1:%d" % port
        digest = hashlib.sha256((ROOT / "dist/index.html").read_bytes()).hexdigest()
        release = Path("/opt/smeta-parser/releases") / digest
        snippet = nginx_location((release / "index.html").read_text(), host=self.service.authority, port=self.server.port)
        config = directory / "nginx.conf"
        config.write_text("user www-data;\nworker_processes 1;\npid " + str(directory / "nginx.pid")
            + ";\nerror_log stderr crit;\nevents {}\nhttp { include /etc/nginx/mime.types; access_log off;"
            + " server { listen 127.0.0.1:" + str(port) + " ssl; ssl_certificate " + str(directory / "test.crt")
            + "; ssl_certificate_key " + str(directory / "test.key") + ";\n" + snippet
            + "\nlocation / { return 404; }\n} }\n")
        self.nginx_error = (directory / "nginx.stderr").open("wb")
        self.nginx = subprocess.Popen(["nginx", "-c", str(config), "-g", "daemon off;"],
            stdout=subprocess.DEVNULL, stderr=self.nginx_error)
        self.front_client = aiohttp.ClientSession(cookie_jar=aiohttp.DummyCookieJar())
        for _ in range(60):
            if self.nginx.poll() is not None:
                raise RuntimeError((directory / "nginx.stderr").read_text())
            try:
                async with self.front_client.get(self.front + "/", ssl=False) as response:
                    if response.status == 404:
                        break
            except aiohttp.ClientError:
                await asyncio.sleep(0.05)
        else:
            raise RuntimeError("изолированный nginx не запустился")

    async def asyncTearDown(self):
        if hasattr(self, "front_client"):
            await self.front_client.close()
        if hasattr(self, "nginx"):
            self.nginx.terminate()
            await asyncio.to_thread(self.nginx.wait, 10)
            self.nginx_error.close()
        await super().asyncTearDown()

    async def begin(self):
        response = await self.client.get(BASE + "/auth/login?" + urlencode({"return_to": BASE + "/"}), allow_redirects=False)
        self.assertEqual(response.status, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        self.assertEqual(query["client_id"], [CLIENT])
        self.assertEqual(query["code_challenge_method"], ["S256"])
        return query

    async def login(self):
        query = await self.begin()
        code = secrets.token_urlsafe(16)
        self.codes[code] = query
        response = await self.client.get(BASE + "/auth/callback?" + urlencode({"state": query["state"][0],
            "code": code, "iss": self.issuer}), allow_redirects=False)
        self.assertEqual(response.status, 303)
        self.assertEqual(response.headers["Location"], BASE + "/")
        self.cookie_value = self.client.session.cookie_jar.filter_cookies(self.server.make_url("/"))[self.service.cookie].value
        return response

    async def get_front(self, path=BASE + "/", *, authenticated=False, method="GET", headers=None):
        headers = {"Sec-Fetch-Mode": "navigate", "Accept": "text/html", **(headers or {})}
        if authenticated:
            headers["Cookie"] = self.service.cookie + "=" + self.cookie_value
        return await self.front_client.request(method, self.front + path, headers=headers,
            allow_redirects=False, ssl=False)

    async def test_box1_anonymous_and_forged_identity_cannot_read_html(self):
        for path in (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"):
            response = await self.get_front(path)
            self.assertEqual(response.status, 303)
            target = urlsplit(response.headers["Location"])
            self.assertEqual(target.path, BASE + "/auth/login")
            self.assertEqual([unquote(v) for v in parse_qs(target.query)["return_to"]], [path])
            self.assertNotIn("/tabel", response.headers["Location"])
            self.assertNotIn(b"smeta-data", await response.read())
        response = await self.get_front(headers={"Cookie": self.service.cookie + "=forged",
            "Authorization": "Bearer forged", "X-User": "admin", "X-Role": "admin"})
        self.assertEqual(response.status, 303)
        response = await self.get_front(AUTH_CHECK)
        self.assertEqual(response.status, 404)

    async def test_box1_valid_native_session_serves_exact_html_and_headers(self):
        await self.login()
        for path in (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"):
            response = await self.get_front(path, authenticated=True)
            self.assertEqual(response.status, 200)
            self.assertEqual(await response.read(), (ROOT / "dist/index.html").read_bytes())
            self.assertIn("sha256-", response.headers["Content-Security-Policy"])
            self.assertIn("connect-src 'none'", response.headers["Content-Security-Policy"])
            self.assertEqual(response.headers["Cache-Control"], "no-store")
            self.assertEqual(response.headers["X-Frame-Options"], "DENY")
            self.assertIn("text/html", response.headers["Content-Type"])
            self.assertEqual(self.last_check_client, CLIENT)
        response = await self.get_front(authenticated=True, method="HEAD")
        self.assertEqual(response.status, 200)
        self.assertEqual(await response.read(), b"")
        self.assertEqual(self.backend_hits, 0)

    async def test_box1_only_registered_files_and_fixed_return_are_exposed(self):
        await self.login()
        for suffix in ("Source", "data/smetas.json", "tests/out/report.xlsx", "Исходные данные/test.xlsx",
                       "app/template.html", "build.py", ".env", "gateway.json", "report.xlsx"):
            response = await self.get_front(BASE + "/" + suffix, authenticated=True)
            self.assertEqual(response.status, 404, suffix)
        response = await self.get_front(authenticated=True, method="POST")
        self.assertEqual(response.status, 403)
        response = await self.get_front(BRIDGE)
        self.assertEqual(response.status, 404)
        response = await self.get_front(BASE + "/?return_to=https://outside.invalid")
        target = urlsplit(response.headers["Location"])
        self.assertEqual(target.path, BASE + "/auth/login")
        self.assertNotIn("/tabel/", response.headers["Location"])

    async def test_box1_revoke_logout_and_identity_outage_close_access(self):
        await self.login()
        self.allowed = False
        response = await self.get_front(authenticated=True)
        self.assertEqual(response.status, 303)
        self.allowed = True
        await self.login()
        self.outage_path = "/cii-security/check"
        self.outage_status = 503
        response = await self.get_front(authenticated=True)
        self.assertEqual(response.status, 503)
        self.assertNotIn(b"smeta-data", await response.read())
        self.outage_path = None
        response = await self.client.get(BASE + "/auth/session")
        session = await response.json()
        response = await self.client.post(BASE + "/auth/logout", headers={"Origin": self.origin,
            "X-CII-CSRF": session["csrf"]}, allow_redirects=False)
        self.assertEqual(response.status, 200)
        response = await self.get_front(authenticated=True)
        self.assertEqual(response.status, 303)

    async def test_box1_chromium_upload_and_excel_with_auth_and_csp(self):
        from playwright.async_api import async_playwright
        from openpyxl import load_workbook
        await self.login()
        errors, requests = [], []
        async with async_playwright() as p:
            installed = sorted((Path.home() / ".cache/ms-playwright").glob("chromium-*/chrome-linux*/chrome"))
            browser = await p.chromium.launch(headless=True, executable_path=str(installed[-1]), args=["--no-sandbox"])
            context = await browser.new_context(ignore_https_errors=True, accept_downloads=True)
            await context.add_cookies([{"name": self.service.cookie, "value": self.cookie_value,
                "url": self.front + "/", "secure": True, "httpOnly": True, "sameSite": "Lax"}])
            page = await context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.on("request", lambda req: requests.append((req.method, req.url)))
            await page.goto(self.front + BASE + "/")
            self.assertEqual(await page.evaluate("SmetaData.get().smetas.length"), 0)
            paths = [ROOT / "Исходные данные/fixture-008.xlsx",
                     ROOT / "Исходные данные/fixture-004.xlsx"]
            await page.locator("#fXlsx").set_input_files([str(p) for p in paths])
            await page.wait_for_function("() => UI.state.computed.length === 2", timeout=60000)
            self.assertEqual(await page.evaluate("SmetaData.get().smetas.length"), 2)
            async with page.expect_download(timeout=60000) as event:
                await page.locator("#bExcel").click()
            download = await event.value
            output = ROOT / "tests/out/box1-auth-report.xlsx"
            await download.save_as(output)
            workbook = load_workbook(output, read_only=True, data_only=True)
            self.assertIn("Сметные расчёты", workbook.sheetnames)
            self.assertIn("Позиции", workbook.sheetnames)
            self.assertEqual(sum(1 for _ in workbook["Позиции"].iter_rows()) - 1, 52)
            workbook.close()
            self.assertEqual(errors, [])
            self.assertTrue(all(method == "GET" and url.startswith(self.front) for method, url in requests))
            await browser.close()


if __name__ == "__main__":
    names = [name for name in Box1AuthTests.__dict__ if name.startswith("test_box1_")]
    result = unittest.TextTestRunner(verbosity=2).run(unittest.TestSuite(Box1AuthTests(name) for name in names))
    raise SystemExit(0 if result.wasSuccessful() else 1)
