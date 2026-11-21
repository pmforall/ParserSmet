#!/usr/bin/env python3
"""Шлюз из действующего Табеля + веб-сборка парсера, только синтетический IdP.

python3 -I -S -B tests/test_sso_gateway.py
Не обращается к корпоративному IdP и не использует его ключи или учётные записи.
"""
import dataclasses
import json
from pathlib import Path
import secrets
import sys
import tempfile
import time
import unittest
from urllib.parse import parse_qs, urlencode, urlsplit

ROOT = Path(__file__).resolve().parents[1]
DIGEST = "273755939965b47b883b9c70f4ad328360ffe82239ccfdaa69e00036de2804b0"
CAPSULE = Path("/opt/cii-tabel-gateway/releases") / DIGEST
sys.path[:0] = [str(ROOT / "deploy"), str(CAPSULE / "source"), str(CAPSULE / "site-packages")]
from prepare_sso import BASE, CLIENT, gateway_config, verify_gateway, client_registration
from gateway.app import GATEWAY
from gateway.config import Route, load
from gateway.tests.test_gateway import GatewayIntegration


class SmetaGatewayTests(GatewayIntegration):
    async def asyncSetUp(self):
        await super().asyncSetUp()
        self.level = 1
        files = {BASE + "/": str(ROOT / "dist/index.html"),
                 BASE + "/index.html": str(ROOT / "dist/index.html"),
                 BASE + "/Парсер_смет.html": str(ROOT / "dist/Парсер_смет.html")}
        self.service = dataclasses.replace(self.service, name="smeta-parser", client_id=CLIENT,
            minimum_level=1, base_path=BASE, home=BASE + "/", launch_path=BASE + "/auth/launch",
            launch_return_prefixes=(BASE,), routes=(Route(BASE, static_files=tuple(files.items())),))
        gateway = self.app[GATEWAY]
        gateway.config = dataclasses.replace(gateway.config, services=(self.service,))
        gateway.oidc.config = gateway.config

    async def begin(self, path=None):
        response = await self.client.get(path or BASE + "/auth/login?" + urlencode({"return_to": BASE + "/"}),
                                         allow_redirects=False)
        self.assertEqual(response.status, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        self.assertEqual(query["client_id"], [CLIENT])
        self.assertEqual(query["code_challenge_method"], ["S256"])
        self.assertEqual(query["redirect_uri"], [self.service.callback])
        return query

    async def finish(self, query):
        code = secrets.token_urlsafe(16)
        self.codes[code] = query
        return await self.client.get(BASE + "/auth/callback?" + urlencode({"state": query["state"][0],
            "code": code, "iss": self.issuer}), allow_redirects=False)

    async def login(self):
        response = await self.finish(await self.begin())
        self.assertEqual(response.status, 303)
        self.assertEqual(response.headers["Location"], BASE + "/")
        return response

    async def test_smeta_files_require_standard_login(self):
        for path in (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"):
            response = await self.client.get(path, allow_redirects=False)
            self.assertEqual(response.status, 401)
            self.assertNotIn(b"smeta-data", await response.read())
        response = await self.client.get(BASE + "/", headers={"Accept": "text/html", "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate"}, allow_redirects=False)
        self.assertEqual(response.status, 303)
        self.assertTrue(response.headers["Location"].startswith(BASE + "/auth/login"))
        self.assertEqual(self.backend_hits, 0)

    async def test_smeta_signed_code_pkce_delivers_only_exact_static_files(self):
        login = await self.login()
        cookie = login.headers["Set-Cookie"]
        for attr in ("__Host-cii-rp-smeta-parser", "Secure", "HttpOnly", "SameSite=Lax"):
            self.assertIn(attr, cookie)
        self.assertNotIn("synthetic-native-refresh", cookie)
        for path in (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"):
            response = await self.client.get(path)
            self.assertEqual(response.status, 200)
            self.assertEqual(await response.read(), (ROOT / "dist/index.html").read_bytes())
            self.assertEqual(response.headers["Cache-Control"], "no-store")
            self.assertEqual(self.last_check_client, CLIENT)
        for suffix in ("Source", "data/smetas.json", "tests/out/report.xlsx", "Исходные данные/test.xlsx",
                       "app/template.html", "build.py", ".env", "gateway.json", "report.xlsx"):
            response = await self.client.get(BASE + "/" + suffix, allow_redirects=False)
            self.assertEqual(response.status, 404, suffix)
        self.assertEqual(self.backend_hits, 0)

    async def test_smeta_forged_headers_and_old_tabel_cookie_are_denied(self):
        response = await self.client.get(BASE + "/", headers={"Authorization": "Bearer forged",
            "X-Role": "admin", "X-User": "admin", "Cookie": "__Host-cii-rp-tabel=forged"}, allow_redirects=False)
        self.assertEqual(response.status, 401)
        self.assertNotIn(b"smeta-data", await response.read())

    async def test_smeta_revocation_and_expiry_deny_new_reads(self):
        await self.login()
        self.allowed = False
        response = await self.client.get(BASE + "/", allow_redirects=False)
        self.assertEqual(response.status, 401)
        self.assertNotIn(b"smeta-data", await response.read())
        self.allowed = True
        await self.login()
        gw = self.app[GATEWAY]
        cookie = self.client.session.cookie_jar.filter_cookies(self.server.make_url("/"))[self.service.cookie].value
        session = gw.store.get("session", cookie, self.service.name)
        session["expires"] = int(time.time()) - 1
        gw.store.put("session", cookie, self.service.name, session, session["expires"])
        response = await self.client.get(BASE + "/", allow_redirects=False)
        self.assertEqual(response.status, 401)

    async def test_smeta_logout_requires_csrf_then_denies_reads(self):
        await self.login()
        response = await self.client.get(BASE + "/auth/session")
        session = await response.json()
        self.assertTrue(session["authenticated"])
        response = await self.client.post(BASE + "/auth/logout", headers={"Origin": self.origin}, allow_redirects=False)
        self.assertEqual(response.status, 403)
        response = await self.client.post(BASE + "/auth/logout", headers={"Origin": self.origin,
            "X-CII-CSRF": session["csrf"]}, allow_redirects=False)
        self.assertEqual(response.status, 200)
        self.assertEqual(self.logout_calls, 1)
        response = await self.client.get(BASE + "/", allow_redirects=False)
        self.assertEqual(response.status, 401)

    async def test_smeta_invalid_claims_and_replayed_state_cannot_open_page(self):
        for flag in ("bad_nonce", "bad_audience", "bad_issuer", "bad_signature"):
            setattr(self, flag, True)
            response = await self.finish(await self.begin())
            self.assertEqual(response.status, 401, flag)
            setattr(self, flag, False)
        query = await self.begin()
        response = await self.finish(query)
        self.assertEqual(response.status, 303)
        response = await self.finish(query)
        self.assertEqual(response.status, 403)

    async def test_smeta_outage_does_not_admit_static_content(self):
        await self.login()
        self.outage_path = "/cii-security/check"
        self.outage_status = 503
        response = await self.client.get(BASE + "/", allow_redirects=False)
        self.assertEqual(response.status, 503)
        self.assertNotIn(b"smeta-data", await response.read())
        self.outage_path = None
        response = await self.client.get(BASE + "/", allow_redirects=False)
        self.assertEqual(response.status, 200)


class PreparedConfigurationTests(unittest.TestCase):
    def test_smeta_exact_candidate_loads_with_own_credentials(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for name, content in (("client-secret", b"synthetic-client"), ("store-key", b"s" * 32)):
                path = root / name
                path.write_bytes(content)
                path.chmod(0o600)
            config = gateway_config(ROOT / "dist", root, root)
            path = root / "gateway.json"
            path.write_text(json.dumps(config))
            parsed = load(path)
            self.assertEqual(parsed.services[0].client_id, CLIENT)
            self.assertEqual(parsed.listen_host, "127.0.0.1")
            self.assertEqual(len(parsed.services[0].routes[0].static_files), 3)

    def test_smeta_policy_matches_tabel_floor_and_uses_own_client(self):
        verify_gateway()
        client = client_registration()
        self.assertEqual(client["clientId"], CLIENT)
        self.assertEqual(client["name"], "Парсер смет")
        self.assertEqual(client["attributes"]["cii.display-name"], "Парсер смет")
        self.assertEqual(client["attributes"]["cii.min-level"], "1")
        self.assertEqual(client["attributes"]["pkce.code.challenge.method"], "S256")
        self.assertFalse(client["directAccessGrantsEnabled"])
        self.assertFalse(client["implicitFlowEnabled"])
        self.assertFalse(client["publicClient"])
        self.assertNotIn("secret", client)
        prepared = json.loads((ROOT / "deploy/prepared/gateway.json").read_text())
        service = prepared["services"][0]
        self.assertNotIn("tabel", json.dumps(prepared))
        self.assertEqual(service["minimum_level"], 1)
        self.assertEqual(prepared["max_session_seconds"], 43200)
        self.assertEqual(set(service["routes"][0]["static_files"]),
            {BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"})


if __name__ == "__main__":
    # The reference class supplies the signed in-process OIDC fixture. Run only
    # this product's methods; its inherited product-specific tests are unrelated.
    names = [name for name in SmetaGatewayTests.__dict__ if name.startswith("test_smeta_")]
    suite = unittest.TestSuite([SmetaGatewayTests(name) for name in names])
    suite.addTests(unittest.defaultTestLoader.loadTestsFromTestCase(PreparedConfigurationTests))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    raise SystemExit(0 if result.wasSuccessful() else 1)
