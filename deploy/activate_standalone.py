#!/usr/bin/env python3
"""Activate only the parser's registered standalone gateway, then verify ingress."""
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request

from prepare_sso import ROOT, CLIENT, ISSUER, ORIGIN, BASE, verify_gateway, GATEWAY_DIGEST
from register_idp_client import digest


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def request(url, *, headers=None, body=None):
    opener = urllib.request.build_opener(NoRedirect)
    req = urllib.request.Request(url, headers=headers or {}, data=body)
    try:
        response = opener.open(req, timeout=20)
    except urllib.error.HTTPError as error:
        response = error
    with response:
        return response.code, dict(response.headers), response.read(1000000)


def main():
    if os.geteuid() != 0:
        raise SystemExit("local root required")
    prepared = ROOT / "deploy/prepared"
    for filename, summary in (("tests/sso-gateway.log", "Ran 9 tests"), ("tests/box1-auth.log", "Ran 5 tests")):
        log = (ROOT / filename).read_text()
        if summary not in log or not log.rstrip().endswith("OK"):
            raise SystemExit("required standalone test has not passed")
    private = Path("/etc/smeta-parser")
    policy = json.loads((private / "client-registration.json").read_text())
    proof = json.loads((private / "client-registration.ok").read_text())
    if (proof.get("client_id") != CLIENT or proof.get("registered") is not True
            or proof.get("issuer") != ISSUER or proof.get("policy_sha256") != digest(policy)):
        raise SystemExit("registration proof mismatch")
    credential = private / "client-secret"
    info = credential.lstat()
    if (not stat.S_ISREG(info.st_mode) or stat.S_IMODE(info.st_mode) != 0o600 or info.st_uid != 0
            or info.st_nlink != 1 or not re.fullmatch(rb"[A-Za-z0-9_-]{24,256}", credential.read_bytes())):
        raise SystemExit("unsafe own credential")
    if (private / "gateway.json").read_bytes() != (prepared / "gateway.json").read_bytes():
        raise SystemExit("gateway configuration mismatch")
    runtime = Path("/opt/smeta-parser/gateway") / GATEWAY_DIGEST
    verify_gateway(runtime)
    snippet = Path("/etc/nginx/snippets/smeta-parser.conf")
    if snippet.read_bytes() != (prepared / "nginx-location.conf").read_bytes() or "/tabel/" in snippet.read_text():
        raise SystemExit("standalone nginx configuration mismatch")
    subprocess.run(["nginx", "-t"], check=True)
    old_pid = subprocess.check_output(["systemctl", "show", "cii-tabel-gateway.service", "--property=MainPID", "--value"], text=True).strip()
    prior = subprocess.run(["systemctl", "is-active", "--quiet", "smeta-parser-gateway.service"]).returncode == 0
    try:
        subprocess.run(["systemctl", "enable", "--now", "smeta-parser-gateway.service"], check=True)
        deadline = time.monotonic() + 30
        while True:
            try:
                code, _, data = request("http://127.0.0.1:18468" + BASE + "/auth/session",
                    headers={"Host": "box1.cii-eks.ru", "Accept": "application/json"})
                if code == 401 and json.loads(data).get("login_url", "").startswith(BASE + "/auth/login"):
                    break
            except (OSError, ValueError):
                pass
            if time.monotonic() > deadline:
                raise RuntimeError("own gateway did not become ready")
            time.sleep(0.25)
        checks = []
        for path in (BASE + "/", BASE + "/index.html", BASE + "/Парсер_смет.html"):
            code, headers, body = request(ORIGIN + urllib.parse.quote(path), headers={"Sec-Fetch-Mode": "navigate", "Accept": "text/html"})
            location = urllib.parse.urlsplit(headers.get("Location", ""))
            if code != 303 or location.path != BASE + "/auth/login" or b"smeta-data" in body:
                raise RuntimeError("public entry not protected by own login")
            checks.append({"path": path, "status": code, "login_path": location.path})
        code, headers, _ = request(ORIGIN + BASE + "/auth/login?return_to=%2Fsmeta-parser%2F")
        target = urllib.parse.urlsplit(headers.get("Location", ""))
        values = urllib.parse.parse_qs(target.query)
        if (code != 302 or target.hostname != "id.cii-eks.ru" or values.get("client_id") != [CLIENT]
                or values.get("redirect_uri") != [ORIGIN + BASE + "/auth/callback"]
                or values.get("code_challenge_method") != ["S256"]):
            raise RuntimeError("own authorization request mismatch")
        # Check actual client authentication without a user account or sending
        # email: an unknown one-use code must fail as invalid_grant, not invalid_client.
        payload = urllib.parse.urlencode({"grant_type": "authorization_code", "client_id": CLIENT,
            "client_secret": credential.read_text(), "redirect_uri": ORIGIN + BASE + "/auth/callback",
            "code": "smeta-registration-credential-check", "code_verifier": "v" * 64}).encode()
        token_code, _, token_body = request(ISSUER + "/protocol/openid-connect/token",
            headers={"Content-Type": "application/x-www-form-urlencoded"}, body=payload)
        if token_code != 400 or json.loads(token_body).get("error") != "invalid_grant":
            raise RuntimeError("own client credential was not accepted")
        after_pid = subprocess.check_output(["systemctl", "show", "cii-tabel-gateway.service", "--property=MainPID", "--value"], text=True).strip()
        if old_pid != after_pid:
            raise RuntimeError("reference gateway process changed")
        active_pid = subprocess.check_output(["systemctl", "show", "smeta-parser-gateway.service", "--property=MainPID", "--value"], text=True).strip()
        receipt = {"schema": "eks.smeta.standalone-live/v1", "url": ORIGIN + BASE + "/", "activated": True,
            "client_id": CLIENT, "display_name": policy["attributes"]["cii.display-name"],
            "client_uuid": proof["client_uuid"], "issuer": ISSUER, "minimum_level": 1,
            "gateway_pid": int(active_pid), "existing_tabel_gateway_pid_unchanged": int(after_pid),
            "policy_sha256": digest(policy), "nginx_sha256": hashlib.sha256(snippet.read_bytes()).hexdigest(),
            "public_checks": checks, "s256_verified": True, "own_client_authentication_verified": True,
            "corporate_account_login_tested": False}
        (prepared / "standalone-live-receipt.json").write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n")
        print("Active: " + receipt["url"] + "; own client and S256 verified")
    except BaseException:
        if not prior:
            subprocess.run(["systemctl", "disable", "--now", "smeta-parser-gateway.service"], check=False)
        raise


if __name__ == "__main__":
    main()
