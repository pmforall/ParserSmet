#!/usr/bin/env python3
"""One-client provisioning through the installed private Keycloak Admin transport.

Run only on eksprod through ~/eksprod sudo. Input is public policy plus the
current manual lease ID. stdout is metadata, except explicit export-secret mode
whose output must go directly to the own 0600 credential file on box1.
"""
from __future__ import annotations

import base64
import fcntl
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import time
import types
import uuid

CLIENT = "cii-smeta-parser"
ORIGIN = "https://box1.cii-eks.ru"
ISSUER = "https://id.cii-eks.ru/realms/cii"
BASE = "/admin/realms/cii"
TRANSPORT_SHA = "443e3cb60641334623fae63b92cbaa274bea4d8fca6927ce9d7ce491b51a6d37"
RECOVERY_SHA = "7d4d4ad916f8a9c29c8da1fc9258b311438da6e0c5bfa17c91cc64fa7c6558e2"


def require(value, code):
    if not value:
        raise ValueError(code)


def canonical(value):
    return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()


def digest(value):
    return hashlib.sha256(canonical(value)).hexdigest()


def validate_policy(policy):
    require(policy.get("clientId") == CLIENT and policy.get("name") == "Парсер смет", "wrong_client")
    require(not ({"id", "secret", "authorizationServicesEnabled"} & policy.keys()), "credential_or_id_in_policy")
    for key, value in {"enabled": True, "protocol": "openid-connect", "publicClient": False,
                       "standardFlowEnabled": True, "directAccessGrantsEnabled": False,
                       "implicitFlowEnabled": False, "serviceAccountsEnabled": False,
                       "fullScopeAllowed": False, "clientAuthenticatorType": "client-secret"}.items():
        require(policy.get(key) == value and type(policy.get(key)) is type(value), "wrong_client_boundary")
    require(policy.get("redirectUris") == [ORIGIN + "/smeta-parser/auth/callback"]
            and policy.get("webOrigins") == [ORIGIN]
            and policy.get("rootUrl") == ORIGIN and policy.get("baseUrl") == ORIGIN + "/smeta-parser/"
            and len(policy.get("defaultClientScopes", [])) == 3
            and set(policy.get("defaultClientScopes", [])) == {"basic", "email", "profile"}
            and policy.get("optionalClientScopes") == [], "wrong_routes_or_scopes")
    attrs = policy.get("attributes", {})
    expected = {"cii.min-level": "1", "cii.security-api": "true", "pkce.code.challenge.method": "S256",
        "cii.display-name": "Парсер смет", "cii.launch-uri": ORIGIN + "/smeta-parser/auth/launch",
        "cii.launch-path-prefixes": "/smeta-parser", "access.token.lifespan": "60",
        "use.refresh.tokens": "true", "post.logout.redirect.uris": ORIGIN + "/smeta-parser/",
        "oauth2.device.authorization.grant.enabled": "false", "oidc.ciba.grant.enabled": "false"}
    require(attrs == expected, "wrong_native_policy")
    mappers = policy.get("protocolMappers", [])
    require(len(mappers) == 2 and {m.get("protocolMapper") for m in mappers} ==
            {"cii-assurance-claims", "oidc-audience-mapper"}, "wrong_native_mappers")
    audience = next(m for m in mappers if m["protocolMapper"] == "oidc-audience-mapper")
    require(audience.get("config", {}).get("included.custom.audience") == "cii-security", "wrong_audience")


def collection(api):
    values = api.request("GET", BASE + "/clients")
    require(type(values) is list and len(values) <= 4096, "invalid_client_collection")
    require(all(type(c) is dict and type(c.get("clientId")) is str and type(c.get("id")) is str
                for c in values), "invalid_client_row")
    result = {c["clientId"]: c for c in values}
    require(len(result) == len(values) and len({c["id"] for c in values}) == len(values), "duplicate_clients")
    return result


def enroll(api, policy, matches, checkpoint, publish):
    """Create only this client; keep full immutable before/intent/outcome evidence."""
    validate_policy(policy)
    before = collection(api)
    publish("before.json", {"schema": "eks.smeta.idp-before/v1", "clients": before})
    created = CLIENT not in before
    if created:
        publish("intent.json", {"schema": "eks.smeta.idp-create/v1", "client_id": CLIENT,
                               "before_sha256": digest(before), "policy": policy})
        api.arm(digest(before), policy)
        checkpoint()
        require(collection(api) == before, "registry_changed_before_create")
        api.request("POST", BASE + "/clients", policy)
    after = collection(api)
    require(set(after) == set(before) | {CLIENT}, "client_collection_changed")
    require(all(after[k] == value for k, value in before.items() if k != CLIENT), "other_clients_changed")
    row = after[CLIENT]
    require(matches(row, policy), "client_readback_mismatch")
    require(re.fullmatch(r"[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}", row["id"]), "invalid_client_uuid")
    detail = api.request("GET", BASE + "/clients/" + row["id"])
    require(detail == row, "client_detail_changed")
    scopes = api.request("GET", BASE + "/clients/" + row["id"] + "/default-client-scopes")
    require({s.get("name") for s in scopes} == {"basic", "email", "profile"}, "scope_readback_mismatch")
    checkpoint()
    proof = {"schema": "eks.smeta.idp-registration/v1", "client_id": CLIENT, "client_uuid": row["id"],
        "created": created, "registered": True, "issuer": ISSUER, "display_name": "Парсер смет",
        "client_count_before": len(before), "client_count_after": len(after), "other_clients_unchanged": True,
        "policy_sha256": digest(policy), "client_row_sha256": digest(row),
        "rollback": {"created_in_this_operation": created, "own_client_uuid": row["id"],
                     "require_exact_row_sha256": digest(row), "other_clients_must_be_preserved": True}}
    publish("outcome.json", {**proof, "client": row})
    return proof


def load_module(path, expected, name):
    raw = path.read_bytes()
    require(hashlib.sha256(raw).hexdigest() == expected, "installed_admin_source_changed")
    module = types.ModuleType(name)
    module.__file__ = str(path)
    exec(compile(raw, str(path), "exec"), module.__dict__)
    return module


def check_file(path):
    i = path.lstat()
    require(stat.S_ISREG(i.st_mode) and stat.S_IMODE(i.st_mode) == 0o600 and i.st_nlink == 1
            and i.st_uid in (0, 1000), "unsafe_private_file")


def checkpoint(activation):
    require(os.geteuid() == 0 and re.fullmatch(r"[a-f0-9]{32}", activation), "root_lease_required")
    path = Path("/run/cii-eksprod-owner-access-20261008/grant.json")
    check_file(path)
    g = json.loads(path.read_text())
    require(g["activation_id"] == activation and g["origin"] == "manual" and g["uid"] == 1000
            and g["boot_id"] == Path("/proc/sys/kernel/random/boot_id").read_text().strip()
            and time.clock_gettime_ns(time.CLOCK_BOOTTIME) < g["deadline_boottime_ns"], "manual_lease_closed_or_changed")


def main(argv):
    require(len(argv) == 3 and argv[0] in ("preflight", "apply", "export-secret"), "arguments")
    mode, encoded, activation = argv
    policy = json.loads(base64.b64decode(encoded, validate=True))
    validate_policy(policy)
    checkpoint(activation)
    source = Path("/opt/cii-id/current/deploy/production")
    admin = load_module(source / "admin_transport.py", TRANSPORT_SHA, "smeta_admin_transport")
    recovery = load_module(source / "native_admin_recovery.py", RECOVERY_SHA, "smeta_admin_compare")
    ids = subprocess.check_output(["docker", "ps", "--filter", "label=com.docker.compose.project=cii-id-prod",
            "--filter", "label=com.docker.compose.service=keycloak", "--format", "{{.ID}}"], text=True).split()
    require(len(ids) == 1, "keycloak_container_not_unique")
    runtime = json.loads(subprocess.check_output(["docker", "inspect", ids[0]], text=True))[0]
    addresses = [v["IPAddress"] for k, v in runtime["NetworkSettings"]["Networks"].items() if k.endswith("_frontend")]
    require(len(addresses) == 1 and runtime["Name"] == "/cii-id-prod-keycloak-1", "keycloak_network_changed")
    api = admin.PrivateAdmin(addresses[0], Path("/var/lib/cii-id/secrets/bootstrap_password"), check_file)
    realm = api.verify_realm()
    require(ORIGIN in realm.get("attributes", {}).get("cii.launch-origins", "").split(","), "origin_not_registered")
    original_private = api._request
    armed = {}

    def arm(before_sha, candidate):
        require(candidate == policy and not armed, "write_already_armed")
        armed.update(before_sha=before_sha)

    def confined(method, path, body=None, **kwargs):
        if method not in ("GET", "HEAD") and path != "/realms/master/protocol/openid-connect/token":
            require(mode == "apply" and method == "POST" and path == BASE + "/clients"
                    and body == policy and "before_sha" in armed, "write_outside_own_client")
            checkpoint(activation)
            require(digest(collection(api)) == armed["before_sha"], "registry_changed_after_authentication")
            armed.clear()
        return original_private(method, path, body, **kwargs)

    api.arm = arm
    api._request = confined
    lock_fd = os.open("/var/lock/cii-id-prod.lock", os.O_RDWR | os.O_CREAT | os.O_NOFOLLOW, 0o600)
    try:
        fcntl.flock(lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        rows = collection(api)
        if mode == "preflight":
            print(json.dumps({"schema": "eks.smeta.idp-preflight/v1", "client_id": CLIENT,
                "target_exists": CLIENT in rows, "client_count": len(rows), "policy_sha256": digest(policy),
                "origin_registered": True, "reference_floor": rows["cii-tabel"]["attributes"]["cii.min-level"]}))
            return
        if mode == "export-secret":
            require(CLIENT in rows and recovery.matches(rows[CLIENT], policy), "target_not_registered")
            result = api.request("GET", BASE + "/clients/" + rows[CLIENT]["id"] + "/client-secret")
            value = result.get("value")
            require(type(value) is str and re.fullmatch(r"[A-Za-z0-9_-]{24,256}", value), "invalid_own_client_secret")
            checkpoint(activation)
            sys.stdout.write(value)
            return
        folder = Path("/var/lib/cii-id/evidence") / ("smeta-registration-" + uuid.uuid4().hex)
        folder.mkdir(mode=0o700)

        def publish(name, value):
            path = folder / name
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY | os.O_NOFOLLOW, 0o600)
            with os.fdopen(fd, "wb") as stream:
                stream.write(canonical(value)); stream.flush(); os.fsync(stream.fileno())

        try:
            proof = enroll(api, policy, recovery.matches, lambda: checkpoint(activation), publish)
        except Exception:
            publish("unconfirmed.json", {"schema": "eks.smeta.idp-registration-unknown/v1", "client_id": CLIENT})
            raise
        proof["evidence_directory"] = str(folder)
        proof["manual_lease_activation"] = activation
        print(json.dumps(proof, ensure_ascii=False))
    finally:
        os.close(lock_fd)


if __name__ == "__main__":
    try:
        main(sys.argv[1:])
    except Exception as error:
        # The transport uses fixed error codes; response bodies and credentials
        # never enter stdout, exception text or command arguments.
        code = str(error) if isinstance(error, (ValueError, RuntimeError)) and re.fullmatch(r"[a-z_0-9]+", str(error)) else "registration_failed"
        print(json.dumps({"ok": False, "code": code}), file=sys.stderr)
        raise SystemExit(1)
