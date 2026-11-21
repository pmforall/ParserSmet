#!/usr/bin/env python3
"""Finite one-client registration, synthetic registry only."""
import copy
import json
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "deploy"))
from register_idp_client import BASE, CLIENT, digest, enroll, validate_policy

POLICY = json.loads((ROOT / "deploy/prepared/client-registration.json").read_text())
OWN_ID = "11111111-1111-4111-8111-111111111111"


class Registry:
    def __init__(self, existing=False, drift=False, fail=False):
        self.rows = {"legacy": {"id": "22222222-2222-4222-8222-222222222222", "clientId": "legacy", "enabled": True}}
        self.writes = []
        self.expected = None
        self.drift, self.fail = drift, fail
        if existing:
            self.rows[CLIENT] = {**copy.deepcopy(POLICY), "id": OWN_ID}

    def arm(self, expected, policy):
        self.expected = expected

    def request(self, method, path, body=None):
        if method == "POST":
            assert path == BASE + "/clients" and body == POLICY
            assert digest(self.rows) == self.expected
            self.writes.append((method, path))
            if self.fail:
                raise RuntimeError("transport_failed")
            self.rows[CLIENT] = {**copy.deepcopy(body), "id": OWN_ID}
            if self.drift:
                self.rows["legacy"]["enabled"] = False
            return None
        if path == BASE + "/clients":
            return copy.deepcopy(list(self.rows.values()))
        if path == BASE + "/clients/" + OWN_ID:
            return copy.deepcopy(self.rows[CLIENT])
        if path.endswith("/default-client-scopes"):
            return [{"name": x} for x in ("basic", "profile", "email")]
        raise AssertionError("Unexpected API request")


class RegistrationTests(unittest.TestCase):
    def run_enroll(self, registry, policy=None):
        self.events = {}
        self.checkpoints = 0
        def checkpoint():
            self.checkpoints += 1
        def publish(name, value):
            self.assertNotIn(name, self.events)
            self.events[name] = copy.deepcopy(value)
        return enroll(registry, policy or POLICY,
            lambda a, p: all(a.get(k) == v for k, v in p.items()), checkpoint, publish)

    def test_creates_only_own_client_with_before_and_readback(self):
        api = Registry()
        legacy = copy.deepcopy(api.rows["legacy"])
        proof = self.run_enroll(api)
        self.assertEqual(api.rows["legacy"], legacy)
        self.assertEqual(api.writes, [("POST", BASE + "/clients")])
        self.assertTrue(proof["created"] and proof["other_clients_unchanged"])
        self.assertEqual(proof["client_count_after"], 2)
        self.assertEqual(set(self.events), {"before.json", "intent.json", "outcome.json"})
        self.assertGreaterEqual(self.checkpoints, 2)

    def test_matching_client_is_not_overwritten_or_rotated(self):
        api = Registry(existing=True)
        before = copy.deepcopy(api.rows)
        proof = self.run_enroll(api)
        self.assertFalse(proof["created"])
        self.assertEqual(api.rows, before)
        self.assertEqual(api.writes, [])

    def test_other_client_target_is_rejected(self):
        policy = copy.deepcopy(POLICY)
        policy["clientId"] = "cii-tabel"
        with self.assertRaisesRegex(ValueError, "wrong_client"):
            validate_policy(policy)

    def test_broader_auth_or_external_callback_is_rejected(self):
        for key, value in (("directAccessGrantsEnabled", True), ("publicClient", True),
                           ("redirectUris", ["https://outside.invalid/callback"]),
                           ("optionalClientScopes", ["offline_access"])):
            policy = copy.deepcopy(POLICY)
            policy[key] = value
            with self.assertRaises(ValueError):
                validate_policy(policy)

    def test_foreign_drift_is_detected_and_never_reverted(self):
        api = Registry(drift=True)
        with self.assertRaisesRegex(ValueError, "other_clients_changed"):
            self.run_enroll(api)
        self.assertFalse(api.rows["legacy"]["enabled"])
        self.assertNotIn("outcome.json", self.events)

    def test_transport_failure_never_claims_success(self):
        api = Registry(fail=True)
        with self.assertRaisesRegex(RuntimeError, "transport_failed"):
            self.run_enroll(api)
        self.assertIn("intent.json", self.events)
        self.assertNotIn("outcome.json", self.events)


if __name__ == "__main__":
    unittest.main()
