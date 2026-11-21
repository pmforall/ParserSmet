#!/usr/bin/env python3
"""Проверка архива передачи: исключения данных, сканирование и контрольные суммы."""
import hashlib
import json
from pathlib import Path
import secrets
import sys
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import pack_source


class SourceArchiveTests(unittest.TestCase):
    def test_forbidden_paths(self):
        for name in (".env", "deploy/prepared/gateway.json", "tests/out/report.xlsx",
                     "keys/server.key", "data/smetas.json", "../private.py", "/private.py",
                     "tests/run.log", ".git/config", "Исходные данные/book.xlsx"):
            with self.subTest(name=name), self.assertRaises(ValueError):
                pack_source.scan(name, b"private data")

    def test_secret_material_rejected_without_printing_value(self):
        examples = (b"-----BEGIN " + b"PRIVATE KEY-----", b'client_secret = "' + b'dummy-sensitive-value"',
                    b"https://operator:" + b"password" + b"@" + b"invalid.example/",
                    b"person" + b"@invalid.example")
        for data in examples:
            with self.subTest(size=len(data)), self.assertRaises(ValueError) as error:
                pack_source.scan("source.py", data)
            self.assertNotIn(data.decode(), str(error.exception))

    def test_data_excluded_and_manifest_matches_archive(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            project = folder / "project"
            project.mkdir()
            for name, source in pack_source.sources(ROOT).items():
                target = project / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(source.read_bytes())
            sentinel = b"PRIVATE_DATA_SENTINEL_" + secrets.token_hex(16).encode()
            for name in (".env", "data/smetas.json", "deploy/prepared/client-secret", "tests/out/report.xlsx",
                         "Исходные данные/private-fixture.xlsx"):
                target = project / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(sentinel)
            template = project / "app/template.html"
            template.write_text(template.read_text() + "\n<!-- private-fixture.xlsx -->\n")
            output = folder / "source.zip"
            result = pack_source.build_archive(project, output)
            with zipfile.ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
                manifest = json.loads(archive.read("smeta-parser/MANIFEST.json"))
                self.assertEqual(len(archive.namelist()), len(manifest["files"]) + 1)
                for name, info in manifest["files"].items():
                    data = archive.read("smeta-parser/" + name)
                    self.assertNotIn(sentinel, data)
                    self.assertNotIn(b"private-fixture.xlsx", data)
                    self.assertTrue(pack_source.allowed_name(name), name)
                    self.assertEqual(info["sha256"], hashlib.sha256(data).hexdigest(), name)
                self.assertIn("CHANGELOG.md", manifest["files"])
                self.assertIn("dist/index.html", manifest["files"])
                self.assertIn("fixture-001.xlsx", archive.read("smeta-parser/app/template.html").decode())
            self.assertEqual(result["sha256"], hashlib.sha256(output.read_bytes()).hexdigest())
            with self.assertRaises(ValueError):
                pack_source.build_archive(project, project / "dist/source.zip")


if __name__ == "__main__":
    unittest.main(verbosity=2)
