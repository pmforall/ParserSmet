#!/usr/bin/env python3
"""Проверки, что книги и отчёты не попадают в веб-сборку."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from package_web import validate


class WebBuildTests(unittest.TestCase):
    def test_data_argument_keeps_only_schema(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            seed = json.loads((ROOT / "app/schema.json").read_text())
            sentinel = "PRIVATE_ESTIMATE_SENTINEL_b2034b"
            seed["smetas"] = [{"name": sentinel}]
            seed["checks"] = [{"message": sentinel}]
            seed["meta"]["files"] = [sentinel]
            seed["meta"]["warnings"] = [sentinel]
            seed["meta"]["reference_source"] = sentinel
            seed["reference_report"] = {"private": sentinel}
            data = folder / "private.json"
            data.write_text(json.dumps(seed))
            output = folder / "public.html"
            subprocess.run([sys.executable, str(ROOT / "build.py"), "--data", str(data),
                            "--out", str(output)], check=True, capture_output=True)
            self.assertNotIn(sentinel, output.read_text())

    def test_cannot_embed_estimates_in_dist(self):
        output = ROOT / "dist/blocked-private-data.html"
        self.assertFalse(output.exists())
        result = subprocess.run([sys.executable, str(ROOT / "build.py"), "--with-data",
                                 "--data", str(ROOT / "data/smetas.json"), "--out", str(output)],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 2)
        self.assertFalse(output.exists())

    def test_cannot_write_development_archive_in_dist(self):
        output = ROOT / "dist/blocked-private-archive.zip"
        self.assertFalse(output.exists())
        candidates = (("pack_source.py", "вне каталога проекта"),
                      ("pack_archive.py", "каталог сайта"))
        for name, message in candidates:
            script = ROOT / "tools" / name
            if not script.exists():
                continue
            with self.subTest(script=name):
                result = subprocess.run([sys.executable, str(script), "--out", str(output)],
                                        capture_output=True, text=True)
                self.assertEqual(result.returncode, 2)
                self.assertIn(message, result.stderr)
                self.assertFalse(output.exists())

    def test_packager_rejects_extra_files_and_symlinks(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            for name in ("index.html", "Парсер_смет.html"):
                (folder / name).write_bytes((ROOT / "dist" / name).read_bytes())
            validate(folder)
            secret = folder / "estimate.xlsx"
            secret.write_bytes(b"private-test-data")
            with self.assertRaises(ValueError):
                validate(folder)
            secret.unlink()
            (folder / "index.html").unlink()
            (folder / "index.html").symlink_to(ROOT / "dist/index.html")
            with self.assertRaises(ValueError):
                validate(folder)


if __name__ == "__main__":
    unittest.main()
