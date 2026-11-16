#!/usr/bin/env python3
"""Tests de check_prerelease_profile.py (#1627).

Se prueba por la interfaz pública: el proceso, su código de salida y su stderr.
"""

import ast
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "check_prerelease_profile.py"


def run_check(content):
    """Escribe `content` (str, bytes u objeto JSON) en un temporal y valida ese fichero."""
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "prerelease.json"
        if isinstance(content, bytes):
            path.write_bytes(content)
        elif isinstance(content, str):
            path.write_text(content, encoding="utf-8")
        else:
            path.write_text(json.dumps(content), encoding="utf-8")
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), str(path)],
            capture_output=True,
            encoding="utf-8",
            errors="replace",
        )
    return proc.returncode, proc.stdout, proc.stderr


class TestValidDeclaration(unittest.TestCase):
    def test_apap_shape_enabled_only_exits_0(self):
        code, _, err = run_check({"enabled": True})
        self.assertEqual(code, 0, err)

    def test_explicit_schema_version_1_exits_0(self):
        code, _, err = run_check({"schema_version": 1, "enabled": True})
        self.assertEqual(code, 0, err)

    def test_apap_web_real_file_fixture_exits_0(self):
        fixture = Path(__file__).resolve().parent / "fixtures" / "apap_web_prerelease.json"
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), str(fixture)],
            capture_output=True,
            encoding="utf-8",
            errors="replace",
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)

    def test_enabled_false_exits_0(self):
        code, _, err = run_check({"enabled": False})
        self.assertEqual(code, 0, err)

    def test_schema_version_above_one_is_an_integer_and_exits_0(self):
        code, _, err = run_check({"schema_version": 2, "enabled": True})
        self.assertEqual(code, 0, err)

    def test_utf8_bom_is_accepted(self):
        code, _, err = run_check(b"\xef\xbb\xbf" + json.dumps({"enabled": True}).encode("utf-8"))
        self.assertEqual(code, 0, err)


class TestInvalidDeclaration(unittest.TestCase):
    def test_enabled_missing_exits_2_and_names_field(self):
        code, _, err = run_check({"schema_version": 1})
        self.assertEqual(code, 2)
        self.assertIn("enabled", err)

    def test_enabled_as_string_exits_2_and_names_field(self):
        code, _, err = run_check({"enabled": "true"})
        self.assertEqual(code, 2)
        self.assertIn("enabled", err)

    def test_unknown_field_exits_2_and_names_it(self):
        code, _, err = run_check({"enabled": True, "gates": []})
        self.assertEqual(code, 2)
        self.assertIn("gates", err)

    def test_every_invalid_field_is_named_in_one_run(self):
        code, _, err = run_check({"enabled": "yes", "extra": 1})
        self.assertEqual(code, 2)
        self.assertIn("enabled", err)
        self.assertIn("extra", err)

    def test_schema_version_zero_exits_2_and_names_field(self):
        code, _, err = run_check({"schema_version": 0, "enabled": True})
        self.assertEqual(code, 2)
        self.assertIn("schema_version", err)

    def test_float_schema_version_is_rejected(self):
        code, _, err = run_check({"schema_version": 1.0, "enabled": True})
        self.assertEqual(code, 2)
        self.assertIn("schema_version", err)

    def test_bool_schema_version_is_rejected(self):
        code, _, err = run_check({"schema_version": True, "enabled": True})
        self.assertEqual(code, 2)
        self.assertIn("schema_version", err)

    def test_invalid_json_exits_2(self):
        code, _, err = run_check("{not json")
        self.assertEqual(code, 2)
        self.assertIn("JSON", err)

    def test_duplicate_keys_exit_2(self):
        code, _, err = run_check('{"enabled": true, "enabled": false}')
        self.assertEqual(code, 2)
        self.assertIn("duplicad", err)

    def test_non_utf8_file_exits_2_without_traceback(self):
        code, _, err = run_check(b"\xff\xfe\x00{\x00")
        self.assertEqual(code, 2)
        self.assertNotIn("Traceback", err)

    def test_utf16_file_exits_2_without_traceback(self):
        code, _, err = run_check(json.dumps({"enabled": True}).encode("utf-16"))
        self.assertEqual(code, 2)
        self.assertNotIn("Traceback", err)

    def test_non_object_message(self):
        code, _, err = run_check([1, 2])
        self.assertEqual(code, 2)
        self.assertIn("objeto JSON", err)

    def test_missing_file_names_the_path(self):
        missing = Path(tempfile.gettempdir()) / "no-existe-1627.json"
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), str(missing)],
            capture_output=True,
            encoding="utf-8",
            errors="replace",
        )
        self.assertEqual(proc.returncode, 2)
        self.assertIn("no se puede leer", proc.stderr)
        self.assertIn("no-existe-1627.json", proc.stderr)


class TestDocstring(unittest.TestCase):
    def test_docstring_does_not_claim_full_apap_compatibility(self):
        doc = ast.get_docstring(ast.parse(SCRIPT.read_text(encoding="utf-8"))) or ""
        self.assertNotIn("que solo lee `enabled`", doc)
        self.assertIn('{"enabled": true}', doc)
        self.assertIn("solo acepta `enabled`", doc)


if __name__ == "__main__":
    unittest.main()
