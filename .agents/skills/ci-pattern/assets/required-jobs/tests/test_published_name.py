#!/usr/bin/env python3
# ci-pattern asset tests — el nombre publicado del check en la causa (DysTelefonica/team-skills#1219)
"""El nombre del check que esta rojo, tal como lo publica el runner, y no solo
el id del job: el que lee el log del agregador tiene que casar la causa con el
check que ve en la UI de la PR.

El mapeo id -> nombre publicado es DATO DECLARADO en la politica del consumer
(``check_names``), no codigo del gate: el nombre publicado lo decide el workflow
del consumer. Un nombre que depende del evento se declara con el marcador
``{event}`` (``name: lint (${{ github.event_name }})``).

Un named test por regla. Las reglas que ya cumplian antes de #1219 quedan
fijadas aqui, para que un cambio posterior no las deje caer.
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_required_jobs.py"

REQUIRED = ["lint", "test"]
PR_ACCEPTED: list[str] = []


def _policy(check_names: object = None) -> dict:
    """La politica del consumer; ``check_names`` solo si el test lo declara."""
    policy: dict = {
        "required_jobs": list(REQUIRED),
        "events": {"pull_request": {"accepted_skips": list(PR_ACCEPTED)}},
    }
    if check_names is not None:
        policy["check_names"] = check_names
    return policy


def _run(needs: dict, *, policy: dict | None = None, event: str = "pull_request",
         extra: list[str] | None = None) -> subprocess.CompletedProcess:
    """Ejecuta el gate con un needs de una conclusions por job."""
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        policy_path = root / "policy.json"
        policy_path.write_text(json.dumps(policy if policy is not None else _policy()),
                               encoding="utf-8")
        needs_path = root / "needs.json"
        needs_path.write_text(json.dumps(
            {job: {"result": result, "outputs": {}} for job, result in needs.items()}),
            encoding="utf-8")
        return subprocess.run(
            [sys.executable, str(SCRIPT), "--policy", str(policy_path), "--event", event,
             "--needs-file", str(needs_path), *(extra or [])],
            capture_output=True, text=True, encoding="utf-8")


class PublishedNameTests(unittest.TestCase):
    """(aviso 3) la causa nombra el check publicado y conserva el id del job."""

    def test_cause_names_the_published_check_and_keeps_the_job_id(self) -> None:
        # El nombre publicado lleva el evento: se declara con {event}.
        needs = {"lint": "failure", "test": "success"}
        proc = _run(needs, policy=_policy({"lint": "lint ({event})"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("lint (pull_request) [lint] (failure)", proc.stdout)

    def test_cause_names_a_published_check_that_is_not_the_job_id(self) -> None:
        # Nombre publicado que no parece el id (00_CADETE: php-lint -> PHP Lint).
        needs = {"lint": "failure", "test": "success"}
        proc = _run(needs, policy=_policy({"lint": "PHP Lint (php -l)"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("PHP Lint (php -l) [lint] (failure)", proc.stdout)

    def test_published_name_equal_to_the_job_id_keeps_the_old_shape(self) -> None:
        # Sin divergencia, la causa sale igual que antes de #1219.
        needs = {"lint": "failure", "test": "success"}
        proc = _run(needs, policy=_policy({"lint": "lint"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("lint (failure):", proc.stdout)
        self.assertNotIn("[lint]", proc.stdout)

    def test_policy_without_check_names_keeps_the_old_shape(self) -> None:
        # La clave es opcional: sin ella la causa nombra el id, como siempre.
        needs = {"lint": "failure", "test": "success"}
        proc = _run(needs)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("lint (failure):", proc.stdout)

    def test_blocked_by_line_names_the_published_check(self) -> None:
        # La linea de causa tambien nombra el check publicado.
        needs = {"lint": "failure", "test": "success"}
        proc = _run(needs, policy=_policy({"lint": "lint ({event})"}))
        self.assertIn("blocked by: lint (pull_request) [lint] (failure)", proc.stdout)

    def test_annotation_is_quoted_with_the_published_check(self) -> None:
        # La anotacion travels con la causa, bajo el nombre publicado.
        needs = {"lint": "failure", "test": "success"}
        with tempfile.TemporaryDirectory() as tmp:
            ann = Path(tmp) / "annotations.json"
            ann.write_text(json.dumps({"lint": "PHP syntax error in src/Foo.php:12"}),
                           encoding="utf-8")
            proc = _run(needs, policy=_policy({"lint": "lint ({event})"}),
                        extra=["--annotations", str(ann)])
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("lint (pull_request) [lint] (failure): PHP syntax error in src/Foo.php:12",
                      proc.stdout)


class CheckNamesPolicyTests(unittest.TestCase):
    """La clave de la politica se valida: una duda no se queda esperando (#1219)."""

    def test_check_name_for_an_unknown_job_fails_closed(self) -> None:
        # Un id que no es job requerido es un error de la politica, no un nombre.
        needs = {"lint": "failure", "test": "success"}
        proc = _run(needs, policy=_policy({"nope": "nope"}))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("check_names", proc.stdout + proc.stderr)

    def test_check_names_must_be_an_object_of_strings(self) -> None:
        for value in (["lint"], {"lint": ""}, {"lint": 3}, {"lint": ["lint"]}):
            with self.subTest(value=value):
                needs = {"lint": "failure", "test": "success"}
                proc = _run(needs, policy=_policy(value))
                self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                self.assertIn("check_names", proc.stdout + proc.stderr)

    def test_check_names_is_not_a_relaxation(self) -> None:
        # Declarar el nombre publicado no cambia lo que el gate exige: no es
        # una relajacion contra la copia base (#1219).
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            base = root / "base.json"
            base.write_text(json.dumps(_policy()), encoding="utf-8")
            candidate = root / "candidate.json"
            candidate.write_text(json.dumps(_policy({"lint": "lint ({event})"})),
                                 encoding="utf-8")
            needs_path = root / "needs.json"
            needs_path.write_text(json.dumps(
                {"lint": {"result": "failure", "outputs": {}},
                 "test": {"result": "success", "outputs": {}}}), encoding="utf-8")
            proc = subprocess.run(
                [sys.executable, str(SCRIPT), "--policy", str(candidate), "--event",
                 "pull_request", "--needs-file", str(needs_path), "--base-policy", str(base)],
                capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("base comparison: performed, no relaxation detected", proc.stdout)


if __name__ == "__main__":
    unittest.main()
