#!/usr/bin/env python3
# ci-pattern asset tests — el generador del fichero de anotaciones (DysTelefonica/team-skills#1219)
"""El fichero que lleva la primera anotacion de error de cada check rojo.

El agregador no tiene red: las anotaciones de un check solo estan en la API de
check runs. Este generador las lee y escribe el fichero que el gate lee con
``--annotations``. Su salida es INFORMATIVA: si no puede producir el fichero,
avisa y degrada a «no annotation» en vez de poner el job en rojo.

Un named test por regla. El doble del API es un servidor HTTP local (stdlib),
el patron de la casa: el asset habla por HTTP con el cliente compartido
``runner-controls/github_rest.py`` y el test apunta ``GITHUB_API_URL`` al loopback.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import urlsplit

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "build_annotations.py"

REPO = "DysTelefonica/team-skills"
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
POLICY = {
    "required_jobs": ["lint", "test"],
    "events": {"pull_request": {"accepted_skips": []}},
    "check_names": {"lint": "lint ({event})"},
}
NEEDS = {"lint": {"result": "failure", "outputs": {}},
         "test": {"result": "success", "outputs": {}}}
CHECK_RUNS = {
    "total_count": 2,
    "check_runs": [
        {"id": 11, "name": "lint (pull_request)", "status": "completed", "conclusion": "failure"},
        {"id": 12, "name": "test (pull_request)", "status": "completed", "conclusion": "success"},
    ],
}
ANNOTATIONS = {
    "11": [
        {"annotation_level": "notice", "message": "un aviso primero"},
        {"annotation_level": "failure", "message": "PHP syntax error in src/Foo.php:12"},
    ],
    "12": [],
}


class _MockApi:
    """Servidor local con los dos endpoints que consume el generador.

    ``check_runs`` y ``annotations`` sustituyen los cuerpos por defecto, para
    poder servir un escenario que no sea el del modulo (un check verde con el
    nombre publicado, por ejemplo).
    """

    def __init__(self, *, broken: bool = False, check_runs: dict | None = None,
                 annotations: dict | None = None) -> None:
        self.broken = broken
        self.calls: list[str] = []
        check_runs_body = CHECK_RUNS if check_runs is None else check_runs
        annotations_body = ANNOTATIONS if annotations is None else annotations

        class _Handler(BaseHTTPRequestHandler):
            def do_GET(self) -> None:  # noqa: N802
                self.server.calls.append(self.path)  # type: ignore[attr-defined]
                path = urlsplit(self.path).path
                if self.server.broken:  # type: ignore[attr-defined]
                    self.send_response(500)
                    self.end_headers()
                    return
                if path.endswith("/check-runs"):
                    body = check_runs_body
                elif "/check-runs/" in path and path.endswith("/annotations"):
                    body = annotations_body.get(path.split("/check-runs/")[1].split("/")[0], [])
                else:
                    self.send_response(404)
                    self.end_headers()
                    return
                raw = json.dumps(body).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)

            def log_message(self, *_a, **_kw) -> None:
                return

        self._server = HTTPServer(("127.0.0.1", 0), _Handler)
        self._server.calls = self.calls  # type: ignore[attr-defined]
        self._server.broken = broken  # type: ignore[attr-defined]
        threading.Thread(target=self._server.serve_forever, daemon=True).start()
        self.base_url = f"http://127.0.0.1:{self._server.server_address[1]}"

    def close(self) -> None:
        self._server.shutdown()
        self._server.server_close()


def _run(api: _MockApi | None, *, policy: dict | None = None, needs: dict | None = None,
         event: str = "pull_request") -> tuple[subprocess.CompletedProcess, Path]:
    """Corre el generador; devuelve el proceso y el path del fichero de salida."""
    tmp = Path(tempfile.mkdtemp())
    policy_path = tmp / "policy.json"
    policy_path.write_text(json.dumps(policy if policy is not None else POLICY),
                           encoding="utf-8")
    needs_path = tmp / "needs.json"
    needs_path.write_text(json.dumps(needs if needs is not None else NEEDS), encoding="utf-8")
    out = tmp / "annotations.json"
    env = {**os.environ, "GITHUB_TOKEN": "test-token"}
    if api is not None:
        env["GITHUB_API_URL"] = api.base_url
    else:
        env.pop("GITHUB_API_URL", None)
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), "--policy", str(policy_path), "--needs-file",
         str(needs_path), "--repo", REPO, "--sha", SHA, "--event", event, "--out", str(out)],
        capture_output=True, text=True, encoding="utf-8", env=env)
    return proc, out


class GeneratorTests(unittest.TestCase):
    """El generador escribe la primera anotacion de error de cada check rojo."""

    def test_writes_the_first_error_annotation_of_the_red_check(self) -> None:
        api = _MockApi()
        try:
            proc, out = _run(api)
        finally:
            api.close()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(json.loads(out.read_text(encoding="utf-8")),
                         {"lint": "PHP syntax error in src/Foo.php:12"})

    def test_the_published_name_is_resolved_with_the_event(self) -> None:
        # La politica declara «lint ({event})»: el check se publica como
        # «lint (pull_request)» y es ese el que se busca.
        api = _MockApi()
        try:
            proc, _ = _run(api)
        finally:
            api.close()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertTrue(any("/commits/" in call and "/check-runs" in call for call in api.calls),
                        api.calls)
        self.assertTrue(any("/check-runs/11/annotations" in call for call in api.calls),
                        api.calls)

    def test_a_green_job_has_no_annotation(self) -> None:
        api = _MockApi()
        try:
            proc, out = _run(api)
        finally:
            api.close()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("test", json.loads(out.read_text(encoding="utf-8")))

    def test_api_down_degrades_to_no_annotation_with_a_warning(self) -> None:
        # La API caida no pone el job en rojo: avisa y no escribe el fichero.
        api = _MockApi(broken=True)
        try:
            proc, out = _run(api)
        finally:
            api.close()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertFalse(out.exists(), "un fichero degradado no se escribe: no se pasa al gate")
        self.assertIn("AVISO", proc.stderr)

    def test_without_a_token_it_degrades_instead_of_failing(self) -> None:
        tmp = Path(tempfile.mkdtemp())
        policy_path = tmp / "policy.json"
        policy_path.write_text(json.dumps(POLICY), encoding="utf-8")
        needs_path = tmp / "needs.json"
        needs_path.write_text(json.dumps(NEEDS), encoding="utf-8")
        out = tmp / "annotations.json"
        env = {k: v for k, v in os.environ.items() if k not in ("GITHUB_TOKEN", "GH_TOKEN")}
        proc = subprocess.run(
            [sys.executable, str(SCRIPT), "--policy", str(policy_path), "--needs-file",
             str(needs_path), "--repo", REPO, "--sha", SHA, "--event", "pull_request",
             "--out", str(out)],
            capture_output=True, text=True, encoding="utf-8", env=env)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertFalse(out.exists())
        self.assertIn("AVISO", proc.stderr)

    def test_a_declared_name_that_does_not_exist_is_reported(self) -> None:
        # El mapeo declarado que no aparece entre los check runs se avisa: el
        # consumer ve que su declaracion esta obsoleta.
        api = _MockApi()
        try:
            proc, out = _run(api, policy={**POLICY, "check_names": {"lint": "lint (otro)"}})
        finally:
            api.close()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("AVISO", proc.stderr)
        self.assertEqual(json.loads(out.read_text(encoding="utf-8")), {})

    def test_a_green_check_never_donates_its_annotation_to_a_red_job(self) -> None:
        # El `needs` dice que `lint` esta rojo, pero el unico check con ese
        # nombre publicado termino en VERDE. Su anotacion no es la causa: el
        # gate la imprimiria como si lo fuera (#1219, criterio 3).
        api = _MockApi(
            check_runs={"total_count": 1, "check_runs": [
                {"id": 11, "name": "lint (pull_request)", "status": "completed",
                 "conclusion": "success"},
            ]},
            annotations={"11": [{"annotation_level": "failure", "message": "todo verde"}]},
        )
        try:
            proc, out = _run(api)
        finally:
            api.close()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("lint", json.loads(out.read_text(encoding="utf-8")),
                         "un check verde no presta su anotacion a un job rojo")
        self.assertFalse([call for call in api.calls if "/annotations" in call], api.calls)
        self.assertIn("AVISO", proc.stderr)

    def test_the_output_is_the_flat_form_the_gate_reads(self) -> None:
        # El formato es el que el gate ya acepta: job -> mensaje.
        api = _MockApi()
        try:
            proc, out = _run(api)
        finally:
            api.close()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        data = json.loads(out.read_text(encoding="utf-8"))
        self.assertTrue(all(isinstance(v, str) for v in data.values()), data)


if __name__ == "__main__":
    unittest.main()
