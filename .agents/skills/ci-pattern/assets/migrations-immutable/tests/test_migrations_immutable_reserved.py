#!/usr/bin/env python3
# ci-pattern asset tests — migrations-immutable, parte 3: prefijos
# reservados por PRs abiertas contra main (#431, comentario del
# mantenedor 2026-10-13).
"""Tests del check de prefijos reservados por PRs abiertas.

Caso real APAP_WEB 2026-10-13: tres PR (#1465, #1565, #1573) anadian
cada una su propia ``011_*.sql``. El gate de rama base no detectaba
la colision; este check la cierra consultando la API REST de GitHub.

Estrategia de mock: un ``ThreadingHTTPServer`` local en el puerto que
asigne el sistema operativo, con un handler que sirve respuestas
predefinidas para ``/repos/{owner}/{repo}/pulls`` y
``/repos/{owner}/{repo}/pulls/{n}/files``. La env var
``GITHUB_API_BASE`` del script apunta al mock.
"""
import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _helpers import BaseMigrationTest, add_files, git, make_repo, run_gate, run_gate_full

# Importa el modulo bajo test para poder parchear ``parse_repo_slug``
# y asi no depender de un remote real en cada test.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import check_migrations_immutable as gate


def _make_mock_handler(state):
    """Fabrica un Handler con estado mutable en ``state`` (dict) (#941).

    ``state`` expone ``pulls_response``, ``files_response`` y ``next_status``.
    Cada test instancia su propio ``state`` (dict) en ``setUp`` y su propia
    clase Handler aqui, asi dos ``ReservedPrefixTest`` (o cualquier otro
    consumidor) no comparten respuestas aunque vivan en el mismo proceso.
    """

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802 (BaseHTTPRequestHandler API)
            if state["next_status"] != 200:
                self.send_error(state["next_status"], "simulated API failure")
                return
            # ``self.path`` incluye el query string (``?per_page=100``); lo
            # separamos para enrutar correctamente.
            path = self.path.split("?", 1)[0]
            if path.endswith("/pulls"):
                self._send_json(state["pulls_response"])
            elif "/pulls/" in path and path.endswith("/files"):
                try:
                    pr_number = int(path.split("/pulls/")[1].split("/")[0])
                except (IndexError, ValueError):
                    self.send_error(404)
                    return
                self._send_json(state["files_response"].get(pr_number, []))
            else:
                self.send_error(404)

        def _send_json(self, data: object) -> None:
            body = json.dumps(data).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, format: str, *args: object) -> None:  # noqa: A002
            pass  # silencia el log del servidor en los tests.

    return Handler


class ReservedPrefixTest(BaseMigrationTest):
    """Cubre el check de PRs abiertas contra main."""

    def setUp(self) -> None:
        super().setUp()
        # Crea un repo git con origin apuntando a un slug de GitHub.
        # parse_repo_slug del script parsea la URL y devuelve owner/repo.
        self.repo = make_repo(self.tmp, {
            "migrations/010_first.sql": "CREATE TABLE a;",
        })
        git(self.repo, "remote", "add", "origin", "https://github.com/owner/repo.git")
        # Estado por-instancia (dict). Sustituye a los antiguos atributos de clase
        # compartidos por todos los casos de la clase de test (#941).
        self.mock_state: dict = {"pulls_response": [], "files_response": {}, "next_status": 200}
        # Arranca el mock server en un puerto libre, con Handler propio.
        self.server = ThreadingHTTPServer(
            ("127.0.0.1", 0), _make_mock_handler(self.mock_state))
        self.server_thread = threading.Thread(
            target=self.server.serve_forever, daemon=True
        )
        self.server_thread.start()
        self.api_base = f"http://127.0.0.1:{self.server.server_address[1]}"
        # Parchea parse_repo_slug para que devuelva el slug esperado
        # sin depender de como el script parsea el remote.
        self.slug_patcher = mock.patch.object(
            gate, "parse_repo_slug", return_value="owner/repo"
        )
        self.slug_patcher.start()

    def tearDown(self) -> None:
        self.slug_patcher.stop()
        self.server.shutdown()
        self.server.server_close()
        super().tearDown()

    # -- Casos de prueba ---------------------------------------------------------

    def test_duplicate_prefix_in_open_pr_is_red(self) -> None:
        """S1+S2: PR con 011_a abierta + nueva migracion 011_b -> rojo."""
        self.mock_state["pulls_response"] = [{"number": 100}]
        self.mock_state["files_response"] = {
            100: [{"filename": "migrations/011_a.sql", "status": "added"}],
        }
        self.mock_state["next_status"] = 200

        add_files(self.repo, {
            "migrations/011_b.sql": "CREATE TABLE b;",
        })
        proc = run_gate_full(
            self.repo, "--base", "HEAD~1",
            env={"GITHUB_TOKEN": "test-token", "GITHUB_API_BASE": self.api_base},
        )
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("reserved", proc.stdout)
        self.assertIn("011", proc.stdout)
        self.assertIn("#100", proc.stdout)

    def test_distinct_prefix_in_open_pr_is_green(self) -> None:
        """PR con 011_a abierta + nueva migracion 012_a -> verde."""
        self.mock_state["pulls_response"] = [{"number": 100}]
        self.mock_state["files_response"] = {
            100: [{"filename": "migrations/011_a.sql", "status": "added"}],
        }
        self.mock_state["next_status"] = 200

        add_files(self.repo, {
            "migrations/012_a.sql": "CREATE TABLE b;",
        })
        proc = run_gate_full(
            self.repo, "--base", "HEAD~1",
            env={"GITHUB_TOKEN": "test-token", "GITHUB_API_BASE": self.api_base},
        )
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_api_failure_is_fail_closed(self) -> None:
        """S8: la API devuelve 403 (rate limit) -> rojo, no degrada a OK."""
        self.mock_state["next_status"] = 403

        add_files(self.repo, {
            "migrations/012_a.sql": "CREATE TABLE b;",
        })
        proc = run_gate_full(
            self.repo, "--base", "HEAD~1",
            env={"GITHUB_TOKEN": "test-token", "GITHUB_API_BASE": self.api_base},
        )
        self.assertEqual(proc.returncode, 3, proc.stdout + proc.stderr)
        self.assertIn("403", proc.stderr + proc.stdout)

    def test_no_token_is_fail_closed(self) -> None:
        """S4: GITHUB_TOKEN ausente -> rojo, no se puede medir."""
        add_files(self.repo, {
            "migrations/012_a.sql": "CREATE TABLE b;",
        })
        proc = run_gate_full(
            self.repo, "--base", "HEAD~1",
            env={"GITHUB_API_BASE": self.api_base},  # sin GITHUB_TOKEN
        )
        self.assertEqual(proc.returncode, 3, proc.stdout + proc.stderr)
        self.assertIn("GITHUB_TOKEN", proc.stderr + proc.stdout)

    def test_opt_out_skips_the_check(self) -> None:
        """Con --no-check-reserved-prefixes, el check no se ejecuta
        aunque haya API y token. Util para migracion gradual."""
        self.mock_state["pulls_response"] = [{"number": 100}]
        self.mock_state["files_response"] = {
            100: [{"filename": "migrations/011_a.sql", "status": "added"}],
        }
        self.mock_state["next_status"] = 200

        add_files(self.repo, {
            "migrations/011_b.sql": "CREATE TABLE b;",
        })
        # run_gate (no run_gate_full) ya pasa --no-check-reserved-prefixes.
        proc = run_gate(
            self.repo, "--base", "HEAD~1",
        )
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_historical_duplicate_does_not_exempt_new_reservation(self) -> None:
        """S6: la tolerancia de duplicados historicos en la base NO
        exonera a una nueva colision contra una PR abierta."""
        # La base tiene 005_a y 005_b (duplicado historico tolerado).
        # La PR abierta #200 reserva 005_c.
        # La nueva migracion 005_d colisiona por dos razones:
        # duplicado nuevo Y reservado. Cualquiera basta para rojo.
        historical_repo = self.tmp / "historical"
        historical_repo.mkdir()
        repo = make_repo(historical_repo, {
            "migrations/005_a.sql": "CREATE TABLE a;",
            "migrations/005_b.sql": "CREATE TABLE b;",
        })
        git(repo, "remote", "add", "origin", "https://github.com/owner/repo.git")

        self.mock_state["pulls_response"] = [{"number": 200}]
        self.mock_state["files_response"] = {
            200: [{"filename": "migrations/005_c.sql", "status": "added"}],
        }
        self.mock_state["next_status"] = 200

        add_files(repo, {
            "migrations/005_d.sql": "CREATE TABLE d;",
        })
        proc = run_gate_full(
            repo, "--base", "HEAD~1",
            env={"GITHUB_TOKEN": "test-token", "GITHUB_API_BASE": self.api_base},
        )
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        # La salida debe mencionar la reservada, no la eximir.
        self.assertIn("reserved", proc.stdout)
        self.assertIn("#200", proc.stdout)

    def test_modified_existing_migration_is_still_red(self) -> None:
        """El check de reservados NO exonera el resto de gates:
        una migracion editada sigue roja por la regla de inmutabilidad,
        aunque no haya PRs abiertas que la reserven."""
        self.mock_state["pulls_response"] = []
        self.mock_state["files_response"] = {}
        self.mock_state["next_status"] = 200

        # Modificar la migracion existente cambia su SHA.
        (self.repo / "migrations/010_first.sql").write_text(
            "CREATE TABLE a_changed;", encoding="utf-8"
        )
        git(self.repo, "add", "-A")
        git(self.repo, "commit", "-q", "-m", "edit")

        proc = run_gate_full(
            self.repo, "--base", "HEAD~1",
            env={"GITHUB_TOKEN": "test-token", "GITHUB_API_BASE": self.api_base},
        )
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("edit", proc.stdout)
