#!/usr/bin/env python3
"""El cliente Python de la API reintenta como el helper bash (#1220).

Recorre la MISMA tabla de casos que la suite del helper bash
(`assets/gh-api-retry/casos.tsv`): las dos implementaciones se prueban contra
las mismas filas, así que no pueden divergir sin que salte un test.

La regla: se reintenta un fallo de llamada y un JSON malformado; una respuesta
válida vacía NO se reintenta; agotados los intentos, se falla cerrado.
"""

import contextlib
import importlib
import io
import json
import os
import sys
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent.parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
github_rest = importlib.import_module("github_rest")
GitHubApiError = github_rest.GitHubApiError

#: La tabla compartida: `assets/gh-api-retry/casos.tsv`.
TABLA = HERE.parent / "gh-api-retry" / "casos.tsv"

SECRET = "tok-secret-123"


def _casos():
    """Filas de la tabla, como diccionarios."""
    filas = []
    for line in TABLA.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        caso, falla, fallos, intentos, respuesta, resultado = line.split("\t")
        filas.append(
            {
                "caso": caso,
                "falla": falla,
                "fallos": int(fallos),
                "intentos": int(intentos),
                "respuesta": respuesta,
                "resultado": resultado,
            }
        )
    return filas


class _Handler(BaseHTTPRequestHandler):
    def _serve(self) -> None:
        path = self.path.split("?")[0]
        self.server.seen.append(path)
        route = self.server.routes.get(path)
        if callable(route):
            status, payload = route()
        elif route is None:
            status, payload = 200, []
        else:
            status, payload = route
        raw = payload if isinstance(payload, str) else json.dumps(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw.encode("utf-8"))))
        self.end_headers()
        self.wfile.write(raw.encode("utf-8"))

    def log_message(self, *args: object) -> None:
        return


_Handler.do_GET = _Handler._serve


class RetryTableTests(unittest.TestCase):
    """Cada fila de la tabla es una expectativa de las dos implementaciones."""

    def setUp(self) -> None:
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
        self.server.routes, self.server.seen = {}, []
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.base = f"http://127.0.0.1:{self.server.server_address[1]}"
        self.env = mock.patch.dict(
            os.environ,
            {"GITHUB_API_URL": self.base, "GITHUB_TOKEN": SECRET, "GH_API_RETRY_ESPERA_S": "0"},
            clear=False,
        )
        self.env.start()
        self.addCleanup(self.env.stop)
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)

    def _route(self, fila):
        """Ruta que falla las primeras `fallos` llamadas y despues responde."""
        llamadas = {"n": 0}

        def route():
            llamadas["n"] += 1
            if llamadas["n"] <= fila["fallos"]:
                if fila["falla"] == "error":
                    return 500, {"message": "boom"}
                if fila["falla"] == "json_malformado":
                    return 200, "<html>502 Bad Gateway</html>"
                return 200, []
            if fila["respuesta"] == "vacia":
                return 200, []
            return 200, [{"number": 1, "merged": True}]

        return route

    def test_cada_caso_de_la_tabla(self) -> None:
        ruta = "/repos/o/r/commits/sha/pulls"
        for fila in _casos():
            with self.subTest(caso=fila["caso"]):
                self.server.seen.clear()
                self.server.routes[ruta] = self._route(fila)
                buf = io.StringIO()
                with contextlib.redirect_stderr(buf):
                    try:
                        data = github_rest.get_json(ruta[1:])
                        error = None
                    except GitHubApiError as exc:
                        data, error = None, exc
                self.assertEqual(
                    fila["intentos"],
                    len(self.server.seen),
                    f"{fila['caso']}: numero de llamadas",
                )
                if fila["resultado"] == "ok":
                    self.assertIsNone(error, f"{fila['caso']}: no deberia fallar")
                    self.assertEqual([{"number": 1, "merged": True}], data)
                elif fila["resultado"] == "vacia":
                    self.assertIsNone(error, f"{fila['caso']}: no deberia fallar")
                    self.assertEqual([], data)
                else:
                    self.assertIsNotNone(error, f"{fila['caso']}: deberia fallar cerrado")
                    self.assertIsInstance(error, GitHubApiError)

    def test_un_intento_no_reintenta(self) -> None:
        """Con un solo intento, un fallo no se reintenta."""
        ruta = "/repos/o/r/commits/sha/pulls"
        self.server.routes[ruta] = lambda: (500, {"message": "boom"})
        with mock.patch.dict(os.environ, {"GH_API_RETRY_INTENTOS": "1"}):
            with self.assertRaises(GitHubApiError):
                github_rest.get_json(ruta[1:])
        self.assertEqual(1, len(self.server.seen))

    def test_agotados_los_intentos_nombra_el_caso_y_el_numero(self) -> None:
        """El fallo final dice qué llamada falló y tras cuántos intentos (B1)."""
        ruta = "/repos/o/r/commits/sha/pulls"
        self.server.routes[ruta] = lambda: (500, {"message": "boom"})
        with mock.patch.dict(os.environ, {"GH_API_RETRY_INTENTOS": "3"}):
            with self.assertRaises(GitHubApiError) as ctx:
                github_rest.get_json(ruta[1:])
        self.assertEqual(3, len(self.server.seen))
        self.assertIn("GET repos/o/r/commits/sha/pulls", str(ctx.exception))
        self.assertIn("tras 3 intentos", str(ctx.exception))
        self.assertEqual(500, ctx.exception.status)

    def test_una_respuesta_valida_vacia_no_se_reintenta(self) -> None:
        """La lista vacia es una respuesta valida: una sola llamada."""
        ruta = "/repos/o/r/commits/sha/pulls"
        self.server.routes[ruta] = (200, [])
        self.assertEqual([], github_rest.get_json(ruta[1:]))
        self.assertEqual(1, len(self.server.seen))


if __name__ == "__main__":
    unittest.main()
