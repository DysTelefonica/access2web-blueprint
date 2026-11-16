#!/usr/bin/env python3
"""El cliente Python de la API reintenta como el helper bash (#1220, #1592).

Recorre la MISMA tabla de casos que la suite del helper bash
(`assets/gh-api-retry/casos.tsv`): las dos implementaciones se prueban contra
las mismas filas, así que no pueden divergir sin que salte un test.

La regla: se reintenta un error de red, un 5xx, un 429 y un JSON malformado
con backoff exponencial; un 4xx distinto de 429 y una respuesta válida vacía NO
se reintentan; agotados los intentos, se falla cerrado. Las esperas se observan
con `github_rest._esperar` sustituida: ningún test duerme.
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
#: Estado del doble que corta la conexión sin responder (error de red).
CORTE = "corte"
#: Respuesta HTTP de cada modo de fallo de la tabla.
FALLOS_HTTP = {"5xx": 503, "429": 429, "4xx": 404}
RUTA = "/repos/o/r/commits/sha/pulls"


def _casos():
    """Filas de la tabla, como diccionarios."""
    filas = []
    for line in TABLA.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        caso, falla, fallos, intentos, respuesta, resultado, esperas = line.split("\t")
        filas.append(
            {
                "caso": caso,
                "falla": falla,
                "fallos": int(fallos),
                "intentos": int(intentos),
                "respuesta": respuesta,
                "resultado": resultado,
                "esperas": [] if esperas == "-" else [float(x) for x in esperas.split(",")],
            }
        )
    return filas


class _Handler(BaseHTTPRequestHandler):
    def _serve(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        if length:
            self.rfile.read(length)
        path = self.path.split("?")[0]
        self.server.seen.append(path)
        route = self.server.routes.get(path)
        if callable(route):
            status, payload = route()
        elif route is None:
            status, payload = 200, []
        else:
            status, payload = route
        if status == CORTE:
            # Error de red: la conexión se cierra sin respuesta HTTP.
            return
        raw = payload if isinstance(payload, str) else json.dumps(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw.encode("utf-8"))))
        self.end_headers()
        self.wfile.write(raw.encode("utf-8"))

    def log_message(self, *args: object) -> None:
        return


_Handler.do_GET = _Handler._serve
_Handler.do_PUT = _Handler._serve
_Handler.do_PATCH = _Handler._serve
_Handler.do_POST = _Handler._serve


class _ApiDoble(unittest.TestCase):
    """API local, token de prueba y la espera inyectada: se observa, no se duerme."""

    def setUp(self) -> None:
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
        self.server.routes, self.server.seen = {}, []
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.base = f"http://127.0.0.1:{self.server.server_address[1]}"
        self.env = mock.patch.dict(
            os.environ,
            {"GITHUB_API_URL": self.base, "GITHUB_TOKEN": SECRET},
            clear=False,
        )
        self.env.start()
        os.environ.pop("GH_API_RETRY_ESPERA_S", None)
        os.environ.pop("GH_API_RETRY_INTENTOS", None)
        self.esperas: list[float] = []
        self.espera = mock.patch.object(github_rest, "_esperar", self.esperas.append, create=True)
        self.espera.start()
        self.addCleanup(self.espera.stop)
        self.addCleanup(self.env.stop)
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)

    @staticmethod
    def _secuencia(*respuestas):
        """Ruta que responde la secuencia dada; la ultima se repite."""
        llamadas = {"n": 0}

        def route():
            llamadas["n"] += 1
            return respuestas[min(llamadas["n"], len(respuestas)) - 1]

        return route


class RetryTableTests(_ApiDoble):
    """Cada fila de la tabla es una expectativa de las dos implementaciones."""

    def _route(self, fila):
        """Ruta que falla las primeras `fallos` llamadas y despues responde."""
        llamadas = {"n": 0}

        def route():
            llamadas["n"] += 1
            if llamadas["n"] <= fila["fallos"]:
                if fila["falla"] == "red":
                    return CORTE, ""
                if fila["falla"] in FALLOS_HTTP:
                    return FALLOS_HTTP[fila["falla"]], {"message": "boom"}
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
                self.esperas.clear()
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
                self.assertEqual(fila["esperas"], self.esperas, f"{fila['caso']}: esperas observadas")
                if not fila["esperas"]:
                    self.assertNotIn("AVISO", buf.getvalue(), f"{fila['caso']}: no avisa de reintento")
                if fila["resultado"] == "ok":
                    self.assertIsNone(error, f"{fila['caso']}: no deberia fallar")
                    self.assertEqual([{"number": 1, "merged": True}], data)
                elif fila["resultado"] == "vacia":
                    self.assertIsNone(error, f"{fila['caso']}: no deberia fallar")
                    self.assertEqual([], data)
                else:
                    self.assertIsNotNone(error, f"{fila['caso']}: deberia fallar cerrado")
                    self.assertIsInstance(error, GitHubApiError)
                    if fila["falla"] in FALLOS_HTTP:
                        self.assertEqual(FALLOS_HTTP[fila["falla"]], error.status)
                    if fila["intentos"] > 1:
                        self.assertIn(f"tras {fila['intentos']} intentos", str(error))

    def test_un_intento_no_reintenta(self) -> None:
        """Con un solo intento, un fallo no se reintenta."""
        ruta = "/repos/o/r/commits/sha/pulls"
        self.server.routes[ruta] = lambda: (500, {"message": "boom"})
        with mock.patch.dict(os.environ, {"GH_API_RETRY_INTENTOS": "1"}):
            with self.assertRaises(GitHubApiError):
                github_rest.get_json(ruta[1:])
        self.assertEqual(1, len(self.server.seen))
        self.assertEqual([], self.esperas)

    def test_agotados_los_intentos_nombra_el_caso_y_el_numero(self) -> None:
        """El fallo final dice qué llamada falló y tras cuántos intentos (B1)."""
        ruta = "/repos/o/r/commits/sha/pulls"
        self.server.routes[ruta] = lambda: (500, {"message": "boom"})
        with mock.patch.dict(os.environ, {"GH_API_RETRY_INTENTOS": "3"}):
            with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(GitHubApiError) as ctx:
                github_rest.get_json(ruta[1:])
        self.assertEqual(3, len(self.server.seen))
        self.assertIn("GET repos/o/r/commits/sha/pulls", str(ctx.exception))
        self.assertIn("tras 3 intentos", str(ctx.exception))
        self.assertEqual(500, ctx.exception.status)
        self.assertEqual([2.0, 4.0], self.esperas)

    def test_una_respuesta_valida_vacia_no_se_reintenta(self) -> None:
        """La lista vacia es una respuesta valida: una sola llamada."""
        ruta = "/repos/o/r/commits/sha/pulls"
        self.server.routes[ruta] = (200, [])
        self.assertEqual([], github_rest.get_json(ruta[1:]))
        self.assertEqual(1, len(self.server.seen))

    def test_la_espera_base_del_entorno_escala_el_backoff(self) -> None:
        """`GH_API_RETRY_ESPERA_S` es la primera espera; cada reintento la dobla (#1592)."""
        self.server.routes[RUTA] = lambda: (503, {"message": "boom"})
        with mock.patch.dict(os.environ, {"GH_API_RETRY_ESPERA_S": "0.5"}):
            with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(GitHubApiError):
                github_rest.get_json(RUTA[1:])
        self.assertEqual([0.5, 1.0, 2.0], self.esperas)

    def test_un_4xx_no_se_reintenta_y_conserva_su_mensaje(self) -> None:
        """Un 403 no es transitorio: un intento, sin espera, el error tal cual (#1592)."""
        self.server.routes[RUTA] = lambda: (403, {"message": "Resource not accessible"})
        buf = io.StringIO()
        with contextlib.redirect_stderr(buf), self.assertRaises(GitHubApiError) as ctx:
            github_rest.get_json(RUTA[1:])
        self.assertEqual(1, len(self.server.seen))
        self.assertEqual([], self.esperas)
        self.assertEqual(403, ctx.exception.status)
        self.assertTrue(str(ctx.exception).startswith("HTTP 403 on GET repos/o/r/commits/sha/pulls"))
        self.assertNotIn("intentos", str(ctx.exception))
        self.assertEqual("", buf.getvalue())

    def test_sin_token_no_se_reintenta(self) -> None:
        """Un fallo de configuración no es transitorio: ni llamada ni espera (#1592)."""
        with mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop("GITHUB_TOKEN", None)
            os.environ.pop("GH_TOKEN", None)
            with contextlib.redirect_stderr(io.StringIO()) as buf, self.assertRaises(GitHubApiError):
                github_rest.get_json("repos/o/r")
        self.assertEqual([], self.server.seen)
        self.assertEqual([], self.esperas)
        self.assertEqual("", buf.getvalue())

    def test_un_409_del_merge_no_se_reintenta(self) -> None:
        """`call_con_reintento`: un 409 en `PUT .../merge` es un 4xx, sin reintento (#1592)."""
        ruta = "/repos/o/r/pulls/7/merge"
        self.server.routes[ruta] = lambda: (409, {"message": "Head branch was modified"})
        buf = io.StringIO()
        with contextlib.redirect_stderr(buf), self.assertRaises(GitHubApiError) as ctx:
            github_rest.call_con_reintento("PUT", ruta[1:], {"sha": "abc"})
        self.assertEqual(409, ctx.exception.status)
        self.assertEqual(1, len(self.server.seen))
        self.assertEqual([], self.esperas)
        self.assertEqual("", buf.getvalue())


class IdempotentWriteRetryTests(_ApiDoble):
    """Las escrituras idempotentes reintentan con la misma regla; las no idempotentes, no (#1592)."""

    def test_put_idempotente_reintenta_un_5xx(self) -> None:
        ruta = "/repos/o/r/pulls/7/merge"
        self.server.routes[ruta] = self._secuencia((503, {"message": "boom"}), (200, {"merged": True}))
        with contextlib.redirect_stderr(io.StringIO()):
            text, _ = github_rest.call_con_reintento("PUT", ruta[1:], {"sha": "abc"})
        self.assertEqual({"merged": True}, json.loads(text))
        self.assertEqual(2, len(self.server.seen))
        self.assertEqual([2.0], self.esperas)

    def test_patch_idempotente_reintenta_un_error_de_red(self) -> None:
        ruta = "/repos/o/r/issues/5"
        self.server.routes[ruta] = self._secuencia((CORTE, ""), (200, {"state": "closed"}))
        with contextlib.redirect_stderr(io.StringIO()):
            github_rest.call_con_reintento("PATCH", ruta[1:], {"state": "closed"})
        self.assertEqual(2, len(self.server.seen))
        self.assertEqual([2.0], self.esperas)

    def test_escritura_idempotente_no_reintenta_un_4xx(self) -> None:
        ruta = "/repos/o/r/issues/5"
        self.server.routes[ruta] = lambda: (422, {"message": "Validation Failed"})
        with self.assertRaises(GitHubApiError) as ctx:
            github_rest.call_con_reintento("PATCH", ruta[1:], {"state": "closed"})
        self.assertEqual(422, ctx.exception.status)
        self.assertEqual(1, len(self.server.seen))
        self.assertEqual([], self.esperas)

    def test_escritura_idempotente_agotada_falla_cerrado(self) -> None:
        ruta = "/repos/o/r/pulls/7/merge"
        self.server.routes[ruta] = lambda: (502, {"message": "boom"})
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(GitHubApiError) as ctx:
            github_rest.call_con_reintento("PUT", ruta[1:], {"sha": "abc"})
        self.assertEqual(4, len(self.server.seen))
        self.assertEqual([2.0, 4.0, 8.0], self.esperas)
        self.assertIn("PUT repos/o/r/pulls/7/merge", str(ctx.exception))
        self.assertIn("tras 4 intentos", str(ctx.exception))
        self.assertEqual(502, ctx.exception.status)

    def test_post_no_idempotente_no_se_reintenta(self) -> None:
        """`call` (un POST de comentario o de issue) hace un solo intento ante un 5xx."""
        ruta = "/repos/o/r/issues/5/comments"
        self.server.routes[ruta] = lambda: (503, {"message": "boom"})
        with self.assertRaises(GitHubApiError) as ctx:
            github_rest.call("POST", ruta[1:], {"body": "x"})
        self.assertEqual(503, ctx.exception.status)
        self.assertEqual(1, len(self.server.seen))
        self.assertEqual([], self.esperas)


if __name__ == "__main__":
    unittest.main()
