#!/usr/bin/env python3
# ci-pattern asset tests — cliente REST de GitHub sin `gh` (DysTelefonica/team-skills#630)
"""Simula la API con un servidor HTTP local real: respuestas 2xx, errores,
timeouts, paginacion y ausencia de token. Cada fallo debe lanzar
GitHubApiError (fallo cerrado), nunca devolver datos parciales."""
import importlib
import json
import os
import socket
import sys
import threading
import time
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent.parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
github_rest = importlib.import_module("github_rest")
GitHubApiError = github_rest.GitHubApiError

SECRET = "tok-secret-123"


class _Handler(BaseHTTPRequestHandler):
    def _serve(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length) if length else b""
        # Un sondeo ajeno (`GET /` sin Authorization, #941) no entra en `seen`; el
        # cliente nunca pide la raiz, asi que el resto de peticiones si cuentan.
        if not (self.path.split("?")[0] == "/" and self.headers.get("Authorization") is None):
            self.server.seen.append({"method": self.command, "path": self.path,
                                     "auth": self.headers.get("Authorization"),
                                     "ctype": self.headers.get("Content-Type"), "body": body})
        route = self.server.routes.get(self.path.split("?")[0], (404, {"message": "Not Found"}))
        if callable(route):
            route = route(self.path)
        status, payload = route[0], route[1]
        extra = route[2] if len(route) > 2 else {}
        if status == "hang":
            time.sleep(2)
            return
        raw = payload if isinstance(payload, str) else json.dumps(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw.encode("utf-8"))))
        for name, value in extra.items():
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(raw.encode("utf-8"))

    def log_message(self, *args: object) -> None:
        return


# http.server exige estos nombres; se asignan para no definir funciones en mayusculas.
_Handler.do_GET = _Handler._serve
_Handler.do_POST = _Handler._serve


class RestTests(unittest.TestCase):
    def setUp(self) -> None:
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
        self.server.routes, self.server.seen = {}, []
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.base = f"http://127.0.0.1:{self.server.server_address[1]}"
        self.env = mock.patch.dict(os.environ, {"GITHUB_API_URL": self.base,
                                                "GITHUB_TOKEN": SECRET}, clear=False)
        self.env.start()
        self.addCleanup(self.env.stop)
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)

    def test_get_document_sends_bearer_token_and_decodes_json(self) -> None:
        self.server.routes["/repos/o/r"] = (200, {"full_name": "o/r", "default_branch": "main"})
        data = github_rest.get_json("repos/o/r")
        self.assertEqual(data, {"full_name": "o/r", "default_branch": "main"})
        self.assertEqual(self.server.seen[0]["auth"], f"Bearer {SECRET}")

    def test_github_token_wins_over_gh_token(self) -> None:
        with mock.patch.dict(os.environ, {"GH_TOKEN": "other"}):
            self.server.routes["/repos/o/r"] = (200, {})
            github_rest.get_json("repos/o/r")
        self.assertEqual(self.server.seen[0]["auth"], f"Bearer {SECRET}")

    def test_gh_token_is_used_when_github_token_is_absent(self) -> None:
        with mock.patch.dict(os.environ, {"GH_TOKEN": "fallback-tok"}):
            os.environ.pop("GITHUB_TOKEN", None)
            self.server.routes["/repos/o/r"] = (200, {})
            github_rest.get_json("repos/o/r")
        self.assertEqual(self.server.seen[0]["auth"], "Bearer fallback-tok")

    def test_a_foreign_probe_is_not_recorded(self) -> None:
        """RED (#941): un `GET /` sin Authorization no cuenta como llamada del cliente."""
        import http.client
        conn = http.client.HTTPConnection("127.0.0.1", self.server.server_address[1], timeout=5)
        conn.request("GET", "/")
        conn.getresponse().read()
        conn.close()
        self.assertEqual(self.server.seen, [])

    def test_an_unauthenticated_api_request_is_recorded(self) -> None:
        """Una peticion sin token a una ruta de API SI se registra: el filtro de sondeos no la oculta (#941)."""
        import http.client
        self.server.routes["/repos/o/r"] = (200, {})
        conn = http.client.HTTPConnection("127.0.0.1", self.server.server_address[1], timeout=5)
        conn.request("GET", "/repos/o/r")
        conn.getresponse().read()
        conn.close()
        self.assertEqual([(s["path"], s["auth"]) for s in self.server.seen], [("/repos/o/r", None)])

    def test_missing_token_fails_closed_without_sending_a_request(self) -> None:
        with mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop("GITHUB_TOKEN", None)
            os.environ.pop("GH_TOKEN", None)
            with self.assertRaises(GitHubApiError):
                github_rest.get_json("repos/o/r")
        self.assertEqual(self.server.seen, [])

    def test_http_error_fails_closed_and_never_leaks_the_token(self) -> None:
        self.server.routes["/repos/o/r"] = (500, {"message": "boom"})
        with self.assertRaises(GitHubApiError) as ctx:
            github_rest.get_json("repos/o/r")
        self.assertIn("HTTP 500", str(ctx.exception))
        self.assertNotIn(SECRET, str(ctx.exception))

    def test_http_error_carries_github_message_status_and_request_headers(self) -> None:
        headers = {"X-GitHub-Request-Id": "REQ-9", "X-Accepted-GitHub-Permissions": "workflows=write"}
        self.server.routes["/repos/o/r"] = (403, {"message": "Resource not accessible"}, headers)
        with self.assertRaises(GitHubApiError) as ctx:
            github_rest.get_json("repos/o/r")
        text = str(ctx.exception)
        self.assertTrue(text.startswith("HTTP 403 on GET repos/o/r"), text)
        for expected in ("Resource not accessible", "REQ-9", "workflows=write"):
            self.assertIn(expected, text)
        self.assertEqual(ctx.exception.status, 403)

    def test_http_error_detail_redacts_the_token_and_survives_a_non_json_body(self) -> None:
        self.server.routes["/repos/o/r"] = (403, {"message": f"echo {SECRET} back"})
        with self.assertRaises(GitHubApiError) as ctx:
            github_rest.get_json("repos/o/r")
        self.assertNotIn(SECRET, str(ctx.exception))
        self.server.routes["/repos/o/r"] = (502, "<html>bad gateway</html>")
        with self.assertRaises(GitHubApiError) as ctx:
            github_rest.get_json("repos/o/r")
        message = str(ctx.exception)
        self.assertTrue(message.startswith("HTTP 502 on GET repos/o/r"), message)
        self.assertTrue(message.endswith("(falló tras 3 intentos)"), message)

    def test_unreachable_api_fails_closed(self) -> None:
        with socket.socket() as probe:
            probe.bind(("127.0.0.1", 0))
            closed = probe.getsockname()[1]
        with mock.patch.dict(os.environ, {"GITHUB_API_URL": f"http://127.0.0.1:{closed}"}):
            with self.assertRaises(GitHubApiError) as ctx:
                github_rest.get_json("repos/o/r")
        self.assertIn("transport error", str(ctx.exception))

    def test_timeout_fails_closed(self) -> None:
        self.server.routes["/repos/o/r"] = ("hang", {})
        with mock.patch.object(github_rest, "API_TIMEOUT", 0.2):
            with self.assertRaises(GitHubApiError) as ctx:
                github_rest.get_json("repos/o/r")
        self.assertIn("transport error", str(ctx.exception))

    def test_invalid_json_fails_closed(self) -> None:
        self.server.routes["/repos/o/r"] = (200, "{not json")
        with self.assertRaises(GitHubApiError) as ctx:
            github_rest.get_json("repos/o/r")
        self.assertIn("not valid JSON", str(ctx.exception))

    def test_pulls_listing_follows_link_next_even_on_short_pages(self) -> None:
        next_url = f"{self.base}/repos/o/r/commits/abc/pulls?per_page=100&page=2"
        link = {"Link": f'<{next_url}>; rel="next", <{next_url}>; rel="last"'}
        self.server.routes["/repos/o/r/commits/abc/pulls"] = (
            lambda path: (200, [{"number": 1}], link) if "page=2" not in path
            else (200, [{"number": 2}]))
        data = github_rest.get_json("repos/o/r/commits/abc/pulls")
        self.assertEqual(data, [{"number": 1}, {"number": 2}])
        self.assertIn("page=2", self.server.seen[-1]["path"])

    def test_pulls_listing_with_query_follows_link_next(self) -> None:
        next_url = f"{self.base}/repos/o/r/pulls?state=closed&page=2"
        link = {"Link": f'<{next_url}>; rel="next"'}
        self.server.routes["/repos/o/r/pulls"] = (
            lambda path: (200, [{"number": 1}], link) if "page=2" not in path
            else (200, [{"number": 2}]))
        data = github_rest.get_json("repos/o/r/pulls?state=closed&base=main")
        self.assertEqual(data, [{"number": 1}, {"number": 2}])
        self.assertIn("page=2", self.server.seen[-1]["path"])

    def test_next_link_keeps_our_host_and_takes_only_the_path(self) -> None:
        evil = {"Link": '<https://evil.example/steal?page=2>; rel="next"'}
        self.server.routes["/repos/o/r/commits/abc/pulls"] = (200, [{"number": 1}], evil)
        self.server.routes["/steal"] = (200, [{"number": 2}])
        data = github_rest.get_json("repos/o/r/commits/abc/pulls")
        self.assertEqual(data, [{"number": 1}, {"number": 2}])
        self.assertEqual(self.server.seen[1]["path"], "/steal?page=2")

    def test_pulls_listing_that_is_not_an_array_fails_closed(self) -> None:
        self.server.routes["/repos/o/r/commits/abc/pulls"] = (200, {"message": "not a list"})
        with self.assertRaises(GitHubApiError):
            github_rest.get_json("repos/o/r/commits/abc/pulls")

    def test_post_sends_json_body_with_content_type(self) -> None:
        self.server.routes["/repos/o/r/issues"] = (201, {"number": 42})
        github_rest.call("POST", "repos/o/r/issues", {"title": "t", "body": "b"})
        sent = self.server.seen[0]
        self.assertEqual(sent["method"], "POST")
        self.assertEqual(sent["ctype"], "application/json")
        self.assertEqual(json.loads(sent["body"]), {"title": "t", "body": "b"})


if __name__ == "__main__":
    unittest.main()
