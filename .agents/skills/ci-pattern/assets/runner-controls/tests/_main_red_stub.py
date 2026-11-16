"""Stub HTTP de main_red.py (#1708, parte A): módulo propio, `_github_stub.py` no se toca.

Registra cada escritura (POST, PATCH) en `state['writes']`.
"""

import contextlib
import json
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PULLS_SUFFIX = "/pulls"


def _route(method: str, path: str, state: dict) -> tuple:
    """(estado, cuerpo) de la petición; las escrituras quedan en state['writes']."""
    parts = path.split("/")
    if method == "GET" and len(parts) == 3 and parts[0] == "repos":
        return 200, {"default_branch": state.get("default_branch", "main")}
    if method == "GET" and "/actions/runs/" in path:
        if state.get("run_error"):
            return 500, {"message": "run_error fixture"}
        return 200, state["run"]
    if method == "GET" and path.endswith("/issues"):
        return 200, state.get("issues", [])
    if method == "GET" and "/commits/" in path and path.endswith(PULLS_SUFFIX):
        if state.get("pulls_error"):
            return 500, {"message": "pulls_error fixture"}
        return 200, state.get("pulls", [])
    if method == "POST" and path.endswith("/issues"):
        if state.get("write_error"):
            return 500, {"message": "write_error fixture"}
        return 201, {"number": state.get("created_number", 42),
                     "html_url": "https://github.com/owner/repo/issues/42"}
    if method == "POST" and path.endswith("/comments"):
        if state.get("write_error"):
            return 500, {"message": "write_error fixture"}
        return 201, {"id": 1}
    if method == "PATCH" and "/issues/" in path:
        if state.get("write_error"):
            return 500, {"message": "write_error fixture"}
        return 200, {"state": "closed"}
    return 404, {"message": "unsupported endpoint"}


class _Handler(BaseHTTPRequestHandler):
    def _serve(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        body = json.loads(raw) if raw else None
        url = urllib.parse.urlsplit(self.path)
        path = url.path.lstrip("/")
        state = self.server.state
        if self.command != "GET":
            state["writes"].append({"method": self.command, "path": path, "body": body})
        status, payload = _route(self.command, path, state)
        text = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(text)))
        self.end_headers()
        self.wfile.write(text)

    do_GET = _serve
    do_POST = _serve
    do_PATCH = _serve

    def log_message(self, *args: object) -> None:
        return


@contextlib.contextmanager
def main_red_stub(state: dict):
    """Arranca el servidor con `state` y da la URL base; `state['writes']` recoge las escrituras."""
    state.setdefault("writes", [])
    server = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
    server.state = state
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_address[1]}"
    finally:
        server.shutdown()
        server.server_close()
