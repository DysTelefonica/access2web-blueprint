"""Stub de las ejecuciones de workflow por SHA, con filtro de evento (#1713).

El stub compartido (`_github_stub.py`) no sirve `GET repos/{repo}/actions/runs`. El
veredicto de `required` necesita saber qué ejecuciones de workflow son del evento
`pull_request` sobre el SHA de cabeza. Este módulo añade esa ruta sin tocar el stub
compartido: si el estado trae `workflow_runs_for_sha`, un `GET .../actions/runs` con
`head_sha` responde con esas ejecuciones filtradas por `event` (si la query lo pide);
`workflow_runs_error` fuerza un 500. El resto de peticiones va al stub compartido.
"""

import contextlib
import json
import threading
import urllib.parse
from http.server import ThreadingHTTPServer
from pathlib import Path

from _github_stub import _atomic_write_json, _Handler


class _SuiteHandler(_Handler):
    def _serve_suites(self) -> None:
        url = urllib.parse.urlsplit(self.path)
        query = urllib.parse.parse_qs(url.query)
        state = self.server.state
        if self.command == "GET" and url.path.endswith("/actions/runs") and "head_sha" in query:
            if state.get("workflow_runs_error"):
                self._reply(500, {"message": "workflow_runs_error fixture"})
                return
            runs = state.get("workflow_runs_for_sha", [])
            wanted = query.get("event", [None])[0]
            if wanted:
                runs = [run for run in runs if run.get("event") == wanted]
            self._reply(200, {"total_count": len(runs), "workflow_runs": runs})
            return
        _Handler._serve(self)

    def _reply(self, status: int, body: object) -> None:
        payload = json.dumps(body).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    do_GET = _serve_suites
    do_POST = _serve_suites
    do_PUT = _serve_suites
    do_PATCH = _serve_suites


@contextlib.contextmanager
def suites_github_stub(state_path: Path, state: dict):
    """Como `github_stub` del stub compartido, con las ejecuciones de workflow por SHA."""
    _atomic_write_json(Path(state_path), state)
    server = ThreadingHTTPServer(("127.0.0.1", 0), _SuiteHandler)
    server.state, server.state_path = state, str(state_path)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_address[1]}"
    finally:
        server.shutdown()
        server.server_close()
