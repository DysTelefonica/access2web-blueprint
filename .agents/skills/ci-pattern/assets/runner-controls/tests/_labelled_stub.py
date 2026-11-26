"""Stub de la API con el listado de issues filtrado por etiqueta (#1708).

El stub compartido (`_github_stub.py`) sirve la misma lista de issues para cualquier
etiqueta. La API real filtra por `labels=`: la señal `prerelease:main-red` y la nocturna
son dos etiquetas distintas. Este módulo añade ESA ruta sin tocar el stub compartido:
si el estado trae `issues_by_label`, un `GET .../issues?labels=<etiqueta>` responde con
`issues_by_label[<etiqueta>]` (o `[]`); el resto de peticiones va al stub compartido.
"""

import contextlib
import json
import threading
import urllib.parse
from http.server import ThreadingHTTPServer
from pathlib import Path

from _github_stub import _atomic_write_json, _Handler


class _LabelledHandler(_Handler):
    def _serve_labelled(self) -> None:
        url = urllib.parse.urlsplit(self.path)
        query = urllib.parse.parse_qs(url.query)
        by_label = self.server.state.get("issues_by_label")
        if self.command == "GET" and by_label is not None and url.path.endswith("/issues") \
                and "labels" in query:
            payload = json.dumps(by_label.get(query["labels"][0], [])).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return
        _Handler._serve(self)

    do_GET = _serve_labelled
    do_POST = _serve_labelled
    do_PUT = _serve_labelled
    do_PATCH = _serve_labelled


@contextlib.contextmanager
def labelled_github_stub(state_path: Path, state: dict):
    """Como `github_stub` del stub compartido, con el listado de issues filtrado por etiqueta."""
    _atomic_write_json(Path(state_path), state)
    server = ThreadingHTTPServer(("127.0.0.1", 0), _LabelledHandler)
    server.state, server.state_path = state, str(state_path)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_address[1]}"
    finally:
        server.shutdown()
        server.server_close()
