"""Servidor HTTP local que simula la API de GitHub para los tests de runner-controls (#630).

Sustituye al `gh` falso: los assets llaman a la API por HTTP (github_rest.py),
asi que el test levanta este servidor en loopback y apunta GITHUB_API_URL a el.
Las rutas y los flags de fallo (api_error, merge_rejected, head_changed...) son
los del shim anterior; el estado se persiste en state.json como antes, de modo
que los tests que lo leen siguen valiendo."""
import contextlib
import copy
import json
import os
import tempfile
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


def _atomic_write_json(path: Path, payload: dict) -> None:
    """Escribe ``payload`` en ``path`` de forma atomica con tempfile + os.replace.

    ``Path.write_text`` abre con O_TRUNC antes de escribir: si el lector del
    test (que corre en otro thread / proceso) abre el fichero durante la
    ventana entre truncate y write, ve un fichero vacio y rompe con
    JSONDecodeError. Bajo carga CPU con UNIT_PARALLEL>1 (#941, #511) esa
    ventana crecia a milisegundos y la rotura era intermitente.
    ``tempfile + os.replace`` es una operacion atomica en el mismo filesystem
    (rename(2) es atomica en POSIX), asi que el lector siempre ve la version
    anterior o la nueva completa, nunca un estado intermedio vacio."""
    data = json.dumps(payload).encode("utf-8")
    tmp_dir = path.parent
    with tempfile.NamedTemporaryFile(
        prefix=path.name + ".", suffix=".tmp", dir=str(tmp_dir), delete=False
    ) as tmp:
        tmp.write(data)
        tmp_path = Path(tmp.name)
    os.replace(tmp_path, path)

MERGED_SHA = "2c7d9b1a4e"


def _route(method: str, path: str, query: dict, body: dict, state: dict, base: str) -> tuple:
    """Devuelve (status, payload, headers) para la petición; muta state si es escritura."""
    state["calls"].append(f"{method} {path}")
    if state.get("api_error"):
        return 500, {"message": "api_error fixture"}, {}
    if method == "POST" and path.endswith("/issues"):
        if state.get("issue_post_error"):
            return 503, {"message": "issue_post_error fixture"}, {}
        state["created_issues"].append({"title": body.get("title", ""), "body": body.get("body", "")})
        repo = path[len("repos/"):-len("/issues")]
        return 201, {"number": 42, "html_url": f"https://github.com/{repo}/issues/42",
                     "title": body.get("title", "")}, {}
    if method == "POST" and path == "graphql":
        if state.get("graphql_error"):
            return 502, {"message": "graphql_error fixture"}, {}
        if state.get("graphql_errors"):
            return 200, {"errors": [{"message": "graphql_errors fixture"}]}, {}
        nodes = state.get("closing_refs", [])
        return 200, {"data": {"repository": {"pullRequest": {"closingIssuesReferences": {"nodes": nodes}}}}}, {}
    if method == "POST" and "/issues/" in path and path.endswith("/comments"):
        state.setdefault("comments", []).append({"path": path, "body": body.get("body", "")})
        return 201, {"id": 1}, {}
    if method == "PATCH" and "/issues/" in path and not path.endswith("/comments"):
        # #835: PATCH `/issues/{n}` puede recibir 422 cuando el cierre
        # automático de GitHub por la palabra `Closes` ya cerró la issue
        # (carrera con el PATCH explícito del gate). El re-read decide.
        if state.get("race_after_close"):
            return 422, {"message": "Validation Failed",
                         "errors": [{"resource": "Issue", "code": "already_closed"}]}, {}
        if state.get("close_error"):
            return 403, {"message": "Resource not accessible by integration"}, {
                "X-Accepted-GitHub-Permissions": "issues=write"}
        state.setdefault("closed_issues", []).append(
            {"number": int(path.rsplit("/", 1)[1]), "state": body.get("state"),
             "state_reason": body.get("state_reason")})
        return 200, {"state": "closed"}, {}
    if method == "PUT" and path.endswith("/merge"):
        if state.get("merge_rejected"):
            return 405, {"message": "merge_rejected fixture"}, {}
        if state.get("merge_forbidden"):
            return 403, {"message": state["merge_forbidden"]}, {
                "X-GitHub-Request-Id": "REQ-1234", "X-Accepted-GitHub-Permissions": "workflows=write"}
        if state.get("head_changed") and "sha" in body:
            return 409, {"message": "Head branch was modified"}, {}
        state["merges"].append(path)
        return 200, {"sha": MERGED_SHA, "merged": True, "message": "Pull Request successfully merged"}, {}
    if method == "POST" and "/actions/workflows/" in path and path.endswith("/dispatches"):
        # #817/#933: post-merge dispatch (`gh workflow run` directo). El cuerpo trae
        # `ref` = la rama base (la API rechaza un SHA con 422) y `inputs` con el
        # rango `before`/`after` cuando el workflow lo declara. 4xx/5xx caen al
        # fallo cerrado del caller.
        name = path.rsplit("/actions/workflows/", 1)[1].rsplit("/dispatches", 1)[0]
        if state.get("dispatch_error"):
            return 422, {"message": "dispatch_error fixture"}, {}
        if state.get("dispatch_forbidden"):
            return 403, {"message": "dispatch_forbidden fixture"}, \
                {"X-Accepted-GitHub-Permissions": "actions=write"}
        if state.get("dispatch_not_found"):
            return 404, {"message": "dispatch_not_found fixture"}, {}
        state.setdefault("dispatches", []).append(
            {"name": name, "ref": body.get("ref"), "inputs": body.get("inputs")})
        return 204, "", {}
    if method != "GET":
        return 404, {"message": "unsupported endpoint"}, {}
    if path.startswith("repos/") and path.count("/") == 2:
        return 200, {"default_branch": state.get("default_branch", "main")}, {}
    if path.endswith("/pulls") and "/commits/" in path:
        sha = path.split("/commits/")[1].split("/")[0]
        if state.get("invalid_body"):
            return 200, "{not valid json", {}
        return _pulls_page(state, sha, query, base)
    if "/check-runs" in path:
        sha = path.split("/commits/")[1].split("/")[0]
        runs = state["check_runs"].get(sha, [])
        script = state.get("check_runs_script", {}).get(sha)
        if script:  # una respuesta por consulta; la ultima se repite
            polls = state.setdefault("check_runs_polls", {})
            runs = script[min(polls.get(sha, 0), len(script) - 1)]
            polls[sha] = polls.get(sha, 0) + 1
        return 200, {"total_count": len(runs), "check_runs": runs}, {}
    if method == "GET" and path.endswith("/actions/workflows") \
            and "/runs" not in path and "/dispatches" not in path:
        # #817: lista de workflows para validar que cada nombre declarado en
        # `post_merge_dispatch` existe y declara `workflow_dispatch`. Sin
        # `state["workflows"]`, devuelve el conjunto vacío. #966: la respuesta
        # lleva la forma real de la API (`{"total_count", "workflows"}`) y se
        # pagina con `page`/`per_page`; `state["workflows_body"]` sustituye el
        # cuerpo entero para los dobles mal formados.
        body = state.get("workflows_body")
        return (200, body, {}) if body is not None else _workflows_page(state, query)
    if method == "GET" and "/contents/" in path:
        # #817: contenidos de un fichero del repo. La clave del fixture es la
        # ruta sin query (la API acepta `?ref=` y se ignora aquí).
        raw_path = path.split("?")[0]
        tail = raw_path.split("/contents/", 1)[1]
        file_path = state.get("workflow_files_paths", {}).get(tail, tail)
        entry = state.get("workflow_files", {}).get(file_path)
        if entry is None:
            return 404, {"message": "Not Found"}, {}
        return 200, entry, {}
    if "/compare/" in path:
        # `compares` fija la respuesta por rango ("base...head"); `compare` es el valor por defecto.
        spec = path.split("/compare/", 1)[1]
        default = state.get("compare", {"total_commits": 0, "commits": [], "files": [], "behind_by": 0})
        return 200, state.get("compares", {}).get(spec, default), {}
    if "/pulls/" in path:
        detail = state["pr_details"].get(path.rsplit("/pulls/", 1)[1])
        return (404, {"message": "Not Found"}, {}) if detail is None else (200, detail, {})
    if "/git/ref/heads/" in path:
        # #933: punta de la rama base antes del merge (`base_tips`); sin ella, 404.
        branch = path.rsplit("/git/ref/heads/", 1)[1]
        tip = state.get("base_tips", {}).get(branch)
        if tip is None:
            return 404, {"message": "Not Found"}, {}
        return 200, {"ref": f"refs/heads/{branch}", "object": {"sha": tip, "type": "commit"}}, {}
    if path.count("/") == 4 and "/commits/" in path:
        commit = state.get("commits", {}).get(path.split("/commits/")[1], {})
        return (200, commit, {}) if commit else (404, {"message": "Not Found"}, {})
    if "/issues/" in path and not path.endswith("/comments") and path.count("/") == 4:
        # #835: re-read de la issue tras un PATCH 422. Default `closed` para
        # reflejar el final típico de la carrera; los tests pueden sobreescribir
        # vía `state["issue_states"][N]`.
        number = int(path.rsplit("/", 1)[1])
        override = state.get("issue_states", {}).get(number)
        if override is not None:
            return 200, override, {}
        return 200, {"state": "closed", "state_reason": "completed"}, {}
    return 404, {"message": "unsupported endpoint"}, {}


def _workflows_page(state: dict, query: dict) -> tuple:
    """Pagina del listado de workflows con la forma real de la API (#966).

    `state["workflows"]` es el listado completo del repo (el doble del test) y el
    corte lo decide la query `page`/`per_page` del cliente; `total_count` es el
    tamaño de ese listado, como en `GET /repos/{repo}/actions/workflows`. Cada
    query servida queda en `state["workflows_queries"]` para que el test pueda
    fijar que se pide `per_page=100` y que se avanza de pagina.
    """
    items = state.get("workflows", [])
    state.setdefault("workflows_queries", []).append(query)
    per_page = int(query.get("per_page", ["30"])[0])
    page = int(query.get("page", ["1"])[0])
    start = max(page - 1, 0) * per_page
    return 200, {"total_count": len(items), "workflows": items[start:start + per_page]}, {}


def _pulls_page(state: dict, sha: str, query: dict, base: str) -> tuple:
    """Pulls de un commit; `pulls_pages` sirve varias paginas enlazadas con Link.

    La cabecera Link solo lleva datos que el servidor ya tiene: su direccion (`base`), el SHA
    guardado en `pulls_pages` (la clave que coincide, no el segmento de la peticion) y una
    plantilla de ruta fija. Nada de lo que envia el cliente llega a una cabecera de respuesta.
    """
    stored = next((key for key in state.get("pulls_pages", {}) if key == sha), None)
    if stored is None:
        return 200, state["pulls"].get(sha, []), {}
    pages = state["pulls_pages"][stored]
    number = int(query.get("page", ["1"])[0])
    headers = {}
    if number < len(pages):
        headers["Link"] = f'<{base}/repos/stub/stub/commits/{stored}/pulls?page={number + 1}>; rel="next"'
    return 200, pages[number - 1] if 0 < number <= len(pages) else [], headers


class _Handler(BaseHTTPRequestHandler):
    def _serve(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        body = json.loads(raw) if raw else {}
        url = urllib.parse.urlsplit(self.path)
        # Un sondeo ajeno (p. ej. un agente local que barre puertos de loopback con
        # `GET /` sin Authorization, #941) se responde sobre una copia y no deja rastro
        # en `calls` ni en state.json. El cliente nunca pide la raiz, asi que cualquier
        # otra peticion, tenga o no token, se registra.
        client = not (urllib.parse.urlsplit(self.path).path == "/"
                      and self.headers.get("Authorization") is None)
        state = self.server.state if client else copy.deepcopy(self.server.state)
        # La URL base sale de la direccion del servidor, no del Host de la
        # peticion: un Host ajeno acabaria copiado en la cabecera Link.
        host, port = self.server.server_address[:2]
        base = f"http://{host}:{port}"
        status, payload, headers = _route(self.command, url.path.lstrip("/"),
                                          urllib.parse.parse_qs(url.query), body, state, base)
        if client:
            _atomic_write_json(Path(self.server.state_path), state)
        text = payload if isinstance(payload, str) else json.dumps(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(text.encode("utf-8"))))
        for name, value in headers.items():
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(text.encode("utf-8"))

    do_GET = _serve
    do_POST = _serve
    do_PUT = _serve
    do_PATCH = _serve

    def log_message(self, *args: object) -> None:
        return


def stub_env(base: str) -> dict:
    """Entorno de los assets: API local y token de prueba (GITHUB_TOKEN gana a GH_TOKEN)."""
    return {**os.environ, "GITHUB_API_URL": base, "GITHUB_TOKEN": "test-token"}


@contextlib.contextmanager
def github_stub(state_path: Path, state: dict):
    """Arranca el servidor con el estado dado y devuelve la URL base para GITHUB_API_URL."""
    _atomic_write_json(Path(state_path), state)
    server = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
    server.state, server.state_path = state, str(state_path)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_address[1]}"
    finally:
        server.shutdown()
        server.server_close()
