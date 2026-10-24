#!/usr/bin/env python3
# ci-pattern asset — cliente REST minimo de GitHub (DysTelefonica/team-skills#630)
"""Peticiones REST a la API de GitHub con `urllib.request` (stdlib).

Sustituye la llamada al CLI `gh`: la imagen del runner no lo trae y su ruta
no es fija. El token sale de `GITHUB_TOKEN` o `GH_TOKEN` del entorno y nunca
aparece en un mensaje de error. HTTPS fijo (HTTP solo hacia loopback, para los
tests) y sin redirecciones: el token no viaja a otro host. Los listados se
pagian siguiendo la cabecera `Link` (rel="next"). Cualquier respuesta
distinta de 2xx, fallo de red, timeout o cuerpo que no sea JSON lanza
`GitHubApiError`: el llamador falla cerrado (HR-3). `GITHUB_API_URL` (la
variable que fija Actions) apunta a la raiz del host; por defecto,
https://api.github.com."""
import contextlib
import json
import os
import ssl
import urllib.error
import urllib.parse
import urllib.request
from typing import NamedTuple, NoReturn

API_TIMEOUT = 30
DETAIL_MAX = 300  # caracteres del `message` de GitHub que viajan al fallo
DIAGNOSTIC_HEADERS = ("X-Accepted-GitHub-Permissions", "X-GitHub-Request-Id")
MAX_PAGES = 50
OK_MIN, OK_MAX = 200, 300
LINK_PARTS = 2  # <url>; rel="..."
DEFAULT_API = "https://api.github.com"
LOOPBACK = ("127.0.0.1", "localhost")
HEADERS_BASE = {
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "ci-pattern-runner-controls",
}


class GitHubApiError(Exception):
    """La API no devolvio una respuesta valida; el llamador debe fallar cerrado.

    `status` lleva el codigo HTTP cuando el fallo es una respuesta no 2xx
    (0 en cualquier otro caso)."""

    def __init__(self, message: str, status: int = 0) -> None:
        super().__init__(message)
        self.status = status


def _fail(message: str, cause: BaseException | None = None, status: int = 0) -> NoReturn:
    raise GitHubApiError(message, status) from cause


def _diagnostic(exc: urllib.error.HTTPError) -> str:
    """`message` del cuerpo de GitHub y cabeceras de permisos, para el mensaje de fallo.

    Nunca lanza: si el cuerpo no se lee o no es JSON, el fallo conserva solo el
    codigo HTTP. El token se elimina del texto por si algun proxy lo reflejara."""
    parts = []
    try:
        body = json.loads(exc.read().decode("utf-8", errors="replace"))
        text = body.get("message") if isinstance(body, dict) else None
        if isinstance(text, str) and text:
            parts.append(text.replace("\n", " ")[:DETAIL_MAX])
    except (ValueError, OSError):
        pass
    parts += [f"{name}: {value}" for name in DIAGNOSTIC_HEADERS
              if (value := exc.headers.get(name))]
    detail = "; ".join(parts)
    tok = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN") or ""
    return detail.replace(tok, "***") if tok else detail


def token() -> str:
    tok = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if not tok:
        _fail("no GITHUB_TOKEN or GH_TOKEN in the environment")
    return tok


def _api_root() -> str:
    base = os.environ.get("GITHUB_API_URL", DEFAULT_API)
    try:
        parts = urllib.parse.urlsplit(base)
    except ValueError as exc:
        _fail(f"GITHUB_API_URL is not a valid URL: {exc}", exc)
    if parts.scheme == "https":
        return base
    if parts.scheme == "http" and (parts.hostname or "") in LOOPBACK:
        return base
    _fail("GITHUB_API_URL must be https (loopback http only for tests)")


class _Payload(NamedTuple):
    """Cuerpo, cabeceras y metodo de una peticion."""

    data: bytes | None
    headers: dict[str, str]
    method: str


def _open(scheme: str, netloc: str, tail: str, payload: _Payload, ctx: ssl.SSLContext):
    """Abre la peticion con el esquema que `_api_root` ya validó.

    Cada rama lleva su literal `https://` o `http://` dentro de la llamada:
    ruff S310 solo acepta una URL con esquema literal, y el origen sigue
    viniendo de GITHUB_API_URL (ya validado; http solo en loopback).
    """
    if scheme == "https":
        return urllib.request.urlopen(
            urllib.request.Request(f"https://{netloc}{tail}", data=payload.data, headers=payload.headers, method=payload.method),
            timeout=API_TIMEOUT, context=ctx)
    if scheme == "http":
        return urllib.request.urlopen(
            urllib.request.Request(f"http://{netloc}{tail}", data=payload.data, headers=payload.headers, method=payload.method),
            timeout=API_TIMEOUT, context=ctx)
    _fail(f"unsupported scheme for GitHub API: {scheme!r}")


def call(method: str, path: str, body: dict | None = None) -> tuple[str, str]:
    """Devuelve (cuerpo, cabecera Link) de una respuesta 2xx."""
    parts = urllib.parse.urlsplit(_api_root())
    tail = parts.path.rstrip("/") + "/" + path.lstrip("/")
    headers = dict(HEADERS_BASE)
    headers["Authorization"] = f"Bearer {token()}"
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    ctx = ssl.create_default_context()
    text = ""
    link = ""
    status = 0
    detail = ""
    try:
        with _open(parts.scheme, parts.netloc, tail, _Payload(data, headers, method), ctx) as response:
            text = response.read().decode("utf-8")
            link = response.headers.get("Link") or ""
            status = response.status
    except urllib.error.HTTPError as exc:
        # Las redirecciones 3xx nunca se siguen: GITHUB_API_URL fija el host y
        # el token no debe viajar a otro origen. Si llegara una 3xx, urlopen
        # la reemitiria; el catch explicito la convierte en fallo cerrado.
        # HTTPError es cerrable: libera la respuesta subyacente sin esperar al GC.
        with contextlib.closing(exc):
            status = exc.code
            detail = _diagnostic(exc)
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        _fail(f"transport error on {method} {path}: {exc}", exc)
    if not OK_MIN <= status < OK_MAX:
        suffix = f": {detail}" if detail else ""
        _fail(f"HTTP {status} on {method} {path}{suffix}", status=status)
    return text, link


def get_json(path: str) -> object:
    """GET de un documento. Un listado se pagina siguiendo `Link` y se fusiona."""
    text, link = call("GET", path)
    if not urllib.parse.urlsplit(path).path.endswith("/pulls"):
        return _parse(text, path)
    merged: list = []
    target = path
    for _ in range(MAX_PAGES):
        chunk = _parse(text, path)
        if not isinstance(chunk, list):
            _fail(f"listing {path} is not a JSON array")
        merged += chunk
        nxt = _next_path(link)
        if nxt is None:
            return merged
        target = nxt
        text, link = call("GET", target)
    _fail(f"listing {path} exceeds {MAX_PAGES} pages; refusing to continue")


def _next_path(link: str) -> str | None:
    """Ruta y query de `rel="next"`; el host lo fija GITHUB_API_URL, nunca la cabecera."""
    for part in link.split(","):
        segments = part.strip().split(";")
        if len(segments) == LINK_PARTS and segments[1].strip() == 'rel="next"':
            parts = urllib.parse.urlsplit(segments[0].strip().strip("<>"))
            return parts.path + (f"?{parts.query}" if parts.query else "")
    return None


def _parse(text: str, path: str) -> object:
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        _fail(f"response to {path} is not valid JSON: {exc}", exc)
