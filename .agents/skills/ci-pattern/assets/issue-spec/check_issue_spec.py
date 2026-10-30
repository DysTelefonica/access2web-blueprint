#!/usr/bin/env python3
# ci-pattern asset — gate issue-spec (DysTelefonica/team-skills#633)
"""Gate de trazabilidad: el cuerpo del PR enlaza una issue del repositorio.

Reglas:

1. El cuerpo lleva `Closes|Fixes|Resolves|Refs #N` fuera de código.
2. La referencia resuelve a una issue, nunca a una pull request.
5. La referencia que cierra apunta al repositorio del PR. Un `Refs` a otro
   repositorio no cierra nada y no se valida (DysTelefonica/team-skills#1093).

Retirado por decisión del dueño (DysTelefonica/team-skills#1093, R1 y R2): la
etiqueta `status:approved` de la issue y su comprobación de quién la puso (la
lista `approvers` del host-contract), la etiqueta `type:*` de la issue y del
PR, las seis secciones obligatorias del cuerpo de la issue y el resto de las
validaciones del cuerpo. La PR bloquea por el enlace a la issue, el tamaño y
los commits convencionales (gates `pr-size` y `pr-contract`, otros assets).

Qué cuenta como referencia sigue a GitHub: las palabras de cierre valen en
prosa, pero no dentro de un span de código ni de un bloque cercado.

Modos:

- `local --body-file <ruta|-> [--repository owner/repo]`: forma del cuerpo. Sin
  payload, sin token ni red.
- `pr-event <event.json>`: el cuerpo del PR se lee por API y las referencias se
  resuelven contra las issues del repositorio. Usa GITHUB_TOKEN y
  GITHUB_API_URL del entorno si existen. Las flags obsoletas
  `--required-sections` y `--host-contract` se aceptan y se ignoran (#1103);
  el gate avisa por stderr y no cambia el veredicto. `--required-sections`
  existia tambien en el modo `local`, asi que ahi se acepta igual.

Un PR de Dependabot (actor `dependabot[bot]` y rama de Dependabot) queda exento
(#1113): sale 0 sin leer el cuerpo ni las issues.
Exit codes: 0 sin hallazgos · 1 hallazgos · 2 entrada inválida.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# Detección compartida con check_branch_name.py (#1113).
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "dependabot"))
from dependabot import is_dependabot_pr  # noqa: E402

DEPENDS_ON_HINT = (
    "Refs/Closes solo apuntan a issues del mismo repositorio; la PR de la que "
    "depende va en depends-on de Chain Context")

# Palabras de cierre que GitHub reconoce, más la referencia no cerrante `Refs`.
CLOSE_KEYWORDS = "close|closes|closed|fix|fixes|fixed|resolve|resolves|resolved"
REFERENCE_RX = re.compile(
    rf"(?i)(?P<kind>\b(?:{CLOSE_KEYWORDS})|\brefs)\s+"
    r"(?P<repo>[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)?#(?P<number>[0-9]+)\b")
FENCE_RX = re.compile(r"^ {0,3}(`{3,}|~{3,})")
PARAGRAPH_BREAK_RX = re.compile(r"\n\s*\n")
# Un span de código de n backticks se cierra con otro run de exactamente n.
CODE_SPAN_RX = re.compile(r"(?<![\\`])(`+)(?!`)(.+?)(?<!`)\1(?!`)", re.DOTALL)
CODE_PLACEHOLDER = "\x00"  # no es espacio: `Closes `x` #1` no forma una referencia


class GitHubApiError(RuntimeError):
    """Raised when GitHub cannot provide authoritative issue data."""


@dataclass(frozen=True)
class Reference:
    kind: str  # "close" | "refs"
    repo: str  # "" cuando la referencia no lleva owner/repo
    number: int


def strip_code(body: str) -> str:
    """Quita bloques cercados y spans de código: GitHub no cierra desde ahí."""
    kept: list[str] = []
    fence: str | None = None
    for line in body.splitlines():
        if fence is None:
            opened = FENCE_RX.match(line)
            if opened:
                fence = opened.group(1)
            else:
                kept.append(line)
            continue
        closing = line.strip()
        if closing and set(closing) == {fence[0]} and len(closing) >= len(fence):
            fence = None
    paragraphs = PARAGRAPH_BREAK_RX.split("\n".join(kept))
    return "\n\n".join(CODE_SPAN_RX.sub(CODE_PLACEHOLDER, p) for p in paragraphs)


def references(body: str) -> tuple[Reference, ...]:
    """Referencias a issues fuera de código, sin duplicados, en orden estable."""
    found = {
        Reference("refs" if match.group("kind").lower() == "refs" else "close",
                  match.group("repo") or "", int(match.group("number")))
        for match in REFERENCE_RX.finditer(strip_code(body or ""))
    }
    return tuple(sorted(found, key=lambda r: (r.number, r.kind, r.repo)))


# --------------------------------------------------------------------------------------------
# Reglas de forma: solo leen el cuerpo. Son las que corre el modo local.
# --------------------------------------------------------------------------------------------


def rule_has_reference(refs: tuple[Reference, ...]) -> str | None:
    """Regla 1: el cuerpo debe enlazar una issue fuera de código."""
    if refs:
        return None
    return ("el cuerpo no enlaza ninguna issue fuera de código: añada `Closes #N` "
            "(punta de cadena) o `Refs #N` (intermedio). Una PR de la que depende "
            "va en depends-on de Chain Context, no aquí")


def rule_same_repository(refs: tuple[Reference, ...], repository: str) -> list[str]:
    """Regla 5: el enlace apunta al repositorio del PR. Sin repositorio, no se juzga."""
    if not repository:
        return []
    return [
        f"{ref.repo}#{ref.number}: la referencia apunta a otro repositorio; el enlace "
        f"va al repositorio del PR ({repository})"
        for ref in refs
        if ref.repo and ref.repo.lower() != repository.lower()
    ]


def judged_references(body: str, repository: str = "") -> tuple[Reference, ...]:
    """Las referencias que el gate juzga: un `Refs` a otro repositorio se ignora (#1093).

    Sin repositorio conocido no hay con qué comparar, así que se juzga todo igual.
    """
    refs = references(body)
    if not repository:
        return refs
    return tuple(
        ref for ref in refs
        if not (ref.kind == "refs" and ref.repo and ref.repo.lower() != repository.lower())
    )


def validate_body(body: str, repository: str = "") -> list[str]:
    """Modo local: forma del cuerpo, sin tocar la red."""
    refs = judged_references(body, repository)
    violation = rule_has_reference(refs)
    if violation:
        return [violation]
    return rule_same_repository(refs, repository)


# --------------------------------------------------------------------------------------------
# Reglas de la issue remota: las que necesitan la API.
# --------------------------------------------------------------------------------------------


def rule_reference_is_issue(issue: Mapping[str, Any], number: int) -> str | None:
    """Regla 2: `GET /issues/{n}` devuelve `pull_request` cuando el número es una PR."""
    if "pull_request" not in issue:
        return None
    return (f"#{number}: la referencia resuelve a una pull request, no a una issue; "
            f"muévala a depends-on de Chain Context. {DEPENDS_ON_HINT}")


def rule_api_error(number: int, exc: GitHubApiError, what: str = "the issue") -> str:
    """Regla 6: un error de la API falla cerrado."""
    return f"#{number}: cannot read {what} ({exc})"


def validate_remote(
    body: str,
    repository: str,
    fetch_issue: Callable[[int], Mapping[str, Any]],
) -> list[str]:
    """Modo pr-event: el cuerpo enlaza una issue del repositorio.

    Regla por PR: el cuerpo lleva al menos una referencia (`Closes`/`Refs`) que
    resuelve a una issue del propio repositorio. Un error de la API al
    consultar una issue es hallazgo, no verde silencioso (regla 6, fail-closed).
    """
    violations = validate_body(body, repository)
    refs = judged_references(body, repository)
    if not refs:
        return violations
    same_repo_refs = tuple(
        ref for ref in refs
        if not ref.repo or ref.repo.lower() == repository.lower()
    )
    fetched: dict[int, Mapping[str, Any]] = {}

    def _get(number: int) -> Mapping[str, Any]:
        if number not in fetched:
            fetched[number] = fetch_issue(number)
        return fetched[number]

    for number in sorted({ref.number for ref in same_repo_refs}):
        try:
            issue = _get(number)
        except GitHubApiError as exc:
            violations.append(rule_api_error(number, exc))
            continue
        if message := rule_reference_is_issue(issue, number):
            violations.append(message)
    return violations


# --------------------------------------------------------------------------------------------
# GitHub client: solo lectura. `issues: read` basta.
# --------------------------------------------------------------------------------------------


class GitHubIssueFailedError(GitHubApiError):
    def __init__(self, repository: str, number: int, exc: BaseException) -> None:
        super().__init__(f"GET /repos/{repository}/issues/{number} failed: {exc}")


class GitHubPullRequestFailedError(GitHubApiError):
    def __init__(self, repository: str, number: int, exc: BaseException) -> None:
        super().__init__(f"GET /repos/{repository}/pulls/{number} failed: {exc}")


class UnsupportedApiUrlError(GitHubApiError):
    def __init__(self, api_url: str) -> None:
        super().__init__(f"api url must be https (loopback http only): {api_url}")


def _open(api_url: str, path: str, headers: Mapping[str, str],
          query: Mapping[str, Any] | None = None):
    """Abre el GET con el esquema validado: https, o http solo en loopback.

    Cada rama lleva su literal dentro de la llamada: ruff S310 solo acepta
    una URL con esquema literal.
    """
    parts = urllib.parse.urlsplit(api_url)
    tail = parts.path.rstrip("/") + path
    if query:
        tail = f"{tail}?{urllib.parse.urlencode(query)}"
    if parts.scheme == "https":
        return urllib.request.urlopen(
            urllib.request.Request(f"https://{parts.netloc}{tail}", headers=dict(headers)),
            timeout=30)
    if parts.scheme == "http" and parts.hostname in ("127.0.0.1", "localhost"):
        return urllib.request.urlopen(
            urllib.request.Request(f"http://{parts.netloc}{tail}", headers=dict(headers)),
            timeout=30)
    raise UnsupportedApiUrlError(api_url)


class GitHubClient:
    def __init__(self, token: str = "", api_url: str = "https://api.github.com") -> None:
        self._token = token
        self._api_url = api_url.rstrip("/")

    def _headers(self) -> dict[str, str]:
        headers = {
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "ci-pattern-issue-spec",
        }
        if self._token:
            headers["Authorization"] = f"Bearer {self._token}"
        return headers

    def issue(self, repository: str, number: int) -> Mapping[str, Any]:
        try:
            with _open(self._api_url, f"/repos/{repository}/issues/{number}",
                       self._headers()) as response:
                return json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubIssueFailedError(repository, number, exc) from exc

    def pull_request(self, repository: str, number: int) -> Mapping[str, Any]:
        # El PR del mode `pr-event` se trae por API (#830): el snapshot del
        # payload queda obsoleto en cuanto el cuerpo se edita, y todo
        # "Re-run jobs" del UI reusa el mismo payload (el workaround de hoy
        # es un commit vacío para forzar `synchronize`, ardelperal/APAP_WEB#2163).
        try:
            with _open(self._api_url, f"/repos/{repository}/pulls/{number}",
                       self._headers()) as response:
                return json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubPullRequestFailedError(repository, number, exc) from exc


def _make_fetcher(repository: str) -> Callable[[int], Mapping[str, Any]]:
    """Fetcher desde `$GITHUB_TOKEN` y `$GITHUB_API_URL`. Los tests lo sustituyen."""
    client = GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )
    return lambda number: client.issue(repository, number)


def _make_body_fetcher() -> Callable[[str, int], Mapping[str, Any]]:
    """Fetcher del cuerpo del PR desde `$GITHUB_TOKEN` y `$GITHUB_API_URL` (#830)."""
    client = GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )
    return lambda repository, number: client.pull_request(repository, number)


def _read_body(path: str) -> str:
    if path == "-":
        return sys.stdin.read()
    return Path(path).read_text(encoding="utf-8")


def _event_subject(payload: Mapping[str, Any]) -> tuple[str, Mapping[str, Any], int] | None:
    """El repositorio, el PR y su número del evento, o ``None`` si el evento no los identifica.

    El motivo del rechazo va a stderr; el código de salida (2) lo decide quien llama.
    Vive aparte porque ``pr_event`` tiene un tope de seis ``return`` (PLR0911 del
    ratchet de los assets vendorizables) y la exención de Dependabot (#1113) ocupa
    uno: aquí la validación del evento no le cuesta ninguno.
    """
    repository = (payload.get("repository") or {}).get("full_name", "")
    pull = payload.get("pull_request") or {}
    pr_number = pull.get("number")
    if "/" not in repository:
        print("issue-spec: el evento no identifica repository.full_name", file=sys.stderr)
        return None
    if not isinstance(pr_number, int) or pr_number <= 0:
        print("issue-spec: pr-event: el payload no identifica pull_request.number "
              "entero positivo; el cuerpo del PR solo se lee por API", file=sys.stderr)
        return None
    return repository, pull, pr_number


def pr_event(
    payload: Mapping[str, Any],
    *,
    body_fetcher: Callable[[str, int], Mapping[str, Any]] | None = None,
    issue_fetcher: Callable[[int], Mapping[str, Any]] | None = None,
) -> int:
    """Modo ``pr-event``: el cuerpo del PR se lee por API, nunca del payload (#830).

    El snapshot ``payload["pull_request"]["body"]`` queda obsoleto desde la primera
    edición del cuerpo y desde cualquier "Re-run jobs" del UI (porque reusa el
    mismo ``$GITHUB_EVENT_PATH``); el workaround de hoy es un commit vacío para
    forzar ``synchronize`` y un payload fresco — ardelperal/APAP_WEB#2163,
    #2164. La fuente autoritativa del cuerpo es la API: ``GET /repos/{owner}/{repo}/pulls/{n}``
    vía el ``GitHubClient`` del asset. Sin ``GITHUB_TOKEN`` o con la API caída,
    el gate falla cerrado con exit 2 y un motivo explícito (nunca fallback
    silencioso al cuerpo del payload, HR-3).

    ``body_fetcher`` e ``issue_fetcher`` son inyectables: ``main()`` los deja a
    ``None`` para usar el cliente real desde el entorno, y los tests los
    sustituyen por dobles.
    """
    subject = _event_subject(payload)
    if subject is None:
        return 2
    repository, pull, pr_number = subject

    # #1113 E: Dependabot no escribe el enlace a una issue; su PR se exime sin leer
    # el cuerpo ni las issues. Actor y rama tienen que coincidir (ver el helper).
    if is_dependabot_pr((pull.get("user") or {}).get("login"), (pull.get("head") or {}).get("ref")):
        print("issue-spec: PR de Dependabot exenta del enlace a issue (#1113)")
        return 0

    if body_fetcher is None:
        token = os.environ.get("GITHUB_TOKEN", "")
        if not token:
            print("issue-spec: pr-event: GITHUB_TOKEN es obligatorio para leer el "
                  "cuerpo actual del PR; el cuerpo del payload queda obsoleto tras "
                  "ediciones del cuerpo o re-runs del job", file=sys.stderr)
            return 2
        try:
            pull_data = _make_body_fetcher()(repository, pr_number)
        except GitHubApiError as exc:
            print(f"issue-spec: pr-event: no se pudo leer el cuerpo actual del PR "
                  f"por API ({exc}); el cuerpo del payload queda obsoleto tras "
                  f"ediciones del cuerpo o re-runs del job", file=sys.stderr)
            return 2
        body = pull_data.get("body") or ""
    else:
        body = (body_fetcher(repository, pr_number) or {}).get("body") or ""

    fetcher = issue_fetcher if issue_fetcher is not None else _make_fetcher(repository)
    violations = validate_remote(body, repository, fetcher)
    for violation in violations:
        print(f"issue-spec: {violation}", file=sys.stderr)
    if violations:
        print(f"issue-spec: fallo con {len(violations)} hallazgo(s)", file=sys.stderr)
        return 1
    print("issue-spec: contrato satisfecho")
    return 0


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    local = subparsers.add_parser("local", help="forma del cuerpo; sin payload, token ni red")
    local.add_argument("--body-file", required=True, help="cuerpo del PR; '-' lee stdin")
    local.add_argument("--repository", default="", help="owner/repo del PR, si se conoce")
    local.add_argument(
        "--required-sections", default=None,
        help=("obsoleta (#1103): se acepta y se ignora; retirela del workflow"))
    event = subparsers.add_parser("pr-event", help="payload de evento de PR, con la API")
    event.add_argument("event", help="ruta al payload JSON del evento")
    event.add_argument(
        "--required-sections", default=None,
        help=("obsoleta (#1103): se acepta y se ignora; retirela del workflow"))
    event.add_argument(
        "--host-contract", default=None,
        help=("obsoleta (#1103): se acepta y se ignora; retirela del workflow"))
    return parser


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def _warn_obsolete_flag(flag: str) -> None:
    """Avisa de una flag que #1103 retiro: se acepta, se ignora y hay que quitarla."""
    print(f"issue-spec: flag obsoleta --{flag}: se ignora; retirela del workflow",
          file=sys.stderr)


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = _build_parser().parse_args(argv)
    if args.command == "pr-event":
        if args.required_sections is not None:
            _warn_obsolete_flag("required-sections")
        if args.host_contract is not None:
            _warn_obsolete_flag("host-contract")
    elif args.required_sections is not None:
        _warn_obsolete_flag("required-sections")
    if args.command == "local":
        try:
            body = _read_body(args.body_file)
        except (OSError, UnicodeDecodeError) as exc:
            print(f"issue-spec: no se puede leer el cuerpo: {exc}", file=sys.stderr)
            return 2
        violations = validate_body(body, args.repository)
        success = ("issue-spec local: forma de las referencias correcta. Si la "
                   "referencia es a una PR lo verifica el gate remoto.")
        for violation in violations:
            print(f"issue-spec: {violation}", file=sys.stderr)
        if violations:
            print(f"issue-spec: fallo con {len(violations)} hallazgo(s)", file=sys.stderr)
            return 1
        print(success)
        return 0
    try:
        payload = json.loads(Path(args.event).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        print(f"issue-spec: payload ilegible: {exc}", file=sys.stderr)
        return 2
    return pr_event(payload, body_fetcher=None, issue_fetcher=None)


if __name__ == "__main__":
    sys.exit(main())
