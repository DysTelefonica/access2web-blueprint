#!/usr/bin/env python3
# ci-pattern asset — gate issue-spec (DysTelefonica/team-skills#633)
"""Gate de trazabilidad: el cuerpo del PR enlaza una issue aprobada y tipada.

Reglas (cada una es una función; la orquestación vive en validate_body y
validate_remote):

1. El cuerpo lleva `Closes|Fixes|Resolves|Refs #N` fuera de código.
2. La referencia resuelve a una issue, nunca a una pull request: la PR de la
   que se depende va en `depends-on` de Chain Context (HR-54).
3. La issue lleva `status:approved` y exactamente un `type:*`.
4. El cuerpo no cierra y referencia la MISMA issue a la vez (HR-54).
5. La referencia apunta al repositorio del PR (HR-54).
6. Un fallo de la API no es un verde: fail-closed (HR-3).
7. El cuerpo (o el cuerpo de la issue citada) lleva las secciones declaradas
   en la configuración del consumer (DysTelefonica/team-skills#644). Opt-in:
   sin `--required-sections` la regla 7 no se aplica. Las secciones viven
   como títulos Markdown (H1..H6) y se ignoran dentro de un span de código o
   de un bloque cercado, igual que las closing keywords de la regla 1.

Qué cuenta como referencia sigue a GitHub: las palabras de cierre valen en
prosa, pero no dentro de un span de código ni de un bloque cercado. Un bloque
con sangría de 4 espacios no se quita: dentro de una lista es continuación del
ítem, y ocultarlo podría esconder una referencia real. Se cuenta, así que el
fallo es ruidoso y no silencioso. El mismo criterio aplica a la regla 7.

Modos:

- `local --body-file <ruta|-> [--repository owner/repo] [--required-sections <path>]`:
  forma del cuerpo. Sin payload, sin token y sin red. No verifica la issue
  (reglas 2 y 3). Con `--required-sections`, la regla 7 se aplica al cuerpo
  leído.
- `pr-event <event.json> [--required-sections <path>]`: todas las reglas. Usa
  GITHUB_TOKEN y GITHUB_API_URL del entorno si existen. Con
  `--required-sections`, la regla 7 se aplica al cuerpo de cada issue
  consultada.

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

APPROVAL_LABEL = "status:approved"
TYPE_PREFIX = "type:"
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
# ATX heading: 1–6 almohadillas, espacio, texto. La barra final de cierre de
# Markdown y el espacio en blanco alrededor no cuentan para la comparación.
HEADING_RX = re.compile(r"^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$")


class GitHubApiError(RuntimeError):
    """Raised when GitHub cannot provide authoritative issue data."""


class IssueSpecConfigError(ValueError):
    """Raised when the required-sections config is unreadable or invalid (regla 7)."""


class ConfigUnreadableError(IssueSpecConfigError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"no se puede leer {path}: {exc}")


class ConfigNotJsonError(IssueSpecConfigError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"{path} no es JSON válido: {exc}")


class ConfigNotObjectError(IssueSpecConfigError):
    def __init__(self, path: str) -> None:
        super().__init__(f"{path} debe ser un objeto JSON con la clave 'required_sections'")


class ConfigSectionsNotListError(IssueSpecConfigError):
    def __init__(self, path: str) -> None:
        super().__init__(f"{path}: 'required_sections' debe ser una lista de strings no vacíos")


class ConfigSectionInvalidError(IssueSpecConfigError):
    def __init__(self, path: str, item: object) -> None:
        super().__init__(f"{path}: cada sección debe ser un string no vacío (recibido: {item!r})")


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


# --------------------------------------------------------------------------------------------
# Regla 7 — secciones obligatorias (DysTelefonica/team-skills#644). Opt-in: la
# lista la declara el consumer; sin configuración, validate_body y
# validate_remote no la aplican.
# --------------------------------------------------------------------------------------------


def extract_sections(body: str) -> dict[str, bool]:
    """Heading Markdown → True si la sección tiene contenido no vacío bajo el heading.

    Recorre el cuerpo fuera de código (vía ``strip_code``, mismo criterio
    que ``references``: dentro de un span de código o de un bloque cercado
    ni el heading ni el contenido cuentan). Para cada
    heading encontrado registra un flag de contenido que pasa a True en
    cuanto aparece una línea no vacía entre este heading y el siguiente (o
    el final del cuerpo). El caso «heading presente, contenido ausente» es
    lo que la regla 7 llama «sección vacía».
    """
    cleaned = strip_code(body or "")
    sections: dict[str, bool] = {}
    current: str | None = None
    has_content = False
    for line in cleaned.splitlines():
        match = HEADING_RX.match(line)
        if match:
            if current is not None:
                sections[current] = has_content
            current = match.group(2).strip()
            has_content = False
        elif current is not None and line.strip():
            has_content = True
    if current is not None:
        sections[current] = has_content
    return sections


def load_required_sections(path: str) -> tuple[str, ...]:
    """Lee el fichero de configuración de secciones obligatorias. Falla cerrado.

    Esquema:

        {"required_sections": ["Sección A", "Sección B", ...]}

    La clave ``required_sections`` es obligatoria; su valor debe ser una lista
    de strings no vacíos. Cualquier desviación del esquema (fichero ilegible,
    JSON inválido, objeto ausente, clave ausente, valor no lista, elemento no
    string o string vacío) lanza ``IssueSpecConfigError`` y devuelve ``exit 2``
    desde ``main``. El criterio fail-closed es el mismo que la regla 6 (HR-3).
    """
    try:
        text = Path(path).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise ConfigUnreadableError(path, exc) from exc
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ConfigNotJsonError(path, exc) from exc
    if not isinstance(data, dict):
        raise ConfigNotObjectError(path)
    sections = data.get("required_sections")
    if not isinstance(sections, list):
        raise ConfigSectionsNotListError(path)
    for item in sections:
        if not isinstance(item, str) or not item.strip():
            raise ConfigSectionInvalidError(path, item)
    return tuple(sections)


def rule_required_sections(
    body: str,
    required_sections: tuple[str, ...],
    prefix: str = "",
) -> list[str]:
    """Regla 7: cada sección declarada debe aparecer como heading con contenido.

    Una sección se considera satisfecha cuando su heading está presente fuera
    de código y hay al menos una línea no vacía entre este heading y el
    siguiente (o el final del cuerpo). Si el heading falta, el hallazgo dice
    «falta la sección obligatoria»; si el heading está pero sin contenido,
    el hallazgo dice «sección obligatoria … está vacía». ``prefix`` antepone
    ``#<n>: `` en modo ``validate_remote``. Sin ``required_sections`` la regla
    es no-op (backward compat).
    """
    if not required_sections:
        return []
    sections = extract_sections(body)
    violations: list[str] = []
    for name in required_sections:
        if name not in sections:
            violations.append(f"{prefix}falta la sección obligatoria «{name}»")
        elif not sections[name]:
            violations.append(f"{prefix}sección obligatoria «{name}» está vacía")
    return violations


def rule_closes_and_refs_same_issue(refs: tuple[Reference, ...]) -> list[str]:
    """Regla 4: Closes y Refs del mismo número son una contradicción."""
    closed = {r.number for r in refs if r.kind == "close"}
    referred = {r.number for r in refs if r.kind == "refs"}
    return [
        f"#{number}: el cuerpo cierra y referencia la MISMA issue a la vez; use "
        "Closes o Refs, nunca los dos"
        for number in sorted(closed & referred)
    ]


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


def validate_body(
    body: str,
    repository: str = "",
    required_sections: tuple[str, ...] = (),
) -> list[str]:
    """Modo local: forma del cuerpo, sin tocar la red.

    La regla 7 (secciones obligatorias) es opt-in: solo se aplica si
    ``required_sections`` no está vacío.
    """
    refs = references(body)
    violation = rule_has_reference(refs)
    if violation:
        return [violation]
    return (
        rule_closes_and_refs_same_issue(refs)
        + rule_same_repository(refs, repository)
        + rule_required_sections(body, required_sections)
    )


# --------------------------------------------------------------------------------------------
# Reglas de la issue remota: las que necesitan la API.
# --------------------------------------------------------------------------------------------


def rule_reference_is_issue(issue: Mapping[str, Any], number: int) -> str | None:
    """Regla 2: `GET /issues/{n}` devuelve `pull_request` cuando el número es una PR."""
    if "pull_request" not in issue:
        return None
    return (f"#{number}: la referencia resuelve a una pull request, no a una issue; "
            f"muévala a depends-on de Chain Context. {DEPENDS_ON_HINT}")


def rule_missing_labels(issue: Mapping[str, Any], number: int) -> list[str]:
    """Regla 3: la issue lleva `status:approved` y un `type:*`."""
    names = {
        label["name"] if isinstance(label, Mapping) else str(label)
        for label in issue.get("labels", [])
    }
    violations: list[str] = []
    if APPROVAL_LABEL not in names:
        violations.append(f"#{number}: requiere {APPROVAL_LABEL}")
    if not any(name.startswith(TYPE_PREFIX) for name in names):
        violations.append(f"#{number}: requiere una etiqueta {TYPE_PREFIX}*")
    return violations


def rule_api_error(number: int, exc: GitHubApiError) -> str:
    """Regla 6: un error de la API falla cerrado."""
    return f"#{number}: cannot read the issue ({exc})"


def validate_remote(
    body: str,
    repository: str,
    fetch_issue: Callable[[int], Mapping[str, Any]],
    required_sections: tuple[str, ...] = (),
) -> list[str]:
    """Modo pr-event: todas las reglas. Solo consulta las issues del propio repositorio.

    La regla 7 se aplica al cuerpo de cada issue consultada: si una referencia
    está dentro de un span de código o de un bloque cercado, la issue no se
    consulta y la regla tampoco se le aplica. Las issues de otro repositorio
    no se consultan (mismo criterio que la regla 5). El cuerpo del PR
    (parámetro ``body``) no se evalúa contra la regla 7: el PR solo lleva
    la referencia ``Closes #N``; las secciones obligatorias viven en el
    cuerpo de la issue.
    """
    refs = references(body)
    # Regla 7 no se aplica al cuerpo del PR: validate_body se llama con la
    # lista vacía para que las reglas 1–5 sí corran y la 7 no.
    violations = validate_body(body, repository, ())
    if not refs:
        return violations
    numbers = sorted({
        ref.number for ref in refs
        if not ref.repo or ref.repo.lower() == repository.lower()
    })
    for number in numbers:
        try:
            issue = fetch_issue(number)
        except GitHubApiError as exc:
            violations.append(rule_api_error(number, exc))
            continue
        if message := rule_reference_is_issue(issue, number):
            violations.append(message)
        else:
            violations.extend(rule_missing_labels(issue, number))
            violations.extend(rule_required_sections(
                issue.get("body", ""), required_sections, prefix=f"#{number}: "))
    return violations


# --------------------------------------------------------------------------------------------
# GitHub client: solo lectura. `issues: read` basta.
# --------------------------------------------------------------------------------------------


class GitHubIssueFailedError(GitHubApiError):
    def __init__(self, repository: str, number: int, exc: BaseException) -> None:
        super().__init__(f"GET /repos/{repository}/issues/{number} failed: {exc}")


class UnsupportedApiUrlError(GitHubApiError):
    def __init__(self, api_url: str) -> None:
        super().__init__(f"api url must be https (loopback http only): {api_url}")


def _open(api_url: str, path: str, headers: Mapping[str, str]):
    """Abre el GET con el esquema validado: https, o http solo en loopback.

    Cada rama lleva su literal dentro de la llamada: ruff S310 solo acepta
    una URL con esquema literal.
    """
    parts = urllib.parse.urlsplit(api_url)
    tail = parts.path.rstrip("/") + path
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

    def issue(self, repository: str, number: int) -> Mapping[str, Any]:
        headers = {
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "ci-pattern-issue-spec",
        }
        if self._token:
            headers["Authorization"] = f"Bearer {self._token}"
        try:
            with _open(self._api_url, f"/repos/{repository}/issues/{number}", headers) as response:
                return json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubIssueFailedError(repository, number, exc) from exc


def _make_fetcher(repository: str) -> Callable[[int], Mapping[str, Any]]:
    """Fetcher desde `$GITHUB_TOKEN` y `$GITHUB_API_URL`. Los tests lo sustituyen."""
    client = GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )
    return lambda number: client.issue(repository, number)


def _read_body(path: str) -> str:
    if path == "-":
        return sys.stdin.read()
    return Path(path).read_text(encoding="utf-8")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    local = subparsers.add_parser("local", help="forma del cuerpo; sin payload, token ni red")
    local.add_argument("--body-file", required=True, help="cuerpo del PR; '-' lee stdin")
    local.add_argument("--repository", default="", help="owner/repo del PR, si se conoce")
    local.add_argument(
        "--required-sections", default="",
        help=("ruta al JSON con la lista de secciones obligatorias del consumer "
              "(DysTelefonica/team-skills#644); si se omite, la regla 7 no se aplica"))
    event = subparsers.add_parser("pr-event", help="payload de evento de PR, con la API")
    event.add_argument("event", help="ruta al payload JSON del evento")
    event.add_argument(
        "--required-sections", default="",
        help=("ruta al JSON con la lista de secciones obligatorias del consumer "
              "(DysTelefonica/team-skills#644); si se omite, la regla 7 no se aplica"))
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    required: tuple[str, ...] = ()
    if args.required_sections:
        try:
            required = load_required_sections(args.required_sections)
        except IssueSpecConfigError as exc:
            print(f"issue-spec: {exc}", file=sys.stderr)
            return 2
    if args.command == "local":
        try:
            body = _read_body(args.body_file)
        except (OSError, UnicodeDecodeError) as exc:
            print(f"issue-spec: no se puede leer el cuerpo: {exc}", file=sys.stderr)
            return 2
        violations = validate_body(body, args.repository, required)
        success = ("issue-spec local: forma de las referencias correcta. La aprobación, "
                   "el tipo y si la referencia es una PR los verifica el gate remoto.")
    else:
        try:
            payload = json.loads(Path(args.event).read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
            print(f"issue-spec: payload ilegible: {exc}", file=sys.stderr)
            return 2
        repository = (payload.get("repository") or {}).get("full_name", "")
        if "/" not in repository:
            print("issue-spec: el evento no identifica repository.full_name", file=sys.stderr)
            return 2
        body = (payload.get("pull_request") or {}).get("body") or ""
        violations = validate_remote(body, repository, _make_fetcher(repository), required)
        success = "issue-spec: contrato satisfecho"
    for violation in violations:
        print(f"issue-spec: {violation}", file=sys.stderr)
    if violations:
        print(f"issue-spec: fallo con {len(violations)} hallazgo(s)", file=sys.stderr)
        return 1
    print(success)
    return 0


if __name__ == "__main__":
    sys.exit(main())
