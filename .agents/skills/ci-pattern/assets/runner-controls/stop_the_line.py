#!/usr/bin/env python3
"""stop_the_line.py: parada de línea del merge gobernado (#1627).

Regla, la misma en todo repo (decisión fijada en #1627):
- ROJO = hay al menos una issue ABIERTA con la etiqueta de rojo nocturno
  (`prerelease:night-red`; `--label` la cambia) Y la última ejecución COMPLETADA de la
  nocturna en la rama por defecto no es `success`. Cuenta la ejecución `schedule`. Un
  `workflow_dispatch` cuenta solo si `--dispatch-marker <texto>` está fijado y el
  `display_title` del run (su `run-name`) lo contiene (#1702); sin marcador, ningún
  dispatch cuenta. Una en curso se omite en favor de la última completada.
- Sin señal (ninguna issue etiquetada) el gate PASA. Con issue etiquetada, un workflow
  nocturno inexistente (404) RETIENE (exit 2) nombrando el workflow y `--nightly-workflow`.
- Con rojo, un PR pasa solo si enlaza una de esas issues del MISMO repositorio.
- Todas las lecturas van a la rama por defecto del repo, nunca a la base del PR: un PR de
  cadena hacia una rama de feature se juzga con la noche de main.
- Perfil pre-release (`.github/prerelease.json` con `enabled: true`, en la rama por
  defecto): una lectura fallida o una nocturna sin ejecución completada RETIENE (exit 2).
  Sin perfil, la nocturna ausente o sin ejecuciones completadas PASA.
- Cualquier lectura de issues o de enlaces que falle RETIENE (fallo cerrado).
  Todo mensaje de retención da el comando para relanzar la nocturna.
- Segunda señal, independiente de la nocturna (#1708): mientras haya una issue ABIERTA con
  la etiqueta `prerelease:main-red` (`main_red_label`), un PR pasa solo si enlaza una de
  esas issues del MISMO repositorio. Sin issue abierta, la señal pasa. Ambas señales se
  evalúan; cualquiera puede bloquear. Se activa por defecto: no necesita ningún flag.
- El consumer puede apagar cada señal por separado desde su `host-contract.json` (#1784):
  la clave `stop_the_line` con un miembro booleano por señal (`night_red_enabled`,
  `main_red_enabled`). Una señal apagada no se consulta (no se lee su etiqueta) ni bloquea;
  el merge gobernado lo dice por `stderr`. Sin la clave, o sin un miembro, la señal sigue
  activa: el comportamiento de hoy.

Solo lee: no escribe nada en GitHub.
"""

import base64
import binascii
import json
from collections.abc import Callable
from dataclasses import dataclass
from typing import NoReturn
from urllib.parse import quote

from check_push_compliance import FailClosedError
from github_rest import GitHubApiError, get_json

NIGHT_RED_LABEL = "prerelease:night-red"
MAIN_RED_LABEL = "prerelease:main-red"
PROFILE_PATH = ".github/prerelease.json"
NIGHT_EVENTS = ("schedule", "workflow_dispatch")
RUNS_PER_PAGE = 100
ISSUES_PER_PAGE = 100
HTTP_NOT_FOUND = 404
MIN_SCHEMA_VERSION = 1
SUCCESS = "success"


@dataclass(frozen=True)
class Target:
    """Lo que el gate necesita saber del PR y de la nocturna."""

    repo: str
    pr: int
    nightly_workflow: str
    label: str = NIGHT_RED_LABEL
    dispatch_marker: str | None = None
    main_red_label: str = MAIN_RED_LABEL
    night_red_enabled: bool = True
    main_red_enabled: bool = True


class GitHubReader:
    """Lector de producción: solo GET por la API REST."""

    def default_branch(self, repo):
        return get_json(f"repos/{repo}")["default_branch"]

    def contents(self, repo, path, ref):
        return get_json(f"repos/{repo}/contents/{path}?ref={quote(ref, safe='')}")

    def workflow_runs(self, repo, workflow, branch, event):
        return get_json(
            f"repos/{repo}/actions/workflows/{quote(workflow, safe='')}/runs"
            f"?branch={quote(branch, safe='')}&event={quote(event, safe='')}"
            f"&status=completed&per_page={RUNS_PER_PAGE}"
        )

    def open_issues(self, repo, label):
        return get_json(
            f"repos/{repo}/issues?labels={quote(label, safe='')}"
            f"&state=open&per_page={ISSUES_PER_PAGE}"
        )


def redispatch_command(target, branch):
    """Comando oficial para volver a lanzar la nocturna en la rama por defecto."""
    return f"gh workflow run {target.nightly_workflow} --repo {target.repo} --ref {branch}"


def unblock_hint(target, branch) -> str:
    """Comando de relanzamiento y la condición para que cuente (coherente con el marcador, #1702)."""
    command = f"`{redispatch_command(target, branch)}`"
    if target.dispatch_marker is None:
        return (f"{command} (un dispatch sin --dispatch-marker no cuenta: solo cuenta la "
                f"ejecución programada)")
    return (f"{command} (cuenta solo si su run-name contiene «{target.dispatch_marker}»)")


def _hold(message: str) -> NoReturn:
    """Retiene el merge (exit 2): el mensaje ya incluye el desbloqueo."""
    raise FailClosedError(message)


def _hold_with_unblock(target, branch, reason) -> NoReturn:
    message = (f"no se puede leer el estado de la nocturna ({reason}); el merge falla "
               f"cerrado. Desbloqueo: {unblock_hint(target, branch)}")
    _hold(message)


def _read(call, what, target, branch):
    """Lee con `call`; un fallo de GitHub retiene con el comando de desbloqueo."""
    try:
        return call()
    except GitHubApiError as exc:
        _hold_with_unblock(target, branch, f"{what}: {exc}")


def _decode_profile(document, target, branch) -> bool:
    """Valida lo que el gate necesita del perfil y devuelve `enabled`."""
    try:
        raw = base64.b64decode(document["content"]).decode("utf-8-sig")
        declaration = json.loads(raw)
    except (KeyError, TypeError, ValueError, binascii.Error) as exc:
        _hold_with_unblock(target, branch, f"{PROFILE_PATH} no se puede leer: {exc}")
    if not isinstance(declaration, dict) or not isinstance(declaration.get("enabled"), bool):
        _hold_with_unblock(target, branch, f"{PROFILE_PATH}: `enabled` debe ser booleano")
    version = declaration.get("schema_version", MIN_SCHEMA_VERSION)
    if isinstance(version, bool) or not isinstance(version, int) or version < MIN_SCHEMA_VERSION:
        _hold_with_unblock(target, branch, f"{PROFILE_PATH}: `schema_version` no válido")
    return declaration["enabled"]


def profile_enabled(reader, target, branch) -> bool:
    """True si el perfil está activo en la rama por defecto. Ausente (404) es perfil apagado."""
    try:
        document = reader.contents(target.repo, PROFILE_PATH, branch)
    except GitHubApiError as exc:
        if exc.status == HTTP_NOT_FOUND:
            return False
        _hold_with_unblock(target, branch, f"{PROFILE_PATH}: {exc}")
    return _decode_profile(document, target, branch)


def _counts(target, event, run) -> bool:
    """Un `schedule` cuenta siempre; un `workflow_dispatch` solo con el marcador en su título (#1702)."""
    if event == "schedule":
        return True
    return target.dispatch_marker in str(run.get("display_title") or "")


def _newest_completed(reader, target, branch, event):
    """Fila (fecha, -posición, ejecución) de la más reciente completada que cuenta; None si no hay."""
    if event == "workflow_dispatch" and target.dispatch_marker is None:
        return None
    try:
        data = reader.workflow_runs(target.repo, target.nightly_workflow, branch, event)
    except GitHubApiError as exc:
        if exc.status == HTTP_NOT_FOUND:
            _hold(f"el workflow nocturno «{target.nightly_workflow}» no existe en {target.repo} "
                  f"(404): indique el suyo con --nightly-workflow")
        _hold_with_unblock(target, branch, f"ejecuciones de {target.nightly_workflow}: {exc}")
    runs = data.get("workflow_runs") if isinstance(data, dict) else None
    if not isinstance(runs, list):
        _hold_with_unblock(target, branch, "la respuesta de ejecuciones no es una lista")
    for position, run in enumerate(runs):
        if (isinstance(run, dict) and run.get("event") == event
                and run.get("status") == "completed" and _counts(target, event, run)):
            return str(run.get("created_at") or ""), -position, run
    return None


def latest_completed_night(reader, target, branch):
    """(conclusión, URL) de la ejecución completada más reciente de cualquier evento nocturno.

    Se pregunta por evento: una página sin filtro podría esconder la nocturna tras muchas
    ejecuciones de push. None si no hay ejecución completada o no hay nocturna.
    """
    rows = [row for row in (_newest_completed(reader, target, branch, event)
                            for event in NIGHT_EVENTS) if row is not None]
    if not rows:
        return None
    run = max(rows, key=lambda row: row[:2])[2]
    return str(run.get("conclusion") or "sin conclusión"), run.get("html_url") or "sin URL"


def red_issue_numbers(reader, target, branch) -> list[int]:
    """Números de las issues abiertas con la etiqueta de rojo (sin pull requests)."""
    items = _read(lambda: reader.open_issues(target.repo, target.label),
                  f"issues con la etiqueta {target.label}", target, branch)
    if not isinstance(items, list):
        _hold_with_unblock(target, branch, f"el listado de issues {target.label} no es una lista")
    return sorted(
        item["number"] for item in items
        if isinstance(item, dict) and "pull_request" not in item and isinstance(item.get("number"), int)
    )


def _default_branch(reader, target) -> str:
    branch = _read(lambda: reader.default_branch(target.repo), "la rama por defecto",
                   target, "main")
    if not isinstance(branch, str) or not branch:
        _hold_with_unblock(target, "main", "el repo no informa su rama por defecto")
    return branch


def main_red_issue_numbers(reader, target) -> list[int]:
    """Números de las issues abiertas con la etiqueta de main en rojo (sin pull requests).

    Un fallo de lectura retiene (exit 2); el desbloqueo es relanzar el merge cuando la API responda.
    """
    try:
        items = reader.open_issues(target.repo, target.main_red_label)
    except GitHubApiError as exc:
        reason = f"no se pueden leer las issues {target.main_red_label} ({exc})"
        _hold(f"{reason}; falla cerrado. Vuelva a lanzar el merge cuando GitHub responda")
    if not isinstance(items, list):
        reason = f"el listado de issues {target.main_red_label} no es una lista"
        _hold(f"{reason}; falla cerrado. Vuelva a lanzar el merge cuando GitHub responda")
    return sorted(
        item["number"] for item in items
        if isinstance(item, dict) and "pull_request" not in item and isinstance(item.get("number"), int)
    )


def _night_findings(target, reader, branch, profile, linked_refs) -> list[str]:
    """Hallazgos de la señal nocturna (#1627); lista vacía = pasa."""
    issues = red_issue_numbers(reader, target, branch)
    if not issues:
        return []
    night = latest_completed_night(reader, target, branch)
    if night is None:
        if profile:
            _hold_with_unblock(target, branch, f"no hay ejecución completada de {target.nightly_workflow}")
        return []
    conclusion, url = night
    if conclusion == SUCCESS:
        return []
    repo = target.repo.lower()
    if any(name.lower() == repo and number in issues for name, number in linked_refs()):
        return []
    listing = ", ".join(f"#{number}" for number in issues)
    message = (
        f"parada de línea: la ejecución nocturna de {target.nightly_workflow} en {branch} "
        f"terminó «{conclusion}» ({url}). El PR #{target.pr} no enlaza ninguna issue abierta "
        f"por ese rojo (etiqueta {target.label}: {listing}) del repositorio {target.repo}. "
        f"Arregle el rojo antes de fusionar. Desbloqueo: {unblock_hint(target, branch)}"
    )
    return [message]


def _main_red_findings(target, reader, linked_refs) -> list[str]:
    """Hallazgos de la señal main-red (#1708); lista vacía = pasa. Solo se mira el enlace si hay rojo."""
    numbers = main_red_issue_numbers(reader, target)
    if not numbers:
        return []
    repo = target.repo.lower()
    if any(name.lower() == repo and number in numbers for name, number in linked_refs()):
        return []
    listing = ", ".join(f"#{number}" for number in numbers)
    message = (
        f"parada de línea: main está en rojo tras un merge: hay issues abiertas con la etiqueta "
        f"{target.main_red_label} ({listing}) del repositorio {target.repo}, y el PR #{target.pr} "
        f"no enlaza ninguna. Desbloqueo: arregle main en el PR que enlaza #{numbers[0]}; "
        f"el siguiente run verde de main la cierra."
    )
    return [message]


def evaluate(target: Target, reader, linked_refs: Callable[[], list[tuple[str, int]]]) -> list[str]:
    """Hallazgos del gate para el PR; lista vacía = pasa.

    Dos señales, ambas evaluadas (#1627, #1708): la nocturna y la de main en rojo.
    `linked_refs` devuelve pares (repositorio, número) de las issues que el PR enlaza; se
    llama solo cuando alguna señal está en rojo.

    Una señal apagada por el consumer (#1784) no se consulta ni bloquea: su etiqueta no se
    lee. La rama apagada se salta entera, con la lectura que solo ella necesitaba.
    """
    findings: list[str] = []
    if target.night_red_enabled:
        branch = _default_branch(reader, target)
        profile = profile_enabled(reader, target, branch)
        findings += _night_findings(target, reader, branch, profile, linked_refs)
    if target.main_red_enabled:
        findings += _main_red_findings(target, reader, linked_refs)
    return findings
