#!/usr/bin/env python3
# ci-pattern asset — el mapa de orígenes que el presupuesto A1 necesita (#1159)
"""A1 solo excluye el árbol vendorizado si el consumer pasa el mapa de orígenes.

El gate de contrato de PR mide el presupuesto con la regla compartida
(`assets/budget/vendored_budget.py`): sin `--origins-file` cuenta TODO, incluido
el árbol vendorizado (`.agents/skills/`) que el propagador materializa byte a
byte, así que una PR de propagación falla con un A1 que no es trabajo de quien
abre la PR (#1133, #1149). Este módulo detecta esa deriva para la fase 4 de la
adopción (G4.2): lee el árbol del consumer y dice qué cablear. No ejecuta `git`,
no lee la red y no invoca el gate. El mapa se exige cuando algún workflow invoca
`check_pr_contract.py` y la política que ese gate recibe con `--policy-file`
tiene `budget` activa; `pr-size` mide A1 por ruta y no lo necesita.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

#: El gate que mide A1 con el mapa (los gates `pr-size` no lo necesitan), el
#: asset del canon que genera el mapa y las banderas con las que se pasan.
GATE_SCRIPT = "check_pr_contract.py"
GENERATOR_SCRIPT = "origins_file.py"
ORIGINS_FLAG = "--origins-file"
POLICY_FLAG = "--policy-file"
OUTPUT_FLAG = "--output"
#: La política por defecto del gate (la de la plantilla canónica).
DEFAULT_POLICY = ".github/pr-contract.policy.json"
#: Dónde viven los workflows del consumer.
WORKFLOWS_DIR = Path(".github") / "workflows"
#: Las extensiones de workflow que GitHub ejecuta.
WORKFLOW_SUFFIXES = (".yml", ".yaml")
#: La plantilla canónica del workflow de contrato de PR.
CANONICAL_TEMPLATE = "assets/templates/.github/workflows/pr-contract.yml.tmpl"
#: La clave de un `run:` de workflow, con su sangría y su resto.
RUN_KEY_RX = re.compile(r"^(?P<indent>[ \t]*)(?:-[ \t]+)?run:(?P<rest>.*)$")
#: Los marcadores de bloque escalar de YAML (`run: |` y sus variantes).
BLOCK_MARKERS = ("|", ">", "|-", ">-", "|+", ">+")
#: El prefijo de un comentario de YAML o de shell.
COMMENT_PREFIX = "#"


def budget_active(policy) -> bool:
    """True si la política del gate de contrato de PR enciende `budget`.

    Sin `checks` están todas las comprobaciones activas; con `checks`, la
    comprobación ausente está activa y `false` la apaga. Un valor que no es
    booleano no se puede interpretar: el gate falla cerrado al correr y aquí se
    exige el mapa (fallo hacia la exigencia, HR-3).
    """
    if not isinstance(policy, dict):
        return False
    checks = policy.get("checks")
    if not isinstance(checks, dict):
        return True
    return checks.get("budget", True) is not False


def _logical_commands(text: str) -> list[str]:
    """Los comandos del workflow: los bloques `run:` con sus continuaciones unidas.

    Un comando es la línea de su `run:` (o del bloque escalar que abre) con las
    líneas de continuación (`\\`) unidas: así la bandera y su valor se leen aunque
    el valor vaya en la línea siguiente. Los comentarios no son comandos.
    """
    lines = text.splitlines()
    raw: list[str] = []
    index = 0
    while index < len(lines):
        match = RUN_KEY_RX.match(lines[index])
        index += 1
        if match is None:
            continue
        rest = match.group("rest").strip()
        if rest not in BLOCK_MARKERS:
            raw.append(rest)
            continue
        # Bloque escalar: las líneas más sangradas que la clave `run:`.
        indent = len(match.group("indent"))
        while index < len(lines):
            line = lines[index]
            if line.strip() and (len(line) - len(line.lstrip())) <= indent:
                break
            raw.append(line)
            index += 1
    commands: list[str] = []
    buffer: list[str] = []
    for line in raw:
        stripped = line.strip()
        if not stripped or stripped.startswith(COMMENT_PREFIX):
            continue
        if stripped.endswith("\\"):
            buffer.append(stripped[:-1].rstrip())
            continue
        buffer.append(stripped)
        commands.append(" ".join(buffer))
        buffer = []
    if buffer:
        commands.append(" ".join(buffer))
    return commands


def _flag_value(command: str, flag: str) -> str | None:
    """El valor de una bandera en un comando; None si no aparece."""
    tokens = command.split()
    for position, token in enumerate(tokens):
        if token.startswith(flag + "="):
            return token[len(flag) + 1:].strip("\"'")
        if token == flag and position + 1 < len(tokens):
            return tokens[position + 1].strip("\"'")
    return None


def _missing_origins(rel: str, policy_rel: str) -> str:
    """El hallazgo del workflow que no pasa el mapa al gate."""
    return (
        f"{rel}: el gate '{GATE_SCRIPT}' mide el presupuesto A1 (la política '{policy_rel}' "
        f"enciende la comprobación 'budget') y el workflow no le pasa '{ORIGINS_FLAG}': sin el "
        "mapa de orígenes el árbol vendorizado ('.agents/skills/') cuenta como líneas de "
        "autoría y una PR de propagación del catálogo revienta el presupuesto con un falso "
        "'el diff supera el presupuesto'. Cablee el mapa: genérelo con el asset del canon "
        "('assets/bin/origins_file.py --base <base> --head <head> --names-file <nombres> "
        f"--log-file <log> --output <mapa>') y pásaselo al gate con '{ORIGINS_FLAG} <mapa>'; "
        f"la plantilla canónica es '{CANONICAL_TEMPLATE}'."
    )


def _ungenerated_origins(rel: str, origins: str) -> str:
    """El hallazgo del gate que recibe un mapa que nadie genera."""
    return (
        f"{rel}: el gate '{GATE_SCRIPT}' recibe el mapa de orígenes con "
        f"'{ORIGINS_FLAG} {origins}' pero ningún paso del workflow lo genera: el gate lo "
        "leerá vacío y fallará cerrado al medir el presupuesto. Añada el paso que lo escribe "
        "con el asset del canon ('assets/bin/origins_file.py --base <base> --head <head> "
        f"--names-file <nombres> --log-file <log> --output {origins}'); la plantilla "
        f"canónica es '{CANONICAL_TEMPLATE}'."
    )


def _unreadable_policy(rel: str, policy_rel: str) -> str:
    """El hallazgo de la política que el gate usa y no se puede leer."""
    return (
        f"{rel}: el gate '{GATE_SCRIPT}' usa la política '{policy_rel}' y esta es ilegible: "
        "sin ella no hay medición de si el presupuesto A1 está activo, así que no se puede "
        "exigir el mapa de orígenes (HR-3)."
    )


def _policy_path(repo: Path, policy_rel: str) -> Path:
    """Where the gate's policy lives on this tree (DysTelefonica/team-skills#1313).

    In CI the canonical template checks the base out into `base/`, which is the
    trusted copy. Outside CI there is no `base/`, and the consumer's committed copy of
    the same file is what the gate would read, so that copy is used. With neither
    present the read fails and the caller reports it unreadable (fail closed).
    """
    if policy_rel.startswith("base/") and not (repo / "base").is_dir():
        return repo / policy_rel[len("base/"):]
    return repo / policy_rel


def _gate_findings(repo: Path, rel: str, gate: str, commands: list[str]) -> list[str]:
    """Los hallazgos de una invocación del gate de contrato de PR."""
    policy_rel = _flag_value(gate, POLICY_FLAG) or DEFAULT_POLICY
    try:
        policy = json.loads(_policy_path(repo, policy_rel).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return [_unreadable_policy(rel, policy_rel)]
    if not budget_active(policy):
        return []
    origins = _flag_value(gate, ORIGINS_FLAG)
    if origins is None:
        return [_missing_origins(rel, policy_rel)]
    # El mapa tiene que existir: lo genera el asset del canon (`--output`).
    generated = {_flag_value(c, OUTPUT_FLAG) for c in commands
                 if GENERATOR_SCRIPT in c}
    if origins not in generated:
        return [_ungenerated_origins(rel, origins)]
    return []


def findings(repo: Path) -> list[str]:
    """Los hallazgos del cableado de A1 sobre el árbol del consumer."""
    directory = repo / WORKFLOWS_DIR
    if not directory.is_dir():
        return []
    out: list[str] = []
    for workflow in sorted(directory.iterdir()):
        if not workflow.is_file() or workflow.suffix not in WORKFLOW_SUFFIXES:
            continue
        try:
            text = workflow.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        commands = _logical_commands(text)
        gates = [command for command in commands if GATE_SCRIPT in command]
        if not gates:
            continue
        rel = workflow.relative_to(repo).as_posix()
        for gate in gates:
            out.extend(_gate_findings(repo, rel, gate, commands))
    return out
