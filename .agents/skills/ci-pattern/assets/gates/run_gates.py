#!/usr/bin/env python3
"""Job de gates con un veredicto único, EN SOMBRA (#1179, fase 1).

Ejecuta los gates declarados del consumer una vez y agrega sus veredictos en
un único veredicto: cuántos gates están en verde y, para cada gate que falla,
su nombre y el mensaje que emitió. Es la fase 1 de la propuesta de #1179: el
job publica su check y NADIE lo exige (no cambia required checks ni protección
de rama); su veredicto se compara con el rollup actual para decidir la fase 2.

Los gates se declaran en un fichero versionado en el repo del consumer, una
línea por gate: `nombre: comando` (`nombre` es `[A-Za-z0-9._-]+`; el comando es
el resto de la línea). Los comandos se ejecutan con `sh -c` en el repo.

Reglas (HR-3, liveness):

- Todos los gates en verde → exit 0 y el veredicto cuenta N/N.
- Algún gate en rojo → exit 1, y el veredicto nombra cada gate que falla con
  su salida (la causa).
- Un gate declarado que NO se ejecutó (timeout, crash del intérprete) → exit 2:
  un gate que no midió nunca dice «OK».
- Una declaración rota (fichero ausente, ilegible, vacío, línea sin nombre) →
  exit 2.

Los gates corren en paralelo (el job espera al más lento, no a la suma) y cada
uno tiene su timeout. Exit: 0 todos en verde · 1 algún rojo · 2 declaración
rota o un gate que no se ejecutó.
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess  # noqa: S404 - el comando lo declara el consumer
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

#: Nombre de gate válido: `[A-Za-z0-9._-]+` (como un job id de Actions).
GATE_NAME_RX = re.compile(r"^[A-Za-z0-9._-]+$")
#: Timeout por gate, en segundos (el job entero tiene su propio techo).
DEFAULT_TIMEOUT = 900
#: Líneas de salida que se citan por gate que falla (la causa, no el ruido).
CAUSE_LINES = 20
#: Los mensajes de error viven como constantes de módulo (TRY003).
ERR_DECLARATION = "declaración de gates rota: {detail}"
ERR_EMPTY = "la lista de gates está vacía: sin gates declarados el job no mide nada"
ERR_NO_NAME = "línea {line}: el gate necesita un nombre antes de ':'"
ERR_BAD_NAME = "línea {line}: '{name}' no es un nombre de gate válido"
ERR_NO_COMMAND = "línea {line}: el gate '{name}' no declara comando"
ERR_TIMEOUT = "el gate '{name}' no terminó en {seconds} s: no midió"

#: Motivos por los que un gate no se ejecutó (no es un rojo: es una ausencia).
NOT_MEASURED = ("timeout", "crash")


class GateDeclarationError(ValueError):
    """La declaración de gates no sirve: el job sale 2 (HR-3)."""


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: the output must not depend on the locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def parse_declaration(path: str) -> list:
    """Lee el fichero de gates: lista de (nombre, comando). Cualquier forma
    rota sale 2."""
    try:
        raw = Path(path).read_text(encoding="utf-8")
    except OSError as exc:
        raise GateDeclarationError(ERR_DECLARATION.format(detail=exc)) from exc
    gates = []
    for number, line in enumerate(raw.splitlines(), 1):
        entry = line.strip()
        if not entry or entry.startswith("#"):
            continue
        name, sep, command = entry.partition(":")
        name = name.strip()
        if not sep:
            raise GateDeclarationError(ERR_NO_NAME.format(line=number))
        if not GATE_NAME_RX.match(name):
            raise GateDeclarationError(ERR_BAD_NAME.format(line=number, name=name))
        command = command.strip()
        if not command:
            raise GateDeclarationError(ERR_NO_COMMAND.format(line=number, name=name))
        gates.append((name, command))
    if not gates:
        raise GateDeclarationError(ERR_EMPTY)
    return gates


def run_gate(name: str, command: str, timeout: int, cwd: str) -> dict:
    """Ejecuta un gate y devuelve su veredicto.

    `timeout` o un crash del intérprete NO son un rojo: son «no midió», y el
    job sale 2 (HR-3)."""
    try:
        done = subprocess.run(  # noqa: S603 - el comando lo declara el consumer
            ["sh", "-c", command],  # noqa: S607 - shell resuelto por el PATH
            cwd=cwd, capture_output=True, text=True,
            timeout=timeout, check=False)
    except subprocess.TimeoutExpired:
        return {"name": name, "command": command, "result": "timeout", "code": None,
                "output": ERR_TIMEOUT.format(name=name, seconds=timeout)}
    except OSError as exc:
        return {"name": name, "command": command, "result": "crash", "code": None,
                "output": str(exc)}
    return {"name": name, "command": command, "result": "passed" if done.returncode == 0
            else "failed", "code": done.returncode,
            "output": (done.stdout or "") + (done.stderr or "")}


def run_gates(gates: list, timeout: int, cwd: str) -> list:
    """Ejecuta los gates en paralelo; devuelve sus veredictos en orden."""
    with ThreadPoolExecutor(max_workers=len(gates)) as pool:
        futures = [pool.submit(run_gate, name, command, timeout, cwd)
                   for name, command in gates]
        return [future.result() for future in futures]


def verdict(results: list) -> dict:
    """El veredicto único: cuántos gates en verde y los que no."""
    passed = [r for r in results if r["result"] == "passed"]
    failed = [r for r in results if r["result"] == "failed"]
    not_measured = [r for r in results if r["result"] in NOT_MEASURED]
    return {"gates": len(results), "passed": len(passed), "failed": len(failed),
            "not_measured": len(not_measured),
            "failing": [{"name": r["name"], "code": r["code"],
                         "cause": cause_of(r["output"])} for r in failed],
            "not_measured_gates": [{"name": r["name"], "cause": cause_of(r["output"])}
                                   for r in not_measured]}


def cause_of(output: str) -> str:
    """La causa que se cita: las últimas líneas con texto del gate que falló."""
    lines = [line for line in (output or "").splitlines() if line.strip()]
    return "\n".join(lines[-CAUSE_LINES:])


def render(verdict_doc: dict) -> str:
    """El veredicto, legible: una línea por gate que no está en verde."""
    total = verdict_doc["gates"]
    passed = verdict_doc["passed"]
    lines = [f"gates: {passed}/{total} en verde"]
    for failing in verdict_doc["failing"]:
        lines.append(f"  ✗ {failing['name']} (exit {failing['code']}):")
        lines.append(f"    {failing['cause']}")
    for absent in verdict_doc["not_measured_gates"]:
        lines.append(f"  ! {absent['name']} no midió: {absent['cause']}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(
        description="Job de gates con veredicto único, en sombra (#1179, fase 1).")
    parser.add_argument("--gates", required=True, help="fichero con `nombre: comando`")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT,
                        help="segundos por gate (default: %(default)s)")
    parser.add_argument("--cwd", default=".", help="raíz del repo donde corren")
    parser.add_argument("--json", action="store_true", help="salida máquina")
    args = parser.parse_args(argv)
    try:
        gates = parse_declaration(args.gates)
    except GateDeclarationError as exc:
        print(f"GATES ERROR: {exc}", file=sys.stderr)
        return 2
    results = run_gates(gates, args.timeout, args.cwd)
    doc = verdict(results)
    if args.json:
        print(json.dumps(doc, indent=2, sort_keys=True))
    else:
        print(render(doc))
    if doc["not_measured"]:
        return 2
    return 1 if doc["failed"] else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
