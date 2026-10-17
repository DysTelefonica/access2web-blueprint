# architecture_imports.py — reglas de imports del documento de arquitectura (HR-68, #714)
"""Adaptador Python: el bloque ``architecture`` es la única fuente de las reglas.

De las capas y de ``allowed`` se genera la configuración de ``import-linter``
(un contrato ``forbidden`` por cada par de capas no permitido). El gate falla si
la configuración versionada no coincide byte a byte con la generada (el documento
no puede mentir), si el linter no está instalado o no se ejecutó (HR-3) y si un
import viola una regla, nombrando ``fichero:línea`` y el contrato incumplido.

Stdlib exclusivamente; el linter se ejecuta como subproceso.
"""

from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path

LINTER = "lint-imports"
_CONTRACT_RE = re.compile(r"^-{3,}$")
_VIOLATION_RE = re.compile(r"^-\s+(\S+) -> \S+ \(l\.(\d+)\)")


def dotted(module_path: str) -> str:
    """``app/routes`` -> ``app.routes``; ``app/main.py`` -> ``app.main``."""
    path = Path(module_path)
    return ".".join(path.with_suffix("").parts if path.suffix else path.parts)


def forbidden_pairs(data: dict) -> list[tuple[str, str]]:
    """Pares (origen, destino) de capas distintas cuya dependencia no se permite."""
    names = [layer["name"] for layer in data["layers"]]
    return [(src, dst) for src in names for dst in names
            if src != dst and dst not in data["allowed"][src]]


def render_importlinter(data: dict) -> str:
    """Configuración de import-linter derivada del bloque; determinista."""
    modules = {layer["name"]: [dotted(m) for m in layer["modules"]] for layer in data["layers"]}
    roots = sorted({m.split(".")[0] for mods in modules.values() for m in mods})
    out = ["# Generado desde el bloque `architecture` del documento (HR-68); no editar a mano.",
           "[importlinter]", "root_packages =", *[f"    {r}" for r in roots]]
    for src, dst in forbidden_pairs(data):
        out += ["", f"[importlinter:contract:{src}-must-not-import-{dst}]",
                f"name = {src} must not import {dst}", "type = forbidden", "source_modules =",
                *[f"    {m}" for m in modules[src]], "forbidden_modules =",
                *[f"    {m}" for m in modules[dst]]]
    return "\n".join(out) + "\n"


def _module_file(repo: Path, module: str) -> str:
    base = Path(*module.split("."))
    for candidate in (base.with_suffix(".py"), base / "__init__.py"):
        if (repo / candidate).is_file():
            return candidate.as_posix()
    return base.as_posix()


def parse_violations(repo: Path, output: str) -> list[str]:
    """``fichero:línea: contrato`` de cada import prohibido que imprime import-linter."""
    findings, contract, previous = [], "", ""
    for line in output.splitlines():
        if _CONTRACT_RE.match(line) and previous.strip() and not previous.startswith(("-", "Broken")):
            contract = previous.strip()
        match = _VIOLATION_RE.match(line)
        if match:
            findings.append(f"{_module_file(repo, match.group(1))}:{match.group(2)}: "
                            f"import prohibido por la regla «{contract}»")
        previous = line
    return findings


def check_imports(repo: Path, data: dict, config: str, command: str = LINTER) -> list[str]:
    """Hallazgos de la configuración y de la ejecución del linter; [] = verde."""
    path = repo / config
    expected = render_importlinter(data)
    if not path.is_file():
        return [f"falta la configuración del linter {config}; genérela con --write-config"]
    if path.read_text(encoding="utf-8").replace("\r\n", "\n") != expected:
        return [f"la configuración {config} no coincide con las capas y reglas del documento; "
                "regenérela con --write-config (el documento es la única fuente)"]
    if shutil.which(command) is None:
        return [f"el linter {command} no está instalado: sin ejecución no hay medición (HR-3)"]
    try:
        proc = subprocess.run([command, "--config", config], cwd=repo, capture_output=True,
                              text=True, check=False, encoding="utf-8")
    except OSError as exc:
        return [f"el linter {command} no se pudo ejecutar: {exc}"]
    if proc.returncode == 0:
        return []
    found = parse_violations(repo, proc.stdout)
    return found or [f"el linter {command} salió con código {proc.returncode} sin violaciones "
                     f"legibles: {proc.stderr.strip() or proc.stdout.strip()[-300:]}"]
