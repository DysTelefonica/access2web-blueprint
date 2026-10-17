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

import ast
import os
import re
import shutil
import subprocess
from pathlib import Path

LINTER = "lint-imports"
_CONTRACT_RE = re.compile(r"^-{3,}$")
_VIOLATION_RE = re.compile(r"^-\s+(\S+) -> \S+ \(l\.(\d+)\)")
_BULLET_RE = re.compile(r"^-\s+\S")
# import-linter parte su salida al ancho de la terminal (80 columnas sin COLUMNS):
# se ejecuta con un ancho que no parte y, además, el análisis tolera líneas partidas.
WIDE_COLUMNS = "10000"
MAX_CONTINUATION = 3


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
    tolerated = tolerated_imports(data)
    for src, dst in forbidden_pairs(data):
        out += ["", f"[importlinter:contract:{src}-must-not-import-{dst}]",
                f"name = {src} must not import {dst}", "type = forbidden", "source_modules =",
                *[f"    {m}" for m in modules[src]], "forbidden_modules =",
                *[f"    {m}" for m in modules[dst]]]
        if tolerated.get((src, dst)):
            out += ["ignore_imports =", *[f"    {i}" for i in tolerated[(src, dst)]]]
    for a, b in independent_pairs(data):
        slug = f"{a}-must-not-import-{b}".replace(".", "-")
        out += ["", f"[importlinter:contract:independent-{slug}]",
                f"name = {a} must not import {b}", "type = forbidden", "source_modules =",
                f"    {a}", "forbidden_modules =", f"    {b}", "allow_indirect_imports = True"]
    return "\n".join(out) + "\n"


def independent_pairs(data: dict) -> list[tuple[str, str]]:
    """Pares ordenados (a, b) de módulos de un mismo grupo ``independent``: ninguno importa a otro."""
    pairs = []
    for group in data.get("independent", []):
        mods = [dotted(m) for m in group["modules"]]
        pairs += [(a, b) for a in mods for b in mods if a != b]
    return pairs


def tolerated_imports(data: dict) -> dict[tuple[str, str], list[str]]:
    """Importaciones heredadas permitidas por par de capas: ``origen -> destino`` con puntos."""
    result: dict[tuple[str, str], list[str]] = {}
    for item in data.get("tolerated", []):
        edges = [" -> ".join(dotted(part.strip()) for part in edge.split("->")) for edge in item["imports"]]
        result.setdefault((item["from"], item["to"]), []).extend(edges)
    return result


def pure_findings(repo: Path, data: dict) -> list[str]:
    """Las capas puras no importan los paquetes externos prohibidos (recorrido AST, stdlib)."""
    pure = data.get("pure")
    if not pure:
        return []
    forbidden = set(pure["forbidden_externals"])
    wanted = set(pure["layers"])
    findings = []
    for layer in data["layers"]:
        if layer["name"] not in wanted:
            continue
        for module in layer["modules"]:
            base = repo / module
            files = [base.with_suffix(".py")] if base.with_suffix(".py").is_file() else sorted(base.rglob("*.py"))
            for path in files:
                try:
                    tree = ast.parse(path.read_text(encoding="utf-8"))
                except (OSError, SyntaxError, UnicodeDecodeError) as exc:
                    findings.append(f"{path.relative_to(repo).as_posix()}: no se pudo analizar ({exc.__class__.__name__}); "
                                    "sin análisis no hay medición (HR-3)")
                    continue
                for node in ast.walk(tree):
                    names = ([a.name for a in node.names] if isinstance(node, ast.Import)
                             else [node.module or ""] if isinstance(node, ast.ImportFrom) and node.level == 0 else [])
                    for name in names:
                        if name.split(".")[0] in forbidden:
                            findings.append(f"{path.relative_to(repo).as_posix()}:{node.lineno}: la capa pura "
                                            f"«{layer['name']}» importa el framework «{name.split('.')[0]}»")
    return findings


def _module_file(repo: Path, module: str) -> str:
    base = Path(*module.split("."))
    for candidate in (base.with_suffix(".py"), base / "__init__.py"):
        if (repo / candidate).is_file():
            return candidate.as_posix()
    return base.as_posix()


def _logical_lines(output: str) -> list[str]:
    """Une la continuación de una violación que el linter partió al ancho de la terminal.

    Una violación empieza en una viñeta ``-   modulo -> `` y sigue, sin sangría, hasta
    cerrar ``(l.N)``. Se prueba primero la unión literal (el linter conserva el espacio
    del corte) y después la unión con un espacio; una línea en blanco, otra viñeta o una
    cabecera de contrato cortan la continuación."""
    lines, out, i = output.splitlines(), [], 0
    while i < len(lines):
        line, i = lines[i], i + 1
        if _BULLET_RE.match(line) and not _VIOLATION_RE.match(line):
            pieces = [line]
            while i < len(lines) and len(pieces) <= MAX_CONTINUATION:
                nxt = lines[i]
                if not nxt.strip() or _BULLET_RE.match(nxt) or _CONTRACT_RE.match(nxt):
                    break
                pieces.append(nxt)
                i += 1
                joined = next((c for c in ("".join(pieces), " ".join(p.strip() for p in pieces))
                               if _VIOLATION_RE.match(c)), None)
                if joined:
                    line = joined
                    break
        out.append(line)
    return out


def parse_violations(repo: Path, output: str) -> list[str]:
    """``fichero:línea: contrato`` de cada import prohibido que imprime import-linter."""
    findings, contract, previous = [], "", ""
    for line in _logical_lines(output):
        if _CONTRACT_RE.match(line) and previous.strip() and not previous.startswith(("-", "Broken")):
            contract = previous.strip()
        match = _VIOLATION_RE.match(line)
        if match:
            findings.append(f"{_module_file(repo, match.group(1))}:{match.group(2)}: "
                            f"import prohibido por la regla «{contract}»")
        previous = line
    return findings


_STALE_RE = re.compile(r"No matches for ignored import (\S+) -> (\S+?)\.?$", re.MULTILINE)


def check_imports(repo: Path, data: dict, config: str, command: str = LINTER) -> list[str]:
    """Hallazgos de las capas puras y de la configuración y ejecución del linter; [] = verde."""
    return pure_findings(repo, data) + _linter_findings(repo, data, config, command)


def _linter_findings(repo: Path, data: dict, config: str, command: str) -> list[str]:
    """Hallazgos de la configuración y de la ejecución del linter."""
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
                              text=True, check=False, encoding="utf-8",
                              env={**os.environ, "COLUMNS": WIDE_COLUMNS})
    except OSError as exc:
        return [f"el linter {command} no se pudo ejecutar: {exc}"]
    if proc.returncode == 0:
        return []
    stale = [f"STALE en tolerated: «{m.group(1)} -> {m.group(2)}» ya no existe; elimine la entrada"
             for m in _STALE_RE.finditer(proc.stdout + proc.stderr)]
    found = stale + parse_violations(repo, proc.stdout)
    return found or [f"el linter {command} salió con código {proc.returncode} sin violaciones "
                     f"legibles: {proc.stderr.strip() or proc.stdout.strip()[-300:]}"]
