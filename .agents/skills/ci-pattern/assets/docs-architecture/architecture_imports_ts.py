# architecture_imports_ts.py — reglas de imports en TypeScript (HR-68, #714)
"""Adaptador TypeScript: el bloque ``architecture`` es la única fuente.

De las capas y de ``allowed`` se genera la configuración de ``dependency-cruiser``
(una regla ``forbidden`` por cada par de capas no permitido, más una que prohíbe
los imports locales sin resolver, para que un import roto no pase por verde). El
gate falla si la configuración versionada diverge de la generada, si el linter no
está instalado o no analizó ningún módulo (HR-3) y si un import viola una regla,
con ``fichero:línea`` y la regla incumplida. Stdlib; el linter va como subproceso.
"""

from __future__ import annotations

import importlib.util
import json
import re
import shutil
import subprocess
from pathlib import Path

LINTER = "depcruise"
SOURCE_SUFFIXES = (".ts", ".tsx", ".mts", ".cts", ".js", ".jsx", ".mjs", ".cjs")
UNRESOLVED_RULE = "no-unresolvable-local-imports"


def _load_python_adapter():
    """Reutiliza ``forbidden_pairs`` del adaptador Python (una sola definición de las reglas)."""
    path = Path(__file__).resolve().parent / "architecture_imports.py"
    spec = importlib.util.spec_from_file_location("architecture_imports", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _path_regex(modules: list[str]) -> str:
    """Regex de ruta: un directorio casa todo lo que cuelga de él; un fichero, exactamente él."""
    parts = [re.escape(m) + ("$" if Path(m).suffix else "/") for m in modules]
    return "^" + (parts[0] if len(parts) == 1 else "(" + "|".join(parts) + ")")


def render_depcruise(data: dict) -> str:
    """Configuración de dependency-cruiser derivada del bloque; determinista."""
    layers = {layer["name"]: layer["modules"] for layer in data["layers"]}
    rules = [{"name": f"{src}-must-not-import-{dst}", "severity": "error",
              "from": {"path": _path_regex(layers[src])}, "to": {"path": _path_regex(layers[dst])}}
             for src, dst in _load_python_adapter().forbidden_pairs(data)]
    rules.append({"name": UNRESOLVED_RULE, "severity": "error", "from": {},
                  "to": {"couldNotResolve": True, "path": "^\\."}})
    options = {"enhancedResolveOptions": {"extensions": [".ts", ".tsx", ".js", ".jsx"]}}
    return ("// Generado desde el bloque `architecture` del documento (HR-68); no editar a mano.\n"
            f"module.exports = {{\n  forbidden: {json.dumps(rules, indent=2)},\n"
            f"  options: {json.dumps(options)}\n}};\n")


def _source_files(repo: Path, data: dict) -> list[str]:
    found = set()
    for layer in data["layers"]:
        for module in layer["modules"]:
            path = repo / module
            candidates = path.rglob("*") if path.is_dir() else [path]
            found.update(p.relative_to(repo).as_posix() for p in candidates
                         if p.is_file() and p.suffix in SOURCE_SUFFIXES and "node_modules" not in p.parts)
    return sorted(found)


def _line_of(repo: Path, source: str, specifier: str) -> int:
    """Primera línea de ``source`` que importa ``specifier`` (0 si no se localiza)."""
    quoted = (f'"{specifier}"', f"'{specifier}'")
    for number, line in enumerate((repo / source).read_text(encoding="utf-8").splitlines(), 1):
        if any(q in line for q in quoted):
            return number
    return 0


def parse_violations(repo: Path, report: dict) -> list[str]:
    findings = []
    for violation in report["summary"]["violations"]:
        if violation["rule"]["severity"] != "error":
            continue
        line = _line_of(repo, violation["from"], violation.get("unresolvedTo", violation["to"]))
        where = f"{violation['from']}:{line}" if line else violation["from"]
        findings.append(f"{where}: import prohibido por la regla «{violation['rule']['name']}»")
    return findings


def check_imports(repo: Path, data: dict, config: str, command: str = LINTER) -> list[str]:
    """Hallazgos de la configuración y de la ejecución del linter; [] = verde."""
    path = repo / config
    if not path.is_file():
        return [f"falta la configuración del linter {config}; genérela con --write-config"]
    if path.read_text(encoding="utf-8").replace("\r\n", "\n") != render_depcruise(data):
        return [f"la configuración {config} no coincide con las capas y reglas del documento; "
                "regenérela con --write-config (el documento es la única fuente)"]
    if shutil.which(command) is None:
        return [f"el linter {command} no está instalado: sin ejecución no hay medición (HR-3)"]
    files = _source_files(repo, data)
    if not files:
        return ["<empty: ningún fichero TypeScript o JavaScript bajo los módulos de las capas>"]
    try:
        proc = subprocess.run([command, "--config", config, "--output-type", "json", *files],
                              cwd=repo, capture_output=True, text=True, check=False, encoding="utf-8")
        report = json.loads(proc.stdout)
    except (OSError, json.JSONDecodeError) as exc:
        return [f"el linter {command} no produjo un informe legible: {exc}"]
    if not report["summary"].get("totalCruised"):
        return [f"el linter {command} no analizó ningún módulo: sin ejecución real no hay medición (HR-3)"]
    return parse_violations(repo, report)
