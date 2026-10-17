#!/usr/bin/env python3
# check_docs_commands.py — comandos documentados, extraídos y probados (HR-65, #711)
"""Extrae (no enumera) los comandos que ``README.md``, ``AGENTS.md`` y ``docs/**``
mandan ejecutar y los prueba contra la CLI del consumer.

Niveles, decididos mecánicamente:

- ``executed``: verbo declarado seguro, sin placeholders ni shell compuesto;
  se ejecuta en un directorio temporal y no puede ser rechazado por el parser.
- ``parse-only``: lleva placeholders (``<valor>``, se rellenan con un valor
  ficticio) o su verbo no es seguro; se ejecuta con ``dry_run_flag`` y el parser
  real no puede rechazarlo. Sin ``dry_run_flag`` declarado pasa a presencia.
- ``presence-only``: shell compuesto (tuberías, redirecciones, expansiones…) o
  sin forma de probarlo sin efectos; solo se cuenta, para que el nivel sea visible.

El nombre de la CLI es parámetro del consumer. Extraer cero comandos falla
(extractor obsoleto, HR-3). Stdlib exclusivamente, sin red.
"""

from __future__ import annotations

import importlib.util
import re
import shlex
import subprocess
import tempfile
from pathlib import Path

ROOTS = ("README.md", "AGENTS.md", "docs")
TIMEOUT = 20
_PLACEHOLDER = re.compile(r"<[^<>\s]+>")
_SHELL = re.compile(r"[|><;&$`*(){}\\]")


def _load_fences():
    """Carga el módulo compartido ``assets/markdown_fences.py`` por ruta, en el mismo proceso."""
    path = Path(__file__).resolve().parent.parent / "markdown_fences.py"
    spec = importlib.util.spec_from_file_location("markdown_fences", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_FENCES = _load_fences()


def _files(repo, roots):
    for root in roots:
        path = repo / root
        if path.is_file():
            yield path
        elif path.is_dir():
            yield from sorted(path.rglob("*.md"))


def extract(repo, config):
    """[{file, line, command}] de los comandos que empiezan por la CLI declarada."""
    repo, name, found = Path(repo), config["cli_name"], []
    starts = re.compile(r"^(?:\$\s+)?(" + re.escape(name) + r"(?:\s[^\n]*)?)$")
    inline = re.compile(r"`(" + re.escape(name) + r"(?:\s[^`\n]*)?)`")
    for path in _files(repo, config.get("roots", ROOTS)):
        rel = path.relative_to(repo).as_posix()
        for number, line, fenced in _FENCES.classify_lines(path.read_text(encoding="utf-8")):
            hits = [m.group(1) for m in [starts.match(line.strip())] if m] if fenced \
                else inline.findall(line)
            found += [{"file": rel, "line": number, "command": h.strip()} for h in hits]
    return found


def _tier(command, config):
    filled = _PLACEHOLDER.sub("valor", command)
    if _SHELL.search(filled):
        return "presence-only", filled
    words = shlex.split(filled)
    safe = len(words) > 1 and words[1] in config.get("executed_verbs", []) \
        and filled == command
    if safe:
        return "executed", filled
    return ("parse-only" if config.get("dry_run_flag") else "presence-only"), filled


def _run(repo, config, filled, tier):
    argv = [str(repo / a) if (repo / a).is_file() else a for a in config["cli_command"]]
    args = shlex.split(filled)[1:]
    if tier == "parse-only":
        args.append(config["dry_run_flag"])
    with tempfile.TemporaryDirectory() as sandbox:
        try:
            proc = subprocess.run(argv + args, cwd=sandbox, capture_output=True, text=True,
                                  timeout=TIMEOUT, check=False)
        except (OSError, subprocess.TimeoutExpired) as exc:
            return f"no se pudo ejecutar: {exc}"
    if proc.returncode in config.get("rejection_exit_codes", [2]):
        lines = proc.stderr.strip().splitlines()
        return f"rechazado por el parser (exit {proc.returncode}): {lines[0] if lines else ''}"
    return ""


def check(repo, config):
    """(hallazgos, resumen por nivel)."""
    repo = Path(repo).resolve()
    summary = {"executed": 0, "parse-only": 0, "presence-only": 0}
    commands = extract(repo, config)
    if not commands:
        return ["HR-65: no se extrajo ningún comando de la documentación (extractor obsoleto; "
                "sin medición no hay veredicto)"], summary
    excluded = set(config.get("exclude", []))
    findings = [f"HR-65: STALE en exclude: «{e}» ya no se documenta; elimine la entrada"
                for e in sorted(excluded - {c["command"] for c in commands})]
    seen = {}
    for item in commands:
        if item["command"] in excluded:
            continue
        tier, filled = _tier(item["command"], config)
        summary[tier] += 1
        if tier == "presence-only":
            continue
        if item["command"] not in seen:
            seen[item["command"]] = _run(repo, config, filled, tier)
        if seen[item["command"]]:
            findings.append(f"HR-65: {item['file']}:{item['line']}: «{item['command']}» "
                            f"{seen[item['command']]}")
    return findings, summary
