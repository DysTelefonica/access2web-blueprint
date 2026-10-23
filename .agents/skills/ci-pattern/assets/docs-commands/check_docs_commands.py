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

Tres causas, tres hallazgos distintos y tres remedios distintos (#1021): el parser
rechazó el comando documentado, la CLI no está en el runner, o el comando no
terminó a tiempo. La segunda es del entorno y sale **una vez**, no una por comando.
Es además la única que baja a **aviso**: si la política no declara ninguna forma de
ejecutar, todo comando es presence-only, el gate no ejecuta nada y no hay medición
que la CLI ausente pueda ocultar (#1038). El aviso va a `stderr` con prefijo `AVISO`
y los hallazgos vuelven al llamador, como el resto de la familia.
"""

from __future__ import annotations

import importlib.util
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOTS = ("README.md", "AGENTS.md", "docs")
TIMEOUT = 20
# Claves que admite la declaración `docs_commands` del contrato de adopción
# (#1020), las mismas que el esquema `adoption-contract.schema.json`. La fase 5
# rechaza lo que no esté aquí, igual que con docs_links y docs_decisions.
CONFIG_KEYS = ("cli_name", "cli_command", "executed_verbs", "dry_run_flag",
               "rejection_exit_codes", "roots", "exclude", "job")
_PLACEHOLDER = re.compile(r"<[^<>\s]+>")
_SHELL = re.compile(r"[|><;&$`*(){}\\]")


class CliAusenteError(Exception):
    """La CLI declarada no arranca. Es del runner, no de la documentación: el gate
    la traduce a un único hallazgo de entorno, no a uno por comando (#1021)."""


def _argv(repo, config):
    """`cli_command` resuelto: lo que es fichero del repo va con su ruta absoluta y
    lo demás se busca en el `PATH` del runner, porque la disponibilidad de la CLI
    no la declara el contrato."""
    return [str(repo / a) if (repo / a).is_file() else a for a in config["cli_command"]]


def _sin_cli(config, argv, por_que):
    """Hallazgo de entorno: la CLI no se puede lanzar donde corre el job. Dice cuál es
    la causa y qué hacer, y no culpa a la documentación; uno solo, no uno por comando
    documentado (#1021)."""
    return (f"entorno: la CLI «{config['cli_name']}» no se puede ejecutar aquí: «{argv[0]}» "
            f"{por_que}. Los comandos documentados quedan sin probar por falta de la "
            f"herramienta, no por estar mal escritos: instale «{config['cli_name']}» en la "
            f"imagen del runner que ejecuta el job docs-commands.")


def _entorno(repo, config):
    """(`argv`, hallazgo de entorno o `None`). La disponibilidad se decide una vez por
    gate: intentarlo por comando convertía una causa única en N líneas idénticas
    (55 con el `make` ausente que midió APAP#2324)."""
    argv = _argv(repo, config)
    if not argv or shutil.which(argv[0]):
        return argv, None
    por_que = ("existe en el repo pero no es ejecutable" if Path(argv[0]).is_file()
               else "no está en el `PATH` ni es un fichero del repo")
    return argv, _sin_cli(config, argv, por_que)


def _nunca_ejecuta(config):
    """Verdadero si la política no declara ninguna forma de ejecutar un comando: sin
    `dry_run_flag` ni verbos en `executed_verbs`, `_tier` devuelve `presence-only`
    para todos ellos y el gate no lanza nada. Lo que el runner no tenga instalado
    no puede estar ocultando una medición, así que su ausencia avisa y no bloquea
    (#1038: el `make` de APAP#2324 son 58 comandos que nadie iba a ejecutar)."""
    return not config.get("dry_run_flag") and not config.get("executed_verbs")


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


def _run(argv, config, filled, tier):
    """Veredicto de un comando, o `""` si la CLI lo acepta. Cada causa dice qué
    hacer. `CliAusenteError` no se traduce aquí: la comparte el gate como un único
    hallazgo de entorno (#1021)."""
    args = shlex.split(filled)[1:]
    if tier == "parse-only":
        args.append(config["dry_run_flag"])
    with tempfile.TemporaryDirectory() as sandbox:
        try:
            proc = subprocess.run(argv + args, cwd=sandbox, capture_output=True, text=True,
                                  timeout=TIMEOUT, check=False)
        except subprocess.TimeoutExpired:
            return (f"timeout: no terminó en {TIMEOUT}s y el runner lo cortó; el parser no lo "
                    f"rechazó. Si el comando documentado es correcto, este gate no puede "
                    f"probarlo: quítelo del corpus con `exclude` o hágalo más rápido; si no lo "
                    f"es, corrija el comando en la documentación.")
        except OSError as exc:
            raise CliAusenteError(exc) from exc
    if proc.returncode in config.get("rejection_exit_codes", [2]):
        lines = proc.stderr.strip().splitlines()
        return (f"rechazado por el parser (exit {proc.returncode}): "
                f"{lines[0] if lines else ''}; la CLI no acepta el comando documentado: "
                f"corrija el comando en la documentación o quítelo de ella.")
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
    argv, ausente = _entorno(repo, config)
    if ausente and _nunca_ejecuta(config):
        # Aviso, no hallazgo: aquí el gate no iba a ejecutar nada, así que la CLI
        # ausente no oculta ninguna medición y el rojo señalaría un problema del
        # runner donde no hay problema (#1038). Mismo texto y mismo remedio, otro
        # canal: `stderr` con `AVISO`, y el job sale 0.
        print(f"AVISO: HR-65: {ausente}", file=sys.stderr)
    elif ausente:
        findings.append(f"HR-65: {ausente}")
    seen = {}
    for item in commands:
        if item["command"] in excluded:
            continue
        tier, filled = _tier(item["command"], config)
        summary[tier] += 1
        if tier == "presence-only" or ausente:
            continue
        if item["command"] not in seen:
            try:
                seen[item["command"]] = _run(argv, config, filled, tier)
            except CliAusenteError as exc:
                # `which` lo dá por lanzable y no arranca (shebang roto, ELF sin
                # permiso…): mismo hallazgo de entorno, y solo una vez.
                findings.append(f"HR-65: {_sin_cli(config, argv, f'no arranca: {exc}')}")
                ausente = True
                continue
        if seen[item["command"]]:
            findings.append(f"HR-65: {item['file']}:{item['line']}: «{item['command']}» "
                            f"{seen[item['command']]}")
    return findings, summary
