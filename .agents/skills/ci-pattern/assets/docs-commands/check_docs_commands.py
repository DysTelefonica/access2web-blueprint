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

Los niveles de arriba prueban la FORMA de un comando: que el parser real lo acepte. La
capa de existencia, mas abajo, comprueba otra cosa: que lo que el comando invoca exista.
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
               "rejection_exit_codes", "roots", "exclude", "job", "existence")
# Claves que admite la declaración `existence` (la capa de existencia, #1374).
EXISTENCE_KEYS = ("makefile", "path_executables")
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


# --- Capa de existencia ---------------------------------------------------
#
# Esta capa comprueba que lo que el comando invoca exista: los niveles de arriba solo
# prueban la forma. Con `dry_run_flag --help` un comando documentado se pasa aunque lo
# que invoca ya no este (el `make test` de la documentacion de BP sigue citando un
# target cuyo `--cov` dejo de ejecutarse, #1374). No ejecuta nada: solo mira ficheros,
# el Makefile y el `PATH`. La activa la declaracion `existence` del contrato.

MAKE = "make"
_INLINE_SPAN = re.compile(r"`([^`\n]+)`")
_RULE = re.compile(r"^([^\t#][^:=]*?)\s*:(?!=)(.*)$")


def _tokens(text):
    """Las palabras de un comando, o `[]` si no se pueden tokenizar."""
    try:
        return shlex.split(text)
    except ValueError:
        return []


def _flags(tokens):
    """Los flags de una lista de palabras, sin placeholders ni shell: los que se
    pueden comparar con la receta de un target."""
    return {t for t in tokens if t.startswith("-") and len(t) > 1
            and not _PLACEHOLDER.search(t) and not _SHELL.search(t)}


def _spans(repo, config):
    """[{file, line, text}] de los fragmentos de codigo de la documentacion: la
    linea entera de un bloque cercado y cada fragmento en linea fuera de el."""
    spans = []
    for path in _files(repo, config.get("roots", ROOTS)):
        rel = path.relative_to(repo).as_posix()
        for number, line, fenced in _FENCES.classify_lines(path.read_text(encoding="utf-8")):
            if fenced:
                text = line.strip().removeprefix("$ ").strip()
                if text:
                    spans.append({"file": rel, "line": number, "text": text})
            else:
                spans += [{"file": rel, "line": number, "text": m.strip()}
                          for m in _INLINE_SPAN.findall(line) if m.strip()]
    return spans


def _make_target(tokens, default_makefile):
    """(target, makefile) del comando `make`, o `(None, None)` cuando no se puede
    determinar de forma determinista: sin target, con target de patron, o con
    `-C`/`-f`, que consultan otro Makefile."""
    target, rest, i = None, tokens[1:], 0
    while i < len(rest):
        tok = rest[i]
        if tok in ("-C", "--directory", "-f", "--file"):
            return None, None
        if tok.startswith("-") or "=" in tok:
            i += 1
            continue
        target = tok
        break
    if not target or "%" in target:
        return None, None
    return target, default_makefile


def _makefile_rules(text):
    """{target: [lineas de receta]} de un Makefile: el nombre de cada regla, las
    lineas de receta (las que empiezan por tabulador) y los targets `.PHONY`. Una
    regla repartida en varias lineas con `\\` se lee como una sola."""
    rules, current, buffer = {}, [], ""
    for raw in text.splitlines():
        line = raw.rstrip()
        if line.endswith("\\"):
            buffer += line[:-1] + " "
            continue
        line, buffer = buffer + line, ""
        if line.startswith("\t"):
            for name in current:
                rules[name].append(line.strip())
            continue
        match = _RULE.match(line)
        names = match.group(1).split() if match else []
        if names == [".PHONY"]:
            names = line.split(":", 1)[1].split()
        elif names and names[-1].endswith(("?", "+", "!")):
            names = []
        current = names
        for name in current:
            rules.setdefault(name, [])
    return rules


def _kind(subject, path_executables):
    """Que existencia se le comprueba a un fragmento, si alguna."""
    if subject == MAKE:
        return "make"
    if subject.startswith("./") or subject.startswith("../"):
        return "script"
    return "path" if subject in path_executables else None


def _hallazgo(span, veredicto):
    """El hallazgo de la capa de existencia, con `fichero:línea` como los demas."""
    return f"HR-65: {span['file']}:{span['line']}: «{span['text']}» {veredicto}"


def _subject_findings(repo, span, kind, subject):
    """El script local que el comando cita tiene que existir, y el ejecutable del
    `PATH` que la politica declara permitido tiene que resolver."""
    if kind == "script":
        rel = Path(subject).as_posix()
        if (repo / rel).is_file():
            return []
        return [_hallazgo(span, f"no existe: el repositorio no tiene el script «{rel}»; "
                               f"corrija la ruta en la documentacion o quite la cita.")]
    if shutil.which(subject):
        return []
    return [_hallazgo(span, f"no se resuelve: «{subject}» no esta en el `PATH` del runner; "
                           f"installe la herramienta en la imagen del runner o quite la cita.")]


def _make_findings(repo, config, span, tokens, group):
    """El target de `make` que el comando cita tiene que existir en el Makefile, y
    los flags que la documentacion le atribuye tienen que estar en su receta.

    La segunda parte es el caso vivo de BP (`docs/adopcion-fase-2-aislamiento.md:41`):
    la documentacion dice que `make test` mide con `--cov`, y el target dejo de
    medirlo. La atribucion es de la misma linea y solo con la CLI declarada, que
    es lo que hace la regla estrecha: una linea que cita un target y otro comando
    que no es la CLI no se cruza con el."""
    declared = config.get("existence", {}).get("makefile", "Makefile")
    target, rel = _make_target(tokens, declared)
    if target is None:
        return []
    makefile = repo / rel
    rules = _makefile_rules(makefile.read_text(encoding="utf-8")) if makefile.is_file() else {}
    if target not in rules:
        return [_hallazgo(span, f"no existe: el target «{target}» no esta en el Makefile «{rel}»; "
                               f"corrija el comando en la documentacion o quite la cita.")]
    recipe = " ".join(rules[target])
    cli, attributed = config.get("cli_name"), set()
    for other in group:
        other_tokens = _tokens(other["text"])
        if other is not span and other_tokens and other_tokens[0] == cli:
            attributed |= _flags(other_tokens)
    missing = sorted(attributed - _flags(_tokens(recipe)))
    if not missing:
        return []
    quoted = " ni ".join(f"«{flag}»" for flag in missing)
    return [_hallazgo(span, f"no hace lo que la documentacion dice: el target «{target}» no "
                           f"ejecuta {quoted}, que la documentacion le atribuye; su receta es "
                           f"«{recipe}»; corrija la documentacion o el Makefile.")]


def _declaracion_existence(config):
    """(hallazgos, ejecutables del PATH) de la declaracion `existence`, o `None` si el
    consumer no la declara. Una declaracion mal formada es un hallazgo, no un
    fallo del gate (HR-3)."""
    decl = config.get("existence")
    if decl is None:
        return None
    if not isinstance(decl, dict):
        return ["HR-65: la declaracion «existence» debe ser un objeto"], []
    findings = [f"HR-65: clave desconocida en docs_commands.existence: «{key}»"
                for key in sorted(set(decl) - set(EXISTENCE_KEYS))]
    path_executables = decl.get("path_executables", [])
    if not isinstance(path_executables, list):
        findings.append("HR-65: la declaracion «existence.path_executables» debe ser una lista")
        path_executables = []
    return findings, path_executables


def _linea_findings(repo, config, group, path_executables, excluded):
    """(hallazgos, textos) de los fragmentos de una linea."""
    findings, documented, makes = [], set(), []
    for span in group:
        tokens = _tokens(span["text"])
        kind = _kind(tokens[0] if tokens else "", path_executables)
        if kind is None:
            continue
        documented.add(span["text"])
        if span["text"] in excluded:
            continue
        if kind == "make":
            makes.append((span, tokens))
            continue
        findings += _subject_findings(repo, span, kind, tokens[0])
    if len(makes) == 1:
        findings += _make_findings(repo, config, makes[0][0], makes[0][1], group)
    return findings, documented


def _existence(repo, config, spans, excluded):
    """(hallazgos, textos documentados) de la capa de existencia."""
    declaracion = _declaracion_existence(config)
    if declaracion is None:
        return [], set()
    findings, path_executables = declaracion
    grouped = {}
    for span in spans:
        grouped.setdefault((span["file"], span["line"]), []).append(span)
    documented = set()
    for group in grouped.values():
        lineas, textos = _linea_findings(repo, config, group, path_executables, excluded)
        findings += lineas
        documented |= textos
    return findings, documented


def check(repo, config):
    """(hallazgos, resumen por nivel)."""
    repo = Path(repo).resolve()
    summary = {"executed": 0, "parse-only": 0, "presence-only": 0}
    commands = extract(repo, config)
    if not commands:
        return ["HR-65: no se extrajo ningún comando de la documentación (extractor obsoleto; "
                "sin medición no hay veredicto)"], summary
    excluded = set(config.get("exclude", []))
    spans = _spans(repo, config)
    existence_findings, existence_texts = _existence(repo, config, spans, excluded)
    findings = [f"HR-65: STALE en exclude: «{e}» ya no se documenta; elimine la entrada"
                for e in sorted(excluded - {c["command"] for c in commands} - existence_texts)]
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
    findings += existence_findings
    return findings, summary
