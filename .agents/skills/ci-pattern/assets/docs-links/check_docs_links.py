#!/usr/bin/env python3
# check_docs_links.py — enlaces locales de la documentación (HR-66, #712)
"""Los enlaces locales de ``README.md``, ``AGENTS.md`` y ``docs/**`` resuelven.

- Enlace local roto o ancla ``#seccion`` inexistente en un destino Markdown:
  rojo con ``fichero:línea`` y destino. No cuentan las URLs con esquema
  (``https:``, ``mailto:``…), las anclas de la misma página ni los bloques cercados.
- Navegación: cada spoke enlaza al hub y el hub a cada spoke.
- Alcanzabilidad (HR-80, #1379): todo ``.md`` del alcance declarado tiene que
  ser alcanzable desde los puntos de entrada siguiendo enlaces markdown.
- Retirados: ficheros que no pueden existir y frases que no pueden aparecer.
- Deuda: ``baseline`` tolera enlaces rotos ya existentes y solo puede encogerse;
  una entrada que ya no coincide con un enlace roto es ``STALE``.

Stdlib exclusivamente, sin red. Recorrer cero ficheros es un fallo (HR-3).
"""

from __future__ import annotations

import fnmatch
import importlib.util
import re
from pathlib import Path

ROOTS = ("README.md", "AGENTS.md", "docs")
CONFIG_KEYS = ("roots", "hub", "spokes", "retired", "baseline", "job", "reachable")
# Distingue «la clave no está declarada» de «la clave vale null»: lo primero
# deja la regla apagada, lo segundo es una declaración mal formada (HR-3).
_NOT_DECLARED = object()
_LINK_RE = re.compile(r"\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
_HEADING_RE = re.compile(r"^#{1,6}\s+(.*?)\s*#*$")


def _load_fences():
    """Carga el módulo compartido ``assets/markdown_fences.py`` por ruta, en el mismo proceso."""
    path = Path(__file__).resolve().parent.parent / "markdown_fences.py"
    spec = importlib.util.spec_from_file_location("markdown_fences", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_FENCES = _load_fences()


def _slug(title):
    return re.sub(r"[^\w\- ]", "", title).strip().lower().replace(" ", "-")


def _walk(repo, roots):
    for root in roots:
        path = repo / root
        if path.is_file():
            yield path
        elif path.is_dir():
            yield from sorted(path.rglob("*.md"))


def _lines(path):
    """(número, línea) fuera de bloques cercados (CommonMark)."""
    return _FENCES.unfenced_lines(path.read_text(encoding="utf-8"))


def _anchors(path):
    return {_slug(m.group(1)) for _n, line in _lines(path) if (m := _HEADING_RE.match(line))}


def _local_links(path):
    for number, line in _lines(path):
        for dest in _LINK_RE.findall(line):
            if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", dest) or dest.startswith("#"):
                continue
            yield number, dest


def _resolve(repo, path, dest):
    target, _, anchor = dest.partition("#")
    base = repo if target.startswith("/") else path.parent
    return (base / target.lstrip("/")).resolve(), anchor


def _under(repo, path):
    """True si ``path`` está dentro del repositorio (o es el propio repositorio)."""
    try:
        path.relative_to(repo)
    except ValueError:
        return False
    return True


def _string_list(value, allow_empty=False):
    """True si ``value`` es una lista de cadenas no vacías."""
    if not isinstance(value, list) or (not value and not allow_empty):
        return False
    return all(isinstance(item, str) and item for item in value)


def _reachable_declaration_findings(decl):
    """Hallazgos de forma de la declaración ``reachable`` (fail-closed, HR-3)."""
    if not isinstance(decl, dict):
        return ["HR-80: reachable debe ser un objeto"]
    findings = []
    if "from" not in decl:
        findings.append("HR-80: reachable incompleto, falta la clave 'from'")
    elif not _string_list(decl["from"]):
        findings.append("HR-80: reachable.from debe ser una lista no vacía de rutas")
    if "scope" in decl and not _string_list(decl["scope"]):
        findings.append("HR-80: reachable.scope debe ser una lista no vacía de rutas")
    if "exceptions" in decl and not _string_list(decl["exceptions"], allow_empty=True):
        findings.append("HR-80: reachable.exceptions debe ser una lista de rutas")
    return findings


def _reachable_entries(repo, decl):
    """(entradas, hallazgos): los puntos de entrada declarados que existen."""
    entries, findings = [], []
    for rel in decl["from"]:
        path = (repo / rel).resolve()
        if not path.is_file():
            findings.append(f"HR-80: el punto de entrada {rel} no existe")
        elif not _under(repo, path):
            findings.append(f"HR-80: el punto de entrada {rel} está fuera del repositorio")
        else:
            entries.append(path)
    return entries, findings


def _reachable_files(repo, entries):
    """(alcanzables, hallazgos): ficheros alcanzables siguiendo enlaces markdown.

    El recorrido es transitivo y sale del repositorio nunca: un enlace a otro
    sitio no trae la documentación de este repo.
    """
    seen, visited, frontier = set(), set(), list(entries)
    findings = []
    while frontier:
        path = frontier.pop()
        if path in visited:
            continue
        visited.add(path)
        if not path.is_file() or path.suffix.lower() != ".md":
            continue
        seen.add(path)
        try:
            links = list(_local_links(path))
        except OSError as exc:
            findings.append(f"HR-80: no se puede leer {path} para seguir sus enlaces: {exc}")
            continue
        for _number, dest in links:
            target, _anchor = _resolve(repo, path, dest)
            if _under(repo, target) and target not in visited:
                frontier.append(target)
    return seen, findings


def _scope_files(repo, scope):
    """(ficheros, hallazgos): los .md del alcance declarado y sus rutas rotas."""
    files, findings = [], []
    for rel in scope:
        path = (repo / rel).resolve()
        if not _under(repo, path):
            findings.append(f"HR-80: el alcance declarado {rel} está fuera del repositorio")
        elif path.is_dir():
            files.extend(sorted(item.resolve() for item in path.rglob("*.md")))
        elif path.is_file():
            files.append(path.resolve())
        else:
            findings.append(f"HR-80: el alcance declarado {rel} no existe")
    return files, findings


def _exempt(rel, exceptions):
    """True si la ruta relativa ``rel`` está exenta por una excepción declarada.

    Una excepción es una ruta (fichero o directorio) o un glob; el directorio
    cubre todo lo que cuelga de él.
    """
    for exception in exceptions:
        normalized = exception.rstrip("/")
        if not normalized:
            continue
        if rel == normalized or rel.startswith(normalized + "/"):
            return True
        if fnmatch.fnmatchcase(rel, exception) or fnmatch.fnmatchcase(rel, normalized):
            return True
    return False


def _reachable_findings(repo, decl):
    """HR-80 (#1379): todo .md del alcance tiene que ser alcanzable desde los
    puntos de entrada siguiendo enlaces markdown (transitivos)."""
    if decl is _NOT_DECLARED:
        return []
    findings = _reachable_declaration_findings(decl)
    if findings:
        return findings
    entries, entry_findings = _reachable_entries(repo, decl)
    findings += entry_findings
    if findings:
        return findings
    scope = decl.get("scope") or ["docs"]
    exceptions = decl.get("exceptions") or []
    files, scope_findings = _scope_files(repo, scope)
    findings += scope_findings
    reachable, read_findings = _reachable_files(repo, entries)
    findings += read_findings
    origins = ", ".join(decl["from"])
    for path in sorted(files, key=lambda item: item.relative_to(repo).as_posix()):
        rel = path.relative_to(repo).as_posix()
        if path in reachable or _exempt(rel, exceptions):
            continue
        findings.append(f"HR-80: {rel} es un documento huérfano: no se alcanza desde "
                        f"ningún punto de entrada ({origins})")
    return findings


def check(repo, config):
    repo = Path(repo).resolve()
    files = list(_walk(repo, config.get("roots", ROOTS)))
    if not files:
        return ["HR-66: no se recorrió ningún fichero de documentación (sin medición no hay veredicto)"]
    baseline = {(b["file"], b["target"]) for b in config.get("baseline", [])}
    used, findings, linked = set(), [], {}
    for path in files:
        rel = path.relative_to(repo).as_posix()
        for number, dest in _local_links(path):
            target, anchor = _resolve(repo, path, dest)
            linked.setdefault(path.resolve(), set()).add(target)
            broken = not target.exists() or (anchor and target.suffix == ".md"
                                             and anchor.lower() not in _anchors(target))
            if not broken:
                continue
            if (rel, dest) in baseline:
                used.add((rel, dest))
            else:
                findings.append(f"HR-66: {rel}:{number} enlaza {dest} y no resuelve")
    findings += [f"HR-66: STALE en baseline: {f} ya no enlaza {t} roto; elimine la entrada"
                 for f, t in sorted(baseline - used)]
    hub = config.get("hub")
    if hub and config.get("spokes"):
        hub_path = (repo / hub).resolve()
        for spoke in sorted((repo / config["spokes"]).rglob("*.md")):
            rel = spoke.relative_to(repo).as_posix()
            if hub_path not in linked.get(spoke.resolve(), set()):
                findings.append(f"HR-66: el spoke {rel} no enlaza de vuelta al hub {hub}")
            if spoke.resolve() not in linked.get(hub_path, set()):
                findings.append(f"HR-66: el hub {hub} no enlaza al spoke {rel}")
    retired = config.get("retired", {})
    findings += [f"HR-66: el fichero retirado {r} existe" for r in retired.get("files", [])
                 if (repo / r).exists()]
    for item in retired.get("phrases", []):
        target = repo / item["path"]
        if target.is_file() and item["phrase"] in target.read_text(encoding="utf-8"):
            findings.append(f"HR-66: {item['path']} sigue diciendo la frase retirada «{item['phrase']}»")
    findings += _reachable_findings(repo, config.get("reachable", _NOT_DECLARED))
    return findings
