#!/usr/bin/env python3
# check_docs_links.py — enlaces locales de la documentación (HR-66, #712)
"""Los enlaces locales de ``README.md``, ``AGENTS.md`` y ``docs/**`` resuelven.

- Enlace local roto o ancla ``#seccion`` inexistente en un destino Markdown:
  rojo con ``fichero:línea`` y destino. No cuentan las URLs con esquema
  (``https:``, ``mailto:``…), las anclas de la misma página ni los bloques cercados.
- Navegación: cada spoke enlaza al hub y el hub a cada spoke.
- Retirados: ficheros que no pueden existir y frases que no pueden aparecer.
- Deuda: ``baseline`` tolera enlaces rotos ya existentes y solo puede encogerse;
  una entrada que ya no coincide con un enlace roto es ``STALE``.

Stdlib exclusivamente, sin red. Recorrer cero ficheros es un fallo (HR-3).
"""

from __future__ import annotations

import importlib.util
import re
from pathlib import Path

ROOTS = ("README.md", "AGENTS.md", "docs")
CONFIG_KEYS = ("roots", "hub", "spokes", "retired", "baseline", "job")
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
    return findings
