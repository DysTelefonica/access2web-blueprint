#!/usr/bin/env python3
# check_docs_manifest.py — gate del manifiesto canónico de documentos (HR-70, #719)
"""Cada documento activado existe y cada sección obligatoria existe como
título Markdown (H1..H6, fuera de bloques cercados) y no está vacía.

El consumer puede AÑADIR documentos y secciones con ``extend``; cualquier otra
clave del contrato (p. ej. ``exclude``) es un hallazgo: lo del canon no se quita.
Stdlib exclusivamente, sin red.
"""

from __future__ import annotations

import importlib.util
import json
import re
from pathlib import Path

CONTRACT_KEYS = ("extend",)
EMBEDDED_BEGIN = "<!-- ci-pattern-agents:begin -->"
EMBEDDED_END = "<!-- ci-pattern-agents:end -->"
_HEADING_RE = re.compile(r"^(#{1,6})\s+(?:\d+\.\s+)?(.*?)\s*#*\s*$")


def _load_fences():
    """Carga el módulo compartido ``assets/markdown_fences.py`` por ruta, en el mismo proceso."""
    path = Path(__file__).resolve().parent.parent / "markdown_fences.py"
    spec = importlib.util.spec_from_file_location("markdown_fences", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_FENCES = _load_fences()


class ManifestError(Exception):
    """Manifiesto ausente, ilegible o con esquema inválido (exit 2)."""


def validate_manifest(manifest):
    docs = manifest.get("documents") if isinstance(manifest, dict) else None
    if not isinstance(docs, list) or not docs:
        raise ManifestError("el manifiesto necesita 'documents' (lista no vacía)")
    for doc in docs:
        if not isinstance(doc, dict):
            raise ManifestError(f"entrada no-objeto: {doc!r}")
        for key in ("path", "purpose", "required_by"):
            if not isinstance(doc.get(key), str) or not doc[key].strip():
                raise ManifestError(f"{doc.get('path', '?')}: falta '{key}'")
        sections = doc.get("sections")
        if not isinstance(sections, list) or not all(isinstance(s, str) and s.strip() for s in sections):
            raise ManifestError(f"{doc['path']}: 'sections' debe ser una lista de títulos")
    return manifest


def load_manifest(path):
    try:
        return validate_manifest(json.loads(Path(path).read_text(encoding="utf-8")))
    except (OSError, json.JSONDecodeError) as exc:
        raise ManifestError(f"manifiesto ilegible: {exc}") from exc


def _sections(text):
    """{título normalizado: cuerpo}; el cuerpo llega hasta el siguiente título
    de nivel igual o superior (los subtítulos pertenecen a su sección)."""
    heads, lines = [], text.splitlines()
    for number, line, fenced in _FENCES.classify_lines(text):
        match = None if fenced else _HEADING_RE.match(line)
        if match:
            heads.append((number - 1, len(match.group(1)), match.group(2).casefold()))
    found = {}
    for k, (number, level, title) in enumerate(heads):
        end = next((n for n, lv, _t in heads[k + 1:] if lv <= level), len(lines))
        body = lines[number + 1:end]
        if not _has_content(body):
            end = max(end, _embedded_block_end(lines, number + 1))
            body = lines[number + 1:end]
        found.setdefault(title, body)
    return found


def _embedded_block_end(lines, start):
    """Fin del bloque gobernado ``ci-pattern-agents`` que abre justo tras el
    título (#738): el flujo operativo generado lo embebe con sus propios títulos
    de nivel 2, y ese bloque es el cuerpo de su sección. ``start`` si no hay."""
    first = next((n for n in range(start, len(lines)) if lines[n].strip()), None)
    if first is None or lines[first].strip() != EMBEDDED_BEGIN:
        return start
    return next((n + 1 for n in range(first, len(lines)) if lines[n].strip() == EMBEDDED_END), start)


def _has_content(lines):
    body = re.sub(r"<!--.*?-->", "", "\n".join(lines), flags=re.S)
    return bool(body.strip())


def _effective(manifest, extend):
    docs = {d["path"]: {**d, "sections": list(d["sections"])} for d in manifest["documents"]}
    for item in extend:
        entry = docs.setdefault(item["path"], {"path": item["path"], "sections": [],
                                                "purpose": item.get("purpose", ""),
                                                "required_by": "contrato del consumer"})
        entry["sections"] += [s for s in item.get("sections", []) if s not in entry["sections"]]
    return docs.values()


def routing_rows(manifest, contract, repo=None):
    """``(cuándo, lee, ruta)`` de los documentos activados que declaran ``route``.

    Mismas reglas de activación que ``check``; las filas salen del manifiesto y
    del contrato, nunca de texto libre, de modo que lista, contenido y enrutado
    no pueden divergir. ``repo`` opcional: cuando un doc se activa por
    ``project_type:<x>`` se necesita leer ``.team-skills.yaml``; sin repo, ese
    tipo de activación se omite (silenciosamente) y la fila no se enruta."""
    extend = (contract.get("docs_manifest") or {}).get("extend", [])
    return [(d["route"]["when"], d["route"]["read"], d["path"])
            for d in _effective(manifest, extend)
            if d.get("route") and _routing_active(d, contract, repo)]


_PROJECT_TYPE_RE = re.compile(r"^\s*project_type:\s*(\S+)\s*$", re.M)


def _resolve_project_type(repo):
    """Lee ``project_type`` de ``.team-skills.yaml``. Devuelve ``(value, error)``.
    Si no se puede resolver, ``value=None`` y ``error`` describe el motivo
    (``missing``, ``unreadable``, ``no_project_type``). Activación por
    ``project_type:<x>`` falla cerrado en cualquiera de los tres casos (HR-71)."""
    yaml_path = Path(repo) / ".team-skills.yaml"
    if not yaml_path.is_file():
        return None, "missing"
    try:
        text = yaml_path.read_text(encoding="utf-8")
    except OSError:
        return None, "unreadable"
    m = _PROJECT_TYPE_RE.search(text)
    if m is None:
        return None, "no_project_type"
    return m.group(1), None


def _active(doc, contract, repo, findings):
    """``True`` si el doc debe medirse. Soporta activación por
    ``project_type:<x>`` leyendo ``.team-skills.yaml``; cuando no se puede
    resolver, emite un hallazgo HR-71 y activa el doc (fail-closed)."""
    gate = doc.get("activated_by")
    if not gate:
        return True
    if gate in contract:
        return True
    if gate.startswith("project_type:"):
        value = gate.split(":", 1)[1].strip()
        pt, err = _resolve_project_type(repo)
        if pt is None:
            findings.append(
                f"HR-71: no se puede resolver '{gate}': .team-skills.yaml "
                f"{err}; fail-closed")
            return True
        return pt == value
    return False


def _routing_active(doc, contract, repo):
    """Variante silenciosa de ``_active`` para ``routing_rows``: cuando no se
    puede resolver la activación, devuelve None (la fila no se enruta). El
    hallazgo de fail-closed lo emite ``check``; aquí la omisión es coherente
    con no saber a qué consumer aplicar el enrutado."""
    gate = doc.get("activated_by")
    if not gate:
        return True
    if gate in contract:
        return True
    if gate.startswith("project_type:"):
        if repo is None:
            return False
        value = gate.split(":", 1)[1].strip()
        pt, _err = _resolve_project_type(repo)
        if pt is None:
            return False
        return pt == value
    return False


def check(repo, manifest, contract):
    declared = contract.get("docs_manifest")
    findings = [f"HR-70: la clave '{k}' del contrato no se admite: lo del canon no se quita"
                for k in declared if k not in CONTRACT_KEYS]
    extend = declared.get("extend", [])
    for doc in _effective(manifest, extend):
        if not _active(doc, contract, repo, findings):
            continue
        paths = sorted(repo.glob(doc["path"])) if "*" in doc["path"] else [repo / doc["path"]]
        paths = [p for p in paths if p.is_file()]
        if not paths:
            findings.append(f"HR-70: {doc['path']}: falta el documento ({doc['purpose']}; "
                            f"lo exige {doc['required_by']})")
        for path in paths:
            found = _sections(path.read_text(encoding="utf-8"))
            rel = path.relative_to(repo).as_posix()
            for title in doc["sections"]:
                body = found.get(title.casefold())
                if body is None:
                    findings.append(f"HR-70: {rel}: falta la sección «{title}»")
                elif not _has_content(body):
                    findings.append(f"HR-70: {rel}: la sección «{title}» está vacía")
    return findings
