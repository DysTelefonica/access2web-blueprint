#!/usr/bin/env python3
# check_docs_decisions.py — registro de decisiones por consumer (HR-67, #713)
"""``docs/decisions/``: un fichero por decisión (``AAAA-MM-DD-<slug>.md``) y un
índice ``README.md``. Cada decisión declara ``estado`` (adoptada, rechazada,
sustituida o aplazada), ``fecha`` (ISO), ``decide`` y ``motivo``; una sustituida
enlaza con ``sustituida-por`` a otra que existe (sin ciclos) y una aplazada
declara su ``revision``. Toda cita ``decisión: <slug>`` de la documentación
apunta a una decisión que existe y el índice lista exactamente el directorio.

Stdlib exclusivamente, sin red. Un directorio sin decisiones es un fallo (HR-3).
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

CONFIG_KEYS = ("dir", "cite_roots", "job")
STATES = ("adoptada", "rechazada", "sustituida", "aplazada")
_FIELD_RE = re.compile(r"^-\s+(estado|fecha|decide|motivo|sustituida-por|revision):\s*(.*?)\s*$")
_CITE_RE = re.compile(r"decisión:\s*([\w\-]+)")
_INDEX_LINK_RE = re.compile(r"\[[^\]]*\]\(([^)#\s]+)\)")


def _fields(text):
    return {m.group(1): (m.group(2), n) for n, line in enumerate(text.splitlines(), 1)
            if (m := _FIELD_RE.match(line))}


def _valid_date(value):
    try:
        return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)) and bool(date.fromisoformat(value))
    except ValueError:
        return False


def _format_findings(rel, fields, known):
    out = [f"HR-67: {rel}: falta el campo «{k}»" for k in ("estado", "fecha", "decide", "motivo")
           if not fields.get(k, ("",))[0]]
    state = fields.get("estado", ("",))[0]
    if state and state not in STATES:
        out.append(f"HR-67: {rel}:{fields['estado'][1]}: estado «{state}» fuera de {', '.join(STATES)}")
    if fields.get("fecha", ("",))[0] and not _valid_date(fields["fecha"][0]):
        out.append(f"HR-67: {rel}:{fields['fecha'][1]}: fecha «{fields['fecha'][0]}» no es ISO (AAAA-MM-DD)")
    target = fields.get("sustituida-por", ("",))[0]
    if state == "sustituida" and not target:
        out.append(f"HR-67: {rel}: una decisión sustituida necesita «sustituida-por»")
    elif target and target not in known:
        out.append(f"HR-67: {rel}:{fields['sustituida-por'][1]}: sustituida-por apunta a «{target}», que no existe")
    if state == "aplazada" and not fields.get("revision", ("",))[0]:
        out.append(f"HR-67: {rel}: una decisión aplazada necesita su «revision» (condición o fecha)")
    return out


def _cycles(decisions):
    out = []
    for slug, fields in sorted(decisions.items()):
        seen, cur = [slug], fields.get("sustituida-por", ("",))[0]
        while cur and cur in decisions:
            if cur in seen:
                out.append(f"HR-67: ciclo de sustitución: {' -> '.join(seen + [cur])}")
                break
            seen.append(cur)
            cur = decisions[cur].get("sustituida-por", ("",))[0]
    return out


def check(repo, config):
    repo = Path(repo)
    base = config.get("dir", "docs/decisions")
    files = sorted(p for p in (repo / base).glob("*.md") if p.name != "README.md")
    if not files:
        return [f"HR-67: {base} no contiene ninguna decisión (sin medición no hay veredicto)"]
    decisions = {p.stem: _fields(p.read_text(encoding="utf-8")) for p in files}
    findings = []
    for path in files:
        findings += _format_findings(path.relative_to(repo).as_posix(), decisions[path.stem], decisions)
    findings += _cycles(decisions)
    index = repo / base / "README.md"
    listed = {Path(t).stem for t in _INDEX_LINK_RE.findall(index.read_text(encoding="utf-8"))} \
        if index.is_file() else set()
    if not index.is_file():
        findings.append(f"HR-67: falta el índice {base}/README.md")
    findings += [f"HR-67: la decisión {s} no figura en el índice {base}/README.md" for s in sorted(set(decisions) - listed)]
    findings += [f"HR-67: el índice {base}/README.md lista {s}, que no existe" for s in sorted(listed - set(decisions))]
    for root in config.get("cite_roots", ["README.md", "AGENTS.md", "docs"]):
        path = repo / root
        for doc in ([path] if path.is_file() else sorted(path.rglob("*.md")) if path.is_dir() else []):
            if (repo / base) in doc.parents:
                continue
            for number, line in enumerate(doc.read_text(encoding="utf-8").splitlines(), 1):
                findings += [f"HR-67: {doc.relative_to(repo).as_posix()}:{number} cita la decisión «{s}», que no existe"
                             for s in _CITE_RE.findall(line) if s not in decisions]
    return findings
