"""Static content of the admin IT handover guide (Citrix + Linux).

Facts come from docs/07-migracion/epica-coexistencia-postgresql.md
(decisions 10, 18-20, 28; H-3, H-4) and docs/architecture.md (D93).
The guide prepares an isolated rehearsal; it never authorises the
production data move. The content is data, not code: it lives in
``guia_it_content.json``, next to this module, which only loads it.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import NamedTuple

from app.src.modules.lanzadera.delivery.http.guia_it_glossary import (
    GLOSSARY,
    GLOSSARY_TERMS,
)

_SOURCE = Path(__file__).with_name("guia_it_content.json")


class Scenario(NamedTuple):
    key: str
    label: str
    title: str
    summary: str
    tag: str
    terms: tuple[str, ...]


class Flow(NamedTuple):
    route: str
    transport: str
    condition: str
    terms: tuple[str, ...]


class Check(NamedTuple):
    key: str
    label: str


class Command(NamedTuple):
    key: str
    label: str
    text: str


def _with_terms(entry: dict[str, object]) -> dict[str, object]:
    if "terms" in entry:
        return {**entry, "terms": tuple(entry["terms"])}
    return entry


def load_content(source: Path = _SOURCE) -> dict[str, object]:
    """Read the guide sections and build the template context."""
    data = json.loads(source.read_text(encoding="utf-8"))
    return {
        "glossary": GLOSSARY,
        "glossary_terms": GLOSSARY_TERMS,
        "scenarios": tuple(Scenario(**_with_terms(e)) for e in data["scenarios"]),
        "flows": tuple(Flow(**_with_terms(e)) for e in data["flows"]),
        "profile_tests": tuple(data["profile_tests"]),
        "checks": tuple(Check(**e) for e in data["checks"]),
        "commands": tuple(Command(**e) for e in data["commands"]),
        "sources": tuple(data["sources"]),
    }


GUIDE_CONTEXT: dict[str, object] = load_content()

__all__ = ["GUIDE_CONTEXT", "load_content"]
