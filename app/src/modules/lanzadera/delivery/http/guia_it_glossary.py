"""Glossary of the admin IT handover guide (Citrix + Linux).

Every technical term the guide links to is defined in
``guia_it_glossary.json``, next to this module, so a reader who does not
know a word (for example, bitness) can look it up. The terms are data,
not code: they live in JSON and this module only loads them.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, NamedTuple

_HERE = Path(__file__).parent


def load_json(name: str) -> Any:
    """Read one of the guide's JSON data files, kept next to this module."""
    return json.loads((_HERE / name).read_text(encoding="utf-8"))


class Term(NamedTuple):
    key: str
    term: str
    definition: str


def load_glossary() -> tuple[Term, ...]:
    """Read the glossary terms, in document order."""
    return tuple(Term(**entry) for entry in load_json("guia_it_glossary.json"))


GLOSSARY: tuple[Term, ...] = load_glossary()

GLOSSARY_TERMS: dict[str, str] = {entry.key: entry.term for entry in GLOSSARY}

__all__ = ["GLOSSARY", "GLOSSARY_TERMS", "Term", "load_glossary", "load_json"]
