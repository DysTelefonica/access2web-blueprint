"""Glossary of the admin IT handover guide (Citrix + Linux).

Every technical term the guide links to is defined in
``guia_it_glossary.json``, next to this module, so a reader who does not
know a word (for example, bitness) can look it up. The terms are data,
not code: they live in JSON and this module only loads them.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import NamedTuple

_SOURCE = Path(__file__).with_name("guia_it_glossary.json")


class Term(NamedTuple):
    key: str
    term: str
    definition: str


def load_glossary(source: Path = _SOURCE) -> tuple[Term, ...]:
    """Read the glossary terms, in document order."""
    entries = json.loads(source.read_text(encoding="utf-8"))
    return tuple(Term(**entry) for entry in entries)


GLOSSARY: tuple[Term, ...] = load_glossary()

GLOSSARY_TERMS: dict[str, str] = {entry.key: entry.term for entry in GLOSSARY}

__all__ = ["GLOSSARY", "GLOSSARY_TERMS", "Term", "load_glossary"]
