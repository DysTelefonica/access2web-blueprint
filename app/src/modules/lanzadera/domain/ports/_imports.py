# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W37 — _imports
"""Shared prelude for the repository port declarations.

The two main ports (AppRepositoryPort, AssignmentRepositoryPort) carry
the same three-statement import block:

- ``from collections.abc import Sequence``
- ``from typing import Protocol``
- ``from uuid import UUID``

DRY BASELINE entry ``dup:08c289bbf440`` flagged the duplication as a
five-statement window (docstring + ``from __future__`` + the three
imports) with occurrences=2. W34/W36 cleared the same shape for the
domain entities; W37 applies it to ``domain/ports/``.

Each port collapses to a single multi-line
``from ..._imports import (...)`` statement. ``from __future__ import
annotations`` stays per-port because Python future-imports are
file-scoped, not module-scoped.
"""

from collections.abc import Sequence
from typing import Protocol
from uuid import UUID

__all__ = ["Protocol", "Sequence", "UUID"]
