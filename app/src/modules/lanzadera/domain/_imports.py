# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W34 — _imports
"""Shared prelude for the pure domain entities.

The five domain entities (assignment, audit_event, profile, reset_token,
session) all carry the same four-statement import block:

- ``from __future__ import annotations``
- ``from dataclasses import dataclass``
- ``from datetime import datetime``
- ``from uuid import UUID``

DRY BASELINE entry ``dup:79d8c5976f0f`` flagged the duplication as a
five-statement window with occurrences=5. W32 (#467) cleared the same
pattern for the Postgres adapters by extracting to ``_pg_imports.py``;
W34 applies the same shape to ``domain/``.

Each entity collapses to a single multi-line
``from ..._imports import (...)`` statement. ``from __future__ import
annotations`` stays per-entity because Python future-imports are
file-scoped, not module-scoped.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID

__all__ = ["UUID", "dataclass", "datetime"]
