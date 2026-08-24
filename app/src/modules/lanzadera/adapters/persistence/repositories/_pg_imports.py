# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W32 — _pg_imports
"""Shared prelude for the Postgres adapters.

The W01..W07 Postgres adapters (user/app/profile/assignment/audit_log/
global_admin/reset_token) all imported the same five-symbol set from
``async_session_factory`` — ``SCHEMA`` and ``AsyncSessionFactoryPort``
— and the same four-statement SQLAlchemy + typing prelude
(``Sequence``, ``Any``, ``sa``, ``select``). Two DRY BASELINE entries
flagged the duplication:

- ``dup:3d62b08337d3`` occurrences=2 (app + profile, no ``uuid.UUID``)
- ``dup:25d63157feb1`` occurrences=3 (assignment + global_admin + user)

This module hoists the common prelude so each adapter collapses to a
single multi-line ``from _pg_imports import (...)``. The dialect
imports (CITEXT, JSONB, PGUUID) stay per-adapter; ``from __future__
import annotations`` stays per-adapter because Python's future-imports
are file-scoped, not module-scoped.

W23 (#452) pinned the runtime prelude-cleave (no manual
``session: AsyncSession = self._factory()``); W32 finishes the
import-side cleanup that W08..W20 deferred.
"""

from collections.abc import Sequence
from typing import Any

import sqlalchemy as sa
from sqlalchemy import select

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    SCHEMA,
    AsyncSessionFactoryPort,
)

__all__ = [
    "Any",
    "AsyncSessionFactoryPort",
    "SCHEMA",
    "Sequence",
    "sa",
    "select",
]
