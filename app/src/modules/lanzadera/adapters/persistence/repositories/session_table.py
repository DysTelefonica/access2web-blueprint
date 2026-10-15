# DA-1, D-W62-2 — SQLAlchemy table definition for the session rows.
"""SQLAlchemy table definition for the ``sessions`` table (W62, #921).

Schema layout (migration 0001_core_schema):

    lanzadera.sessions
    ├── id          UUID PK
    ├── user_id     UUID NOT NULL  (FK users.id, ON DELETE CASCADE)
    ├── created_at  timestamptz NOT NULL
    └── expires_at  timestamptz NOT NULL

Revocation keeps the row and shortens ``expires_at``; the audit consumer
(DA-11) reads the full lifecycle.
"""

from __future__ import annotations

from sqlalchemy.dialects.postgresql import UUID as PGUUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)

SESSIONS_TABLE = sa.Table(
    "sessions",
    sa.MetaData(),
    sa.Column("id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("user_id", PGUUID(as_uuid=True), nullable=False),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("expires_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


__all__ = ["SESSIONS_TABLE"]
