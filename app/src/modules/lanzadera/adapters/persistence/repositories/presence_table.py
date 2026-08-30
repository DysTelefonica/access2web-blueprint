# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — presence table definition.
"""SQLAlchemy table definition for the ``presence_sessions`` table.

W60 (#522) splits this out of ``presence_repository_pg.py`` so the
adapter file's mutation-site count stays below the 100-site ceiling —
the pattern is the same as ``user_table.py`` (W51, #502) and
``app_table.py`` (W52, #504). The adapter's four methods
(``track`` / ``heartbeat`` / ``disconnect`` / ``list_connected``) live
in ``presence_repository_pg.py`` and import ``PRESENCE_TABLE`` from
here.

Schema layout (W60, #522):

    lanzadera.presence_sessions
    ├── user_id        UUID PK          (no FK — sessions outlive a user row)
    ├── email          text NOT NULL    (cached at track-time)
    ├── connected_at   timestamptz NOT NULL  (set on first INSERT, untouched on conflict)
    └── last_seen      timestamptz NOT NULL  (bumped on every track/heartbeat)

The FK to ``users.id`` is deliberately omitted: a presence row records
"this user was here at that time", and the row should survive the user
removal so the audit viewer still sees the historical heartbeat. The
``user_id`` PK guarantees a single row per user (the ``ON CONFLICT``
clause in the adapter relies on it).
"""

from __future__ import annotations

from sqlalchemy.dialects.postgresql import UUID as PGUUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)

PRESENCE_TABLE = sa.Table(
    "presence_sessions",
    sa.MetaData(),
    sa.Column("user_id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("email", sa.Text(), nullable=False),
    sa.Column("connected_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("last_seen", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


__all__ = ["PRESENCE_TABLE"]
