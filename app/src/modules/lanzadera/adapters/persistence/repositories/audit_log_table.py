# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""SQLAlchemy table definition for the ``audit`` table.

W48 (#496) split this out of ``audit_log_pg.py`` so the file's
mutation-site count stays below the 100-site ceiling. The table
definition itself is the bulk of the original file; the
``AuditLogPg`` class (with the append / list_for_actor methods) is
in ``audit_log_pg.py`` and imports ``AUDIT_TABLE`` from here.

Schema layout (D27, D55):

    lanzadera.audit
    ├── id               UUID PK
    ├── event_type       text NOT NULL
    ├── actor_id         UUID NULL (system events: NULL)
    ├── target_id        text NOT NULL
    ├── module           text NOT NULL
    ├── result           text NOT NULL
    ├── correlation_id   UUID NOT NULL
    ├── payload          jsonb NOT NULL  (Postgres → Python dict on read)
    └── created_at       timestamptz NOT NULL

D55 forbids every telemetry column the legacy ``TbConexiones`` /
``TbAplicacionesAperturas`` used to carry (SSID, BSSID, machine name,
coordinates, IP). ``scripts/check_legacy_hashes.py`` rejects any
reintroduction.
"""

from __future__ import annotations

from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.dialects.postgresql import UUID as PGUUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)

AUDIT_TABLE = sa.Table(
    "audit",
    sa.MetaData(),
    sa.Column("id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("event_type", sa.Text(), nullable=False),
    sa.Column("actor_id", PGUUID(as_uuid=True), nullable=True),
    sa.Column("target_id", sa.Text(), nullable=False),
    sa.Column("module", sa.Text(), nullable=False),
    sa.Column("result", sa.Text(), nullable=False),
    sa.Column("correlation_id", PGUUID(as_uuid=True), nullable=False),
    sa.Column("payload", JSONB(astext_type=sa.Text()), nullable=False),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


__all__ = ["AUDIT_TABLE"]
