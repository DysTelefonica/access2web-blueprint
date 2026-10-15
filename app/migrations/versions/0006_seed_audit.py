# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #621
"""seed audit: insert ~200 sample events from data/fixtures/lanzadera/audit_events.json.

Revision ID: 0006
Revises: 0005
Create Date: 2026-09-16

DA-11: ``TbConexiones`` → ``auth.login.success`` / ``auth.login.failure``.
DA-11: ``TbAplicacionesAperturas`` → ``app.open``.
Sin columnas de telemetría (SSID, BSSID, coordenadas, máquina, IP) — DA-11
retira esa información del schema de auditoría."""

from __future__ import annotations

import json
from collections.abc import Sequence
from datetime import datetime
from pathlib import Path

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0007"
down_revision: str | None = "0006"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"

# Module value the upgrade() forces on every seeded row. downgrade()
# uses it as part of the identity check so a real row whose id happens
# to collide with a seeded one is not deleted. See #928.
SEEDED_MODULE = "lanzadera"


def _fixture_path() -> Path:
    return (
        Path(__file__).resolve().parent.parent.parent.parent
        / "data"
        / "fixtures"
        / "lanzadera"
        / "audit_events.json"
    )


def _seeded_event_ids() -> list[int]:
    """Ids the upgrade() inserts, taken from the same fixture source.

    downgrade() uses this list together with ``SEEDED_MODULE`` to delete
    exactly the seeded rows and nothing else. Issue #928: prior heuristic
    (``WHERE actor_id IS NULL``) wiped unrelated audit rows with a null
    actor and could not remove seeded rows with a non-null actor (which
    in turn blocked the 0005 downgrade via ``fk_audit_actor_id_users``).
    """
    return [event["id"] for event in json.loads(_fixture_path().read_text())]


def upgrade() -> None:
    events = json.loads(_fixture_path().read_text())

    bind = op.get_bind()
    count = bind.execute(sa.text(f"SELECT COUNT(*) FROM {SCHEMA}.audit")).scalar() or 0
    if count >= len(events):
        return

    for event in events:
        bind.execute(
            sa.text(
                f"""INSERT INTO {SCHEMA}.audit
                    (id, event_type, actor_id, target_id, module,
                     result, correlation_id, payload, created_at)
                    VALUES
                    (:id, :event_type, :actor_id, :target_id, '{SEEDED_MODULE}',
                     :result, :correlation_id, :payload, :created_at)"""
            ),
            {
                "id": event["id"],
                "event_type": event["event_type"],
                "actor_id": event.get("actor_id"),
                "target_id": event["target_id"],
                "result": event["result"],
                "correlation_id": event["correlation_id"],
                "payload": json.dumps(event.get("payload", {})),
                # Issue #913: the column is `sa.DateTime(timezone=True)`
                # (see 0001_core_schema.py). The previous binding passed
                # the raw JSON string `event["created_at"]` and raised
                # `DataError: invalid input for query argument: ...
                # expected a datetime.date or datetime.datetime
                # instance, got 'str'` the first time the integration
                # job ran alembic against a real Postgres service (PR-3
                # of #882, CI run 38787186647, after #912 fixed 0005).
                # Parse the ISO string with `datetime.fromisoformat` so
                # asyncpg receives a timezone-aware ``datetime``.
                "created_at": datetime.fromisoformat(event["created_at"]),
            },
        )


def downgrade() -> None:
    bind = op.get_bind()
    # Issue #928: delete only the rows this migration seeded. The id alone
    # is not enough — ``upgrade()`` is a no-op when ``audit`` already
    # holds at least ``len(events)`` rows, so a real row whose id happens
    # to collide with a seeded one would otherwise be deleted here.
    # ``SEEDED_MODULE`` is the literal the upgrade() forces on every
    # inserted row, so ``(id, module)`` uniquely identifies a seeded row
    # among the rows the upgrade actually wrote.
    seeded_ids = _seeded_event_ids()
    if not seeded_ids:
        return
    bind.execute(
        sa.text(f"DELETE FROM {SCHEMA}.audit WHERE id = ANY(:ids) AND module = :module"),
        {"ids": seeded_ids, "module": SEEDED_MODULE},
    )
