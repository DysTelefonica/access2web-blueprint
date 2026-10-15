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

# Row-count expectations derived from the fixture file (deterministic seed=42).
# Keep these in sync with the JSON fixture so downgrade() can target exact
# ids and never touch real audit rows. See #928.
EXPECTED_AUDIT_EVENTS = 200


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

    downgrade() uses this list to delete exactly the seeded rows and
    nothing else. Issue #928: prior heuristic (``WHERE actor_id IS NULL``)
    wiped unrelated audit rows with a null actor and could not remove
    seeded rows with a non-null actor (which in turn blocked the 0005
    downgrade via ``fk_audit_actor_id_users``).
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
                    (:id, :event_type, :actor_id, :target_id, 'lanzadera',
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
    # Issue #928: delete only the rows this migration seeded, by id.
    # ``actor_id IS NULL`` is a heuristic that wiped real audit rows with
    # a null actor and left seeded rows with a non-null actor blocking the
    # 0005 -> 0004 downgrade via ``fk_audit_actor_id_users``.
    seeded_ids = _seeded_event_ids()
    if not seeded_ids:
        return
    bind.execute(
        sa.text(f"DELETE FROM {SCHEMA}.audit WHERE id = ANY(:ids)"),
        {"ids": seeded_ids},
    )
