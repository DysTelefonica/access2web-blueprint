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


def _fixture_path() -> Path:
    return (
        Path(__file__).resolve().parent.parent.parent.parent
        / "data"
        / "fixtures"
        / "lanzadera"
        / "audit_events.json"
    )


def _seeded_events() -> list[dict[str, object]]:
    """The full event list the upgrade() inserts, from the same fixture.

    downgrade() matches the discriminating columns the upgrade() writes
    (id, event_type, actor_id, target_id, created_at). ``module``,
    ``result``, ``correlation_id`` and ``payload`` are intentionally
    excluded: the application also writes ``module='lanzadera'`` and the
    other fields on real events (login.py, create_user.py,
    set_password.py), so they do not discriminate a seeded row from a
    real one. Issue #928.
    """
    return json.loads(_fixture_path().read_text())


def upgrade() -> None:
    events = _seeded_events()

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
    # Issue #928: delete only the rows this migration seeded, matching
    # every column the upgrade() wrote that the application does not also
    # write (``id``, ``event_type``, ``actor_id``, ``target_id``,
    # ``created_at``). ``IS NOT DISTINCT FROM`` treats NULLs in
    # ``actor_id`` as equal to NULLs in the fixture, so a seeded event
    # with no actor is still matched.
    events = _seeded_events()
    if not events:
        return
    bind.execute(
        sa.text(
            f"""DELETE FROM {SCHEMA}.audit
                WHERE (id, event_type, actor_id, target_id, created_at)
                      IS NOT DISTINCT FROM
                      (:id, :event_type, :actor_id, :target_id, :created_at)"""
        ),
        [
            {
                "id": event["id"],
                "event_type": event["event_type"],
                "actor_id": event.get("actor_id"),
                "target_id": event["target_id"],
                # Issue #913: the column is `sa.DateTime(timezone=True)`.
                # Parse the ISO string with ``datetime.fromisoformat`` so
                # asyncpg receives a timezone-aware ``datetime``.
                "created_at": datetime.fromisoformat(event["created_at"]),
            }
            for event in events
        ],
    )
