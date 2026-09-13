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
from pathlib import Path

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0007"
down_revision: str | None = "0006"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"


def upgrade() -> None:
    fixture = (
        Path(__file__).parent.parent.parent.parent
        / "data" / "fixtures" / "lanzadera" / "audit_events.json"
    )
    events = json.loads(fixture.read_text())

    bind = op.get_bind()
    count = (
        bind.execute(sa.text(f"SELECT COUNT(*) FROM {SCHEMA}.audit"))
        .scalar()
        or 0
    )
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
                "created_at": event["created_at"],
            },
        )


def downgrade() -> None:
    bind = op.get_bind()
    # Delete seeded events (those with module='lanzadera' and no actor).
    bind.execute(
        sa.text(f"DELETE FROM {SCHEMA}.audit WHERE actor_id IS NULL")
    )
