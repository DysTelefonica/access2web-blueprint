# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #620
"""seed assignments: insert 622 rows from data/fixtures/lanzadera/assignments.json.

Revision ID: 0005
Revises: 0004
Create Date: 2026-09-16

DA-12: the fixture was generated applying the SinAcceso exclusive rule
(resolve_legacy_roles cortocircuito). Each row carries the pre-resolved
``profile_id`` — no further mapping needed here.

Cardinality target: exactly 622 active rows.
Idempotent: deletes rows seeded from the fixture before re-inserting
(so it is safe to re-run after a downgrade of 0004)."""

from __future__ import annotations

import json
from collections.abc import Sequence
from pathlib import Path

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0005"
down_revision: str | None = "0004"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"


def upgrade() -> None:
    fixture = (
        Path(__file__).parent.parent.parent.parent
        / "data" / "fixtures" / "lanzadera" / "assignments.json"
    )
    assignments = json.loads(fixture.read_text())

    bind = op.get_bind()

    # Delete any rows we may have seeded before (safe to re-run).
    bind.execute(
        sa.text(
            f"DELETE FROM {SCHEMA}.user_app_assignments"
            " WHERE revoked_at IS NULL"
            " AND granted_by IS NULL"
        )
    )

    for row in assignments:
        if row.get("revoked_at") is not None:
            continue
        bind.execute(
            sa.text(
                f"""INSERT INTO {SCHEMA}.user_app_assignments
                    (id, user_id, app_id, profile_id, granted_by, granted_at, revoked_at)
                    VALUES
                    (:id, :user_id, :app_id, :profile_id, NULL, :granted_at, NULL)"""
            ),
            {
                "id": row["id"],
                "user_id": row["user_id"],
                "app_id": row["app_id"],
                "profile_id": row["profile_id"],
                "granted_at": row["granted_at"],
            },
        )


def downgrade() -> None:
    bind = op.get_bind()
    bind.execute(
        sa.text(
            f"DELETE FROM {SCHEMA}.user_app_assignments"
            " WHERE granted_by IS NULL"
        )
    )
