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
from datetime import datetime
from pathlib import Path

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0006"
down_revision: str | None = "0005"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"


def upgrade() -> None:
    fixture = (
        Path(__file__).parent.parent.parent.parent
        / "data"
        / "fixtures"
        / "lanzadera"
        / "assignments.json"
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
                # Issue #911: the column is `sa.DateTime(timezone=True)`
                # (see 0001_core_schema.py). The previous binding
                # passed the raw JSON string `row["granted_at"]` and
                # raised `DataError: invalid input for query argument:
                # ... expected a datetime.date or datetime.datetime
                # instance, got 'str'` the first time the integration
                # job ran alembic against a real Postgres service (PR-3
                # of #882, CI run 38782719315). Parse the ISO string
                # with `datetime.fromisoformat` so asyncpg receives a
                # timezone-aware `datetime`.
                "granted_at": datetime.fromisoformat(row["granted_at"]),
            },
        )


def downgrade() -> None:
    bind = op.get_bind()
    bind.execute(sa.text(f"DELETE FROM {SCHEMA}.user_app_assignments WHERE granted_by IS NULL"))
