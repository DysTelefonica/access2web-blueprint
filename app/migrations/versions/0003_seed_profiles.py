# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #619
"""seed profiles: insert ≥20 rows from data/fixtures/lanzadera/profiles.json.

Revision ID: 0003
Revises: 0002
Create Date: 2026-09-16

Cada app recibe:
  - 1 fila con ``code='DEFAULT'`` (usuario sin rol específico)
  - 7 filas con los legacy codes (ADMIN, CALIDAD, CALIDAD_AVISOS,
    TECNICO, ECONOMIA, SECRETARIA, SIN_ACCESO)

G-2 (ABIERTO): ``profiles.capabilities`` se siembra vacío ``{}``.
El set canónico de capabilities por app se decide después; la estructura
JSONB está lista para recibirlo en una release posterior.

DA-1: ``migrations/`` es driven; no importa ``application/`` ni ``delivery/``.
DA-12: ``legacy_role_map`` es dominio puro y se puede importar aquí.
"""

from __future__ import annotations

import json
from collections.abc import Sequence
from pathlib import Path

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0004"
down_revision: str | None = "0003"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"


def upgrade() -> None:
    fixture = (
        Path(__file__).parent.parent.parent.parent
        / "data" / "fixtures" / "lanzadera" / "profiles.json"
    )
    profiles = json.loads(fixture.read_text())

    bind = op.get_bind()
    count = (
        bind.execute(sa.text(f"SELECT COUNT(*) FROM {SCHEMA}.profiles"))
        .scalar()
        or 0
    )
    if count >= len(profiles):
        # Idempotent: already seeded.
        return

    for profile in profiles:
        bind.execute(
            sa.text(
                f"""INSERT INTO {SCHEMA}.profiles
                    (id, app_id, code, name, capabilities, active, created_at, updated_at)
                    VALUES
                    (:id, :app_id, :code, :name, :capabilities, 'active', now(), now())"""
            ),
            {
                "id": profile["id"],
                "app_id": profile["app_id"],
                "code": profile["code"],
                "name": profile["name"],
                # G-2 ABIERTO: capabilities provisional {}.
                "capabilities": json.dumps(profile.get("capabilities", {})),
            },
        )


def downgrade() -> None:
    bind = op.get_bind()
    # Remove all legacy codes (not DEFAULT, which may be used in production later).
    bind.execute(
        sa.text(
            f"DELETE FROM {SCHEMA}.profiles WHERE code != 'DEFAULT'"
        )
    )
