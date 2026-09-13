# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #621
"""seed apps: insert 20 rows from data/fixtures/lanzadera/apps.json.

Revision ID: 0003
Revises: 0002
Create Date: 2026-09-16

DA-7: fixtures derivadas del walkthrough. Sin Dysflow, sin pyodbc.
G-3: campos NO migrados: Pass, Comando, URLDIrectorioIconoAplicacion,
  NombreEjecutable, NombreArchivoDatos, NombreCarpeta, NombreCarpetaTemporal,
  TituloAplicacion, NombreIconoParaArbol, NombreIcono, NombreIconoLanzadera,
  NombreCarpetaDocumentacion, NombreDirectorioIconos, NombreDirectorioAyuda,
  NombreDirectorioRecursos, EnPruebas, ConIconoEnLanzadera,
  NombreFuncionPublicacion. La decisión de qué se migra vive en
  apps/spec.md §Seed 20 apps.

D52, D53, D85, D58: las 20 apps se siembran con
  registration_status='active'; deployment_topology y
  requires_office_presence derivan de EjecucionEnOficina en la fixture."""

from __future__ import annotations

import json
from collections.abc import Sequence
from pathlib import Path

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0003"
down_revision: str | None = "0002"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"


def upgrade() -> None:
    fixture = (
        Path(__file__).parent.parent.parent.parent
        / "data" / "fixtures" / "lanzadera" / "apps.json"
    )
    apps = json.loads(fixture.read_text())

    bind = op.get_bind()
    count = (
        bind.execute(sa.text(f"SELECT COUNT(*) FROM {SCHEMA}.apps"))
        .scalar()
        or 0
    )
    if count >= 20:
        return

    for app in apps:
        bind.execute(
            sa.text(
                f"""INSERT INTO {SCHEMA}.apps
                    (id, name, short_code, deployment_topology, requires_office_presence,
                     registration_status, created_at, updated_at)
                    VALUES
                    (:id, :name, :short_code, :deployment_topology, :requires_office_presence,
                     'active'::{SCHEMA}.app_registration_status, now(), now())"""
            ),
            {
                "id": app["id"],
                "name": app["name"],
                "short_code": app["short_code"],
                "deployment_topology": app["deployment_topology"],
                "requires_office_presence": app["requires_office_presence"],
            },
        )


def downgrade() -> None:
    bind = op.get_bind()
    bind.execute(sa.text(f"DELETE FROM {SCHEMA}.apps WHERE id <= 20"))
