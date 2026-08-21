"""create base schema (expedientes): shared catalogs.

Revision ID: 0002
Revises: 0001
Create Date: 2026-08-20

Schema layout (F04, issue #225) — the ``app/src/modules/expedientes/``
module's initial DDL. This migration lands the **shared catalogs only**:
the lookup tables that the lifecycle, related-data and catalogs
verticals all reference as foreign keys. The 37 vertical-specific
tables (TbExpedientes, TbExpedientesAnexos, TbExpedientesAnualidades,
TbExpedientesComerciales, TbExpedientesCPVs, TbExpedientesE2E,
TbExpedientesHitos, TbExpedientesJefaturas, TbExpedientesJuridicas,
TbExpedientesLugaresEjecucion, TbExpedientesModificados, TbExpedientesPECAL,
TbExpedientesRACS, TbExpedientesResponsables, TbExpedientesSuministradores,
TbDatosEconomicosExpedientes, TbCambios, TbUltimoCambio,
TbConfMostrarEstado, TbE2EExportBatch, TbE2EExportBatchDetalle,
TbE2EExportSeleccionTemp, TbE2EJsonDestinationUserConfig, etc.)
land in their own migrations (0003..0049, one per vertical) — this
WU delivers the foundation (UoW + ENUMs + lookup catalogs) so the
verticals can run.

The full DDL for all 49 tables is in
``docs/03-aplicaciones/expedientes/ERD/schema.sql`` (PR #393); this
migration is a deliberately small, focused slice of that DDL. The
``downgrade()`` reverses the slice in one transaction so the schema can
be re-rolled cleanly during the UAT phase (D82, Expand and Contract).
"""

from __future__ import annotations

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "0002"
down_revision: str | None = "0001"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


# ---------------------------------------------------------------------------
# Named schema object helpers.
#
# Same pattern as 0001_core_schema.py for ``lanzadera``: a single constant
# so a typo in a future table is caught at module load, not at migration.
# ---------------------------------------------------------------------------
SCHEMA = "expedientes"

# ---------------------------------------------------------------------------
# ENUMs. ``create_type=False`` lets Alembic own the type lifecycle; the
# downgrade drops them alongside the schema. The vertical-specific
# migrations register their own ENUMs (F02, F03, ...) without touching
# the shared catalog types below.
# ---------------------------------------------------------------------------
# No shared ENUMs in the catalog slice — each catalog table uses ``VARCHAR``
# (matching the legacy ``TEXT(255)`` columns in ``schema.sql``). ENUMs
# are added by the verticals that need a closed set (e.g. F04/CAP-001
# lifecycle adds ``expediente_status`` and ``expediente_type``).
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Tables — the 12 lookup catalogs that the vertical chain (C01..C04,
# R01..R07, Q01..Q02, M01..M04) references as foreign keys.
#
# Each catalog maps 1-to-1 to a ``Tb*`` table in the legacy
# ``Expedientes_datos.accdb``. The column types come from
# ``docs/03-aplicaciones/expedientes/ERD/schema.sql`` and follow the
# PostgreSQL mapping rules in the design (long PK → ``BIGSERIAL``,
# text → ``VARCHAR(N) NOT NULL`` / ``TEXT NULL``, memo → ``TEXT NULL``).
# The ID columns keep the legacy names (e.g. ``IDEstado``) so the
# downstream verticals can reference them as-is.
# ---------------------------------------------------------------------------
CATALOG_TABLES: tuple[sa.Table, ...] = (
    sa.Table(
        "TbEstados",
        sa.MetaData(),
        sa.Column("IDEstado", sa.BigInteger, primary_key=True),
        sa.Column("Estado", sa.String(255), nullable=False),
        sa.Column("DESCRIPCION", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbGradosClasificacion",
        sa.MetaData(),
        sa.Column("IDGradoClasificacion", sa.BigInteger, primary_key=True),
        sa.Column("Grado", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbComerciales",
        sa.MetaData(),
        sa.Column("IDComercial", sa.BigInteger, primary_key=True),
        sa.Column("Comercial", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbCPV",
        sa.MetaData(),
        sa.Column("IDCPV", sa.BigInteger, primary_key=True),
        sa.Column("CPV", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbLugaresEjecucion",
        sa.MetaData(),
        sa.Column("IDLugarEjecucion", sa.BigInteger, primary_key=True),
        sa.Column("LugarEjecucion", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbOrganosContratacion",
        sa.MetaData(),
        sa.Column("IDOrganoContratacion", sa.BigInteger, primary_key=True),
        sa.Column("OrganoContratacion", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbOficinasPrograma",
        sa.MetaData(),
        sa.Column("IDOficinaPrograma", sa.BigInteger, primary_key=True),
        sa.Column("OficinaPrograma", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbEjercitos",
        sa.MetaData(),
        sa.Column("IDEjercito", sa.BigInteger, primary_key=True),
        sa.Column("Ejercito", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbJefaturas",
        sa.MetaData(),
        sa.Column("IDJefatura", sa.BigInteger, primary_key=True),
        sa.Column("Jefatura", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbJuridicas",
        sa.MetaData(),
        sa.Column("IDJuridica", sa.BigInteger, primary_key=True),
        sa.Column("Juridica", sa.String(255), nullable=False),
        sa.Column("DESCRIPCION", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbResponsablesPorRol",
        sa.MetaData(),
        sa.Column("IDResponsable", sa.BigInteger, primary_key=True),
        sa.Column("Responsable", sa.String(255), nullable=False),
        sa.Column("Rol", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbSuministradores",
        sa.MetaData(),
        sa.Column("IDSuministrador", sa.BigInteger, primary_key=True),
        sa.Column("Suministrador", sa.String(255), nullable=False),
        sa.Column("NIF", sa.String(20), nullable=True),
        sa.Column("CORREO", sa.String(255), nullable=True),
        sa.Column("DESCRIPCION", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbPECAL",
        sa.MetaData(),
        sa.Column("IDPECAL", sa.BigInteger, primary_key=True),
        sa.Column("PECAL", sa.String(255), nullable=False),
        sa.Column("Descripcion", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
    sa.Table(
        "TbRACS",
        sa.MetaData(),
        sa.Column("IDRAC", sa.BigInteger, primary_key=True),
        sa.Column("RAC", sa.String(255), nullable=False),
        sa.Column("CORREO", sa.String(255), nullable=True),
        sa.Column("DESCRIPCION", sa.Text, nullable=True),
        schema=SCHEMA,
    ),
)


def upgrade() -> None:
    """Create the ``expedientes`` schema + the 13 shared catalog tables.

    The schema is created in this migration so a parallel ``downgrade()``
    drops the whole subtree (D82, Expand and Contract). The verticals
    add their own tables (expediente, anexos, anualidades, …) in
    follow-up migrations that all declare ``down_revision = "0002"``.
    """
    op.execute(f'CREATE SCHEMA IF NOT EXISTS "{SCHEMA}"')
    for table in CATALOG_TABLES:
        op.create_table(table.name, *table.columns, schema=SCHEMA)


def downgrade() -> None:
    """Drop the catalog tables and the schema in one transaction."""
    for table in reversed(CATALOG_TABLES):
        op.drop_table(table.name, schema=SCHEMA)
    op.execute(f'DROP SCHEMA IF EXISTS "{SCHEMA}"')
