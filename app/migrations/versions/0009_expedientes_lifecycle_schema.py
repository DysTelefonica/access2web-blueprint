"""create the Expedientes lifecycle schema (TbExpedientes, TbCambios, TbUltimoCambio).

Issue #801 (hole 2 of the #802 epic) and #803: migration ``0002`` opened the
``expedientes`` schema with the fourteen shared catalogs and sent the vertical
tables to ``0003..0049``, which never landed. The lifecycle use cases
(C01-C04) write three of those tables — the header, the change history and
the last-change marker — and the eighteen ports of the module have nowhere to
persist without them.

The layout is pinned by the code contract, not by the Access dump:

- ``TbExpedientes`` — the aggregate header and its invariants
  (``app/src/modules/expedientes/domain/expediente.py``): identity, type,
  state, optimistic-locking version, parent link, contract dates and
  timestamps.
- ``TbCambios`` — one row per mutation, the value object
  ``ChangeRecord`` of ``app/src/modules/expedientes/ports/audit_log.py``.
- ``TbUltimoCambio`` — the last-change marker written in the same unit of
  work as the header (``create_expediente/service.py``).

Key decision (owner, 2026-10-24): identifiers are UUID, like the deployed
modern schema of ``0001`` (``users``, ``sessions``, ``audit``) and like the
eighteen ports of the module. The ``BIGSERIAL``/``ID*`` criterion of #803
applies to the legacy catalogs of ``0002``, which this revision does not
touch.

No foreign keys: the schema of ``0002`` declares none, and an audit row must
survive the removal of its expediente (CAP-003, D82: expand and contract).

The ENUM types are created explicitly and dropped in the downgrade, so
``downgrade -1`` leaves no residue behind.
"""

from __future__ import annotations

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op
from sqlalchemy.dialects import postgresql

# revision identifiers, used by Alembic.
revision: str = "0009"
down_revision: str | None = "0008"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "expedientes"

# The ENUM literals mirror the domain enums `ExpedienteTipo` and
# `ExpedienteEstado` (`domain/expediente_tipo.py`,
# `domain/expediente_estado.py`). `create_type=False` keeps Alembic from
# emitting a second `CREATE TYPE` for a type this migration already owns.
EXPEDIENTE_TIPO = postgresql.ENUM(
    "AM",
    "LOTE",
    "BASED",
    name="expediente_tipo",
    schema=SCHEMA,
    create_type=False,
)

EXPEDIENTE_ESTADO = postgresql.ENUM(
    "BORRADOR",
    "ADJUDICADO",
    "FORMALIZADO",
    "ARCHIVADO",
    name="expediente_estado",
    schema=SCHEMA,
    create_type=False,
)


def upgrade() -> None:
    """Create the three tables the lifecycle use cases write."""
    # ENUMs first — the table columns reference them.
    EXPEDIENTE_TIPO.create(op.get_bind(), checkfirst=True)
    EXPEDIENTE_ESTADO.create(op.get_bind(), checkfirst=True)

    op.create_table(
        "TbExpedientes",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("tipo", EXPEDIENTE_TIPO, nullable=False),
        sa.Column("estado", EXPEDIENTE_ESTADO, nullable=False),
        sa.Column("version", sa.Integer(), nullable=False),
        sa.Column("id_expediente_padre", postgresql.UUID(as_uuid=True), nullable=True),
        sa.Column("fecha_inicio_contrato", sa.Date(), nullable=True),
        sa.Column("fecha_fin_contrato", sa.Date(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        schema=SCHEMA,
    )

    op.create_table(
        "TbCambios",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("nombre_tabla", sa.Text(), nullable=False),
        sa.Column("id_expediente", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("nombre_campo", sa.Text(), nullable=True),
        sa.Column("valor_inicial", sa.Text(), nullable=True),
        sa.Column("valor_final", sa.Text(), nullable=True),
        sa.Column("fecha_cambio", sa.DateTime(timezone=True), nullable=False),
        sa.Column("id_usuario_cambio", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("accion", sa.Text(), nullable=False),
        schema=SCHEMA,
    )

    op.create_table(
        "TbUltimoCambio",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("id_expediente", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("fecha_cambio", sa.DateTime(timezone=True), nullable=False),
        sa.Column("id_usuario_cambio", postgresql.UUID(as_uuid=True), nullable=False),
        schema=SCHEMA,
    )


def downgrade() -> None:
    """Drop the three tables and their ENUM types.

    The fourteen catalogs of ``0002`` stay: they belong to another revision
    and the schema itself survives.
    """
    op.drop_table("TbUltimoCambio", schema=SCHEMA)
    op.drop_table("TbCambios", schema=SCHEMA)
    op.drop_table("TbExpedientes", schema=SCHEMA)
    EXPEDIENTE_ESTADO.drop(op.get_bind(), checkfirst=True)
    EXPEDIENTE_TIPO.drop(op.get_bind(), checkfirst=True)
