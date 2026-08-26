# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""SQLAlchemy table definition for the ``apps`` table.

W52 (#504) split this out of ``app_repository_pg.py`` so the file's
mutation-site count drops below the 100-site ceiling. The table
definition is the bulk of the original file; the ``AppRepositoryPg``
class (with the list_active / list_visible_to / activate / create
methods) is in ``app_repository_pg.py`` and imports ``APPS_TABLE``
from here.

Schema layout (DA-7, D52, D85):

    lanzadera.apps
    ├── id                       SERIAL PK
    ├── name                     text NOT NULL
    ├── short_code               text NOT NULL
    ├── deployment_topology      enum app_topology
    ├── requires_office_presence bool NOT NULL default FALSE
    ├── registration_status      enum app_registration_status
    ├── created_at               timestamptz NOT NULL
    └── updated_at               timestamptz NOT NULL

UNIQUE (short_code) is enforced at the application layer (D52). The
launcher fields the legacy ``TbAplicaciones`` used to carry
(``Pass``, ``Comando``, ``URLDIrectorioIconoAplicacion``) are NOT
migrated — they belong to the runtime launch system, retired by D52.
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)
from app.src.modules.lanzadera.domain.app import (
    AppRegistrationStatus,
    AppTopology,
)

APPS_TABLE = sa.Table(
    "apps",
    sa.MetaData(),
    sa.Column("id", sa.Integer, sa.Identity(always=False), primary_key=True),
    sa.Column("name", sa.Text(), nullable=False),
    sa.Column("short_code", sa.Text(), nullable=False),
    sa.Column("deployment_topology", sa.Enum(AppTopology, name="app_topology")),
    sa.Column("requires_office_presence", sa.Boolean, nullable=False, default=False),
    sa.Column(
        "registration_status",
        sa.Enum(AppRegistrationStatus, name="app_registration_status"),
    ),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


__all__ = ["APPS_TABLE"]
