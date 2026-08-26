# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""SQLAlchemy table definition for the ``profiles`` table.

W49 (#498) split this out of ``profile_repository_pg.py`` so the
file's mutation-site count stays below the 100-site ceiling. The
table definition itself is the bulk of the original file; the
``ProfileRepositoryPg`` class (with the create / list_for_app /
get_by_code / set_active / update_capabilities methods) is in
``profile_repository_pg.py`` and imports ``PROFILES_TABLE`` from here.

Schema layout (DA-12, D22, D45, D46):

    lanzadera.profiles
    ├── id           UUID PK
    ├── app_id       int NOT NULL  (FK apps.id)
    ├── code         text NOT NULL
    ├── name         text NOT NULL
    ├── capabilities jsonb NOT NULL  (str | int | bool leaves)
    ├── active       bool NOT NULL default TRUE
    ├── created_at   timestamptz NOT NULL
    └── updated_at   timestamptz NOT NULL

UNIQUE (app_id, code) constraint is enforced at the adapter level
via ``ProfileRepositoryPg.create`` (DA-12: serialises the
capabilities map atomically per call so concurrent creates cannot
race on the (app_id, code) pair).
"""

from __future__ import annotations

from sqlalchemy.dialects.postgresql import JSONB, UUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)

PROFILES_TABLE = sa.Table(
    "profiles",
    sa.MetaData(),
    sa.Column("id", UUID(as_uuid=True), primary_key=True),
    sa.Column("app_id", sa.Integer, nullable=False),
    sa.Column("code", sa.Text(), nullable=False),
    sa.Column("name", sa.Text(), nullable=False),
    sa.Column("capabilities", JSONB(astext_type=sa.Text()), nullable=False),
    sa.Column("active", sa.Boolean, nullable=False, default=True),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


__all__ = ["PROFILES_TABLE"]
