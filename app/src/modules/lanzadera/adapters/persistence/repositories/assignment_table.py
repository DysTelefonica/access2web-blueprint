# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""SQLAlchemy table definitions for the ``assignment`` repository.

W53 (#506) split this out of ``assignment_repository_pg.py`` so the
file's mutation-site count drops below the 100-site ceiling. The
``assignment_repository_pg.py`` file declared two sa.Tables:
``USER_APP_ASSIGNMENTS_TABLE`` (the assignment join) and a local
``PROFILES_TABLE`` duplicate. Both are hoisted here; the local
``PROFILES_TABLE`` is replaced with a re-export of the canonical
``PROFILES_TABLE`` from ``profile_table.py``.

Schema layout (DA-12, D22, D42, H11):

    lanzadera.user_app_assignments
    ├── id          UUID PK
    ├── user_id     UUID NOT NULL  (FK users.id)
    ├── app_id      int NOT NULL   (FK apps.id)
    ├── profile_id  UUID NOT NULL  (FK profiles.id)
    ├── granted_by  UUID NULL      (FK users.id)
    ├── granted_at  timestamptz NOT NULL
    └── revoked_at  timestamptz NULL

D42: revocation is a soft-delete (``revoked_at`` is set; the row stays
so audit reads see the full lifecycle).
"""

from __future__ import annotations

from sqlalchemy.dialects.postgresql import UUID as PGUUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.profile_table import (
    PROFILES_TABLE,
)

USER_APP_ASSIGNMENTS_TABLE = sa.Table(
    "user_app_assignments",
    sa.MetaData(),
    sa.Column("id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("user_id", PGUUID(as_uuid=True), nullable=False),
    sa.Column("app_id", sa.Integer, nullable=False),
    sa.Column("profile_id", PGUUID(as_uuid=True), nullable=False),
    sa.Column("granted_by", PGUUID(as_uuid=True), nullable=True),
    sa.Column("granted_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("revoked_at", sa.DateTime(timezone=True), nullable=True),
    schema=SCHEMA,
)


__all__ = ["PROFILES_TABLE", "USER_APP_ASSIGNMENTS_TABLE"]