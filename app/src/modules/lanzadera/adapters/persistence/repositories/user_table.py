# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""SQLAlchemy table definition for the ``users`` table.

W51 (#502) split this out of ``user_repository_pg.py`` so the file's
mutation-site count drops below the 100-site ceiling. The table
definition is the bulk of the original file; the ``UserRepositoryPg``
class (with the create / update_status / update_password_and_activate /
list_all / get_by_email methods) is in ``user_repository_pg.py`` and
imports ``USERS_TABLE`` from here.

Schema layout (DA-1, D88, D89):

    lanzadera.users
    ├── id              UUID PK
    ├── email           CITEXT NOT NULL  (lowercased on insert)
    ├── name            text NOT NULL
    ├── dni_encrypted   bytea NOT NULL  (PGP-encrypted, never plaintext)
    ├── password_hash   text NULL  (NULL until consume_reset_token)
    ├── status          enum user_status NOT NULL
    ├── failed_attempts int NOT NULL default 0
    ├── last_login_at   timestamptz NULL
    ├── created_at      timestamptz NOT NULL
    └── updated_at      timestamptz NOT NULL
"""

from __future__ import annotations

from sqlalchemy.dialects.postgresql import CITEXT
from sqlalchemy.dialects.postgresql import UUID as PGUUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)
from app.src.modules.lanzadera.domain.user import UserStatus

USERS_TABLE = sa.Table(
    "users",
    sa.MetaData(),
    sa.Column("id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("email", CITEXT(), nullable=False),
    sa.Column("name", sa.Text(), nullable=False),
    sa.Column("dni_encrypted", sa.LargeBinary(), nullable=False),
    sa.Column("password_hash", sa.Text(), nullable=True),
    sa.Column("status", sa.Enum(UserStatus, name="user_status")),
    sa.Column("failed_attempts", sa.Integer(), nullable=False, default=0),
    sa.Column("last_login_at", sa.DateTime(timezone=True), nullable=True),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


__all__ = ["USERS_TABLE"]
