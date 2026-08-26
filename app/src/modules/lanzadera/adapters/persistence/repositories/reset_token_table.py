# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""SQLAlchemy table definition for the ``reset_tokens`` table.

W50 (#500) split this out of ``reset_token_repository_pg.py`` so the
file's mutation-site count drops below the 100-site ceiling. The
table definition is the bulk of the original file; the
``ResetTokenRepositoryPg`` class (with the insert / find_unused /
mark_consumed / mark_superseded / purge_expired methods) is in
``reset_token_repository_pg.py`` and imports ``RESET_TOKENS_TABLE``
from here.

Schema layout (DA-4, D90):

    lanzadera.reset_tokens
    ├── id             UUID PK
    ├── user_id        UUID NOT NULL  (FK users.id)
    ├── token_hash     text NOT NULL  (UNIQUE)
    ├── expires_at     timestamptz NOT NULL
    ├── consumed_at    timestamptz NULL
    ├── superseded_at  timestamptz NULL
    └── created_at     timestamptz NOT NULL

The D90 supersession rule is enforced at the application layer
(``mark_superseded`` stamps every live row for a user_id).
"""

from __future__ import annotations

from sqlalchemy.dialects.postgresql import UUID as PGUUID

from app.src.modules.lanzadera.adapters.persistence.repositories._pg_imports import (
    SCHEMA,
    sa,
)

RESET_TOKENS_TABLE = sa.Table(
    "reset_tokens",
    sa.MetaData(),
    sa.Column("id", PGUUID(as_uuid=True), primary_key=True),
    sa.Column("user_id", PGUUID(as_uuid=True), nullable=False),
    sa.Column("token_hash", sa.Text(), nullable=False),
    sa.Column("expires_at", sa.DateTime(timezone=True), nullable=False),
    sa.Column("consumed_at", sa.DateTime(timezone=True), nullable=True),
    sa.Column("superseded_at", sa.DateTime(timezone=True), nullable=True),
    sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    schema=SCHEMA,
)


__all__ = ["RESET_TOKENS_TABLE"]
