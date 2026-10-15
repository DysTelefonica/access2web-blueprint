# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #771
"""create presence_sessions table.

Issue #771: the presence feature (W60, #522) ships the
``PRESENCE_TABLE`` definition in
``app/src/modules/lanzadera/adapters/persistence/repositories/presence_table.py``
and four adapter methods in ``presence_repository_pg.py`` (``track``,
``heartbeat``, ``disconnect``, ``list_connected``), but no migration
creates the underlying ``lanzadera.presence_sessions`` table. A clean
database crashes on the first ``presence_repository_pg.list_connected()``
call::

    sqlalchemy.exc.ProgrammingError: relation
    "lanzadera.presence_sessions" does not exist

This migration creates the table with the layout pinned by
``presence_table.py`` byte-for-byte (W60, #522):

    lanzadera.presence_sessions
    ├── user_id        UUID PK          (no FK — sessions outlive a user row)
    ├── email          text NOT NULL    (cached at track-time)
    ├── connected_at   timestamptz NOT NULL  (set on first INSERT,
    │                                         untouched on conflict)
    └── last_seen      timestamptz NOT NULL  (bumped on every
                                              track/heartbeat)

The FK to ``users.id`` is deliberately omitted: a presence row records
"this user was here at that time", and the row should survive the user
removal so the audit viewer still sees the historical heartbeat. The
``user_id`` PK guarantees a single row per user (the ``ON CONFLICT``
clause in the adapter relies on it).
"""

from __future__ import annotations

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op
from sqlalchemy.dialects import postgresql

# revision identifiers, used by Alembic.
revision: str = "0008"
down_revision: str | None = "0007"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

SCHEMA = "lanzadera"


def upgrade() -> None:
    """Create ``lanzadera.presence_sessions`` with the pinned layout.

    The migration is idempotent: ``op.create_table`` raises if the
    table already exists, so a replayed ``upgrade head`` is a no-op for
    this revision.
    """
    op.create_table(
        "presence_sessions",
        sa.Column(
            "user_id",
            postgresql.UUID(as_uuid=True),
            primary_key=True,
        ),
        sa.Column(
            "email",
            sa.Text(),
            nullable=False,
        ),
        sa.Column(
            "connected_at",
            sa.DateTime(timezone=True),
            nullable=False,
        ),
        sa.Column(
            "last_seen",
            sa.DateTime(timezone=True),
            nullable=False,
        ),
        schema=SCHEMA,
    )


def downgrade() -> None:
    """Drop ``lanzadera.presence_sessions``.

    No data-preservation branch: the table only carries ephemeral
    presence state (no audit trail; ``audit`` covers that), and the
    down-revision has to be replayable to a clean DB.
    """
    op.drop_table("presence_sessions", schema=SCHEMA)
