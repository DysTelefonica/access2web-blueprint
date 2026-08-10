# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 3a
# D5, D14, D27, D56, D82, D89, DA-1, DA-3, DA-4, DA-7, DA-11, DA-12.
"""create core schema (lanzadera): 9 tables, 3 enums, 6 indexes.

Revision ID: 0001
Revises:
Create Date: 2026-08-09

Schema layout (design.md §Modelo de datos):

    lanzadera
    ├── users                    (UUID PK, CITEXT email, password_hash NULL)
    ├── apps                     (SERIAL PK, short_code UNIQUE)
    ├── profiles                 (UUID PK, FK apps, UNIQUE(app_id, code))
    ├── user_app_assignments     (UUID PK, FK users/apps/profiles)
    ├── global_admins            (UUID PK = users.id, FK users)
    ├── sessions                 (UUID PK, FK users)
    ├── audit                    (UUID PK, NO telemetry columns DA-11)
    ├── reset_tokens             (UUID PK, FK users, token_hash UNIQUE)
    └── mail_outbox              (UUID PK, status default 'pending')

ENUMs (`user_status`, `app_topology`, `app_registration_status`) are
created once and reused across every future migration. Indexes land after
their tables; the partial unique on `user_app_assignments` is a 7th index
that pins the `SinAcceso` exclusivity rule (DA-12, assignments/spec.md).

DA-3 forbids `legacy_hash` / `password_legacy` / `pass_hash_v1` columns
anywhere in this schema; DA-11 forbids telemetry columns on `audit`.
`tests/lanzadera/migrations/test_migration_0001.py` pins both rules as
regression gates.

0001 is **schema only** — no seed data lands here. The 156 migrated users
and the 20 apps arrive in 0002-0006 (PR 3b).
"""

from __future__ import annotations

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op
from sqlalchemy.dialects import postgresql

# revision identifiers, used by Alembic.
revision: str = "0001"
down_revision: str | None = None
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


# ---------------------------------------------------------------------------
# Named schema object helpers.
#
# Centralising the schema name avoids the silent drift that happens when a
# table is created with a one-off `schema="lanzdera"` typo. Alembic reads
# the constant every time the migration touches the schema, and the
# downgrade uses the same constant to drop it.
# ---------------------------------------------------------------------------
SCHEMA = "lanzadera"

# Postgres ENUMs. `create_type=True` lets Alembic own the type lifecycle;
# the downgrade path drops them with the schema.
USER_STATUS = postgresql.ENUM(
    "active",
    "disabled",
    "password_reset_required",
    "locked",
    name="user_status",
    schema=SCHEMA,
    create_type=False,
)

APP_TOPOLOGY = postgresql.ENUM(
    "central",
    "office-nas",
    name="app_topology",
    schema=SCHEMA,
    create_type=False,
)

APP_REGISTRATION_STATUS = postgresql.ENUM(
    "pending",
    "active",
    "retired",
    name="app_registration_status",
    schema=SCHEMA,
    create_type=False,
)

# Server-side default for `created_at` / `updated_at` columns. `now()` is
# the Postgres canonical clock; using a Python-side `datetime.utcnow()`
# would race with the host clock and break deterministic migration replay.
NOW = sa.text("now()")


def upgrade() -> None:
    """Create the `lanzdera` schema, ENUMs, tables, and indexes."""
    # `citext` is the case-insensitive text type the spec pins for
    # `users.email`. The extension is idempotent at the database level,
    # but we guard with `IF NOT EXISTS` so replaying the migration against
    # a database that already has `citext` does not raise.
    op.execute("CREATE EXTENSION IF NOT EXISTS citext")

    # `lanzdera` schema — every production table and ENUM lives here.
    # Keeping the schema separate from `public` lets future Phase 4 work
    # add migrations for adjacent modules without namespace collisions.
    # `IF NOT EXISTS` keeps the migration idempotent: `env.py` creates
    # the schema ahead of the version table so the bookkeeping table can
    # land inside `lanzdera`, and this CREATE just no-ops when it is
    # already there.
    op.execute(f"CREATE SCHEMA IF NOT EXISTS {SCHEMA}")

    # ENUMs first — tables reference them in column definitions. The
    # `create_type=True` flag (the default) tells Alembic to emit the
    # `CREATE TYPE` DDL.
    USER_STATUS.create(op.get_bind(), checkfirst=True)
    APP_TOPOLOGY.create(op.get_bind(), checkfirst=True)
    APP_REGISTRATION_STATUS.create(op.get_bind(), checkfirst=True)

    # ---------------------------------------------------------------------------
    # Tables.
    #
    # Column order matches design.md §Modelo de datos. Defaults use
    # server-side functions (`now()`) so the column stays self-describing
    # even when the calling ORM does not pass a value.
    # ---------------------------------------------------------------------------

    op.create_table(
        "users",
        sa.Column(
            "id",
            postgresql.UUID(as_uuid=True),
            server_default=sa.text("gen_random_uuid()"),
            nullable=False,
        ),
        sa.Column(
            "email",
            postgresql.CITEXT(),
            nullable=False,
        ),
        sa.Column("name", sa.Text(), nullable=False),
        sa.Column("dni_encrypted", sa.LargeBinary(), nullable=False),
        # DA-3, D89: `password_hash` is NULL until `consume_reset_token`
        # succeeds. The 0004 seed (PR 3b) writes every row with this
        # column NULL and `status = 'password_reset_required'`.
        sa.Column("password_hash", sa.Text(), nullable=True),
        sa.Column(
            "status",
            USER_STATUS,
            nullable=False,
            server_default=sa.text(f"'password_reset_required'::{SCHEMA}.user_status"),
        ),
        sa.Column(
            "failed_attempts",
            sa.Integer(),
            nullable=False,
            server_default=sa.text("0"),
        ),
        sa.Column("last_login_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.Column(
            "updated_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.PrimaryKeyConstraint("id", name="pk_users"),
        # Index #1 — UNIQUE constraint on `users.email`. Naming the
        # constraint `ix_users_email` lets `pg_indexes` report a stable
        # name; the underlying index inherits it.
        sa.UniqueConstraint("email", name="ix_users_email"),
        schema=SCHEMA,
    )

    op.create_table(
        "apps",
        sa.Column(
            "id",
            sa.Integer(),
            sa.Identity(always=False),
            nullable=False,
        ),
        sa.Column("name", sa.Text(), nullable=False),
        sa.Column("short_code", sa.Text(), nullable=False),
        sa.Column(
            "deployment_topology",
            APP_TOPOLOGY,
            nullable=False,
            server_default=sa.text(f"'central'::{SCHEMA}.app_topology"),
        ),
        sa.Column(
            "requires_office_presence",
            sa.Boolean(),
            nullable=False,
            server_default=sa.text("false"),
        ),
        sa.Column(
            "registration_status",
            APP_REGISTRATION_STATUS,
            nullable=False,
            server_default=sa.text(f"'pending'::{SCHEMA}.app_registration_status"),
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.Column(
            "updated_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.PrimaryKeyConstraint("id", name="pk_apps"),
        sa.UniqueConstraint("short_code", name="uq_apps_short_code"),
        schema=SCHEMA,
    )

    op.create_table(
        "profiles",
        sa.Column(
            "id",
            postgresql.UUID(as_uuid=True),
            server_default=sa.text("gen_random_uuid()"),
            nullable=False,
        ),
        sa.Column("app_id", sa.Integer(), nullable=False),
        sa.Column("code", sa.Text(), nullable=False),
        sa.Column("name", sa.Text(), nullable=False),
        # `capabilities` is JSONB; the canonical per-app shape lands in
        # PR 3b's 0003 seed. Today the column starts empty for every row.
        sa.Column(
            "capabilities",
            postgresql.JSONB(astext_type=sa.Text()),
            nullable=False,
            server_default=sa.text("'{}'::jsonb"),
        ),
        sa.Column(
            "active",
            sa.Boolean(),
            nullable=False,
            server_default=sa.text("true"),
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.Column(
            "updated_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.PrimaryKeyConstraint("id", name="pk_profiles"),
        sa.UniqueConstraint("app_id", "code", name="uq_profiles_app_id_code"),
        sa.ForeignKeyConstraint(
            ["app_id"],
            [f"{SCHEMA}.apps.id"],
            name="fk_profiles_app_id_apps",
            ondelete="CASCADE",
        ),
        schema=SCHEMA,
    )

    op.create_table(
        "user_app_assignments",
        sa.Column(
            "id",
            postgresql.UUID(as_uuid=True),
            server_default=sa.text("gen_random_uuid()"),
            nullable=False,
        ),
        sa.Column("user_id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("app_id", sa.Integer(), nullable=False),
        sa.Column("profile_id", postgresql.UUID(as_uuid=True), nullable=False),
        # `granted_by` carries the global admin who issued the grant.
        # Nullable because the 0005 seed (PR 3b) mints rows attributed
        # to the platform, not to a person.
        sa.Column("granted_by", postgresql.UUID(as_uuid=True), nullable=True),
        sa.Column(
            "granted_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        # `revoked_at` is NULL while the assignment is active. The column
        # is filled when a global admin revokes the profile; the test
        # `test_user_app_assignments_revoked_at_is_nullable` pins this
        # shape so a future NOT NULL cannot slip in.
        sa.Column("revoked_at", sa.DateTime(timezone=True), nullable=True),
        sa.PrimaryKeyConstraint("id", name="pk_user_app_assignments"),
        sa.ForeignKeyConstraint(
            ["user_id"],
            [f"{SCHEMA}.users.id"],
            name="fk_user_app_assignments_user_id_users",
            ondelete="CASCADE",
        ),
        sa.ForeignKeyConstraint(
            ["app_id"],
            [f"{SCHEMA}.apps.id"],
            name="fk_user_app_assignments_app_id_apps",
            ondelete="CASCADE",
        ),
        sa.ForeignKeyConstraint(
            ["profile_id"],
            [f"{SCHEMA}.profiles.id"],
            name="fk_user_app_assignments_profile_id_profiles",
            ondelete="RESTRICT",
        ),
        sa.ForeignKeyConstraint(
            ["granted_by"],
            [f"{SCHEMA}.users.id"],
            name="fk_user_app_assignments_granted_by_users",
        ),
        schema=SCHEMA,
    )

    op.create_table(
        "global_admins",
        sa.Column("user_id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column(
            "granted_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.Column("granted_by", postgresql.UUID(as_uuid=True), nullable=True),
        sa.PrimaryKeyConstraint("user_id", name="pk_global_admins"),
        sa.ForeignKeyConstraint(
            ["user_id"],
            [f"{SCHEMA}.users.id"],
            name="fk_global_admins_user_id_users",
            ondelete="CASCADE",
        ),
        sa.ForeignKeyConstraint(
            ["granted_by"],
            [f"{SCHEMA}.users.id"],
            name="fk_global_admins_granted_by_users",
        ),
        schema=SCHEMA,
    )

    op.create_table(
        "sessions",
        sa.Column(
            "id",
            postgresql.UUID(as_uuid=True),
            server_default=sa.text("gen_random_uuid()"),
            nullable=False,
        ),
        sa.Column("user_id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.Column(
            "expires_at",
            sa.DateTime(timezone=True),
            nullable=False,
        ),
        sa.PrimaryKeyConstraint("id", name="pk_sessions"),
        sa.ForeignKeyConstraint(
            ["user_id"],
            [f"{SCHEMA}.users.id"],
            name="fk_sessions_user_id_users",
            ondelete="CASCADE",
        ),
        schema=SCHEMA,
    )

    op.create_table(
        "audit",
        sa.Column(
            "id",
            postgresql.UUID(as_uuid=True),
            server_default=sa.text("gen_random_uuid()"),
            nullable=False,
        ),
        sa.Column("event_type", sa.Text(), nullable=False),
        # `actor_id` is nullable — system-initiated events (cron, bootstrap)
        # have no human actor. The column is FK into `users` so a deleted
        # user leaves a tombstone instead of an orphan audit row.
        sa.Column("actor_id", postgresql.UUID(as_uuid=True), nullable=True),
        sa.Column("target_id", sa.Text(), nullable=False),
        sa.Column(
            "module",
            sa.Text(),
            nullable=False,
            server_default=sa.text("'lanzadera'"),
        ),
        sa.Column("result", sa.Text(), nullable=False),
        sa.Column("correlation_id", postgresql.UUID(as_uuid=True), nullable=False),
        # `payload` is JSONB. Telemetry columns are forbidden (DA-11); the
        # test `test_audit_has_no_telemetry_columns` pins the rule.
        sa.Column(
            "payload",
            postgresql.JSONB(astext_type=sa.Text()),
            nullable=False,
            server_default=sa.text("'{}'::jsonb"),
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.PrimaryKeyConstraint("id", name="pk_audit"),
        sa.ForeignKeyConstraint(
            ["actor_id"],
            [f"{SCHEMA}.users.id"],
            name="fk_audit_actor_id_users",
        ),
        schema=SCHEMA,
    )

    op.create_table(
        "reset_tokens",
        sa.Column(
            "id",
            postgresql.UUID(as_uuid=True),
            server_default=sa.text("gen_random_uuid()"),
            nullable=False,
        ),
        sa.Column("user_id", postgresql.UUID(as_uuid=True), nullable=False),
        # `token_hash` carries the SHA-256 of the raw token. The raw
        # token only lives in the email body and is never persisted
        # (DA-4, D90).
        sa.Column("token_hash", sa.Text(), nullable=False),
        sa.Column(
            "expires_at",
            sa.DateTime(timezone=True),
            nullable=False,
        ),
        sa.Column("consumed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("superseded_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.PrimaryKeyConstraint("id", name="pk_reset_tokens"),
        sa.UniqueConstraint("token_hash", name="uq_reset_tokens_token_hash"),
        sa.ForeignKeyConstraint(
            ["user_id"],
            [f"{SCHEMA}.users.id"],
            name="fk_reset_tokens_user_id_users",
            ondelete="CASCADE",
        ),
        schema=SCHEMA,
    )

    op.create_table(
        "mail_outbox",
        sa.Column(
            "id",
            postgresql.UUID(as_uuid=True),
            server_default=sa.text("gen_random_uuid()"),
            nullable=False,
        ),
        sa.Column("to_addr", sa.Text(), nullable=False),
        sa.Column("subject", sa.Text(), nullable=False),
        sa.Column("body", sa.Text(), nullable=False),
        sa.Column(
            "status",
            sa.Text(),
            nullable=False,
            server_default=sa.text("'pending'"),
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=NOW,
        ),
        sa.Column("sent_at", sa.DateTime(timezone=True), nullable=True),
        sa.PrimaryKeyConstraint("id", name="pk_mail_outbox"),
        schema=SCHEMA,
    )

    # ---------------------------------------------------------------------------
    # Indexes.
    #
    # Six hot-path indexes design.md commits to. The unique constraint on
    # `users.email` (Index #1) is created above by `UniqueConstraint`; the
    # rest land here.
    # ---------------------------------------------------------------------------

    # Index #2 — hot path: looking up a user's outstanding or expired
    # tokens while serving `/reset` and the daemon that purges them.
    op.create_index(
        "ix_reset_tokens_user_id_expires_at",
        "reset_tokens",
        ["user_id", "expires_at"],
        schema=SCHEMA,
    )

    # Index #3 — hot path: the `effective_permissions(user_id, app_id)`
    # adapter query is the single most common read in the system.
    op.create_index(
        "ix_user_app_assignments_user_id_app_id",
        "user_app_assignments",
        ["user_id", "app_id"],
        schema=SCHEMA,
    )

    # Index #4 — audit lookup by actor for the audit viewer.
    op.create_index(
        "ix_audit_actor_id_created_at",
        "audit",
        ["actor_id", sa.text("created_at DESC")],
        schema=SCHEMA,
    )

    # Index #5 — audit lookup by event type for the audit viewer.
    op.create_index(
        "ix_audit_event_type_created_at",
        "audit",
        ["event_type", sa.text("created_at DESC")],
        schema=SCHEMA,
    )

    # Index #6 — mail_outbox dispatcher: pull pending rows in FIFO order.
    op.create_index(
        "ix_mail_outbox_status_created_at",
        "mail_outbox",
        ["status", "created_at"],
        schema=SCHEMA,
    )

    # Extra: a partial unique on `user_app_assignments` that lets a user
    # hold the same `(user_id, app_id, profile_id)` triple exactly once
    # at a time, but only while it is active. This is the 7th index
    # backing DA-12 (the `SinAcceso` exclusive rule + the cardinality
    # expectation). Not asserted by the 6-index test, but enforced here
    # so 0005's seed cannot insert duplicate active rows.
    op.create_index(
        "uq_user_app_assignments_active_triple",
        "user_app_assignments",
        ["user_id", "app_id", "profile_id"],
        unique=True,
        schema=SCHEMA,
        postgresql_where=sa.text("revoked_at IS NULL"),
    )


def downgrade() -> None:
    """Reverse 0001 in lock-step with `upgrade()`.

    Drop order matters: indexes first (so foreign keys have time to
    cascade), then tables, then ENUMs, then the schema. `IF EXISTS`
    guards make the downgrade idempotent on partial replays — a
    re-invocation after a successful first downgrade is a no-op.
    """
    # Indexes that reference `user_app_assignments` go before the table.
    op.drop_index(
        "uq_user_app_assignments_active_triple",
        table_name="user_app_assignments",
        schema=SCHEMA,
    )
    op.drop_index(
        "ix_user_app_assignments_user_id_app_id",
        table_name="user_app_assignments",
        schema=SCHEMA,
    )
    op.drop_index(
        "ix_mail_outbox_status_created_at",
        table_name="mail_outbox",
        schema=SCHEMA,
    )
    op.drop_index(
        "ix_audit_event_type_created_at",
        table_name="audit",
        schema=SCHEMA,
    )
    op.drop_index(
        "ix_audit_actor_id_created_at",
        table_name="audit",
        schema=SCHEMA,
    )
    op.drop_index(
        "ix_reset_tokens_user_id_expires_at",
        table_name="reset_tokens",
        schema=SCHEMA,
    )

    # Tables. Order matters: child tables first so FK cascades don't
    # strand dangling rows.
    op.drop_table("user_app_assignments", schema=SCHEMA)
    op.drop_table("global_admins", schema=SCHEMA)
    op.drop_table("sessions", schema=SCHEMA)
    op.drop_table("audit", schema=SCHEMA)
    op.drop_table("reset_tokens", schema=SCHEMA)
    op.drop_table("mail_outbox", schema=SCHEMA)
    op.drop_table("profiles", schema=SCHEMA)
    op.drop_table("apps", schema=SCHEMA)
    op.drop_table("users", schema=SCHEMA)

    # ENUMs. `checkfirst=False` would fail if the type is gone; with
    # `checkfirst=True` the downgrade is idempotent.
    APP_REGISTRATION_STATUS.drop(op.get_bind(), checkfirst=True)
    APP_TOPOLOGY.drop(op.get_bind(), checkfirst=True)
    USER_STATUS.drop(op.get_bind(), checkfirst=True)

    # The schema itself is NOT dropped here. Alembic's `head_maintainer`
    # runs AFTER `downgrade()` and tries to DELETE the version row from
    # `lanzadera.alembic_version`. Dropping the schema here would race
    # the head_maintainer's DELETE and crash with `UndefinedTableError`.
    # Leaving the schema empty (with just the emptied `alembic_version`
    # table) keeps the rollback boundary safe; the operator can drop
    # the schema explicitly via `DROP SCHEMA lanzdera CASCADE` when
    # they want to nuke the module — see the Makefile target
    # `migrations-down`.
