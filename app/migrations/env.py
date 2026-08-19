# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 3a
# DA-1 (composition root owns config), DA-3 (no legacy_hash in schema),
# D82 (Expand and Contract — every migration is reversible).
"""Alembic environment for `lanzadera-mvp` Phase 2 migrations.

This is the bridge between Alembic's CLI and the project's runtime config.
The orchestrator pre-resolved decision pins `DATABASE_URL` (loaded by the
composition root from `docker-compose.yml` or the test harness) as the
single source of truth for the database connection. Hard-coding a URL
here would split config across two files and break the `make migrations-up`
target when the developer swaps envs.

Both online (sync migration against an async engine) and offline (raw
SQL to stdout) modes are supported. Online mode is the production path;
offline is for SQL review and for the CI gate that confirms the rendered
DDL matches the committed migration.
"""

from __future__ import annotations

import asyncio
import os
import sys
from logging.config import fileConfig
from pathlib import Path

from alembic import context
from sqlalchemy import pool
from sqlalchemy.engine import Connection
from sqlalchemy.ext.asyncio import async_engine_from_config

# Alembic's runtime config object. The `.ini` file lives alongside this
# module and is the source of truth for `script_location` / logging.
config = context.config

# Configure Python logging from the `[loggers]` / `[handlers]` blocks in
# `alembic.ini`. Skip silently when no config file is provided (e.g. when
# `env.py` is imported by a test fixture that bypasses the CLI).
if config.config_file_name is not None:
    fileConfig(config.config_file_name)

# `target_metadata` stays None: Phase 2 ships DDL-only migrations with no
# SQLAlchemy declarative models. When PR 4 introduces `User`, `App`, etc.
# as declarative classes, this assignment moves to the model import block
# and autogenerate becomes available.
target_metadata = None


def _resolve_database_url() -> str:
    """Return the database URL from `DATABASE_URL`, raising if missing.

    A missing URL is a configuration bug — Alembic must not guess. The
    docker-compose service injects the same env var into the backend
    container, so a missing value here almost always means the developer
    forgot `export DATABASE_URL=...` (or the WSL shell ran outside
    `docker compose up`). The error message is loud on purpose.
    """
    url = os.environ.get("DATABASE_URL")
    if not url:
        raise RuntimeError(
            "DATABASE_URL is unset; alembic env.py refuses to guess. "
            "Export it from the docker-compose service or the test harness "
            "before invoking `alembic upgrade head`."
        )
    return url


def run_migrations_offline() -> None:
    """Render migrations as raw SQL without connecting to the database.

    Offline mode is the CI gate that proves the migration script's text is
    deterministic. The `--sql` flag triggers this branch.
    """
    url = _resolve_database_url()
    context.configure(
        url=url,
        target_metadata=target_metadata,
        literal_binds=True,
        dialect_opts={"paramstyle": "named"},
        # `compare_type=True` is irrelevant without `target_metadata`; we
        # keep it commented to avoid misleading future readers. When the
        # declarative models land in Phase 4, uncomment alongside the
        # `target_metadata = Base.metadata` assignment above.
        # compare_type=True,
    )
    with context.begin_transaction():
        context.run_migrations()


# Pin the version table inside the `lanzadera` schema. Alembic creates
# `alembic_version` automatically before running any migration; without
# an explicit `version_table_schema`, Postgres' default search_path
# (`$user, public`) resolves the unqualified `alembic_version` to the
# `lanzadera` schema because the migration's role is `lanzadera` and we
# pre-create the schema in `env.py`. Pinning it explicitly avoids the
# implicit-search-path dependency and documents the choice.
_VERSION_SCHEMA = "lanzadera"


def do_run_migrations(connection: Connection) -> None:
    """Run the migration scripts against an open sync connection."""
    context.configure(
        connection=connection,
        target_metadata=target_metadata,
        # Render DDL in the `lanzadera` schema by default so migrations
        # need not qualify every `create_table` call. `include_schemas`
        # tells Alembic that `lanzadera` exists when autogenerate wakes
        # up in Phase 4; today the parameter is harmless because the
        # 0001 migration creates the schema before any table.
        include_schemas=True,
        # The bookkeeping table (`alembic_version`) lives inside the
        # `lanzadera` schema. The migration's `downgrade()` does NOT
        # drop the schema; instead Alembic's `head_maintainer` runs
        # AFTER the migration's `downgrade()` and empties the version
        # table itself. Dropping the schema in the migration's
        # `downgrade()` would race the head_maintainer's DELETE and
        # crash with `UndefinedTableError`. The schema is dropped
        # explicitly by `make migrations-down` via `DROP SCHEMA
        # lanzdera CASCADE` when the operator chooses to nuke the
        # module — see the Makefile target. Until then, the schema
        # remains with an empty `alembic_version` row.
        version_table_schema=_VERSION_SCHEMA,
    )
    with context.begin_transaction():
        context.run_migrations()


async def run_async_migrations() -> None:
    """Drive Alembic's sync migration runner from an async engine."""
    # Inject the runtime URL into the alembic config so
    # `async_engine_from_config` picks it up via its `sqlalchemy.url`
    # prefix. The cast from `postgresql+asyncpg://` happens here so the
    # rest of the file stays oblivious to the dialect.
    config.set_main_option("sqlalchemy.url", _resolve_database_url())

    connectable = async_engine_from_config(
        config.get_section(config.config_ini_section, {}),
        prefix="sqlalchemy.",
        poolclass=pool.NullPool,
    )
    async with connectable.connect() as connection:
        # Alembic creates its bookkeeping table (`alembic_version`) BEFORE
        # running the first migration, so the schema that hosts it must
        # exist. `CREATE SCHEMA IF NOT EXISTS` is idempotent; the
        # migration's own `CREATE SCHEMA` is also guarded so re-running
        # the migration after a manual schema drop does not trip over an
        # existing one.
        await connection.exec_driver_sql(f'CREATE SCHEMA IF NOT EXISTS "{_VERSION_SCHEMA}"')
        # Run the migration inside its own transaction so alembic's
        # bookkeeping (creation/deletion of `alembic_version`) and the
        # migration's DDL land atomically.
        await connection.run_sync(do_run_migrations)
        # asyncpg/SQLAlchemy does NOT auto-commit on clean `async with`
        # exit; without an explicit commit, every change rolls back when
        # the connection returns to the pool and the migration appears
        # to succeed (no traceback, "Running upgrade" log line) but the
        # state is gone.
        await connection.commit()
    await connectable.dispose()


def run_migrations_online() -> None:
    """Entry point for the online (live database) branch."""
    asyncio.run(run_async_migrations())


# Dispatch on Alembic's CLI mode. `is_offline_mode()` returns True for
# `alembic --sql` and for `alembic upgrade --sql`; everything else lands in
# the async branch.
if context.is_offline_mode():
    run_migrations_offline()
else:
    run_migrations_online()
