"""Shared PostgreSQL helpers for the Expedientes migration tests.

The helpers mirror ``tests/lanzadera/migrations/test_migration_0001.py``: a
fresh database per module, ``alembic`` invoked as a subprocess against it and
a fail-closed ``APAP_INTEGRATION_ENABLED=1`` gate (issue #882). They live here
so the migration tests of this module share one implementation instead of
copying the database plumbing per revision.
"""

from __future__ import annotations

import os
import secrets
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlsplit

import asyncpg
import pytest

# Database creation borrows the credentials shipped by docker-compose.yml.
PG_USER = "lanzadera"
PG_PASSWORD = "lanzadera-dev-only"


def _pg_endpoint() -> tuple[str, int]:
    """Host and port of the integration Postgres.

    The CI job publishes PostgreSQL on a dynamic host port and exports the
    real DSN in ``APAP_INTEGRATION_URL`` (or ``INTEGRATION_DSN``). Without a
    DSN the loopback default ``127.0.0.1:5432`` applies (local docker
    compose).
    """
    dsn = os.environ.get("APAP_INTEGRATION_URL") or os.environ.get("INTEGRATION_DSN")
    if not dsn:
        return "127.0.0.1", 5432
    parts = urlsplit(dsn)
    return parts.hostname or "127.0.0.1", parts.port or 5432


PG_HOST, PG_PORT = _pg_endpoint()


async def fetch(db_name: str, sql: str, *args: object) -> list[asyncpg.Record]:
    """Run one query against the test database and return its rows."""
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        return await conn.fetch(sql, *args)
    finally:
        await conn.close()


async def _admin(sql: str, *args: object) -> None:
    """Run one statement against the admin database (CREATE/DROP DATABASE)."""
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database="lanzadera",
    )
    try:
        await conn.execute(sql, *args)
    finally:
        await conn.close()


async def create_test_database(name: str) -> None:
    """Create a fresh, empty database for the test run."""
    await _admin(f'DROP DATABASE IF EXISTS "{name}"')
    await _admin(f'CREATE DATABASE "{name}"')


async def drop_test_database(name: str) -> None:
    """Force-disconnect any lingering sessions and drop the test database."""
    await _admin(
        "SELECT pg_terminate_backend(pid) FROM pg_stat_activity "
        "WHERE datname = $1 AND pid <> pg_backend_pid()",
        name,
    )
    await _admin(f'DROP DATABASE IF EXISTS "{name}"')


async def tables(db_name: str, schema: str) -> set[str]:
    """The base tables of one schema."""
    rows = await fetch(
        db_name,
        "SELECT table_name FROM information_schema.tables "
        "WHERE table_schema = $1 AND table_type = 'BASE TABLE'",
        schema,
    )
    return {row["table_name"] for row in rows}


async def columns(db_name: str, schema: str) -> dict[str, dict[str, str]]:
    """Every table of the schema with its columns and their SQL types."""
    rows = await fetch(
        db_name,
        "SELECT table_name, column_name, data_type FROM information_schema.columns "
        "WHERE table_schema = $1",
        schema,
    )
    result: dict[str, dict[str, str]] = {}
    for row in rows:
        result.setdefault(row["table_name"], {})[row["column_name"]] = row["data_type"]
    return result


async def enums(db_name: str, schema: str) -> dict[str, tuple[str, ...]]:
    """Every ENUM type of the schema with its labels in declaration order."""
    rows = await fetch(
        db_name,
        "SELECT t.typname, e.enumlabel FROM pg_enum e "
        "JOIN pg_type t ON t.oid = e.enumtypid "
        "JOIN pg_namespace n ON n.oid = t.typnamespace "
        "WHERE n.nspname = $1 ORDER BY t.typname, e.enumsortorder",
        schema,
    )
    result: dict[str, list[str]] = {}
    for row in rows:
        result.setdefault(row["typname"], []).append(row["enumlabel"])
    return {name: tuple(labels) for name, labels in result.items()}


def worktree_root() -> Path:
    """Locate the worktree root from this module."""
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / ".git"
        if candidate.is_dir() or candidate.is_file():
            return parent
    raise AssertionError("worktree root not found (no .git)")


def alembic_ini(root: Path) -> Path:
    """The Alembic config of this checkout."""
    for candidate in (root / "app" / "migrations" / "alembic.ini", root / "alembic.ini"):
        if candidate.is_file():
            return candidate
    pytest.skip("alembic.ini has not shipped yet")


def run_alembic(root: Path, config: Path, db_name: str, *args: str) -> None:
    """Invoke `alembic` against the test database."""
    database_url = f"postgresql+asyncpg://{PG_USER}:{PG_PASSWORD}@{PG_HOST}:{PG_PORT}/{db_name}"
    env = os.environ.copy()
    env["DATABASE_URL"] = database_url
    env["PYTHONPATH"] = str(root) + os.pathsep + env.get("PYTHONPATH", "")
    result = subprocess.run(
        [sys.executable, "-m", "alembic", "-c", str(config), *args],
        cwd=root,
        env=env,
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )
    if result.returncode != 0:
        pytest.fail(
            f"alembic {' '.join(args)} failed (exit {result.returncode}):\n"
            f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        )


def require_integration() -> None:
    """Fail closed when the integration gate is off (issue #882)."""
    if os.environ.get("APAP_INTEGRATION_ENABLED") != "1":
        pytest.fail(
            "Issue #882: APAP_INTEGRATION_ENABLED is not 1. The integration "
            "DSN gate must fail closed, not skip, so a missing Postgres "
            "shows up as a red CI instead of a silently green one."
        )


def unique_db_name(prefix: str) -> str:
    """A test database name, unique per run."""
    return f"expedientes_test_{prefix}_{secrets.token_hex(4)}"
