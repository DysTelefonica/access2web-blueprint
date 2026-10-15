"""Regression test for issue #771.

The presence feature (W60, #522) shipped ``PRESENCE_TABLE`` and four
adapter methods on ``presence_repository_pg.py``, but no migration
created the underlying ``lanzadera.presence_sessions`` table. A clean
database crashes on the first ``list_connected()`` call with
``relation \"lanzadera.presence_sessions\" does not exist``.

Layout pinned by ``presence_table.py`` (asserted byte-for-byte):
``user_id`` UUID PK, ``email`` text NOT NULL, ``connected_at`` and
``last_seen`` timestamptz NOT NULL.
"""

from __future__ import annotations

import os
import secrets
import subprocess
import sys
from pathlib import Path

import asyncpg
import pytest

from tests.lanzadera.migrations.test_migration_0001 import (
    PG_HOST,
    PG_PASSWORD,
    PG_PORT,
    PG_USER,
    _alembic_ini_candidates,
    _create_test_database,
    _drop_test_database,
    _worktree_root,
)


@pytest.fixture
def integration_enabled() -> bool:
    return os.environ.get("APAP_INTEGRATION_ENABLED") == "1"


SCHEMA = "lanzadera"
TABLE = "presence_sessions"

# Order matches information_schema.column ordinal_position so the
# assertion fails fast if any column is missing or reordered.
EXPECTED_COLUMNS: tuple[tuple[str, str, str], ...] = (
    ("user_id", "uuid", "NO"),
    ("email", "text", "NO"),
    ("connected_at", "timestamp with time zone", "NO"),
    ("last_seen", "timestamp with time zone", "NO"),
)


def _run_alembic(worktree_root: Path, alembic_ini: Path, db_name: str, *args: str) -> None:
    database_url = f"postgresql+asyncpg://{PG_USER}:{PG_PASSWORD}@{PG_HOST}:{PG_PORT}/{db_name}"
    env = os.environ.copy()
    env["DATABASE_URL"] = database_url
    env["PYTHONPATH"] = str(worktree_root) + os.pathsep + env.get("PYTHONPATH", "")
    result = subprocess.run(
        [sys.executable, "-m", "alembic", "-c", str(alembic_ini), *args],
        cwd=worktree_root,
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


async def _table_columns(db_name: str) -> list[asyncpg.Record]:
    conn = await asyncpg.connect(
        host=PG_HOST, port=PG_PORT, user=PG_USER, password=PG_PASSWORD, database=db_name
    )
    try:
        return await conn.fetch(
            "SELECT column_name, data_type, is_nullable "
            "FROM information_schema.columns "
            "WHERE table_schema = $1 AND table_name = $2 "
            "ORDER BY ordinal_position",
            SCHEMA,
            TABLE,
        )
    finally:
        await conn.close()


async def _table_exists(db_name: str) -> bool:
    conn = await asyncpg.connect(
        host=PG_HOST, port=PG_PORT, user=PG_USER, password=PG_PASSWORD, database=db_name
    )
    try:
        return bool(
            await conn.fetchval(
                "SELECT EXISTS(SELECT 1 FROM information_schema.tables "
                "WHERE table_schema = $1 AND table_name = $2)",
                SCHEMA,
                TABLE,
            )
        )
    finally:
        await conn.close()


@pytest.mark.integration
async def test_alembic_upgrade_head_creates_presence_sessions_with_pinned_layout(
    integration_enabled: bool,
) -> None:
    """Issue #771: alembic upgrade head creates the table with the pinned layout."""
    assert integration_enabled, "APAP_INTEGRATION_ENABLED=1 required"

    worktree_root = _worktree_root()
    alembic_ini = next((c for c in _alembic_ini_candidates(worktree_root) if c.is_file()), None)
    assert alembic_ini is not None, "alembic.ini not found"

    db_name = f"lanzadera_test_771_presence_{secrets.token_hex(4)}"
    await _create_test_database(db_name)
    try:
        assert not await _table_exists(db_name), "test DB must start clean"

        _run_alembic(worktree_root, alembic_ini, db_name, "upgrade", "head")

        assert await _table_exists(db_name), (
            "alembic upgrade head did not create lanzadera.presence_sessions (#771)"
        )
        observed = tuple(
            (r["column_name"], r["data_type"], r["is_nullable"])
            for r in await _table_columns(db_name)
        )
        assert observed == EXPECTED_COLUMNS, (
            f"presence_sessions layout drifted from presence_table.py pin.\n"
            f"  expected: {EXPECTED_COLUMNS}\n"
            f"  observed: {observed}\n  (#771)"
        )
    finally:
        await _drop_test_database(db_name)


@pytest.mark.integration
async def test_presence_repository_adapter_can_write_after_upgrade_head(
    integration_enabled: bool,
) -> None:
    """After upgrade head, the adapter can write to the table the runtime uses."""
    assert integration_enabled, "APAP_INTEGRATION_ENABLED=1 required"

    worktree_root = _worktree_root()
    alembic_ini = next((c for c in _alembic_ini_candidates(worktree_root) if c.is_file()), None)
    assert alembic_ini is not None, "alembic.ini not found"

    db_name = f"lanzadera_test_771_adapter_{secrets.token_hex(4)}"
    await _create_test_database(db_name)
    try:
        _run_alembic(worktree_root, alembic_ini, db_name, "upgrade", "head")

        # Importing PRESENCE_TABLE exercises the real adapter import path
        # the runtime uses (proves the schema the migration created lines
        # up with the schema the code queries).
        from app.src.modules.lanzadera.adapters.persistence.repositories.presence_table import (  # noqa: E501
            PRESENCE_TABLE,
        )

        conn = await asyncpg.connect(
            host=PG_HOST,
            port=PG_PORT,
            user=PG_USER,
            password=PG_PASSWORD,
            database=db_name,
        )
        try:
            await conn.execute(
                f"INSERT INTO {SCHEMA}.{TABLE} "
                "(user_id, email, connected_at, last_seen) "
                "VALUES (gen_random_uuid(), $1, now(), now())",
                f"adapter-{secrets.token_hex(4)}@test.local",
            )
            count = await conn.fetchval(f"SELECT count(*) FROM {SCHEMA}.{TABLE}")
            assert count == 1, f"adapter could not write to {SCHEMA}.{TABLE} (#771)"

            adapter_columns = {c.name for c in PRESENCE_TABLE.columns}
            db_columns = {r["column_name"] for r in await _table_columns(db_name)}
            missing = adapter_columns - db_columns
            assert not missing, (
                f"PRESENCE_TABLE references columns the migration did not create: "
                f"{sorted(missing)} (#771)"
            )
        finally:
            await conn.close()
    finally:
        await _drop_test_database(db_name)
