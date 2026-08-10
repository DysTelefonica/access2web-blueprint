# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 3a
# D82 (Expand and Contract) — every migration is downgradable; this test
# pins the contract for 0001 by running `alembic downgrade base` and
# asserting the production schema returns to its empty pre-migration
# state.
"""Integration test for the downgradable boundary of Alembic migration 0001.

The Expand and Contract principle (D82) bans destructive one-shot
migrations; instead, every migration must reverse cleanly so a hotfix can
`alembic downgrade -1` without nuking the `.accdb` legacy. This module
runs `alembic upgrade head` followed by `alembic downgrade base` and
asserts that every production table is gone and Alembic's bookkeeping
table is empty. The schema itself remains — Alembic's `head_maintainer`
deletes the version row during the downgrade and the schema's lifecycle
is the operator's responsibility (`DROP SCHEMA lanzadera CASCADE` via
the Makefile target). The test is gated on `APAP_INTEGRATION_ENABLED=1`
for the same reason as the upgrade tests: it requires a live Postgres.

Sync `def test_*` because the project's `pyproject.toml` does not enable
`pytest-asyncio`; async I/O runs through `asyncio.run()` inside each test
body so the gate is the env flag, not the test runner.
"""

from __future__ import annotations

import asyncio
import os
import secrets
import subprocess
import sys
from pathlib import Path

import asyncpg
import pytest

from tests.lanzadera.migrations.test_migration_0001 import (
    EXPECTED_TABLES,
    PG_HOST,
    PG_PASSWORD,
    PG_PORT,
    PG_USER,
    _alembic_ini_candidates,
    _drop_test_database,
    _fetch_table_names,
    _worktree_root,
)


def _run_alembic(
    worktree_root: Path, alembic_ini: Path, db_name: str, *args: str
) -> None:
    """Invoke `alembic` against the test database."""
    database_url = (
        f"postgresql+asyncpg://{PG_USER}:{PG_PASSWORD}@{PG_HOST}:{PG_PORT}/{db_name}"
    )
    env = os.environ.copy()
    env["DATABASE_URL"] = database_url
    env["PYTHONPATH"] = str(worktree_root) + os.pathsep + env.get("PYTHONPATH", "")
    cmd = [
        sys.executable,
        "-m",
        "alembic",
        "-c",
        str(alembic_ini),
        *args,
    ]
    result = subprocess.run(
        cmd,
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


@pytest.fixture(scope="module")
def integration_enabled() -> bool:
    return os.environ.get("APAP_INTEGRATION_ENABLED") == "1"


@pytest.fixture(scope="module")
def worktree_root() -> Path:
    return _worktree_root()


@pytest.fixture(scope="module")
def alembic_ini(worktree_root: Path) -> Path:
    for candidate in _alembic_ini_candidates(worktree_root):
        if candidate.is_file():
            return candidate
    # See the sibling module: Hard Rule 18. A check with no subject announces the
    # skip instead of erroring the suite for a dependency that is tracked (PR #87).
    pytest.skip("alembic.ini has not shipped yet — it arrives with PR #87")


@pytest.fixture(scope="module")
def test_db_name() -> str:
    return f"lanzadera_test_0001_down_{secrets.token_hex(4)}"


async def _admin_conn() -> asyncpg.Connection:
    return await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database="lanzadera",
    )


async def _create_test_database(name: str) -> None:
    conn = await _admin_conn()
    try:
        await conn.execute(f'DROP DATABASE IF EXISTS "{name}"')
        await conn.execute(f'CREATE DATABASE "{name}"')
    finally:
        await conn.close()


async def _fetch_alembic_version_row(db_name: str) -> str | None:
    """Return the `version_num` row from `lanzadera.alembic_version`.

    Returns `None` when the table does not exist or when the table is
    empty (the downgrade base case).
    """
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        try:
            row = await conn.fetchrow(
                "SELECT version_num FROM lanzadera.alembic_version LIMIT 1"
            )
        except asyncpg.UndefinedTableError:
            return None
    finally:
        await conn.close()
    return row["version_num"] if row else None


@pytest.fixture(scope="module")
def roundtripped_db(
    integration_enabled: bool,
    worktree_root: Path,
    alembic_ini: Path,
    test_db_name: str,
) -> str:
    """Apply 0001 then downgrade to base.

    The downgrade roundtrip is the only way to prove D82 in an automated
    test: if `downgrade()` forgets a DROP, the assertion below catches it.
    The fixture yields the database name so the test can re-query
    `information_schema` directly.
    """
    if not integration_enabled:
        pytest.skip("APAP_INTEGRATION_ENABLED=1 required for migration tests")
    asyncio.run(_create_test_database(test_db_name))
    try:
        _run_alembic(worktree_root, alembic_ini, test_db_name, "upgrade", "head")
        _run_alembic(worktree_root, alembic_ini, test_db_name, "downgrade", "base")
        yield test_db_name
    finally:
        asyncio.run(_drop_test_database(test_db_name))


@pytest.mark.integration
def test_downgrade_base_removes_all_nine_tables(roundtripped_db: str) -> None:
    """Every production table is gone after `alembic downgrade base`.

    Catches a `downgrade()` that forgets a `drop_table` call and serves
    as a regression test against a future PR that introduces a new
    production table without a matching drop in `downgrade()`.

    The Alembic bookkeeping table (`alembic_version`) is allowed to
    remain — Alembic's `head_maintainer` empties it during the downgrade
    but does not drop it (see the rationale in `app/migrations/env.py`).
    """
    try:
        tables = asyncio.run(_fetch_table_names(roundtripped_db))
    except asyncpg.UndefinedTableError:
        tables = set()
    assert not (set(EXPECTED_TABLES) & tables), (
        f"production tables {sorted(set(EXPECTED_TABLES) & tables)} "
        f"survived `alembic downgrade base`; downgrade() missed a DROP"
    )


@pytest.mark.integration
def test_downgrade_base_empties_alembic_version(roundtripped_db: str) -> None:
    """Alembic's bookkeeping table is empty after `downgrade base`.

    The version row must be gone so a subsequent `alembic upgrade head`
    is a no-op (it stays at base). Alembic's `head_maintainer` performs
    this DELETE during the downgrade; this test pins the contract.
    """
    version = asyncio.run(_fetch_alembic_version_row(roundtripped_db))
    assert version is None, (
        f"alembic_version row '{version}' still present after "
        f"`alembic downgrade base`; head_maintainer must have failed"
    )
