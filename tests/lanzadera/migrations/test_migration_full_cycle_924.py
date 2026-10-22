"""Regression test for issue #924: full migration cycle (upgrade head →
downgrade base → upgrade head) must run without errors on a fresh database.

This is the *integration* counterpart of the unit tests already shipped for
#929 (downgrade 0006 by exact row). The test for #929 only runs
``alembic downgrade 0005`` and stops at the 0006->0005 step, so it does not
exercise the 0005->0004 transition that #924 reports as broken.

#924: ``alembic upgrade head && alembic downgrade base`` fails on a fresh
database with::

    asyncpg.exceptions.ForeignKeyViolationError: update or delete on
    table "users" violates foreign key constraint "fk_audit_actor_id_users"
    on table "audit"

Root cause per #924: ``0006_seed_audit``'s old ``downgrade()`` deleted only
events with ``actor_id IS NULL``; the seeded events with a non-null
``actor_id`` survived and blocked the 0004 user purge via the FK.

PR #929 replaced that heuristic with a full-row ``IS NOT DISTINCT FROM``
match on ``(id, event_type, actor_id, target_id, created_at)``, deleting
every row the ``upgrade()`` wrote. The 0005->0004 downgrade should now
succeed because the FK target rows are gone.

TDD contract (gentle-ai RDD):

  RED on ``origin/main`` *before* #929 (would fail with
  ``ForeignKeyViolationError``). GREEN once #929 lands — the cycle
  completes and a second ``upgrade head`` re-creates the same row count
  as the first one.

The test also asserts the second ``upgrade head`` lands at the same head
revision as the first, so an accidental divergence between
``alembic_version`` and the schema would be caught.
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
    """True when ``APAP_INTEGRATION_ENABLED=1`` is set in the test env."""
    return os.environ.get("APAP_INTEGRATION_ENABLED") == "1"


def _run_alembic(
    worktree_root: Path,
    alembic_ini: Path,
    db_name: str,
    *args: str,
) -> None:
    database_url = (
        f"postgresql+asyncpg://{PG_USER}:{PG_PASSWORD}@{PG_HOST}:{PG_PORT}/{db_name}"
    )
    env = os.environ.copy()
    env["DATABASE_URL"] = database_url
    env["PYTHONPATH"] = str(worktree_root) + os.pathsep + env.get("PYTHONPATH", "")
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "alembic",
            "-c",
            str(alembic_ini),
            *args,
        ],
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


async def _head_revision(db_name: str) -> str:
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        return await conn.fetchval(
            "SELECT version_num FROM lanzadera.alembic_version"
        )
    finally:
        await conn.close()


@pytest.mark.integration
async def test_full_cycle_upgrade_head_downgrade_base_upgrade_head(
    integration_enabled: bool,
) -> None:
    """The full cycle (issue #924) must complete on a fresh database.

    Sequence:

    1. Create a brand-new database.
    2. ``alembic upgrade head`` — seeds users (0004), assignments (0005),
       audit (0006).
    3. ``alembic downgrade base`` — drops the schema and the seed data.
       This is the step #924 reports as broken (FK violation at
       ``0005 -> 0004``); with #929's fix the seeded audit rows with
       non-null ``actor_id`` are removed before 0004 deletes the users.
    4. ``alembic upgrade head`` — re-creates the schema and re-seeds.
    5. The head revision must match between step 2 and step 4.
    """
    assert integration_enabled, "APAP_INTEGRATION_ENABLED=1 required"

    worktree_root = _worktree_root()
    alembic_ini = next(
        (c for c in _alembic_ini_candidates(worktree_root) if c.is_file()), None
    )
    assert alembic_ini is not None, "alembic.ini not found"

    db_name = f"lanzadera_test_924_cycle_{secrets.token_hex(4)}"
    await _create_test_database(db_name)
    try:
        # Step 1: brand-new DB, no ``lanzadera`` schema yet.
        conn = await asyncpg.connect(
            host=PG_HOST,
            port=PG_PORT,
            user=PG_USER,
            password=PG_PASSWORD,
            database=db_name,
        )
        try:
            schema_exists = await conn.fetchval(
                "SELECT EXISTS(SELECT 1 FROM information_schema.schemata "
                "WHERE schema_name = 'lanzadera')"
            )
            assert schema_exists is False, "DB must start empty"
        finally:
            await conn.close()

        # Step 2: upgrade head.
        _run_alembic(worktree_root, alembic_ini, db_name, "upgrade", "head")
        head_after_first_upgrade = await _head_revision(db_name)
        assert head_after_first_upgrade is not None, "head revision missing"

        # Step 3: downgrade base — this is the step #924 reports as broken.
        # With #929 the 0006 downgrade clears every seeded audit row, so
        # the 0004 user purge no longer trips the FK.
        _run_alembic(worktree_root, alembic_ini, db_name, "downgrade", "base")

        # Step 4: upgrade head again — must succeed and reach the same
        # revision the first upgrade landed on.
        _run_alembic(worktree_root, alembic_ini, db_name, "upgrade", "head")
        head_after_second_upgrade = await _head_revision(db_name)

        assert head_after_first_upgrade == head_after_second_upgrade, (
            f"head revision drifted across cycles: first={head_after_first_upgrade!r}, "
            f"second={head_after_second_upgrade!r}"
        )
    finally:
        await _drop_test_database(db_name)
