# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #928
"""Regression test for issue #928: 0006 downgrade must target seeded ids only.

The previous downgrade() used a heuristic (``DELETE WHERE actor_id IS NULL``)
that wiped unrelated audit rows with a null actor and left seeded rows with a
non-null actor blocking the 0005 -> 0004 downgrade via
``fk_audit_actor_id_users``. The fix makes downgrade() read the same fixture
as upgrade() and delete exactly those ids.

TDD contract (gentle-ai RDD):
  RED on ``origin/main`` — ``alembic downgrade -1`` raises
  ``ForeignKeyViolationError`` because seeded audit rows still reference users.
  GREEN once the fix lands — ``alembic downgrade -1`` succeeds and the two
  unrelated rows inserted by the test survive untouched.
"""

from __future__ import annotations

import json
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


async def _admin_conn() -> asyncpg.Connection:
    return await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database="lanzadera",
    )


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


def _seeded_event_ids(worktree_root: Path) -> list[int]:
    fixture = worktree_root / "data" / "fixtures" / "lanzadera" / "audit_events.json"
    return [event["id"] for event in json.loads(fixture.read_text())]


@pytest.mark.integration
async def test_0006_downgrade_keeps_unrelated_audit_rows_and_lets_0004_run(
    integration_enabled: bool,
) -> None:
    """Insert two unrelated audit rows, downgrade to 0005, then to base.

    The fix should:

    * Not raise ``ForeignKeyViolationError`` at the 0006 -> 0005 step (no
      seeded row is left behind with a non-null ``actor_id`` pointing at a
      user that 0004 will delete).
    * Leave the two rows we inserted here untouched by the downgrade.
    """
    assert integration_enabled, "APAP_INTEGRATION_ENABLED=1 required"

    worktree_root = _worktree_root()
    for candidate in _alembic_ini_candidates(worktree_root):
        alembic_ini = candidate if candidate.is_file() else None
        if alembic_ini is not None:
            break
    assert alembic_ini is not None, "alembic.ini not found"

    db_name = f"lanzadera_test_0006_fk_{secrets.token_hex(4)}"
    await _create_test_database(db_name)
    try:
        _run_alembic(worktree_root, alembic_ini, db_name, "upgrade", "head")

        # Pick ids that are far from anything the fixture uses. We do not
        # know the fixture's id range, so use a distinct high range and let
        # the assertion confirm they are not in the seeded set.
        seeded = set(_seeded_event_ids(worktree_root))
        unrelated_with_actor = 900_000_001
        unrelated_without_actor = 900_000_002
        assert unrelated_with_actor not in seeded
        assert unrelated_without_actor not in seeded

        conn = await asyncpg.connect(
            host=PG_HOST,
            port=PG_PORT,
            user=PG_USER,
            password=PG_PASSWORD,
            database=db_name,
        )
        try:
            # We need a real user to satisfy the FK for the actor_id row.
            user_id_row = await conn.fetchrow(
                "SELECT id FROM lanzadera.users ORDER BY id LIMIT 1"
            )
            assert user_id_row is not None, "0004 must have seeded at least one user"
            real_user_id = user_id_row["id"]

            await conn.execute(
                """
                INSERT INTO lanzadera.audit
                  (id, event_type, actor_id, target_id, module,
                   result, correlation_id, payload, created_at)
                VALUES
                  ($1, 'app.open', $2, 1, 'real',
                   'success', 'real-actor', '{}'::jsonb, now()),
                  ($3, 'app.open', NULL, 1, 'real',
                   'success', 'real-anon', '{}'::jsonb, now())
                """,
                unrelated_with_actor,
                real_user_id,
                unrelated_without_actor,
            )
        finally:
            await conn.close()

        # The old downgrade() left a non-null actor_id row behind. 0005
        # then tried to delete users and the FK raised. The new downgrade()
        # deletes only seeded ids, so the FK is clear.
        _run_alembic(worktree_root, alembic_ini, db_name, "downgrade", "base")

        conn = await asyncpg.connect(
            host=PG_HOST,
            port=PG_PORT,
            user=PG_USER,
            password=PG_PASSWORD,
            database=db_name,
        )
        try:
            surviving_with_actor = await conn.fetchval(
                "SELECT count(*) FROM lanzadera.audit WHERE id = $1",
                unrelated_with_actor,
            )
            surviving_without_actor = await conn.fetchval(
                "SELECT count(*) FROM lanzadera.audit WHERE id = $1",
                unrelated_without_actor,
            )
        finally:
            await conn.close()

        assert surviving_with_actor == 1, (
            "downgrade wiped an unrelated audit row with a non-null actor "
            "(destructive — issue #928)"
        )
        assert surviving_without_actor == 1, (
            "downgrade wiped an unrelated audit row with a null actor "
            "(destructive — issue #928)"
        )
    finally:
        await _drop_test_database(db_name)
