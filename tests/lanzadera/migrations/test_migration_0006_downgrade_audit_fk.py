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


# Local copy of the ``integration_enabled`` fixture from
# ``test_migration_0001``. Pytest does not auto-discover fixtures from
# sibling test modules; each module that uses a fixture must either
# define it or import it explicitly. Importing caused a redefinition
# warning (F811) and the import shape (``from ... import integration_enabled``)
# shadowed the parameter, so the fixture was not registered. Issue #901
# (runs 38807752784 and 38808296486).
@pytest.fixture
def integration_enabled() -> bool:
    """True when ``APAP_INTEGRATION_ENABLED=1`` is set in the test env."""
    return os.environ.get("APAP_INTEGRATION_ENABLED") == "1"


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
    database_url = f"postgresql+asyncpg://{PG_USER}:{PG_PASSWORD}@{PG_HOST}:{PG_PORT}/{db_name}"
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
            user_id_row = await conn.fetchrow("SELECT id FROM lanzadera.users ORDER BY id LIMIT 1")
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
            "downgrade wiped an unrelated audit row with a null actor (destructive — issue #928)"
        )
    finally:
        await _drop_test_database(db_name)


@pytest.mark.integration
async def test_0006_downgrade_does_not_touch_real_rows_when_upgrade_was_noop(
    integration_enabled: bool,
) -> None:
    """Simulate the case where ``upgrade()`` is a no-op (audit already populated).

    upgrade() returns without inserting when ``audit`` already has at least
    ``len(events)`` rows. downgrade() must still only target rows the
    upgrade actually wrote — the discriminating identity is the full
    matching of ``(id, event_type, actor_id, target_id, created_at)`` with
    the fixture (``module`` is not a discriminator because the application
    also writes ``module='lanzadera'`` on real events).

    Sequence:

    1. ``alembic upgrade head`` seeds the 200 audit rows from the fixture.
    2. Insert 200+ real rows whose ``id`` collides with seeded ids, using
       ``module='lanzadera'`` (what the app actually writes — see
       login.py, create_user.py, set_password.py) and a different
       ``event_type`` and ``created_at`` so the row does not match the
       fixture fingerprint.
    3. ``alembic downgrade 0005`` runs 0006's ``downgrade()``.
    4. The 200+ real rows must still be there, and the 200 seeded rows
       must have been removed (so the 0004 downgrade does not hit
       ``fk_audit_actor_id_users``).
    """
    assert integration_enabled, "APAP_INTEGRATION_ENABLED=1 required"

    worktree_root = _worktree_root()
    for candidate in _alembic_ini_candidates(worktree_root):
        if candidate.is_file():
            alembic_ini = candidate
            break
    else:
        pytest.skip("alembic.ini not found")

    db_name = f"lanzadera_test_0006_noop_{secrets.token_hex(4)}"
    await _create_test_database(db_name)
    try:
        _run_alembic(worktree_root, alembic_ini, db_name, "upgrade", "head")

        seeded = _seeded_event_ids(worktree_root)
        assert len(seeded) >= 200, (
            "fixture must yield at least 200 rows for this test to be meaningful"
        )

        # Pick the first 200 seeded ids and insert a real row for each one
        # with a colliding id, ``module='lanzadera'`` (what the app writes),
        # and a different ``event_type`` / ``created_at`` so the row does
        # not match the fixture fingerprint. ``module`` alone cannot
        # discriminate, hence the full-row match in the downgrade query.
        colliding_ids = seeded[:200]
        rows = [
            (
                id_,
                "auth.real-event",  # different from any fixture event_type
                None,
                999_999,  # different from any fixture target_id
                "lanzadera",  # what the app actually writes
                "success",
                f"real-{id_}",
                "{}",
                "2099-01-01T00:00:00+00:00",  # different from any fixture created_at
            )
            for id_ in colliding_ids
        ]

        conn = await asyncpg.connect(
            host=PG_HOST,
            port=PG_PORT,
            user=PG_USER,
            password=PG_PASSWORD,
            database=db_name,
        )
        try:
            await conn.executemany(
                """
                INSERT INTO lanzadera.audit
                  (id, event_type, actor_id, target_id, module,
                   result, correlation_id, payload, created_at)
                VALUES
                  ($1, $2, $3, $4, $5, $6, $7, $8::jsonb, $9::timestamptz)
                """,
                rows,
            )

            total_before = await conn.fetchval("SELECT count(*) FROM lanzadera.audit")
            real_count_before = await conn.fetchval(
                "SELECT count(*) FROM lanzadera.audit WHERE event_type = 'auth.real-event'"
            )
            assert real_count_before == 200
            assert total_before >= 400
        finally:
            await conn.close()

        # Downgrade only 0006 (so the schema above 0006 is preserved). The
        # important step is 0006's downgrade(), which the new fix narrows to
        # the full row match against the fixture.
        _run_alembic(worktree_root, alembic_ini, db_name, "downgrade", "0005")

        conn = await asyncpg.connect(
            host=PG_HOST,
            port=PG_PORT,
            user=PG_USER,
            password=PG_PASSWORD,
            database=db_name,
        )
        try:
            real_count_after = await conn.fetchval(
                "SELECT count(*) FROM lanzadera.audit WHERE event_type = 'auth.real-event'"
            )
            seeded_count_after = await conn.fetchval(
                "SELECT count(*) FROM lanzadera.audit "
                "WHERE event_type IN ('auth.login.success', 'auth.login.failure', 'app.open')"
            )
        finally:
            await conn.close()

        assert real_count_after == 200, (
            "downgrade() wiped a real audit row whose id collides with a "
            "seeded id (destructive when upgrade() was a no-op — issue #928)"
        )
        assert seeded_count_after == 0, (
            "downgrade() left seeded audit rows behind (would block 0004 via "
            "fk_audit_actor_id_users — issue #928)"
        )
    finally:
        await _drop_test_database(db_name)
