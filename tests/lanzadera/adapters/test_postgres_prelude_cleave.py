"""Architectural test for the W-series async-session preamble.

The W08..W20 prelude-cleave swept every Postgres adapter from the
4-line ``try/except/finally`` pattern to the ``async with
self._factory.read_only_session() as session:`` and ``async with
self._factory.transaction() as session:`` context-manager shapes.

This test pins that invariant: every Postgres adapter must build
its session lifecycle through the typed context-manager helper. A
regression that introduces a manual ``session = self._factory();
try: ... finally: await session.close()`` block fails this test,
because the new block is a string the test greps for.

The test is intentionally cheap (a single AST-string scan) so it
runs on every CI invocation. The migration cost (W08..W20) was
already paid in production code; this is the docstring-level
gate that prevents drift back to the old shape.
"""

from __future__ import annotations

from pathlib import Path

# Adapter file paths relative to the repo root. Mirrors the layout
# documented in openspec/changes/lanzadera-mvp/design.md §Layout.
ADAPTER_FILES = [
    "app/src/modules/lanzadera/adapters/persistence/repositories/user_repository_pg.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/app_repository_pg.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/profile_repository_pg.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/assignment_repository_pg.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/audit_log_pg.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/global_admin_repository_pg.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/reset_token_repository_pg.py",
    "app/src/modules/lanzadera/adapters/persistence/repositories/mail_queue_table_adapter.py",
]


def test_no_legacy_session_lifecycle_in_postgres_adapters() -> None:
    """None of the 8 Postgres adapters may carry the manual ``session: AsyncSession = self._factory()`` shape.

    The W08 prelude-cleave (PR #436) introduced
    ``AsyncSessionFactory.read_only_session()`` and ``transaction()``.
    Every Postgres adapter migrated to one of the two helpers in
    W09..W20 (PR #437..#449). A regression to the manual
    session-acquire shape would re-introduce the 4-line
    ``try/except/finally`` boilerplate and undermine DA-11
    (the helper owns commit-or-rollback semantics; a manual block
    silently bypasses it).

    The pattern we ban is the bare ``session: AsyncSession =
    self._factory()`` line. The legal uses are the two context-manager
    calls on the factory (``self._factory.read_only_session()`` and
    ``self._factory.transaction()``). The AdapterContextManager
    ``async with`` blocks are exactly where the session is consumed;
    the bare ``session:`` assignment no longer has a legal home.
    """
    repo_root = Path(__file__).resolve().parents[3]
    banned = "session: AsyncSession = self._factory()"
    offenders: list[tuple[str, int]] = []
    for rel in ADAPTER_FILES:
        path = repo_root / rel
        for line_no, line in enumerate(path.read_text().splitlines(), 1):
            if banned in line:
                offenders.append((rel, line_no))
    assert not offenders, (
        f"{len(offenders)} Postgres adapter(s) still carry the manual "
        f"session-lifecycle shape: {offenders}. Re-apply the W08 prelude-cleave "
        f"(PR #436) to migrate them to ``async with "
        f"self._factory.read_only_session() as session:`` (reads) or "
        f"``async with self._factory.transaction() as session:`` "
        f"(writes)."
    )
