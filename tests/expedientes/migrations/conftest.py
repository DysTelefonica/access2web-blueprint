"""Fixtures for the Expedientes migration tests (issue #801, #803).

The database plumbing lives in ``_pg.py``; the fixture is declared here so
the test modules use ``migrated_db`` without importing it (an imported name
would shadow the fixture parameter of every test).
"""

from __future__ import annotations

import asyncio

import pytest

from tests.expedientes.migrations._pg import (
    alembic_ini,
    create_test_database,
    drop_test_database,
    require_integration,
    run_alembic,
    unique_db_name,
    worktree_root,
)


@pytest.fixture(scope="module")
def migrated_db() -> str:
    """Run `alembic upgrade head` once per module against a fresh database."""
    require_integration()
    root = worktree_root()
    config = alembic_ini(root)
    name = unique_db_name("0009")
    asyncio.run(create_test_database(name))
    try:
        run_alembic(root, config, name, "upgrade", "head")
        yield name
    finally:
        asyncio.run(drop_test_database(name))
