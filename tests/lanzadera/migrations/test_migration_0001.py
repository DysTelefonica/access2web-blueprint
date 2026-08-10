# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 3a
# DA-3 (no legacy_hash), DA-11 (no audit telemetry), D82 (Expand and Contract).
"""Integration test for Alembic migration 0001_core_schema.

The migration creates the `lanzadera` schema, three ENUMs (`user_status`,
`app_topology`, `app_registration_status`), nine tables (`users`, `apps`,
`profiles`, `user_app_assignments`, `global_admins`, `sessions`, `audit`,
`reset_tokens`, `mail_outbox`), and the six indexes the design pins
(D5/D14/D27/D56, design.md §Modelo de datos). No seed data lands here:
0001 is schema only; seeds arrive in 0002-0006 (PR 3b).

Each test creates its own fresh database via `CREATE DATABASE`, runs
`alembic upgrade head` against it, queries `information_schema` to assert
structure, and drops the database on teardown. Tests are gated on
`APAP_INTEGRATION_ENABLED=1` so CI defaults to skipping them; the local
runtime contract for PR 3a runs them with the flag set.

The tests use sync `def test_*` because the project's `pyproject.toml`
does not enable `pytest-asyncio`. Async I/O runs through `asyncio.run()`
inside the test body so the gate is the env flag, not the test runner.
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

# All production tables live under the `lanzadera` schema.
SCHEMA = "lanzadera"

# The nine tables that migration 0001 creates, in the order the design pins
# them (design.md §Modelo de datos).
EXPECTED_TABLES: tuple[str, ...] = (
    "users",
    "apps",
    "profiles",
    "user_app_assignments",
    "global_admins",
    "sessions",
    "audit",
    "reset_tokens",
    "mail_outbox",
)

# The three ENUMs created by 0001, with the exact values pinned by the spec.
EXPECTED_ENUMS: dict[str, tuple[str, ...]] = {
    "user_status": ("active", "disabled", "password_reset_required", "locked"),
    "app_topology": ("central", "office-nas"),
    "app_registration_status": ("pending", "active", "retired"),
}

# Six indexes design.md commits to. The first is the UNIQUE constraint on
# `users.email`; the rest are composite indexes for hot read paths.
EXPECTED_INDEX_NAMES: tuple[str, ...] = (
    "ix_users_email",
    "ix_reset_tokens_user_id_expires_at",
    "ix_user_app_assignments_user_id_app_id",
    "ix_audit_actor_id_created_at",
    "ix_audit_event_type_created_at",
    "ix_mail_outbox_status_created_at",
)

# DA-3: the canonical column names that MUST NOT appear in any of the nine
# tables. A reintroduction of any of these would be a regression to the
# legacy SHA256-without-salt path (H1) and is pinned here as a hard gate.
FORBIDDEN_COLUMNS: tuple[str, ...] = (
    "legacy_hash",
    "password_legacy",
    "pass_hash_v1",
    "old_password",
    "verify_legacy",
)

# DA-11: telemetry columns the design explicitly bans from `audit`. D55
# retired the inherited Access columns; a reintroduction would leak network
# identity into a table that lives behind `audit_log_pg` (DA-11).
FORBIDDEN_AUDIT_COLUMNS: tuple[str, ...] = (
    "ssid",
    "bssid",
    "gps_lat",
    "gps_lon",
    "machine_name",
    "ip_address",
)

# Database creation borrows the credentials shipped by docker-compose.yml
# (Phase 0). The port stays on the loopback because the test runs in WSL
# alongside the `docker compose up -d postgres` container.
PG_HOST = "127.0.0.1"
PG_PORT = 5432
PG_USER = "lanzadera"
PG_PASSWORD = "lanzadera-dev-only"


def _worktree_root() -> Path:
    """Locate the worktree root from this test file.

    Walks up until it finds `.git` (a FILE in worktrees, a DIRECTORY in full
    clones). Raises if neither is present — a missing marker means the test
    is being executed outside the worktree, which is a configuration bug.
    """
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / ".git"
        if candidate.is_dir() or candidate.is_file():
            return parent
    raise AssertionError("worktree root not found (no .git)")


def _alembic_ini_candidates(root: Path) -> list[Path]:
    """Alembic config locations to probe, in preference order.

    The orchestrator pre-resolved decision lands `script_location` in
    `app/pyproject.toml` under `[tool.alembic]`, but alembic 1.13 still
    expects an `alembic.ini` to bootstrap. We honour whichever the worktree
    ships.
    """
    return [
        root / "app" / "migrations" / "alembic.ini",
        root / "alembic.ini",
    ]


@pytest.fixture(scope="module")
def integration_enabled() -> bool:
    """Gate every test in this module on the integration env flag.

    `APAP_INTEGRATION_ENABLED=1` signals that a postgres container is
    reachable at `127.0.0.1:5432`. Without it, the suite stays green in CI
    while still requiring the tests to be opt-in.
    """
    return os.environ.get("APAP_INTEGRATION_ENABLED") == "1"


@pytest.fixture(scope="module")
def worktree_root() -> Path:
    return _worktree_root()


@pytest.fixture(scope="module")
def alembic_ini(worktree_root: Path) -> Path:
    for candidate in _alembic_ini_candidates(worktree_root):
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(
        "alembic.ini not found; PR 3a must ship app/migrations/alembic.ini"
    )


@pytest.fixture(scope="module")
def test_db_name() -> str:
    """Per-module test database name, unique per run.

    The 8-hex suffix avoids collisions when the suite runs twice without
    cleanup, and keeps the identifier well under Postgres' 63-char limit.
    """
    return f"lanzadera_test_0001_{secrets.token_hex(4)}"


async def _admin_conn() -> asyncpg.Connection:
    """Connection to the default `lanzadera` database (for CREATE/DROP DATABASE)."""
    return await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database="lanzadera",
    )


async def _create_test_database(name: str) -> None:
    """Create a fresh, empty database for the test run.

    The connection is closed before issuing the CREATE so asyncpg does not
    bind the new database to the active session. We also tolerate the
    "already exists" race by issuing DROP first; the secret-suffixed name
    makes accidental overlap effectively impossible.
    """
    conn = await _admin_conn()
    try:
        await conn.execute(f'DROP DATABASE IF EXISTS "{name}"')
        await conn.execute(f'CREATE DATABASE "{name}"')
    finally:
        await conn.close()


async def _drop_test_database(name: str) -> None:
    """Force-disconnect any lingering sessions and drop the test database.

    Postgres refuses to drop a database with active connections, so we
    terminate them via `pg_terminate_backend` first. Idempotent on a
    second invocation because the DROP is wrapped in `IF EXISTS`.
    """
    conn = await _admin_conn()
    try:
        await conn.execute(
            """
            SELECT pg_terminate_backend(pid)
              FROM pg_stat_activity
             WHERE datname = $1 AND pid <> pg_backend_pid()
            """,
            name,
        )
        await conn.execute(f'DROP DATABASE IF EXISTS "{name}"')
    finally:
        await conn.close()


def _run_alembic(
    worktree_root: Path, alembic_ini: Path, db_name: str, *args: str
) -> None:
    """Invoke `alembic` against the test database.

    We prepend `PYTHONPATH` so `app/migrations/env.py` can import
    `app.src...` symbols in future migrations. The subprocess inherits the
    test runner's environment minus a forcibly rewritten `DATABASE_URL`.
    """
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
def migrated_db(
    integration_enabled: bool,
    worktree_root: Path,
    alembic_ini: Path,
    test_db_name: str,
) -> str:
    """Run `alembic upgrade head` once per module against a fresh database.

    Skips when the integration flag is off so the suite stays green in CI.
    The fixture yields the test database name; tests connect directly to
    assert schema structure, then a module-finaliser drops the database.
    """
    if not integration_enabled:
        pytest.skip("APAP_INTEGRATION_ENABLED=1 required for migration tests")
    asyncio.run(_create_test_database(test_db_name))
    try:
        _run_alembic(worktree_root, alembic_ini, test_db_name, "upgrade", "head")
        yield test_db_name
    finally:
        asyncio.run(_drop_test_database(test_db_name))


async def _fetch_table_names(db_name: str) -> set[str]:
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        rows = await conn.fetch(
            """
            SELECT table_name
              FROM information_schema.tables
             WHERE table_schema = $1 AND table_type = 'BASE TABLE'
            """,
            SCHEMA,
        )
    finally:
        await conn.close()
    return {row["table_name"] for row in rows}


async def _fetch_enum_values(db_name: str, enum_name: str) -> tuple[str, ...]:
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        rows = await conn.fetch(
            """
            SELECT enumlabel
              FROM pg_enum e
              JOIN pg_type t ON t.oid = e.enumtypid
             WHERE t.typname = $1
             ORDER BY e.enumsortorder
            """,
            enum_name,
        )
    finally:
        await conn.close()
    return tuple(row["enumlabel"] for row in rows)


async def _fetch_index_names(db_name: str) -> set[str]:
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        rows = await conn.fetch(
            """
            SELECT indexname
              FROM pg_indexes
             WHERE schemaname = $1
            """,
            SCHEMA,
        )
    finally:
        await conn.close()
    return {row["indexname"] for row in rows}


async def _fetch_column_names(db_name: str, table: str) -> set[str]:
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        rows = await conn.fetch(
            """
            SELECT column_name
              FROM information_schema.columns
             WHERE table_schema = $1 AND table_name = $2
            """,
            SCHEMA,
            table,
        )
    finally:
        await conn.close()
    return {row["column_name"] for row in rows}


async def _fetch_column(db_name: str, table: str, column: str) -> asyncpg.Record | None:
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        return await conn.fetchrow(
            """
            SELECT column_name, data_type, is_nullable, column_default
              FROM information_schema.columns
             WHERE table_schema = $1 AND table_name = $2 AND column_name = $3
            """,
            SCHEMA,
            table,
            column,
        )
    finally:
        await conn.close()


async def _fetch_primary_key(db_name: str, table: str) -> list[str]:
    conn = await asyncpg.connect(
        host=PG_HOST,
        port=PG_PORT,
        user=PG_USER,
        password=PG_PASSWORD,
        database=db_name,
    )
    try:
        rows = await conn.fetch(
            """
            SELECT a.attname
              FROM pg_index i
              JOIN pg_attribute a
                ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey)
             WHERE i.indrelid = ($1 || '.' || $2)::regclass
               AND i.indisprimary
             ORDER BY array_position(i.indkey, a.attnum)
            """,
            SCHEMA,
            table,
        )
    finally:
        await conn.close()
    return [row["attname"] for row in rows]


# ---------------------------------------------------------------------------
# RED-first assertions: each test fails when the migration does not exist
# (the `migrated_db` fixture triggers alembic upgrade and fails the suite).
# ---------------------------------------------------------------------------


@pytest.mark.integration
def test_schema_lanzadera_exists(migrated_db: str) -> None:
    """The migration creates the `lanzadera` schema; assert it is present."""

    async def _check() -> str | None:
        conn = await asyncpg.connect(
            host=PG_HOST,
            port=PG_PORT,
            user=PG_USER,
            password=PG_PASSWORD,
            database=migrated_db,
        )
        try:
            row = await conn.fetchrow(
                "SELECT schema_name FROM information_schema.schemata WHERE schema_name = $1",
                SCHEMA,
            )
        finally:
            await conn.close()
        return row["schema_name"] if row else None

    assert asyncio.run(_check()) == SCHEMA, (
        "schema `lanzadera` missing after migration 0001"
    )


@pytest.mark.integration
def test_all_nine_tables_created(migrated_db: str) -> None:
    """All nine production tables exist under `lanzadera` after 0001.

    Alembic's bookkeeping table (`alembic_version`) also lives inside
    `lanzadera` — see `app/migrations/env.py` for the rationale. The
    test asserts the production set is a strict subset of the schema.
    """
    actual = asyncio.run(_fetch_table_names(migrated_db))
    expected = set(EXPECTED_TABLES)
    assert expected <= actual, (
        f"missing tables in `lanzadera`: {sorted(expected - actual)}; "
        f"got {sorted(actual)}"
    )
    # Hard pin: nothing extra inside `lanzadera` apart from the nine
    # production tables plus Alembic's bookkeeping.
    assert actual - expected == {"alembic_version"}, (
        f"unexpected extra tables: {sorted(actual - expected)}"
    )


@pytest.mark.integration
def test_three_enums_created_with_expected_values(migrated_db: str) -> None:
    """The three ENUMs have exactly the values the spec pins."""
    for enum_name, expected_values in EXPECTED_ENUMS.items():
        actual_values = asyncio.run(_fetch_enum_values(migrated_db, enum_name))
        assert tuple(actual_values) == expected_values, (
            f"enum `{enum_name}` values differ: "
            f"expected {expected_values}, got {tuple(actual_values)}"
        )


@pytest.mark.integration
def test_users_password_hash_is_nullable(migrated_db: str) -> None:
    """`users.password_hash` must be nullable (DA-3, D89).

    The column is filled by `consume_reset_token`; until then it stays NULL.
    A NOT NULL constraint here would make migration 0004 (PR 3b) crash when
    seeding the 156 users with `status='password_reset_required'`.
    """
    column = asyncio.run(_fetch_column(migrated_db, "users", "password_hash"))
    assert column is not None, "users.password_hash column missing"
    assert column["is_nullable"] == "YES", (
        f"users.password_hash must be nullable; got is_nullable="
        f"{column['is_nullable']!r}"
    )


@pytest.mark.integration
def test_users_status_default_is_password_reset_required(migrated_db: str) -> None:
    """`users.status` defaults to `password_reset_required` (D89).

    The default guarantees that every seeded user starts in the recovery
    state, regardless of who calls the create-user use case.
    """
    column = asyncio.run(_fetch_column(migrated_db, "users", "status"))
    assert column is not None, "users.status column missing"
    assert column["data_type"] == "USER-DEFINED", (
        f"users.status must be an ENUM; got data_type={column['data_type']!r}"
    )
    assert column["column_default"] is not None, (
        "users.status must have a server-side DEFAULT"
    )
    # Postgres renders ENUM defaults as `<enum-cast>'label'::text` or similar;
    # the substring check is robust to either form.
    assert "password_reset_required" in column["column_default"], (
        f"users.status default must be `password_reset_required`; "
        f"got {column['column_default']!r}"
    )


@pytest.mark.integration
def test_no_legacy_hash_column_anywhere(migrated_db: str) -> None:
    """DA-3: NO `legacy_hash` / `password_legacy` / `pass_hash_v1` columns.

    This is the regression pin for H1 (SHA256-without-salt legacy). A
    reintroduction would resurrect the attack surface D89 deliberately
    closes; we refuse to ship 0001 if any forbidden column appears.
    """
    for table in EXPECTED_TABLES:
        columns = asyncio.run(_fetch_column_names(migrated_db, table))
        forbidden_present = columns & set(FORBIDDEN_COLUMNS)
        assert not forbidden_present, (
            f"forbidden columns {sorted(forbidden_present)} found in "
            f"lanzadera.{table}; DA-3 bans legacy hash storage"
        )


@pytest.mark.integration
def test_audit_has_no_telemetry_columns(migrated_db: str) -> None:
    """DA-11: `audit` MUST NOT carry SSID/BS coordinates/machine/IP columns.

    D55 retired the inherited Access columns; reintroducing any of them
    would resurrect the privacy regression the design closes.
    """
    columns = asyncio.run(_fetch_column_names(migrated_db, "audit"))
    forbidden_present = columns & set(FORBIDDEN_AUDIT_COLUMNS)
    assert not forbidden_present, (
        f"forbidden telemetry columns {sorted(forbidden_present)} found in "
        f"lanzadera.audit; DA-11 bans telemetry at rest"
    )


@pytest.mark.integration
def test_all_six_indexes_created(migrated_db: str) -> None:
    """Six indexes design.md commits to exist in the `lanzadera` schema."""
    actual = asyncio.run(_fetch_index_names(migrated_db))
    expected = set(EXPECTED_INDEX_NAMES)
    assert expected <= actual, (
        f"missing indexes: {sorted(expected - actual)}; got {sorted(actual)}"
    )


@pytest.mark.integration
def test_global_admins_user_id_is_primary_key(migrated_db: str) -> None:
    """`global_admins.user_id` is the sole primary key column (D21, D42).

    The table is a membership relation with one row per global admin; its
    primary key is `user_id` so the membership test (`is_global_admin`) is
    O(1) and `grant` is upsert-friendly.
    """
    pk_columns = asyncio.run(_fetch_primary_key(migrated_db, "global_admins"))
    assert pk_columns == ["user_id"], (
        f"global_admins primary key must be [user_id]; got {pk_columns}"
    )


@pytest.mark.integration
def test_user_app_assignments_revoked_at_is_nullable(migrated_db: str) -> None:
    """`user_app_assignments.revoked_at` is nullable (DA-12, assignments/spec.md).

    Active assignments carry `revoked_at = NULL`; the column is filled when
    a global admin revokes a profile. NOT NULL would force every assignment
    to carry a sentinel timestamp and break the "active" view.
    """
    column = asyncio.run(
        _fetch_column(migrated_db, "user_app_assignments", "revoked_at")
    )
    assert column is not None, "user_app_assignments.revoked_at missing"
    assert column["is_nullable"] == "YES", (
        f"user_app_assignments.revoked_at must be nullable; "
        f"got is_nullable={column['is_nullable']!r}"
    )
