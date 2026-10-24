"""Integration test for Alembic migration 0009 (Expedientes lifecycle schema).

Issue #801 (hole 2 of the #802 epic) and #803: migration 0009 creates the three
tables the lifecycle use cases write — the ``TbExpedientes`` header, the
``TbCambios`` history and the ``TbUltimoCambio`` marker — inside the
``expedientes`` schema that ``0002`` opened. The test runs against a real
PostgreSQL (never SQLite): it creates a fresh database, runs ``alembic
upgrade head``, asserts the structure through ``information_schema`` and
``pg_enum``, proves that ``alembic downgrade -1`` leaves no residue and drops
the database on teardown.

The suite is gated on ``APAP_INTEGRATION_ENABLED=1`` and fails closed when the
flag is off (issue #882). The layout follows
``tests/lanzadera/migrations/test_migration_0001.py``.
"""

from __future__ import annotations

import asyncio
import os
import secrets
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlsplit

import asyncpg
import pytest

SCHEMA = "expedientes"

# The three tables this migration creates.
LIFECYCLE_TABLES = ("TbExpedientes", "TbCambios", "TbUltimoCambio")

# The fourteen catalog tables of migration 0002: they must survive untouched
# (#803: the schema never duplicates a table of 0002).
CATALOG_TABLES = tuple(
    "TbEstados TbGradosClasificacion TbComerciales TbCPV TbLugaresEjecucion "
    "TbOrganosContratacion TbOficinasPrograma TbEjercitos TbJefaturas TbJuridicas "
    "TbResponsablesPorRol TbSuministradores TbPECAL TbRACS".split()
)

# Columns pinned by the domain and the ports: TbExpedientes by
# ``domain/expediente.py``; TbCambios and TbUltimoCambio by
# ``ports/audit_log.py``.
EXPECTED_COLUMNS = {
    "TbExpedientes": (
        "id",
        "tipo",
        "estado",
        "version",
        "id_expediente_padre",
        "fecha_inicio_contrato",
        "fecha_fin_contrato",
        "created_at",
        "updated_at",
    ),
    "TbCambios": (
        "id",
        "nombre_tabla",
        "id_expediente",
        "nombre_campo",
        "valor_inicial",
        "valor_final",
        "fecha_cambio",
        "id_usuario_cambio",
        "accion",
    ),
    "TbUltimoCambio": (
        "id",
        "id_expediente",
        "fecha_cambio",
        "id_usuario_cambio",
    ),
}

# The ENUM literals mirror the domain enums ExpedienteTipo and ExpedienteEstado.
EXPECTED_ENUM_VALUES = {
    "expediente_tipo": ("AM", "LOTE", "BASED"),
    "expediente_estado": ("BORRADOR", "ADJUDICADO", "FORMALIZADO", "ARCHIVADO"),
}

# Columns that must be UUID: the eighteen ports of the module type the
# identifiers as UUID, and the deployed modern schema (0001) does the same.
UUID_COLUMNS = (
    ("TbExpedientes", "id"),
    ("TbExpedientes", "id_expediente_padre"),
    ("TbCambios", "id"),
    ("TbCambios", "id_expediente"),
    ("TbCambios", "id_usuario_cambio"),
    ("TbUltimoCambio", "id"),
    ("TbUltimoCambio", "id_expediente"),
    ("TbUltimoCambio", "id_usuario_cambio"),
)

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


async def _fetch(db_name: str, sql: str, *args: object) -> list[asyncpg.Record]:
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


async def _create_test_database(name: str) -> None:
    """Create a fresh, empty database for the test run."""
    await _admin(f'DROP DATABASE IF EXISTS "{name}"')
    await _admin(f'CREATE DATABASE "{name}"')


async def _drop_test_database(name: str) -> None:
    """Force-disconnect any lingering sessions and drop the test database."""
    await _admin(
        "SELECT pg_terminate_backend(pid) FROM pg_stat_activity "
        "WHERE datname = $1 AND pid <> pg_backend_pid()",
        name,
    )
    await _admin(f'DROP DATABASE IF EXISTS "{name}"')


async def _table_names(db_name: str) -> set[str]:
    rows = await _fetch(
        db_name,
        "SELECT table_name FROM information_schema.tables "
        "WHERE table_schema = $1 AND table_type = 'BASE TABLE'",
        SCHEMA,
    )
    return {row["table_name"] for row in rows}


async def _column_names(db_name: str, table: str) -> set[str]:
    rows = await _fetch(
        db_name,
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema = $1 AND table_name = $2",
        SCHEMA,
        table,
    )
    return {row["column_name"] for row in rows}


async def _column_data_type(db_name: str, table: str, column: str) -> str | None:
    rows = await _fetch(
        db_name,
        "SELECT data_type FROM information_schema.columns "
        "WHERE table_schema = $1 AND table_name = $2 AND column_name = $3",
        SCHEMA,
        table,
        column,
    )
    return rows[0]["data_type"] if rows else None


async def _enum_values(db_name: str, enum_name: str) -> tuple[str, ...]:
    rows = await _fetch(
        db_name,
        "SELECT enumlabel FROM pg_enum e JOIN pg_type t ON t.oid = e.enumtypid "
        "WHERE t.typname = $1 ORDER BY e.enumsortorder",
        enum_name,
    )
    return tuple(row["enumlabel"] for row in rows)


async def _type_names(db_name: str) -> set[str]:
    """Every ENUM type of the schema: the residue check of a downgrade."""
    rows = await _fetch(
        db_name,
        "SELECT t.typname FROM pg_type t JOIN pg_namespace n ON n.oid = t.typnamespace "
        "WHERE n.nspname = $1 AND t.typtype = 'e'",
        SCHEMA,
    )
    return {row["typname"] for row in rows}


def _worktree_root() -> Path:
    """Locate the worktree root from this test file."""
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / ".git"
        if candidate.is_dir() or candidate.is_file():
            return parent
    raise AssertionError("worktree root not found (no .git)")


def _alembic_ini(root: Path) -> Path:
    """The Alembic config of this checkout."""
    for candidate in (root / "app" / "migrations" / "alembic.ini", root / "alembic.ini"):
        if candidate.is_file():
            return candidate
    pytest.skip("alembic.ini has not shipped yet")


def _run_alembic(root: Path, alembic_ini: Path, db_name: str, *args: str) -> None:
    """Invoke `alembic` against the test database."""
    database_url = f"postgresql+asyncpg://{PG_USER}:{PG_PASSWORD}@{PG_HOST}:{PG_PORT}/{db_name}"
    env = os.environ.copy()
    env["DATABASE_URL"] = database_url
    env["PYTHONPATH"] = str(root) + os.pathsep + env.get("PYTHONPATH", "")
    result = subprocess.run(
        [sys.executable, "-m", "alembic", "-c", str(alembic_ini), *args],
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


@pytest.fixture(scope="module")
def integration_enabled() -> bool:
    """Gate every test in this module on the integration env flag."""
    return os.environ.get("APAP_INTEGRATION_ENABLED") == "1"


@pytest.fixture(scope="module")
def worktree_root() -> Path:
    return _worktree_root()


@pytest.fixture(scope="module")
def alembic_ini(worktree_root: Path) -> Path:
    return _alembic_ini(worktree_root)


@pytest.fixture(scope="module")
def test_db_name() -> str:
    """Per-module test database name, unique per run."""
    return f"expedientes_test_0009_{secrets.token_hex(4)}"


@pytest.fixture(scope="module")
def migrated_db(
    integration_enabled: bool,
    worktree_root: Path,
    alembic_ini: Path,
    test_db_name: str,
) -> str:
    """Run `alembic upgrade head` once per module against a fresh database."""
    if not integration_enabled:
        pytest.fail(
            "Issue #882: APAP_INTEGRATION_ENABLED is not 1. The integration "
            "DSN gate must fail closed, not skip, so a missing Postgres "
            "shows up as a red CI instead of a silently green one."
        )
    asyncio.run(_create_test_database(test_db_name))
    try:
        _run_alembic(worktree_root, alembic_ini, test_db_name, "upgrade", "head")
        yield test_db_name
    finally:
        asyncio.run(_drop_test_database(test_db_name))


@pytest.mark.integration
def test_upgrade_head_creates_the_lifecycle_schema(migrated_db: str) -> None:
    """The schema holds the fourteen catalogs of 0002 plus the three new tables."""
    tables = asyncio.run(_table_names(migrated_db))
    assert tables == set(CATALOG_TABLES) | set(LIFECYCLE_TABLES), (
        "the expedientes schema must hold exactly the fourteen catalogs of 0002 "
        f"plus the three tables of 0009, found {sorted(tables)}"
    )


@pytest.mark.integration
@pytest.mark.parametrize("table", sorted(EXPECTED_COLUMNS))
def test_columns_match_the_domain_and_the_ports(migrated_db: str, table: str) -> None:
    """Every table exposes exactly the columns its contract declares."""
    columns = asyncio.run(_column_names(migrated_db, table))
    assert columns == set(EXPECTED_COLUMNS[table]), (
        f"{table} columns are {sorted(columns)}; "
        f"the contract declares {sorted(EXPECTED_COLUMNS[table])}"
    )


@pytest.mark.integration
@pytest.mark.parametrize(("table", "column"), UUID_COLUMNS)
def test_identifiers_are_uuid_columns(migrated_db: str, table: str, column: str) -> None:
    """The eighteen ports type identifiers as UUID; the columns must match."""
    data_type = asyncio.run(_column_data_type(migrated_db, table, column))
    assert data_type == "uuid", (
        f"{table}.{column} is {data_type!r}; the ports declare UUID identifiers"
    )


@pytest.mark.integration
@pytest.mark.parametrize("enum_name", sorted(EXPECTED_ENUM_VALUES))
def test_enum_values_match_the_domain(migrated_db: str, enum_name: str) -> None:
    """The ENUM literals round-trip with the domain enums."""
    values = asyncio.run(_enum_values(migrated_db, enum_name))
    assert values == EXPECTED_ENUM_VALUES[enum_name], (
        f"{enum_name} values are {values}; the domain declares {EXPECTED_ENUM_VALUES[enum_name]}"
    )


@pytest.mark.integration
def test_downgrade_removes_the_new_tables_without_residue(
    integration_enabled: bool,
    worktree_root: Path,
    alembic_ini: Path,
) -> None:
    """`downgrade -1` leaves the catalogs and no trace of the new tables."""
    if not integration_enabled:
        pytest.fail(
            "Issue #882: APAP_INTEGRATION_ENABLED is not 1. The integration "
            "DSN gate must fail closed, not skip."
        )
    name = f"expedientes_test_0009_down_{secrets.token_hex(4)}"
    asyncio.run(_create_test_database(name))
    try:
        _run_alembic(worktree_root, alembic_ini, name, "upgrade", "head")
        upgraded = asyncio.run(_table_names(name))
        for table in LIFECYCLE_TABLES:
            assert table in upgraded, (
                f"migration 0009 must create {table} before its downgrade can "
                f"remove it; the schema has {sorted(upgraded)}"
            )
        _run_alembic(worktree_root, alembic_ini, name, "downgrade", "-1")
        tables = asyncio.run(_table_names(name))
        for table in LIFECYCLE_TABLES:
            assert table not in tables, (
                f"downgrade must drop {table}; the schema still has {sorted(tables)}"
            )
        assert tables == set(CATALOG_TABLES), (
            f"downgrade must leave the fourteen catalogs of 0002 untouched, found {sorted(tables)}"
        )
        types = asyncio.run(_type_names(name))
        for enum_name in EXPECTED_ENUM_VALUES:
            assert enum_name not in types, (
                f"downgrade must drop the ENUM type {enum_name}; found {sorted(types)}"
            )
    finally:
        asyncio.run(_drop_test_database(name))
