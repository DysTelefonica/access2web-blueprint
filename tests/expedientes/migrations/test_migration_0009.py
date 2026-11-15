"""Integration test for Alembic migration 0009 (Expedientes header).

Issue #801 (hole 2 of the #802 epic) and #803: migration 0009 creates the
header of the Expedientes aggregate (``TbExpedientes``) inside the
``expedientes`` schema that ``0002`` opened. The test runs against a real
PostgreSQL (never SQLite): it creates a fresh database, runs ``alembic
upgrade head``, asserts the structure through ``information_schema`` and
``pg_enum``, proves that ``alembic downgrade -1`` leaves no residue and drops
the database on teardown.

The database plumbing lives in ``_pg.py``; the gate on
``APAP_INTEGRATION_ENABLED=1`` fails closed (issue #882).
"""

from __future__ import annotations

import asyncio

import pytest

from tests.expedientes.migrations._pg import (
    alembic_ini,
    columns,
    create_test_database,
    drop_test_database,
    enums,
    require_integration,
    run_alembic,
    tables,
    unique_db_name,
    worktree_root,
)

SCHEMA = "expedientes"

# The table this migration creates.
HEADER_TABLE = "TbExpedientes"

# The fourteen catalog tables of migration 0002: they must survive untouched
# (#803: the schema never duplicates a table of 0002).
CATALOG_TABLES = tuple(
    "TbEstados TbGradosClasificacion TbComerciales TbCPV TbLugaresEjecucion "
    "TbOrganosContratacion TbOficinasPrograma TbEjercitos TbJefaturas TbJuridicas "
    "TbResponsablesPorRol TbSuministradores TbPECAL TbRACS".split()
)

# Columns pinned by the aggregate and its invariants
# (``app/src/modules/expedientes/domain/expediente.py``).
EXPECTED_COLUMNS = {
    "id": "uuid",
    "tipo": "USER-DEFINED",
    "estado": "USER-DEFINED",
    "version": "integer",
    "id_expediente_padre": "uuid",
    "fecha_inicio_contrato": "date",
    "fecha_fin_contrato": "date",
    "created_at": "timestamp with time zone",
    "updated_at": "timestamp with time zone",
}

# The ENUM literals mirror the domain enums ExpedienteTipo and ExpedienteEstado.
EXPECTED_ENUM_VALUES = {
    "expediente_tipo": ("AM", "LOTE", "BASED"),
    "expediente_estado": ("BORRADOR", "ADJUDICADO", "FORMALIZADO", "ARCHIVADO"),
}


@pytest.mark.integration
def test_upgrade_head_adds_the_header_without_touching_the_catalogs(migrated_db: str) -> None:
    """The schema holds the fourteen catalogs of 0002 plus the new header."""
    found = asyncio.run(tables(migrated_db, SCHEMA))
    assert found == set(CATALOG_TABLES) | {HEADER_TABLE}, (
        "the expedientes schema must hold exactly the fourteen catalogs of 0002 "
        f"plus {HEADER_TABLE}, found {sorted(found)}"
    )


@pytest.mark.integration
def test_columns_match_the_aggregate(migrated_db: str) -> None:
    """The header exposes exactly the columns the aggregate declares."""
    found = asyncio.run(columns(migrated_db, SCHEMA)).get(HEADER_TABLE, {})
    assert found == EXPECTED_COLUMNS, (
        f"{HEADER_TABLE} columns are {found}; the aggregate declares {EXPECTED_COLUMNS}"
    )


@pytest.mark.integration
def test_enum_values_match_the_domain(migrated_db: str) -> None:
    """The ENUM literals round-trip with the domain enums."""
    found = asyncio.run(enums(migrated_db, SCHEMA))
    for enum_name, expected in EXPECTED_ENUM_VALUES.items():
        assert found.get(enum_name) == expected, (
            f"{enum_name} values are {found.get(enum_name)}; the domain declares {expected}"
        )


@pytest.mark.integration
def test_downgrade_removes_the_header_without_residue() -> None:
    """`downgrade -1` leaves the catalogs and no trace of the header."""
    require_integration()
    root = worktree_root()
    config = alembic_ini(root)
    name = unique_db_name("0009_down")
    asyncio.run(create_test_database(name))
    try:
        run_alembic(root, config, name, "upgrade", "head")
        upgraded = asyncio.run(tables(name, SCHEMA))
        assert HEADER_TABLE in upgraded, (
            f"migration 0009 must create {HEADER_TABLE} before its downgrade "
            f"can remove it; the schema has {sorted(upgraded)}"
        )
        run_alembic(root, config, name, "downgrade", "-1")
        left = asyncio.run(tables(name, SCHEMA))
        assert HEADER_TABLE not in left, (
            f"downgrade must drop {HEADER_TABLE}; the schema still has {sorted(left)}"
        )
        assert left == set(CATALOG_TABLES), (
            f"downgrade must leave the fourteen catalogs of 0002 untouched, found {sorted(left)}"
        )
        types = asyncio.run(enums(name, SCHEMA))
        for enum_name in EXPECTED_ENUM_VALUES:
            assert enum_name not in types, (
                f"downgrade must drop the ENUM type {enum_name}; found {sorted(types)}"
            )
    finally:
        asyncio.run(drop_test_database(name))
