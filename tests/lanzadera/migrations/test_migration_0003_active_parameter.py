# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #904
"""Regression test for issue #904: 0003_seed_profiles.py must bind `active` as a parameter.

The original migration hardcoded a SQL literal ``'active'`` (a string)
in a ``sa.Boolean()`` column, raising
``asyncpg.exceptions.InvalidTextRepresentationError: invalid input
syntax for type boolean: "active"`` the first time the integration job
ran alembic against a real Postgres service (PR-3 of #882, CI run
38773849842). The fix reads the per-row boolean from the fixture
(``data/fixtures/lanzadera/profiles.json``) and binds it through the
``:active`` placeholder, so per-row fidelity is preserved.

TDD contract (gentle-ai RDD):
  RED on ``origin/main`` — the literal ``'active'`` sits inside the
  ``INSERT INTO ... VALUES`` clause, so the test fails.
  GREEN once the fix lands — the literal is gone from the VALUES
  clause, the ``:active`` placeholder is present, and the parameter
  dict carries ``"active": profile["active"]``.

This is a static unit test so the RED → GREEN cycle runs without a
live Postgres. The integration behaviour
(``SELECT active FROM lanzadera.profiles LIMIT 1`` returning a
``bool``) is exercised by ``test_seed_migrations.py::TestSeedProfiles``
in CI under ``APAP_INTEGRATION_ENABLED=1``; once 0003 is fixed that
test goes from RED to GREEN alongside the unit test here.
"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
MIGRATION = ROOT / "app" / "migrations" / "versions" / "0003_seed_profiles.py"

# Match the single ``INSERT INTO lanzadera.profiles ... VALUES (...)``
# statement inside the ``upgrade()`` body. ``re.DOTALL`` so ``.`` spans
# the multi-line raw string literal that holds the SQL.
_INSERT_PATTERN = re.compile(
    r"INSERT\s+INTO\s+\{SCHEMA\}\.profiles.*?VALUES\s*\((?P<values>.*?)\)\s*" "",
    re.DOTALL | re.IGNORECASE,
)


def test_0003_active_column_uses_parameter_binding() -> None:
    """0003 must bind the ``active`` column through ``:active``, not a literal.

    Issue #904: the original migration hardcoded ``'active'`` (a string)
    in a ``sa.Boolean()`` column, raising
    ``InvalidTextRepresentationError`` from asyncpg. The fix replaces
    the literal with a parameter binding fed from the fixture's
    per-row boolean.
    """
    assert MIGRATION.is_file(), f"migration file missing: {MIGRATION}"
    src = MIGRATION.read_text(encoding="utf-8")

    match = _INSERT_PATTERN.search(src)
    assert match is not None, (
        "Could not find `INSERT INTO {SCHEMA}.profiles ... VALUES (...)` "
        "in 0003_seed_profiles.py; update the test if the SQL shape changed."
    )
    values_clause = match.group("values")

    # The bug: a hardcoded SQL literal 'active' inside the VALUES clause.
    # The fix removes the literal and substitutes the :active placeholder.
    assert "'active'" not in values_clause, (
        "0003_seed_profiles.py hardcodes the string 'active' inside the "
        "VALUES clause. The `active` column is sa.Boolean() (see "
        "0001_core_schema.py); the value must come from the fixture "
        "via the `:active` parameter binding. See issue #904."
    )
    assert ":active" in values_clause, (
        "0003_seed_profiles.py must bind the active column with `:active` "
        "so the fixture's per-row boolean reaches the database. See issue #904."
    )


def test_0003_active_value_is_read_from_fixture() -> None:
    """0003 must pass ``profile['active']`` to the bind dict.

    The parameter binding alone is not enough — the value has to come
    from the fixture, otherwise every row collapses to the same default.
    Issue #904: 160 profiles × per-row boolean fidelity.
    """
    src = MIGRATION.read_text(encoding="utf-8")
    assert '"active": profile["active"]' in src or ("'active': profile['active']" in src), (
        "0003_seed_profiles.py must pass the fixture's per-row boolean "
        'to the bind dict (`"active": profile["active"]`). See issue #904.'
    )
