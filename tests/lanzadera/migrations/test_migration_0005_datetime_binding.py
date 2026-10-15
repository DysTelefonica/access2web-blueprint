# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #911
"""Regression test for issue #911: 0005_seed_assignments.py must bind ``granted_at`` as a datetime, not a raw string.

The original migration bound ``row["granted_at"]`` (an ISO 8601 string
from the JSON fixture) directly to a ``sa.DateTime(timezone=True)``
column. asyncpg rejects strings for ``TIMESTAMPTZ``, raising
``sqlalchemy.exc.DBAPIError``:

  invalid input for query argument $5: '2025-06-01T00:00:00+00:00'
  (expected a datetime.date or datetime.datetime instance, got 'str')

The fix parses the ISO string with ``datetime.fromisoformat`` before
binding, so asyncpg receives a timezone-aware ``datetime``. The runner
ships Python 3.12.11 (verified in CI run 38782719315), and
``datetime.fromisoformat`` understands the ``+00:00`` suffix from
Python 3.11 onwards.

TDD contract (gentle-ai RDD):
  RED on ``origin/main`` — the migration binds ``row["granted_at"]``
  raw; the test fails.
  GREEN once the fix lands — the migration binds
  ``datetime.fromisoformat(row["granted_at"])``; the test passes.

This is a static unit test so the RED → GREEN cycle runs without a
live Postgres. The integration behaviour (``SELECT granted_at FROM
lanzadera.user_app_assignments LIMIT 1`` returning a ``datetime``
instance, not a ``str``) is exercised by
``tests/lanzadera/migrations/test_seed_migrations.py::TestSeedAssignments``
in CI under ``APAP_INTEGRATION_ENABLED=1``; once 0005 is fixed that
test goes from RED to GREEN alongside the unit test here.
"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
MIGRATION = ROOT / "app" / "migrations" / "versions" / "0005_seed_assignments.py"

# The fix introduces this exact binding. The single-quote spelling is
# the natural form when ``row["granted_at"]`` is already double-quoted.
_EXPECTED_BINDING = "datetime.fromisoformat(row['granted_at'])"
# Accept the equally correct double-quote spelling.
_EXPECTED_BINDING_ALT = 'datetime.fromisoformat(row["granted_at"])'

# The bug: passing the raw JSON string straight to the bind. asyncpg
# rejects strings for TIMESTAMPTZ, so any spelling of this pattern is
# forbidden after the fix lands.
_FORBIDDEN_DQUOTE = '"granted_at": row["granted_at"]'
_FORBIDDEN_SQUOTE = "'granted_at': row['granted_at']"


def test_0005_granted_at_parsed_as_datetime() -> None:
    """The ``granted_at`` parameter must be parsed before binding."""
    assert MIGRATION.is_file(), f"migration file missing: {MIGRATION}"
    src = MIGRATION.read_text(encoding="utf-8")

    # RED condition: the raw string reaches the bind call.
    assert _FORBIDDEN_DQUOTE not in src, (
        "0005_seed_assignments.py binds `row[\"granted_at\"]` (a JSON "
        "string) directly to a TIMESTAMPTZ column. asyncpg rejects "
        "strings for TIMESTAMPTZ. Parse with `datetime.fromisoformat` "
        "before binding. See issue #911."
    )
    assert _FORBIDDEN_SQUOTE not in src, (
        "0005_seed_assignments.py binds `row['granted_at']` raw to a "
        "TIMESTAMPTZ column. See issue #911."
    )

    # GREEN condition: the migration binds a parsed datetime.
    assert (
        _EXPECTED_BINDING in src or _EXPECTED_BINDING_ALT in src
    ), (
        "0005_seed_assignments.py must parse the `granted_at` ISO "
        "string with `datetime.fromisoformat(row[\"granted_at\"])` "
        "before binding. See issue #911."
    )


def test_0005_datetime_imported() -> None:
    """The migration must import ``datetime`` to parse the timestamp."""
    src = MIGRATION.read_text(encoding="utf-8")
    # ``datetime.fromisoformat`` only resolves when ``datetime`` is in
    # scope. Accept either spelling of the import.
    assert (
        "from datetime import datetime" in src
        or "import datetime" in src
    ), (
        "0005_seed_assignments.py must import `datetime` (either "
        "`from datetime import datetime` or `import datetime`) so "
        "`datetime.fromisoformat(...)` resolves. See issue #911."
    )