# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #904
"""Regression test for issue #904: 0004_seed_users.py must tolerate a missing PLATFORM_SECRET_KEY.

The original migration read the Fernet key with
``os.environ["PLATFORM_SECRET_KEY"]`` (index access), which raises
``KeyError`` when the env var is unset. The first CI run that
actually executed alembic against a real Postgres service (PR-3 of
#882, run 38775229896) crashed at the 0004 -> 0005 step with:

  File ".../0004_seed_users.py", line 47, in _load_key
      raw = os.environ["PLATFORM_SECRET_KEY"]
  KeyError: 'PLATFORM_SECRET_KEY'

The dev fixture
(``tests/lanzadera/migrations/test_seed_migrations.py::_FIXTURE_KEY``)
and ``docker-compose.yml`` both export the key, so the migration
worked locally; the bug only surfaces for callers (like the new
integration job in PR-3 of #882) that forget the export.

The fix gives the migration a default key — the same Fernet-format
value the test fixture uses, ``base64.urlsafe_b64encode(b"0" * 32)
.decode()`` — so the migration applies cleanly on a clean database
without any environment setup. When ``PLATFORM_SECRET_KEY`` IS set
(``docker-compose`` dev, the test fixture, or any production
deployment), the explicit value still wins via ``os.environ.get``.

TDD contract (gentle-ai RDD):
  RED on ``origin/main`` — the migration uses index access, so the
  test fails.
  GREEN once the fix lands — the migration uses ``os.environ.get``
  with a default that decodes to 32 zero bytes (the same Fernet key
  the test fixture uses).
"""

from __future__ import annotations

import base64
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
MIGRATION = ROOT / "app" / "migrations" / "versions" / "0004_seed_users.py"

# The default the fix introduces: a Fernet-format key derived from 32
# zero bytes. Must equal the test fixture's `_FIXTURE_KEY` so the
# integration tests (which decrypt DNI values with that key) keep
# round-tripping once alembic has run.
EXPECTED_DEFAULT_KEY = base64.urlsafe_b64encode(b"0" * 32).decode()


def test_0004_load_key_uses_get_with_default() -> None:
    """``_load_key`` must use ``os.environ.get`` (tolerant), not index access."""
    assert MIGRATION.is_file(), f"migration file missing: {MIGRATION}"
    src = MIGRATION.read_text(encoding="utf-8")

    # The bug: index access raises KeyError when the env var is unset.
    # The fix: ``os.environ.get(..., default)`` returns the default.
    assert 'os.environ["PLATFORM_SECRET_KEY"]' not in src, (
        "0004_seed_users.py uses `os.environ['PLATFORM_SECRET_KEY']` "
        "(index access) which raises KeyError when the env var is unset. "
        "See issue #904 — use `os.environ.get('PLATFORM_SECRET_KEY', "
        "<default>)` so the migration applies on a clean database."
    )
    assert "os.environ.get" in src and "PLATFORM_SECRET_KEY" in src, (
        "0004_seed_users.py must read PLATFORM_SECRET_KEY with "
        "`os.environ.get(...)` so a missing env var falls back to a "
        "default Fernet key. See issue #904."
    )


def test_0004_default_key_matches_test_fixture() -> None:
    """The default key must be the same Fernet key the test fixture uses.

    The integration tests in ``test_seed_migrations.py`` decrypt
    ``dni_encrypted`` with ``_FIXTURE_KEY``. If the migration uses a
    different default, the round-trip breaks and CI fails downstream.
    Either spelling is acceptable — a literal string or the computed
    ``base64.urlsafe_b64encode(b"0" * 32).decode()`` form (single or
    double quotes) — as long as the value equals the test fixture's
    key.
    """
    src = MIGRATION.read_text(encoding="utf-8")
    literal_match = EXPECTED_DEFAULT_KEY in src
    computed_dquote = 'base64.urlsafe_b64encode(b"0" * 32)' in src
    computed_squote = "base64.urlsafe_b64encode(b'0' * 32)" in src
    has_decode = ".decode()" in src
    computed_match = (computed_dquote or computed_squote) and has_decode
    assert literal_match or computed_match, (
        f"0004_seed_users.py must default PLATFORM_SECRET_KEY to the "
        f"same Fernet key the test fixture uses. Either spell the "
        f"value as the literal `{EXPECTED_DEFAULT_KEY}` or compute it "
        f"as `base64.urlsafe_b64encode(b'0' * 32).decode()` (single or "
        f"double quotes). See issue #904."
    )
