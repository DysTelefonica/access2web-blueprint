# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp #916
"""Regression test for issue #916: migration tests read the integration DSN.

The integration job publishes Postgres on a dynamic host port and exports
the real DSN in ``APAP_INTEGRATION_URL`` and ``INTEGRATION_DSN``. The
migration tests must take host and port from that DSN, not from a
hardcoded ``PG_PORT = 5432``, which has no listener on the CI runner.

TDD contract (gentle-ai RDD):
  RED on ``origin/main`` — ``_pg_endpoint`` does not exist; the import fails.
  GREEN once the fix lands — the helper parses the DSN and defaults to
  ``127.0.0.1:5432`` when no variable is set.
"""

from __future__ import annotations

import pytest

from tests.lanzadera.migrations.test_migration_0001 import _pg_endpoint

DSN_FROM_JOB = "postgresql+asyncpg://lanzadera:lanzadera-dev-only@127.0.0.1:32768/lanzadera"


@pytest.fixture
def clean_env(monkeypatch: pytest.MonkeyPatch) -> pytest.MonkeyPatch:
    monkeypatch.delenv("APAP_INTEGRATION_URL", raising=False)
    monkeypatch.delenv("INTEGRATION_DSN", raising=False)
    return monkeypatch


def test_pg_endpoint_reads_port_from_apap_integration_url(clean_env: pytest.MonkeyPatch) -> None:
    clean_env.setenv("APAP_INTEGRATION_URL", DSN_FROM_JOB)
    assert _pg_endpoint() == ("127.0.0.1", 32768)


def test_pg_endpoint_falls_back_to_integration_dsn(clean_env: pytest.MonkeyPatch) -> None:
    clean_env.setenv("INTEGRATION_DSN", DSN_FROM_JOB)
    assert _pg_endpoint() == ("127.0.0.1", 32768)


def test_pg_endpoint_defaults_to_loopback_5432_without_env(clean_env: pytest.MonkeyPatch) -> None:
    assert _pg_endpoint() == ("127.0.0.1", 5432)
