"""Strict TDD — ExpedientesContainer production guard (F03, issue #224).

D-EXP-3: "Fakes deterministas solo en tests/dev y readiness falla si se
enlazan en producción".

These tests verify the deny-by-default boundary between dev and prod: the
container must REFUSE to bind fakes when ``app_env == "production"``.
"""

from __future__ import annotations

import pytest

from app.src.modules.expedientes.di.config import AppEnv, ExpedientesConfig, ProductionFakeRefused
from app.src.modules.expedientes.di.container import ExpedientesContainer
from tests.expedientes.adapters._fakes import (
    FakeAuditLog,
    FakeCatalogRepository,
    FakeDocumentStorage,
    FakeExpedienteRepository,
    FakeHitoRepository,
    FakeNotificationDelivery,
    FakeReadiness,
)


def _all_fakes() -> dict[str, object]:
    """Build the full set of in-memory fakes for F01 ports."""
    return {
        "expediente_repo": FakeExpedienteRepository(),
        "hito_repo": FakeHitoRepository(),
        "catalog_repo": FakeCatalogRepository(),
        "audit_log": FakeAuditLog(),
        "readiness": FakeReadiness(),
        "document_storage": FakeDocumentStorage(),
        "notification_delivery": FakeNotificationDelivery(),
    }


# ---------------------------------------------------------------------------
# Production guard — the F03 deny-by-default contract
# ---------------------------------------------------------------------------


def test_build_production_with_only_fakes_raises() -> None:
    """In production, the container MUST refuse to start with only fakes.

    D-EXP-3 + issue #224 acceptance: deny/no fake prod. This is the
    headline test of the F03 contract. The config is well-formed (db_url
    is set) so the rejection comes from the container's own guard, not
    from the config validator.
    """
    config = ExpedientesConfig(
        app_env=AppEnv.PRODUCTION,
        db_url="postgresql+asyncpg://user:pass@host/db",
    )
    with pytest.raises(ProductionFakeRefused):
        ExpedientesContainer.build(config=config, ports=_all_fakes())


def test_build_production_with_partial_fakes_raises() -> None:
    """Even one fake in production is enough to refuse the build."""
    config = ExpedientesConfig(
        app_env=AppEnv.PRODUCTION,
        db_url="postgresql+asyncpg://user:pass@host/db",
    )
    real_only = _all_fakes()
    real_only["audit_log"] = object()  # pretend this is a real adapter
    with pytest.raises(ProductionFakeRefused):
        ExpedientesContainer.build(config=config, ports=real_only)


def test_build_development_with_fakes_succeeds() -> None:
    """Dev environment accepts fakes so work can proceed without DB."""
    config = ExpedientesConfig(app_env=AppEnv.DEVELOPMENT)
    container = ExpedientesContainer.build(config=config, ports=_all_fakes())
    assert container.readiness is not None


def test_build_test_with_fakes_succeeds() -> None:
    """Test environment is the same as dev for wiring purposes."""
    config = ExpedientesConfig(app_env=AppEnv.TEST)
    container = ExpedientesContainer.build(config=config, ports=_all_fakes())
    assert container.expediente_repo is not None
