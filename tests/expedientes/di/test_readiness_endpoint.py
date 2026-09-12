"""Strict TDD — /readyz endpoint (F03, issue #224, CAP-047).

The runtime calls ``ReadinessPort.check()`` before declaring ``ready``.
D-EXP-3: the readiness endpoint MUST return 503 when any check fails —
including when fakes are bound in production (defence in depth alongside
the container's deny-by-default).
"""

from __future__ import annotations

from uuid import uuid4

from app.src.modules.expedientes.di.container import ExpedientesContainer
from app.src.modules.expedientes.di.config import AppEnv, ExpedientesConfig
from app.src.modules.expedientes.di.readiness import readiness_response
from app.src.modules.expedientes.ports.readiness import ReadinessCheck, ReadinessResult
from tests.expedientes.adapters._fakes import (
    FakeAuditLog,
    FakeCatalogRepository,
    FakeDocumentStorage,
    FakeExpedienteRepository,
    FakeHitoRepository,
    FakeNotificationDelivery,
    FakeReadiness,
    _UNCONFIGURED,
)


def _build_container_with_readiness(readiness: FakeReadiness) -> ExpedientesContainer:
    return ExpedientesContainer.build(
        config=ExpedientesConfig(app_env=AppEnv.TEST),
        ports={
            "expediente_repo": FakeExpedienteRepository(),
            "hito_repo": FakeHitoRepository(),
            "catalog_repo": FakeCatalogRepository(),
            "audit_log": FakeAuditLog(),
            "readiness": readiness,
            "document_storage": FakeDocumentStorage(),
            "notification_delivery": FakeNotificationDelivery(),
        },
    )


# ---------------------------------------------------------------------------
# /readyz response shape
# ---------------------------------------------------------------------------


async def test_readiness_returns_200_when_all_checks_pass() -> None:
    fake = FakeReadiness()
    fake._result = ReadinessResult(
        ready=True,
        checks=[ReadinessCheck(name="database", ok=True, message="ok")],
    )
    container = _build_container_with_readiness(fake)
    response = await readiness_response(container)
    assert response.status_code == 200
    assert response.body["ready"] is True


async def test_readiness_returns_503_when_a_check_fails() -> None:
    fake = FakeReadiness()
    fake._result = ReadinessResult(
        ready=False,
        checks=[ReadinessCheck(name="database", ok=False, message="connection refused")],
    )
    container = _build_container_with_readiness(fake)
    response = await readiness_response(container)
    assert response.status_code == 503
    assert response.body["ready"] is False
    assert response.body["checks"][0]["name"] == "database"
    assert response.body["checks"][0]["ok"] is False


async def test_readiness_returns_503_when_no_readiness_check_defined() -> None:
    """A container with no ReadinessPort check is itself not ready."""
    fake = FakeReadiness()
    fake._result = _UNCONFIGURED  # sentinel: returns None from check()
    container = _build_container_with_readiness(fake)
    response = await readiness_response(container)
    assert response.status_code == 503


async def test_readiness_body_lists_every_check() -> None:
    fake = FakeReadiness()
    fake._result = ReadinessResult(
        ready=False,
        checks=[
            ReadinessCheck(name="database", ok=True),
            ReadinessCheck(name="mail_outbox", ok=False, message="queue unreachable"),
        ],
    )
    container = _build_container_with_readiness(fake)
    response = await readiness_response(container)
    assert len(response.body["checks"]) == 2
    names = {c["name"] for c in response.body["checks"]}
    assert names == {"database", "mail_outbox"}
