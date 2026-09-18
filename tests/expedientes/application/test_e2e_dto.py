"""Strict TDD for EXP-CAP-033 stable DTO contract (E01)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.e2e_dto.command import (
    E2EDtoAuthorizationError,
    E2EDtoCommand,
    E2EDtoError,
    E2EDtoResult,
    E2EDtoValidationError,
)
from app.src.modules.expedientes.application.e2e_dto.service import E2EDtoService
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


def _expediente() -> Expediente:
    return Expediente(
        id=uuid4(),
        tipo=ExpedienteTipo.AM,
        estado=ExpedienteEstado.BORRADOR,
        version=1,
        fecha_inicio_contrato=date(2026, 1, 1),
        fecha_fin_contrato=date(2026, 12, 31),
    )


def _make_service() -> tuple[E2EDtoService, _Audit]:
    audit = _Audit()
    service = E2EDtoService(audit_log=audit, permissions=set())
    return service, audit


def _command(expediente: Expediente, **overrides: Any) -> E2EDtoCommand:
    values: dict[str, Any] = {
        "expediente": expediente,
        "actor_id": uuid4(),
        "api_version": "1.0",
    }
    values.update(overrides)
    return E2EDtoCommand(**values)


async def test_dto_round_trip_preserves_domain_fields() -> None:
    service, _ = _make_service()
    service.grant("e2e.exchange")
    expediente = _expediente()

    result = await service.exchange(_command(expediente))

    assert isinstance(result, E2EDtoResult)
    assert result.dto.api_version == "1.0"
    assert result.dto.tipo == "AM"
    assert result.dto.estado == "BORRADOR"
    assert result.dto.fecha_inicio_contrato == "2026-01-01"
    assert result.dto.fecha_fin_contrato == "2026-12-31"
    assert result.reconstructed.tipo is ExpedienteTipo.AM
    assert result.reconstructed.estado is ExpedienteEstado.BORRADOR


async def test_dto_does_not_expose_technical_ids() -> None:
    service, _ = _make_service()
    service.grant("e2e.exchange")
    expediente = _expediente()

    result = await service.exchange(_command(expediente))
    payload = result.dto.to_payload()

    assert "id" not in payload
    assert "version" not in payload
    assert "id_expediente_padre" not in payload


async def test_dto_rejects_wrong_api_version() -> None:
    service, _ = _make_service()
    service.grant("e2e.exchange")

    with pytest.raises(E2EDtoValidationError, match="apiVersion"):
        await service.exchange(_command(_expediente(), api_version="2.0"))


async def test_dto_denies_actor_without_permission() -> None:
    service, _ = _make_service()

    with pytest.raises(E2EDtoAuthorizationError, match="e2e.exchange"):
        await service.exchange(_command(_expediente()))


async def test_dto_rejects_invalid_date_format() -> None:
    """A bad inbound date must be rejected at the DTO level."""
    from app.src.modules.expedientes.application.e2e_dto.command import (
        ExpedienteDTO,
    )

    service, _ = _make_service()
    bad = ExpedienteDTO(
        api_version="1.0",
        tipo="AM",
        estado="BORRADOR",
        fecha_inicio_contrato="not-a-date",
        fecha_fin_contrato=None,
    )

    with pytest.raises(E2EDtoValidationError, match="fecha"):
        bad.to_aggregate()


async def test_dto_audits_each_exchange() -> None:
    service, audit = _make_service()
    service.grant("e2e.exchange")
    expediente = _expediente()

    await service.exchange(_command(expediente))

    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.dto.exchanged"
    assert audit.events[0].capacidad == "EXP-CAP-033"


async def test_dto_rejects_invalid_combination() -> None:
    """Invalid tipo/estado strings must be rejected at the DTO level
    rather than silently coerce to the wrong enum value.
    """
    from app.src.modules.expedientes.application.e2e_dto.command import (
        ExpedienteDTO,
    )

    with pytest.raises(E2EDtoError):
        ExpedienteDTO(
            api_version="1.0",
            tipo="NOT-A-TYPE",
            estado="BORRADOR",
            fecha_inicio_contrato=None,
            fecha_fin_contrato=None,
        ).to_aggregate()
