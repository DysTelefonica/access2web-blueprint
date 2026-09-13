"""Strict TDD for EXP-CAP-034 canonical JSON (E02)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.json_canonical.command import (
    JsonCanonicalAuthorizationError,
    JsonCanonicalValidationError,
)
from app.src.modules.expedientes.application.json_canonical.service import (
    JsonCanonicalService,
)

EXPECTED_COLLECTIONS: tuple[str, ...] = (
    "expedientes",
    "hitos",
    "modificados",
    "anualidades",
    "anexos",
    "juridicas",
    "responsables",
    "suministradores",
    "lugares_ejecucion",
)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


def _expediente(
    *,
    eid: UUID | None = None,
    ordinal: int = 1,
    tipo: str = "AM",
    estado: str = "BORRADOR",
    fecha: datetime | None = "__default__",
) -> dict[str, Any]:
    if fecha == "__default__":
        fecha = datetime(2026, 1, 1, 0, 0, 0, tzinfo=__import__("datetime").UTC)
    return {
        "id": eid or uuid4(),
        "tipo": tipo,
        "estado": estado,
        "version": 1,
        "ordinal_e2e": ordinal,
        "fecha_inicio_contrato": fecha,
    }


def _make_service() -> tuple[JsonCanonicalService, _Audit]:
    audit = _Audit()
    return JsonCanonicalService(audit_log=audit, permissions=set()), audit


def _actor() -> UUID:
    return uuid4()


async def test_envelope_has_meta_and_data() -> None:
    service, _ = _make_service()
    service.grant("e2e.export")

    payload = await service.render(actor_id=_actor(), collections={"expedientes": [_expediente()]})

    assert "meta" in payload
    assert "data" in payload
    assert payload["meta"]["apiVersion"] == "1.0"
    assert set(payload["data"]) == set(EXPECTED_COLLECTIONS)


async def test_dates_serialised_as_iso_z() -> None:
    service, _ = _make_service()
    service.grant("e2e.export")
    payload = await service.render(actor_id=_actor(), collections={"expedientes": [_expediente()]})
    value = payload["data"]["expedientes"][0]["fecha_inicio_contrato"]
    assert value.endswith("Z")
    assert "T00:00:00" in value


async def test_null_fields_serialised_explicitly() -> None:
    service, _ = _make_service()
    service.grant("e2e.export")
    payload = await service.render(
        actor_id=_actor(), collections={"expedientes": [_expediente(fecha=None)]}
    )
    value = payload["data"]["expedientes"][0]["fecha_inicio_contrato"]
    assert value is None


async def test_collections_ordered_by_ordinal_then_id() -> None:
    service, _ = _make_service()
    service.grant("e2e.export")
    e1 = _expediente(ordinal=2, eid=UUID("00000000-0000-0000-0000-000000000001"))
    e2 = _expediente(ordinal=1, eid=UUID("00000000-0000-0000-0000-000000000002"))
    e3 = _expediente(ordinal=1, eid=UUID("00000000-0000-0000-0000-000000000000"))

    payload = await service.render(actor_id=_actor(), collections={"expedientes": [e1, e2, e3]})
    ordered = [row["id"] for row in payload["data"]["expedientes"]]
    assert ordered == [
        "00000000-0000-0000-0000-000000000000",
        "00000000-0000-0000-0000-000000000002",
        "00000000-0000-0000-0000-000000000001",
    ]


async def test_missing_collection_is_filled_with_empty_list() -> None:
    service, _ = _make_service()
    service.grant("e2e.export")

    payload = await service.render(actor_id=_actor(), collections={"expedientes": [_expediente()]})

    assert payload["data"]["hitos"] == []
    assert payload["data"]["anexos"] == []


async def test_unknown_collection_rejected() -> None:
    service, _ = _make_service()
    service.grant("e2e.export")

    with pytest.raises(JsonCanonicalValidationError, match="unknown collection"):
        await service.render(actor_id=_actor(), collections={"inventado": []})


async def test_render_requires_permission() -> None:
    service, _ = _make_service()

    with pytest.raises(JsonCanonicalAuthorizationError, match="e2e.export"):
        await service.render(actor_id=_actor(), collections={})


async def test_render_audits_each_invocation() -> None:
    service, audit = _make_service()
    service.grant("e2e.export")

    await service.render(actor_id=_actor(), collections={"expedientes": [_expediente()]})

    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.canonical.rendered"
    assert audit.events[0].capacidad == "EXP-CAP-034"
