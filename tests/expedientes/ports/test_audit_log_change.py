"""Strict TDD — AuditLogPort.record_change (issue #668, CAP-001 §Camino feliz).

CAP-001 §Camino feliz exige "se persisten cabecera, hijos, read-model
y último cambio en una transacción confirmada". C01 (#669)
implementó cabecera + hijos + audit event. Falta:

- ``TbUltimoCambio``: una fila por expediente con
  ``(IDExpediente, FechaCambio, IDUsuarioCambio)`` — el último cambio
  registrado.
- ``TbCambios``: log histórico con
  ``(IDCambio, NombreTabla, IDExpediente, NombreCampo, ValorInicial,
  ValorFinal, FechaCambio, IDUsuarioCambio, Accion)`` — fila por
  cambio.

Estos tests verifican el comportamiento del puerto
``AuditLogPort.record_change`` que el servicio llama dentro del UoW
para materializar ambos registros en una sola transacción (DA-11).
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.expedientes.ports.audit_log import (
    AuditLogPort,
    ChangeRecord,
)

# ---------------------------------------------------------------------------
# Doubles
# ---------------------------------------------------------------------------


class _FakeAuditLog(AuditLogPort):
    """Minimal in-memory audit log that records ``record_change`` calls.

    Production adapter lives in a follow-up vertical; the use case
    exercises the protocol structurally.
    """

    def __init__(self) -> None:
        self.events: list[Any] = []
        self.changes: list[ChangeRecord] = []
        self.calls: list[str] = []

    async def append(self, event: Any) -> None:  # type: ignore[override]
        self.calls.append("append")
        self.events.append(event)

    async def list_for_actor(  # type: ignore[override]
        self, actor_id: Any, since: datetime
    ) -> list[Any]:
        self.calls.append("list_for_actor")
        return []

    async def record_change(self, change: ChangeRecord) -> None:  # type: ignore[override]
        self.calls.append("record_change")
        self.changes.append(change)


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


def _make_change(**overrides: Any) -> ChangeRecord:
    defaults: dict[str, Any] = {
        "id": uuid4(),
        "nombre_tabla": "expedientes",
        "id_expediente": uuid4(),
        "nombre_campo": None,
        "valor_inicial": None,
        "valor_final": None,
        "fecha_cambio": datetime.now(UTC),
        "id_usuario_cambio": uuid4(),
        "accion": "alta",
    }
    defaults.update(overrides)
    return ChangeRecord(**defaults)


def test_change_record_carries_all_required_fields() -> None:
    """TbCambios has 10 fields; the value object captures them all."""
    change = _make_change()
    assert change.nombre_tabla == "expedientes"
    assert change.accion == "alta"
    assert change.fecha_cambio.tzinfo is not None


def test_audit_log_records_the_change() -> None:
    audit = _FakeAuditLog()
    change = _make_change()

    asyncio_run(audit.record_change(change))

    assert "record_change" in audit.calls
    assert len(audit.changes) == 1
    assert audit.changes[0].accion == "alta"


def test_audit_log_appends_independent_of_record_change() -> None:
    """``append`` (audit event) and ``record_change`` (change log) are
    two distinct writes. A successful alta produces both: one audit
    event (who did it, when) and one change row (what happened)."""
    audit = _FakeAuditLog()
    audit_event_id = uuid4()
    change = _make_change(accion="alta")

    asyncio_run(audit.append(_dummy_event(audit_event_id)))
    asyncio_run(audit.record_change(change))

    assert audit.calls == ["append", "record_change"]


def _dummy_event(event_id: Any) -> Any:
    from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

    return ExpedienteAuditEvent(
        id=event_id,
        event_type="expediente.created",
        actor_id=uuid4(),
        target_id=uuid4(),
        capacidad="EXP-CAP-001",
    )


def asyncio_run(coro: Any) -> None:
    """Helper to run async code in a synchronous test."""
    import asyncio

    asyncio.run(coro)


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------


def test_change_record_rejects_empty_nombre_tabla() -> None:
    with pytest.raises(ValueError):
        _make_change(nombre_tabla="")


def test_change_record_rejects_empty_accion() -> None:
    with pytest.raises(ValueError):
        _make_change(accion="")


def test_change_record_requires_tz_aware_fecha() -> None:
    """Same timezone strictness as RetentionPolicy (CAP-010 follow-up).

    Naive datetimes would let different servers disagree on the
    timestamp of the change. Audit logs must be reproducible.
    """
    with pytest.raises(ValueError):
        _make_change(fecha_cambio=datetime(2024, 1, 1))  # noqa: DTZ001


def test_change_record_rejects_null_id_expediente() -> None:
    """An audit row without an associated expediente is orphaned."""
    with pytest.raises(ValueError):
        _make_change(id_expediente=None)


def test_change_record_rejects_null_id_usuario() -> None:
    """D-EXP-3 deny-by-default requires an actor for every audit row."""
    with pytest.raises(ValueError):
        _make_change(id_usuario_cambio=None)
