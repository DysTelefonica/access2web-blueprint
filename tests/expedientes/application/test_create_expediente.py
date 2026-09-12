"""Strict TDD — CreateExpediente use case (C01, issue #226).

CAP-001: alta transaccional. Strict TDD: RED -> GREEN.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

import pytest

if TYPE_CHECKING:
    from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent

from app.src.modules.expedientes.application.create_expediente import (
    CreateExpedienteError,
    create_expediente,
)
from app.src.modules.expedientes.domain import ExpedienteEstado, ExpedienteTipo


def _now() -> datetime:
    return datetime(2026, 9, 12, 12, 0, 0, tzinfo=UTC)


# ---------------------------------------------------------------------------
# Stubs
# ---------------------------------------------------------------------------


class _RepoStub:
    """Duck-typed ExpedienteRepositoryPort for unit tests."""

    calls: list[str] = []
    persisted: object | None = None

    async def create(self, agg: object) -> object:
        self.calls.append("create")
        self.persisted = agg
        return agg

    async def get_by_id(self, id_: UUID) -> object | None:
        return None

    async def update(self, agg: object) -> object:
        return agg

    async def delete(self, id_: UUID) -> None:
        pass

    async def list_by_state(self, estado: str, limit: int, offset: int) -> tuple[list[object], int]:
        return [], 0


@dataclass
class _AuditStub:
    """Duck-typed AuditLogPort for unit tests."""

    calls: list[str] = field(default_factory=list)
    appended: list[ExpedienteAuditEvent] = field(default_factory=list)

    async def append(self, event: ExpedienteAuditEvent) -> None:
        self.calls.append("append")
        self.appended.append(event)

    async def list_for_actor(self, actor_id: UUID, since: datetime) -> list[ExpedienteAuditEvent]:
        return []


# ---------------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------------


class TestCreateExpedienteHappyPath:
    """Creates Expediente AM (root, no parent) and persists it."""

    async def test_creates_am_expediente(self) -> None:
        repo = _RepoStub()
        audit = _AuditStub()

        exp = await create_expediente(
            tipo=ExpedienteTipo.AM,
            estado=ExpedienteEstado.BORRADOR,
            repository=repo,
            audit=audit,
            actor_id=uuid4(),
            now=_now(),
        )

        assert exp.tipo is ExpedienteTipo.AM
        assert exp.estado is ExpedienteEstado.BORRADOR
        assert exp.version == 1
        assert exp.id_expediente_padre is None

    async def test_persists_via_repository(self) -> None:
        repo = _RepoStub()
        audit = _AuditStub()

        exp = await create_expediente(
            tipo=ExpedienteTipo.AM,
            repository=repo,
            audit=audit,
            actor_id=uuid4(),
            now=_now(),
        )

        assert "create" in repo.calls
        assert repo.persisted is exp

    async def test_audits_event(self) -> None:
        repo = _RepoStub()
        audit = _AuditStub()
        actor_id = uuid4()

        await create_expediente(
            tipo=ExpedienteTipo.AM,
            repository=repo,
            audit=audit,
            actor_id=actor_id,
            now=_now(),
        )

        assert "append" in audit.calls
        evt = audit.appended[0]
        assert evt.actor_id == actor_id
        assert evt.event_type == "expedientes.create"
        assert evt.capacidad == "EXP-CAP-001"
        assert evt.result == "ok"

    async def test_returns_created_expediente(self) -> None:
        repo = _RepoStub()
        audit = _AuditStub()

        exp = await create_expediente(
            tipo=ExpedienteTipo.AM,
            repository=repo,
            audit=audit,
            actor_id=uuid4(),
            now=_now(),
        )

        assert exp is repo.persisted


# ---------------------------------------------------------------------------
# Lote requires parent
# ---------------------------------------------------------------------------


class TestCreateExpedienteLoteRequiresParent:
    """CAP-006: Lote MUST have id_expediente_padre."""

    async def test_lote_without_parent_raises(self) -> None:
        repo = _RepoStub()
        audit = _AuditStub()

        with pytest.raises(CreateExpedienteError, match="LOTE"):
            await create_expediente(
                tipo=ExpedienteTipo.LOTE,
                id_expediente_padre=None,
                repository=repo,
                audit=audit,
                actor_id=uuid4(),
                now=_now(),
            )

    async def test_lote_with_parent_succeeds(self) -> None:
        repo = _RepoStub()
        audit = _AuditStub()
        parent_id = uuid4()

        exp = await create_expediente(
            tipo=ExpedienteTipo.LOTE,
            id_expediente_padre=parent_id,
            repository=repo,
            audit=audit,
            actor_id=uuid4(),
            now=_now(),
        )

        assert exp.tipo is ExpedienteTipo.LOTE
        assert exp.id_expediente_padre == parent_id


# ---------------------------------------------------------------------------
# Audit trail on rejection
# ---------------------------------------------------------------------------


class TestCreateExpedienteAuditOnRejection:
    """Rejections are auditable with result != 'ok'."""

    async def test_rejected_lote_audit_event(self) -> None:
        repo = _RepoStub()
        audit = _AuditStub()
        actor_id = uuid4()

        with pytest.raises(CreateExpedienteError):
            await create_expediente(
                tipo=ExpedienteTipo.LOTE,
                id_expediente_padre=None,
                repository=repo,
                audit=audit,
                actor_id=actor_id,
                now=_now(),
            )

        assert len(audit.appended) == 1
        evt = audit.appended[0]
        assert evt.actor_id == actor_id
        assert evt.event_type == "expedientes.create"
        assert evt.result == "denied"
