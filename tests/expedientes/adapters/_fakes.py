"""In-memory port fakes for Expedientes adapter tests (F01, issue #222).

Each fake satisfies the corresponding Port contract structurally (duck-typed
in Python). ``calls`` tracks only the method name for deterministic assertions.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

if TYPE_CHECKING:
    from app.src.modules.expedientes.ports.audit_log import ExpedienteAuditEvent
    from app.src.modules.expedientes.ports.document_storage import StoredDocument


@dataclass
class FakeExpedienteRepository:
    by_id: dict[UUID, object] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)

    async def get_by_id(self, id_: UUID) -> object | None:
        self.calls.append("get_by_id")
        return self.by_id.get(id_)

    async def create(self, agg: object) -> object:
        self.calls.append("create")
        self.by_id[getattr(agg, "id", uuid4())] = agg
        return agg

    async def update(self, agg: object) -> object:
        self.calls.append("update")
        self.by_id[getattr(agg, "id", uuid4())] = agg
        return agg

    async def delete(self, id_: UUID) -> None:
        self.calls.append("delete")
        self.by_id.pop(id_, None)

    async def list_by_state(self, estado: str, limit: int, offset: int) -> tuple[list[object], int]:
        self.calls.append("list_by_state")
        rows = [v for v in self.by_id.values() if getattr(v, "estado", None) == estado]
        return rows[offset : offset + limit], len(rows)


@dataclass
class FakeHitoRepository:
    by_exp: dict[UUID, list[object]] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)

    async def get_by_expediente(self, id_: UUID) -> list[object]:
        self.calls.append("get_by_expediente")
        return list(self.by_exp.get(id_, []))

    async def upsert(self, hito: object) -> object:
        self.calls.append("upsert")
        eid = getattr(hito, "id_expediente", None)
        if eid:
            self.by_exp.setdefault(eid, []).append(hito)
        return hito

    async def delete(self, id_: UUID) -> None:
        self.calls.append("delete")


@dataclass
class FakeCatalogRepository:
    entries: dict[int, object] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)

    async def list_all(self) -> list[object]:
        self.calls.append("list_all")
        return list(self.entries.values())

    async def get_by_id(self, id_: int) -> object | None:
        self.calls.append("get_by_id")
        return self.entries.get(id_)

    async def search(self, query: str, limit: int = 20) -> list[object]:
        self.calls.append("search")
        q = query.lower()
        return [
            v
            for v in self.entries.values()
            if q in str(getattr(v, "descripcion", "")).lower()
        ][:limit]


@dataclass
class FakeAuditLog:
    events: list[ExpedienteAuditEvent] = field(default_factory=list)
    calls: list[str] = field(default_factory=list)

    async def append(self, event: ExpedienteAuditEvent) -> None:
        self.calls.append("append")
        self.events.append(event)

    async def list_for_actor(self, actor_id: UUID, since: datetime) -> list[ExpedienteAuditEvent]:
        self.calls.append("list_for_actor")
        return [
            e
            for e in self.events
            if e.actor_id == actor_id and (e.created_at or datetime.min) >= since
        ]


@dataclass
class FakeReadiness:
    _result: object = field(default=None)

    async def check(self) -> object:
        if self._result is None:
            from app.src.modules.expedientes.ports.readiness import ReadinessResult

            return ReadinessResult(ready=True, checks=[])
        return self._result


@dataclass
class FakeDocumentStorage:
    by_ref: dict[str, tuple[bytes, str, int]] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)

    async def prepare_upload(
        self,
        exp_id: UUID,
        filename: str,
        content_type: str,
        size_bytes: int,
    ) -> str:
        self.calls.append("prepare_upload")
        ref = f"fake://{uuid4()}"
        self.by_ref[ref] = (b"", content_type, size_bytes)
        return ref

    async def confirm_upload(self, ref: str) -> StoredDocument:
        self.calls.append("confirm_upload")
        _c, ct, sz = self.by_ref[ref]
        from app.src.modules.expedientes.ports.document_storage import StoredDocument

        return StoredDocument(storage_ref=ref, content_type=ct, size_bytes=sz)

    async def download(self, ref: str) -> bytes:
        self.calls.append("download")
        return self.by_ref.get(ref, (b"", "application/octet-stream", 0))[0]

    async def delete(self, ref: str) -> None:
        self.calls.append("delete")
        self.by_ref.pop(ref, None)


@dataclass
class FakeNotificationDelivery:
    sent: list[tuple[str, str, str]] = field(default_factory=list)
    calls: list[str] = field(default_factory=list)

    async def send(self, to: str, subject: str, body: str) -> None:
        self.calls.append("send")
        self.sent.append((to, subject, body))
