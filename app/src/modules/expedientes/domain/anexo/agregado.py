"""Anexo aggregate root — D-EXP-1, CAP-010.

Legacy ``TbExpedientesAnexos`` keeps three columns:

- IDDocumento (PK)
- IDExpediente (FK, nullable)
- NombreDocumento (255, nullable)

The domain extends that with the storage reference, content type and size
needed to honor CAP-010 §Camino feliz ("se registra la referencia y su
retención sin exponer rutas locales") and CAP-010 §Validación
("dentro del límite acordado"). The anexo carries its retention policy
so the audit event can prove it has not expired before deletion.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from uuid import UUID

from app.src.modules.expedientes.domain.anexo.reference import AnexoReference
from app.src.modules.expedientes.domain.anexo.retention import RetentionPolicy


@dataclass
class Anexo:
    """Anexo aggregate root attached to an Expediente (CAP-010).

    The aggregate owns its reference, retention policy and lifecycle.
    Anexos are append-only at creation; deletion is a separate transition
    that emits its own audit event and history row.
    """

    id: UUID
    id_expediente: UUID
    nombre: str | None
    referencia: AnexoReference
    retention: RetentionPolicy
    created_at: datetime | None = None
    deleted_at: datetime | None = None
    _history: list[AnexoEvent] = field(default_factory=list)

    @property
    def history(self) -> tuple[AnexoEvent, ...]:
        """Immutable view of the anexo's own event log (in-memory only)."""
        return tuple(self._history)

    def is_deleted(self) -> bool:
        return self.deleted_at is not None

    def record_event(self, event: AnexoEvent) -> None:
        """Append a domain event to the anexo's history."""
        self._history.append(event)


@dataclass(frozen=True)
class AnexoEvent:
    """Domain event carried by an Anexo.

    Lifecycle: ``created`` at first persistence, ``deleted`` when the
    anexo is removed via the use case. Future events (restored, archived)
    can extend the enum without breaking the persistence shape.
    """

    kind: str
    at: datetime
    actor_id: UUID

    def __post_init__(self) -> None:
        if self.kind not in {"created", "deleted"}:
            raise ValueError(f"unknown anexo event kind: {self.kind!r}")
