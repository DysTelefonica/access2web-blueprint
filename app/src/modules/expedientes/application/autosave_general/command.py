"""Boundary types for CAP-030 autosave core (#234)."""

from dataclasses import dataclass
from datetime import date, datetime
from uuid import UUID


class AutosaveGeneralError(Exception):
    """Base autosave failure."""


class AutosaveGeneralValidationError(AutosaveGeneralError):
    """Malformed input or aggregate invariant violation."""


class AutosaveGeneralAuthorizationError(AutosaveGeneralError):
    """Missing actor under D-EXP-3 deny-by-default."""


class AutosaveGeneralConflictError(AutosaveGeneralError):
    """Stale version or idempotency-key payload mismatch."""


@dataclass(frozen=True)
class AutosaveGeneralCommand:
    expediente_id: UUID
    expected_version: int
    idempotency_key: UUID
    fecha_inicio_contrato: date | None
    fecha_fin_contrato: date | None
    actor_id: UUID

    def signature(self) -> str:
        start = self.fecha_inicio_contrato.isoformat() if self.fecha_inicio_contrato else ""
        end = self.fecha_fin_contrato.isoformat() if self.fecha_fin_contrato else ""
        return f"{self.expediente_id}|{self.expected_version}|{start}|{end}|{self.actor_id}"


@dataclass(frozen=True)
class AutosaveGeneralResult:
    expediente_id: UUID
    new_version: int
    modified_at: datetime
    fields_changed: tuple[str, ...]
