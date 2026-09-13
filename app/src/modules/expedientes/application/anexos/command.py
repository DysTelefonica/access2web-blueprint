"""Commands and outcomes for EXP-CAP-010 anexos."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from uuid import UUID

from app.src.modules.expedientes.domain.anexo.reference import AnexoReference
from app.src.modules.expedientes.domain.anexo.retention import RetentionPolicy


class AnexosError(Exception):
    """Base error for the anexos use case."""


class RegisterAnexoError(AnexosError):
    """A persistence dependency failed; no specific subclass applies."""


class RegisterAnexoValidationError(RegisterAnexoError):
    """The command or referenced expediente is invalid."""


class RegisterAnexoSizeLimitError(RegisterAnexoError):
    """The reference size exceeds the configured limit."""


class CreateAnexoAuthorizationError(RegisterAnexoError):
    """The caller has no authenticated actor."""


class DeleteAnexoAuthorizationError(AnexosError):
    """The caller has no authenticated actor for deletion."""


class DeleteAnexoError(AnexosError):
    """A persistence dependency failed during deletion."""


@dataclass(frozen=True)
class CreateAnexoCommand:
    anexo_id: UUID
    expediente_id: UUID
    nombre: str | None
    referencia: AnexoReference
    retention: RetentionPolicy
    actor_id: UUID


@dataclass(frozen=True)
class RegisterAnexoResult:
    anexo_id: UUID
    expediente_id: UUID
    created_at: datetime


@dataclass(frozen=True)
class DeleteAnexoCommand:
    anexo_id: UUID
    expected_version: int
    actor_id: UUID | None


@dataclass(frozen=True)
class DeleteAnexoResult:
    anexo_id: UUID
    deleted_at: datetime
