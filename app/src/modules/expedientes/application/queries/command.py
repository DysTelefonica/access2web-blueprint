"""Commands and outcomes for EXP-CAP-025..029 query services."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from uuid import UUID


class QueriesError(Exception):
    """Base error for the queries use case."""


class QueriesValidationError(QueriesError):
    """The query is malformed or the result is missing."""


class QueriesAuthorizationError(QueriesError):
    """The caller has no authenticated actor or lacks permission."""


@dataclass(frozen=True)
class BandejaQuery:
    actor_id: UUID
    estado: str | None
    codigo: str | None
    responsable_id: UUID | None = None
    juridica_id: UUID | None = None
    suministrador_id: UUID | None = None
    limit: int = 20
    offset: int = 0


@dataclass(frozen=True)
class BandejaPage:
    rows: list[object]
    total: int
    limit: int
    offset: int

    @property
    def has_more(self) -> bool:
        return self.offset + len(self.rows) < self.total


@dataclass(frozen=True)
class ExportacionExcelCommand:
    actor_id: UUID
    filtros: dict[str, Any]
    formato: str = "xlsx"


@dataclass(frozen=True)
class ExportacionExcelJob:
    job_id: UUID
    filtros: dict[str, Any]
    formato: str
    requested_at: datetime


@dataclass(frozen=True)
class TareasCalculadasQuery:
    actor_id: UUID
    now: datetime


@dataclass(frozen=True)
class TareasResult:
    contadores: dict[str, int] = field(default_factory=dict)
    total: int = 0
