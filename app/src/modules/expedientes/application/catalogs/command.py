"""Commands and outcomes for EXP-CAP-015..019 catalogs."""

from dataclasses import dataclass
from typing import Any
from uuid import UUID

from app.src.modules.expedientes.domain.cpv.code import CPVCode


class CatalogsError(Exception):
    """Base error for the catalogs use case."""


class CatalogsValidationError(CatalogsError):
    """The query is malformed or the entry is missing."""


class CatalogsAuthorizationError(CatalogsError):
    """The caller has no authenticated actor or lacks permission."""


@dataclass(frozen=True)
class CatalogQuery:
    actor_id: UUID
    text: str | None
    limit: int = 20


@dataclass(frozen=True)
class CpvQuery:
    actor_id: UUID
    raw_code: CPVCode


@dataclass(frozen=True)
class CpvDecision:
    raw_code: CPVCode
    accepted: bool
    decision: str
    row: Any | None
    reason: str | None
