"""Commands and outcomes for EXP-CAP-037 E2E package."""

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID


class ExportPackageError(Exception):
    """Base error for the E2E package use case."""


class ExportPackageValidationError(ExportPackageError):
    """The package is empty or malformed."""


class ExportPackageAuthorizationError(ExportPackageError):
    """The caller has no authenticated actor or lacks permission."""


@dataclass(frozen=True)
class ExportItem:
    """A single expediente to be included in the package."""

    id: UUID
    ordinal: int
    data: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ExportPackageCommand:
    actor_id: UUID | None
    items: list[ExportItem]


@dataclass(frozen=True)
class ExportPackageResult:
    """The result of an E2E package assembly + sink dispatch."""

    package: dict[str, Any]
    manifest_count: int


__all__ = [
    "ExportItem",
    "ExportPackageAuthorizationError",
    "ExportPackageCommand",
    "ExportPackageError",
    "ExportPackageResult",
    "ExportPackageValidationError",
]
