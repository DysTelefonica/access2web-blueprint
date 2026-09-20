"""Commands and results for EXP-CAP-053 Riesgos adapter (H03)."""

from dataclasses import dataclass
from uuid import UUID


class RiesgosError(Exception):
    """Base error for the Riesgos use case."""


class RiesgosAuthorizationError(RiesgosError):
    """Deny-by-default: raised before any audit emission."""


class RiesgosDependencyError(RiesgosError):
    """Wraps the underlying adapter exception (``__cause__``)."""


@dataclass(frozen=True)
class RiesgosLookupCommand:
    actor_id: UUID | None
    expediente_id: UUID


__all__ = [
    "RiesgosAuthorizationError",
    "RiesgosDependencyError",
    "RiesgosError",
    "RiesgosLookupCommand",
]
