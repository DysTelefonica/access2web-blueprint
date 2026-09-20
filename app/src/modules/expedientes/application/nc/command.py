"""Commands for EXP-CAP-054 NC adapter (H04)."""

from dataclasses import dataclass
from uuid import UUID


class NCError(Exception):
    """Base error for the NC use case."""


class NCAuthorizationError(NCError):
    """Deny-by-default: raised before any audit emission."""


class NCDependencyError(NCError):
    """Wraps the underlying adapter exception (``__cause__``)."""


@dataclass(frozen=True)
class NCLookupCommand:
    actor_id: UUID | None
    expediente_id: UUID | None = None
    s4h_code: str | None = None


__all__ = ["NCAuthorizationError", "NCDependencyError", "NCError", "NCLookupCommand"]
