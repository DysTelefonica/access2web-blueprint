"""Commands and outcomes for EXP-CAP-036 E2E hash."""

from dataclasses import dataclass
from typing import Any
from uuid import UUID

from app.src.modules.expedientes.domain.hash.versioning import (
    HashRegistry,
    HashVersioned,
    canonicalize,
    compute_hash,
)


class E2EHashError(Exception):
    """Base error for the E2E hash use case."""


class E2EHashValidationError(E2EHashError):
    """The algorithm is missing, unsupported or otherwise invalid."""


class E2EHashAuthorizationError(E2EHashError):
    """The caller has no authenticated actor or lacks permission."""


@dataclass(frozen=True)
class E2EHashCommand:
    actor_id: UUID
    payload: object
    algorithm: str | None = None


@dataclass(frozen=True)
class E2EHashService:
    """Application-level wrapper around the domain hash registry.

    The service enforces the E03 policy:
    - permission ``e2e.hash`` is required;
    - only the algorithms in ``supported_algorithms`` are accepted,
      even if the underlying :class:`HashRegistry` holds more entries;
    - one audit event is appended per successful compute.
    """

    registry: HashRegistry
    audit_log: Any
    permissions: set[str]
    default_algorithm: str
    supported_algorithms: tuple[str, ...]
    canonicalize_fn: Any = canonicalize
    compute_fn: Any = compute_hash

    def grant(self, permission: str) -> None:
        self.permissions.add(permission)

    async def compute(self, command: E2EHashCommand) -> HashVersioned:
        if command.actor_id is None:
            raise E2EHashAuthorizationError("actor_id is required (deny-by-default)")
        if "e2e.hash" not in self.permissions:
            raise E2EHashAuthorizationError("actor lacks permission 'e2e.hash'")
        algorithm = command.algorithm or self.default_algorithm
        if algorithm not in self.supported_algorithms:
            raise E2EHashValidationError(
                f"hash algorithm {algorithm!r} not in supported set {self.supported_algorithms!r}"
            )
        try:
            result: HashVersioned = self.compute_fn(
                command.payload,
                algorithm=algorithm,
                registry=self.registry,
            )
        except E2EHashError:
            raise
        except Exception as exc:
            raise E2EHashError(f"hash failed for algorithm {algorithm!r}: {exc}") from exc
        from app.src.modules.expedientes.application.e2e_hash._evidence import (
            hash_event,
        )

        await self.audit_log.append(hash_event(command.actor_id, result))
        return result


__all__ = [
    "E2EHashAuthorizationError",
    "E2EHashCommand",
    "E2EHashError",
    "E2EHashService",
    "E2EHashValidationError",
]
