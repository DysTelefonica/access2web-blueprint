"""Commands, reports and error hierarchy for EXP-CAP-047..050 (T01).

``EXPEDIENTES_APP_ID`` lives in ``_evidence.py`` (not here) — see that
module's docstring for the DRY-ratchet rationale (``dup:79d8c5976f0f``).
"""

from dataclasses import dataclass


class RuntimeError_(Exception):
    """Base error for the runtime-policies use case."""


class RuntimeAuthorizationError(RuntimeError_):
    """Deny-by-default: raised before any audit emission."""


class RuntimeDependencyError(RuntimeError_):
    """Wraps the underlying driver exception (``__cause__``) so the
    caller's UoW can roll back without losing the dependency-failure
    context.
    """


@dataclass(frozen=True)
class ReadinessCommand:
    actor_id: object


@dataclass(frozen=True)
class ConfigCommand:
    actor_id: object
    environment: str


@dataclass(frozen=True)
class CacheCommand:
    actor_id: object
    target: str


@dataclass(frozen=True)
class BindingCommand:
    actor_id: object
    deployment: str


@dataclass(frozen=True)
class ReadinessReport:
    actor_id: object
    ready: bool
    missing: tuple[str, ...] = ()


@dataclass(frozen=True)
class RuntimeConfig:
    actor_id: object
    environment: str


@dataclass(frozen=True)
class CacheReport:
    actor_id: object
    target: str
    rebuilt: bool


@dataclass(frozen=True)
class BindingReport:
    actor_id: object
    deployment: str
    backend: str


__all__ = [
    "BindingCommand",
    "BindingReport",
    "CacheCommand",
    "CacheReport",
    "ConfigCommand",
    "ReadinessCommand",
    "ReadinessReport",
    "RuntimeAuthorizationError",
    "RuntimeConfig",
    "RuntimeDependencyError",
    "RuntimeError_",
]
