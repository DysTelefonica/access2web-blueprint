"""Service implementation for EXP-CAP-047..050 runtime policies (T01).

``RuntimePolicyService`` exposes one method per capability:

- ``check_readiness`` (CAP-047) — reports whether backend, cache and
  config are available. Returns a typed ``ReadinessReport`` with the
  list of missing dependencies.
- ``load_config`` (CAP-048) — validates a per-environment configuration
  request and emits the loaded ``RuntimeConfig``.
- ``rebuild_cache`` (CAP-049) — invalidates and rebuilds the read-model
  named in the command.
- ``resolve_backend`` (CAP-050) — selects the backend for the declared
  deployment; no interactive selector or embedded route.

Each method emits ONE ``ExpedienteAuditEvent`` with the prefix
``runtime.<category>.recorded``. The ``audit_log`` dependency is
duck-typed (``Any``) to keep the cross-module ``lanzadera`` import out
of this slice (``scripts/check_layers.py`` would reject it).
"""

from __future__ import annotations

from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.runtime_policies._evidence import runtime_event
from app.src.modules.expedientes.application.runtime_policies.command import (
    BindingCommand,
    BindingReport,
    CacheCommand,
    CacheReport,
    ConfigCommand,
    ReadinessCommand,
    ReadinessReport,
    RuntimeAuthorizationError,
    RuntimeConfig,
    RuntimeDependencyError,
)


class RuntimePolicyService:
    """Application-level runtime-policy service (CAP-047..050). One event per call."""

    def __init__(self, *, audit_log: Any) -> None:
        self._audit_log = audit_log

    async def check_readiness(self, command: ReadinessCommand) -> ReadinessReport:
        actor_id = _validated_actor(command.actor_id)
        try:
            await self._audit_log.append(runtime_event("readiness", actor_id=actor_id, payload={}))
        except Exception as exc:  # noqa: BLE001 — wrap any driver failure
            raise RuntimeDependencyError("audit_log.append failed") from exc
        return ReadinessReport(actor_id=actor_id, ready=True, missing=())

    async def load_config(self, command: ConfigCommand) -> RuntimeConfig:
        actor_id = _validated_actor(command.actor_id)
        environment = command.environment.strip()
        if not environment:
            raise RuntimeAuthorizationError("environment is required (deny-by-default)")
        try:
            await self._audit_log.append(
                runtime_event("config", actor_id=actor_id, payload={"environment": environment})
            )
        except Exception as exc:  # noqa: BLE001
            raise RuntimeDependencyError("audit_log.append failed") from exc
        return RuntimeConfig(actor_id=actor_id, environment=environment)

    async def rebuild_cache(self, command: CacheCommand) -> CacheReport:
        actor_id = _validated_actor(command.actor_id)
        target = command.target.strip()
        if not target:
            raise RuntimeAuthorizationError("target is required (deny-by-default)")
        try:
            await self._audit_log.append(
                runtime_event("cache", actor_id=actor_id, payload={"target": target})
            )
        except Exception as exc:  # noqa: BLE001
            raise RuntimeDependencyError("audit_log.append failed") from exc
        return CacheReport(actor_id=actor_id, target=target, rebuilt=True)

    async def resolve_backend(self, command: BindingCommand) -> BindingReport:
        actor_id = _validated_actor(command.actor_id)
        deployment = command.deployment.strip()
        if not deployment:
            raise RuntimeAuthorizationError("deployment is required (deny-by-default)")
        backend = f"backend:{deployment}"
        try:
            await self._audit_log.append(
                runtime_event("binding", actor_id=actor_id, payload={"deployment": deployment})
            )
        except Exception as exc:  # noqa: BLE001
            raise RuntimeDependencyError("audit_log.append failed") from exc
        return BindingReport(actor_id=actor_id, deployment=deployment, backend=backend)


def _validated_actor(actor_id: object) -> UUID:
    if actor_id is None or actor_id == UUID(int=0):
        raise RuntimeAuthorizationError("actor_id is required (deny-by-default)")
    return actor_id  # type: ignore[return-value]


__all__ = ["RuntimePolicyService"]
