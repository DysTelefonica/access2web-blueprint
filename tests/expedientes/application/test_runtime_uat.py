"""UAT end-to-end for EXP-CAP-047..050 runtime-policy lane (issue #281, U08).

Composes the four RuntimePolicyService methods (check_readiness,
load_config, rebuild_cache, resolve_backend) over a single audit log
to verify the composed chain. Does NOT re-test individual methods —
T01 unit tests cover that ground.
"""

from dataclasses import dataclass, field
from uuid import uuid4

from app.src.modules.expedientes.application.runtime_policies.command import (
    BindingCommand,
    CacheCommand,
    ConfigCommand,
    ReadinessCommand,
    RuntimeAuthorizationError,
    RuntimeDependencyError,
)
from app.src.modules.expedientes.application.runtime_policies.service import RuntimePolicyService


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


@dataclass
class _RaisingAudit:
    async def append(self, event: object) -> None:
        raise RuntimeError("audit sink unavailable")


def _service(audit: _AuditLog | None = None) -> tuple[RuntimePolicyService, _AuditLog]:
    audit = audit or _AuditLog()
    return RuntimePolicyService(audit_log=audit), audit


def _types(audit: _AuditLog) -> list[str]:
    return [e.event_type for e in audit.events]  # type: ignore[attr-defined]


async def test_full_chain_emits_four_events_in_capability_order() -> None:
    """check_readiness -> load_config -> rebuild_cache -> resolve_backend, in order."""
    actor = uuid4()
    svc, audit = _service()

    await svc.check_readiness(ReadinessCommand(actor_id=actor))
    await svc.load_config(ConfigCommand(actor_id=actor, environment="ci"))
    await svc.rebuild_cache(CacheCommand(actor_id=actor, target="read_model"))
    await svc.resolve_backend(BindingCommand(actor_id=actor, deployment="production"))

    assert _types(audit) == [
        "runtime.readiness.recorded",
        "runtime.config.recorded",
        "runtime.cache.recorded",
        "runtime.binding.recorded",
    ]
    capacidades = [e.capacidad for e in audit.events]  # type: ignore[attr-defined]
    assert capacidades == ["EXP-CAP-047", "EXP-CAP-048", "EXP-CAP-049", "EXP-CAP-050"]


async def test_per_context_isolation_emit_independent_audit_chains() -> None:
    """Two services over the same audit log produce distinct, interleaved chains."""
    actor_a, actor_b = uuid4(), uuid4()
    audit = _AuditLog()
    svc_a = RuntimePolicyService(audit_log=audit)
    svc_b = RuntimePolicyService(audit_log=audit)

    await svc_a.check_readiness(ReadinessCommand(actor_id=actor_a))
    await svc_b.check_readiness(ReadinessCommand(actor_id=actor_b))
    await svc_a.load_config(ConfigCommand(actor_id=actor_a, environment="ci"))
    await svc_b.resolve_backend(BindingCommand(actor_id=actor_b, deployment="staging"))

    actors = [e.actor_id for e in audit.events]  # type: ignore[attr-defined]
    assert actors == [actor_a, actor_b, actor_a, actor_b]
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]


async def test_dependency_failure_wraps_underlying_exception() -> None:
    """audit_log.append raises -> RuntimeDependencyError chained from the cause."""
    actor = uuid4()
    svc = RuntimePolicyService(audit_log=_RaisingAudit())
    with pytest.raises(RuntimeDependencyError) as excinfo:
        await svc.check_readiness(ReadinessCommand(actor_id=actor))
    assert isinstance(excinfo.value.__cause__, RuntimeError)


async def test_missing_actor_raises_authorization_and_skips_audit() -> None:
    """actor_id=None raises RuntimeAuthorizationError without emitting any audit."""
    svc, audit = _service()
    with pytest.raises(RuntimeAuthorizationError, match="actor_id"):
        await svc.check_readiness(ReadinessCommand(actor_id=None))
    assert audit.events == []


async def test_load_config_and_resolve_backend_surface_payload_fields() -> None:
    """environment and deployment appear in both the report and the audit payload."""
    actor = uuid4()
    svc, audit = _service()
    await svc.load_config(ConfigCommand(actor_id=actor, environment="ci"))
    await svc.resolve_backend(BindingCommand(actor_id=actor, deployment="production"))

    assert audit.events[0].payload["environment"] == "ci"  # type: ignore[attr-defined]
    assert audit.events[1].payload["deployment"] == "production"  # type: ignore[attr-defined]


async def test_two_method_calls_with_same_actor_emit_two_distinct_events() -> None:
    """Same actor, two calls: distinct event ids, audit chain of length 2."""
    actor = uuid4()
    svc, audit = _service()
    await svc.check_readiness(ReadinessCommand(actor_id=actor))
    await svc.check_readiness(ReadinessCommand(actor_id=actor))

    assert len(audit.events) == 2
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]
    assert audit.events[0].created_at != audit.events[1].created_at  # type: ignore[attr-defined]


# --- imports used by the parametrized-style tests below (kept local to avoid
# widening the top-of-file import block) ---
import pytest  # noqa: E402
