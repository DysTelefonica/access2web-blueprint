"""Strict TDD for EXP-CAP-047..050 RuntimePolicyService (T01)."""

from dataclasses import dataclass, field
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.runtime_policies._evidence import EXPEDIENTES_APP_ID
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
from app.src.modules.expedientes.application.runtime_policies.service import RuntimePolicyService


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


@dataclass
class _RaisingAudit:
    async def append(self, event: object) -> None:
        raise RuntimeError("audit sink unavailable")


_CAPS = {
    "readiness": "EXP-CAP-047",
    "config": "EXP-CAP-048",
    "cache": "EXP-CAP-049",
    "binding": "EXP-CAP-050",
}
_REPORTS = {
    "readiness": ReadinessReport,
    "config": RuntimeConfig,
    "cache": CacheReport,
    "binding": BindingReport,
}


def _cmd(category: str, actor_id: object) -> object:
    if category == "readiness":
        return ReadinessCommand(actor_id=actor_id)
    if category == "config":
        return ConfigCommand(actor_id=actor_id, environment="ci")
    if category == "cache":
        return CacheCommand(actor_id=actor_id, target="read_model")
    return BindingCommand(actor_id=actor_id, deployment="production")


@pytest.mark.parametrize("category", list(_CAPS))
async def test_happy_path_emits_one_event_with_correct_cap(category: str) -> None:
    actor = uuid4()
    audit = _Audit()
    svc = RuntimePolicyService(audit_log=audit)
    method_name = {
        "readiness": "check_readiness",
        "config": "load_config",
        "cache": "rebuild_cache",
        "binding": "resolve_backend",
    }[category]

    result = await getattr(svc, method_name)(_cmd(category, actor_id=actor))

    assert isinstance(result, _REPORTS[category])
    event = audit.events[0]
    assert event.event_type == f"runtime.{category}.recorded"
    assert event.actor_id == actor and event.capacidad == _CAPS[category]
    assert event.payload["app_id"] == EXPEDIENTES_APP_ID


@pytest.mark.parametrize("category", list(_CAPS))
async def test_missing_actor_raises_and_skips_audit(category: str) -> None:
    audit = _Audit()
    svc = RuntimePolicyService(audit_log=audit)
    method_name = {
        "readiness": "check_readiness",
        "config": "load_config",
        "cache": "rebuild_cache",
        "binding": "resolve_backend",
    }[category]

    with pytest.raises(RuntimeAuthorizationError, match="actor_id"):
        await getattr(svc, method_name)(_cmd(category, actor_id=None))
    assert audit.events == []


async def test_dependency_failure_wraps_underlying_exception() -> None:
    svc = RuntimePolicyService(audit_log=_RaisingAudit())
    with pytest.raises(RuntimeDependencyError) as excinfo:
        await svc.check_readiness(ReadinessCommand(actor_id=uuid4()))
    assert isinstance(excinfo.value.__cause__, RuntimeError)


async def test_two_independent_services_emit_distinct_events() -> None:
    actor_a, actor_b = uuid4(), uuid4()
    audit = _Audit()
    svc_a, svc_b = RuntimePolicyService(audit_log=audit), RuntimePolicyService(audit_log=audit)

    await svc_a.check_readiness(ReadinessCommand(actor_id=actor_a))
    await svc_b.check_readiness(ReadinessCommand(actor_id=actor_b))

    assert [e.actor_id for e in audit.events] == [actor_a, actor_b]
    assert audit.events[0].id != audit.events[1].id


async def test_load_config_reports_environment() -> None:
    audit = _Audit()
    svc = RuntimePolicyService(audit_log=audit)
    result = await svc.load_config(ConfigCommand(actor_id=uuid4(), environment="ci"))
    assert result.environment == "ci"
    assert audit.events[0].payload["environment"] == "ci"  # type: ignore[attr-defined]


async def test_resolve_backend_reports_deployment() -> None:
    audit = _Audit()
    svc = RuntimePolicyService(audit_log=audit)
    result = await svc.resolve_backend(BindingCommand(actor_id=uuid4(), deployment="production"))
    assert isinstance(result, BindingReport) and result.deployment == "production"
    assert audit.events[0].payload["deployment"] == "production"  # type: ignore[attr-defined]
