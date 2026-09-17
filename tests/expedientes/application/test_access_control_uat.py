"""UAT end-to-end for EXP-CAP-043..046 access-control lane (issue #280, U07).

Composes A01+A02+A03+A04 over one audit log. Does NOT re-test individual
capabilities — A01..A04 unit tests cover that ground.
"""

from dataclasses import dataclass, field
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.audit.command import AuditCommand
from app.src.modules.expedientes.application.audit.service import AuditService
from app.src.modules.expedientes.application.authorization.command import AuthorizationCommand
from app.src.modules.expedientes.application.authorization.service import AuthorizationService
from app.src.modules.expedientes.application.current_principal.command import (
    CurrentPrincipalCommand,
    Principal,
)
from app.src.modules.expedientes.application.current_principal.service import (
    CurrentPrincipalService,
)
from app.src.modules.expedientes.application.session_state.command import SessionStateCommand
from app.src.modules.expedientes.application.session_state.service import SessionStateService


@dataclass
class _AuditLog:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


@dataclass
class _Lanz:
    p: Principal | None = None
    exc: BaseException | None = None

    async def get_principal(self, actor_id: UUID) -> Principal | None:
        if self.exc is not None:
            raise self.exc
        return self.p


def _stack(
    audit: _AuditLog, *, p: Principal | None = None, exc: BaseException | None = None
) -> tuple[AuthorizationService, CurrentPrincipalService, SessionStateService, AuditService]:
    lanz = _Lanz(p=p, exc=exc)
    return (
        AuthorizationService(audit_log=audit),
        CurrentPrincipalService(lanzadera_principal=lanz, audit_log=audit),
        SessionStateService(audit_log=audit),
        AuditService(audit_log=audit),
    )


def _types(audit: _AuditLog) -> list[str]:
    return [e.event_type for e in audit.events]  # type: ignore[attr-defined]


async def test_happy_path_emits_loaded_bound_allowed() -> None:
    actor = uuid4()
    perms = frozenset({"EXP-CAP-045"})
    audit = _AuditLog()
    authz, current, session, _ = _stack(audit, p=Principal(actor, 19, perms))

    await current.get_current(CurrentPrincipalCommand(actor_id=actor))
    await session.bind(SessionStateCommand(actor_id=actor))
    decision = await authz.authorize(
        AuthorizationCommand(actor_id=actor, capability="EXP-CAP-045", permissions=set(perms))
    )

    assert decision.allowed is True
    assert _types(audit) == ["authz.principal.loaded", "authz.session.bound", "e2e.authz.allowed"]


async def test_deny_path_emits_loaded_bound_denied() -> None:
    actor = uuid4()
    audit = _AuditLog()
    authz, current, session, _ = _stack(audit, p=Principal(actor, 19, frozenset({"EXP-CAP-001"})))
    await current.get_current(CurrentPrincipalCommand(actor_id=actor))
    await session.bind(SessionStateCommand(actor_id=actor))
    decision = await authz.authorize(
        AuthorizationCommand(actor_id=actor, capability="EXP-CAP-045", permissions={"EXP-CAP-001"})
    )

    assert decision.allowed is False and decision.denied_reason == "deny_by_default"
    assert _types(audit) == ["authz.principal.loaded", "authz.session.bound", "e2e.authz.denied"]


async def test_lanzadera_failure_fail_closed_emits_only_failed_audit() -> None:
    actor = uuid4()
    audit = _AuditLog()
    _, current, session, _ = _stack(audit, exc=RuntimeError("lanzadera down"))

    result = await current.get_current(CurrentPrincipalCommand(actor_id=actor))

    assert result.loaded is False and result.principal.permissions == frozenset()
    assert session.get_current() is None
    assert _types(audit) == ["authz.principal.failed"]
    assert audit.events[0].payload["reason"] == "lanzadera_error"  # type: ignore[attr-defined]


async def test_per_request_contexts_emit_independent_chains() -> None:
    actor_a, actor_b = uuid4(), uuid4()
    perms = frozenset({"EXP-CAP-045"})
    audit = _AuditLog()
    _, current_a, session_a, _ = _stack(audit, p=Principal(actor_a, 19, perms))
    _, current_b, session_b, _ = _stack(audit, p=Principal(actor_b, 19, perms))

    await current_a.get_current(CurrentPrincipalCommand(actor_id=actor_a))
    await session_a.bind(SessionStateCommand(actor_id=actor_a))
    await current_b.get_current(CurrentPrincipalCommand(actor_id=actor_b))
    await session_b.bind(SessionStateCommand(actor_id=actor_b))

    actors = [e.actor_id for e in audit.events]  # type: ignore[attr-defined]
    assert actors == [actor_a, actor_a, actor_b, actor_b]
    assert session_a.get_current() != session_b.get_current()


async def test_full_session_lifecycle_emits_bound_access_cleared() -> None:
    actor = uuid4()
    perms = frozenset({"EXP-CAP-045"})
    audit = _AuditLog()
    _, current, session, audit_svc = _stack(audit, p=Principal(actor, 19, perms))

    await current.get_current(CurrentPrincipalCommand(actor_id=actor))
    await session.bind(SessionStateCommand(actor_id=actor))
    await audit_svc.record_access(
        AuditCommand(actor_id=actor, target_id=uuid4(), correlation_id=None, payload={})
    )
    await session.clear()

    assert _types(audit) == [
        "authz.principal.loaded",
        "authz.session.bound",
        "audit.access.recorded",
        "authz.session.cleared",
    ]


async def test_audit_correlation_id_propagates_to_recorded_event() -> None:
    actor, corr = uuid4(), uuid4()
    audit = _AuditLog()
    _, _, _, audit_svc = _stack(audit, p=Principal(actor, 19, frozenset({"EXP-CAP-045"})))

    await audit_svc.record_access(
        AuditCommand(actor_id=actor, target_id=actor, correlation_id=corr, payload={})
    )

    assert audit.events[0].correlation_id == corr  # type: ignore[attr-defined]
