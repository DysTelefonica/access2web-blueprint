"""Strict TDD for EXP-CAP-042 E2E ordinal / family cycle expansion (E04)."""

from __future__ import annotations

from dataclasses import dataclass, field
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.ordinal_e2e.command import (
    FamilyNode,
    OrdinalE2EAuthorizationError,
    OrdinalE2EService,
    OrdinalE2EValidationError,
)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


def _node(
    id: UUID | None = None,
    parent_id: UUID | None = None,
    ordinal: int | None = None,
) -> FamilyNode:
    return FamilyNode(
        id=id or uuid4(),
        parent_id=parent_id,
        ordinal=ordinal,
    )


def _make_service() -> tuple[OrdinalE2EService, _Audit]:
    audit = _Audit()
    service = OrdinalE2EService(
        audit_log=audit,
        permissions=set(),
    )
    return service, audit


async def test_assigns_ordinals_in_dfs_pre_order() -> None:
    service, _ = _make_service()
    service.grant("e2e.ordinal")

    root = _node()
    a = _node()
    a1 = _node()
    b = _node()

    result = await service.assign(
        actor_id=uuid4(),
        family=[
            _node(id=root.id, ordinal=None),
            _node(id=a.id, parent_id=root.id, ordinal=None),
            _node(id=a1.id, parent_id=a.id, ordinal=None),
            _node(id=b.id, parent_id=root.id, ordinal=None),
        ],
    )

    assert result.root_ordinal == 1
    by_id = {assignment.node_id: assignment.ordinal for assignment in result.assignments}
    assert by_id[root.id] == 1
    assert by_id[a.id] == 2
    assert by_id[a1.id] == 3
    assert by_id[b.id] == 4


async def test_cycle_is_detected_and_no_ordinal_assigned() -> None:
    service, audit = _make_service()
    service.grant("e2e.ordinal")
    self_loop_root = _node()

    result = await service.assign(
        actor_id=uuid4(),
        family=[
            # Root is its own parent — the only cycle possible with a
            # single declared root.
            _node(id=self_loop_root.id, parent_id=self_loop_root.id, ordinal=None),
        ],
    )

    assert result.assignments == ()
    assert self_loop_root.id in result.cycle
    assert not audit.events


async def test_unknown_parent_is_rejected() -> None:
    service, _ = _make_service()
    service.grant("e2e.ordinal")
    orphan = _node(parent_id=uuid4())

    with pytest.raises(OrdinalE2EValidationError, match="unknown parent"):
        await service.assign(
            actor_id=uuid4(),
            family=[_node(), orphan],
        )


async def test_assign_requires_permission() -> None:
    service, _ = _make_service()

    with pytest.raises(OrdinalE2EAuthorizationError, match="e2e.ordinal"):
        await service.assign(actor_id=uuid4(), family=[_node()])


async def test_assign_audits_successful_expansion() -> None:
    service, audit = _make_service()
    service.grant("e2e.ordinal")
    root = _node()
    child = _node(parent_id=root.id, ordinal=None)

    result = await service.assign(
        actor_id=uuid4(),
        family=[root, child],
    )

    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.ordinal.assigned"
    assert audit.events[0].capacidad == "EXP-CAP-042"
    assert audit.events[0].payload["root_id"] == str(result.root_ordinal)


async def test_two_roots_are_rejected() -> None:
    service, _ = _make_service()
    service.grant("e2e.ordinal")

    with pytest.raises(OrdinalE2EValidationError, match="single root"):
        await service.assign(
            actor_id=uuid4(),
            family=[_node(), _node()],
        )


async def test_empty_family_is_rejected() -> None:
    service, _ = _make_service()
    service.grant("e2e.ordinal")

    with pytest.raises(OrdinalE2EValidationError, match="empty"):
        await service.assign(actor_id=uuid4(), family=[])


async def test_assignments_are_deterministic_across_calls() -> None:
    service, _ = _make_service()
    service.grant("e2e.ordinal")
    root = _node()
    family = [
        root,
        _node(parent_id=root.id, ordinal=None),
    ]
    # Re-run: same input must produce same ordinals.
    first = await service.assign(actor_id=uuid4(), family=family)
    second = await service.assign(actor_id=uuid4(), family=family)
    assert first.assignments == second.assignments
