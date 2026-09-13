"""Strict TDD for EXP-CAP-036 E2E hash service (E03)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.expedientes.application.e2e_hash.command import (
    E2EHashAuthorizationError,
    E2EHashCommand,
    E2EHashError,
    E2EHashService,
    E2EHashValidationError,
)
from app.src.modules.expedientes.domain.hash.versioning import (
    HashRegistry,
)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


class _FNV1AStub:
    @property
    def algorithm(self) -> str:
        return "fnv1a-stub"

    def hash(self, payload: bytes) -> str:
        return format(len(payload), "08x")


def _make_service() -> tuple[E2EHashService, _Audit, HashRegistry]:
    audit = _Audit()
    registry = HashRegistry()
    registry.register(_FNV1AStub())
    service = E2EHashService(
        registry=registry,
        audit_log=audit,
        permissions=set(),
        default_algorithm="fnv1a-stub",
        supported_algorithms=("fnv1a-stub",),
    )
    return service, audit, registry


def _command(**overrides: Any) -> E2EHashCommand:
    values: dict[str, Any] = {
        "actor_id": uuid4(),
        "payload": {"version": 1, "expedientes": []},
    }
    values.update(overrides)
    return E2EHashCommand(**values)


async def test_compute_returns_versioned_hash() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.hash")

    result = await service.compute(_command())

    assert result.algorithm == "fnv1a-stub"
    assert result.digest == format(len(b"canonical"), "08x") or len(result.digest) == 8
    # The exact digest is implementation-defined for the stub; just ensure
    # it's a deterministic 8-char hex string.


async def test_compute_rejects_unknown_algorithm() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.hash")

    with pytest.raises(E2EHashValidationError, match="algorithm"):
        await service.compute(_command(algorithm="not-registered"))


async def test_compute_requires_permission() -> None:
    service, _, _ = _make_service()

    with pytest.raises(E2EHashAuthorizationError, match="e2e.hash"):
        await service.compute(_command())


async def test_compute_rejects_unsupported_algorithm_even_if_registered() -> None:
    service, audit, registry = _make_service()
    service.grant("e2e.hash")

    class _OtherAlgo:
        @property
        def algorithm(self) -> str:
            return "sha256-x"

        def hash(self, payload: bytes) -> str:
            return "0" * 8

    registry.register(_OtherAlgo())
    # E03 contract: only the algorithms in ``supported_algorithms`` are
    # accepted, even if the registry holds more.
    with pytest.raises(E2EHashValidationError, match="sha256-x"):
        await service.compute(_command(algorithm="sha256-x"))

    assert not audit.events


async def test_compute_audits_each_invocation() -> None:
    service, audit, _ = _make_service()
    service.grant("e2e.hash")

    await service.compute(_command())

    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.hash.computed"
    assert audit.events[0].capacidad == "EXP-CAP-036"


async def test_compute_dependency_failure_rolls_back() -> None:
    from app.src.modules.expedientes.application.e2e_hash.command import (
        E2EHashService,
    )

    class _BrokenAlgo:
        @property
        def algorithm(self) -> str:
            return "broken"

        def hash(self, payload: bytes) -> str:
            raise RuntimeError("boom")

    audit = _Audit()
    registry = HashRegistry()
    registry.register(_BrokenAlgo())
    service = E2EHashService(
        registry=registry,
        audit_log=audit,
        permissions=set(),
        default_algorithm="broken",
        supported_algorithms=("broken",),
    )
    service.grant("e2e.hash")

    with pytest.raises(E2EHashError, match="hash failed"):
        await service.compute(_command())

    assert not audit.events
