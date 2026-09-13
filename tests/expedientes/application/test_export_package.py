"""Strict TDD for EXP-CAP-037 E2E package + sink (E06)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.export_package.command import (
    ExportItem,
    ExportPackageAuthorizationError,
    ExportPackageCommand,
    ExportPackageError,
    ExportPackageValidationError,
)
from app.src.modules.expedientes.application.export_package.service import (
    ExportPackageService,
)


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)
    changes: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)

    async def record_change(self, change: object) -> None:
        self.changes.append(change)


@dataclass
class _Sink:
    name: str = "in_memory"
    calls: list[dict[str, Any]] = field(default_factory=list)
    fail_on: set[UUID] = field(default_factory=set)

    async def dispatch(self, package: dict[str, Any]) -> str:
        if package["package_id"] in self.fail_on:
            raise RuntimeError(f"sink failed for {package['package_id']}")
        self.calls.append(package)
        return f"{self.name}:{package['package_id']}"


def _make_service() -> tuple[ExportPackageService, _Audit, _Sink]:
    audit = _Audit()
    sink = _Sink()
    service = ExportPackageService(
        audit_log=audit,
        sink=sink,
        permissions=set(),
    )
    return service, audit, sink


def _command(items: list[ExportItem], **overrides: Any) -> ExportPackageCommand:
    values: dict[str, Any] = {
        "actor_id": uuid4(),
        "items": items,
    }
    values.update(overrides)
    return ExportPackageCommand(**values)


def _item(ordinal: int = 1, data: dict[str, Any] | None = None) -> ExportItem:
    return ExportItem(id=uuid4(), ordinal=ordinal, data=data or {})


async def test_assembles_package_with_manifest_and_items() -> None:
    service, audit, sink = _make_service()
    service.grant("e2e.package")
    items = [_item(ordinal=1), _item(ordinal=2), _item(ordinal=3)]

    result = await service.assemble(_command(items))

    assert result.manifest_count == 3
    assert len(result.package["items"]) == 3
    assert result.package["api_version"] == "1.0"
    assert len(audit.events) == 1
    assert audit.events[0].event_type == "e2e.package.assembled"
    assert audit.events[0].capacidad == "EXP-CAP-037"


async def test_package_items_sorted_by_ordinal() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.package")
    items = [_item(ordinal=3), _item(ordinal=1), _item(ordinal=2)]

    result = await service.assemble(_command(items))

    ordinals = [item["ordinal"] for item in result.package["items"]]
    assert ordinals == [1, 2, 3]


async def test_sink_dispatch_called_with_assembled_package() -> None:
    service, _, sink = _make_service()
    service.grant("e2e.package")

    await service.assemble(_command([_item()]))

    assert len(sink.calls) == 1
    assert sink.calls[0]["package_id"]  # service-generated id is present


async def test_sink_failure_rolls_back_assembly() -> None:
    service, audit, sink = _make_service()
    service.grant("e2e.package")
    items = [_item()]
    sink.fail_on = set()
    # Use the package_id once we know it: actually we cannot because
    # it's service-generated. To exercise the failure we use a sink
    # that always fails.

    class _AlwaysFailsSink(_Sink):
        async def dispatch(self, package: dict[str, Any]) -> str:
            raise RuntimeError("boom")

    service._sink = _AlwaysFailsSink()

    with pytest.raises(ExportPackageError, match="sink failed"):
        await service.assemble(_command(items))

    assert not audit.events


async def test_empty_package_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.package")

    with pytest.raises(ExportPackageValidationError, match="empty"):
        await service.assemble(_command([]))


async def test_missing_actor_rejected() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.package")

    with pytest.raises(ExportPackageAuthorizationError):
        await service.assemble(_command([_item()], actor_id=None))


async def test_missing_permission_rejected() -> None:
    service, _, _ = _make_service()
    items = [_item()]

    with pytest.raises(ExportPackageAuthorizationError, match="e2e.package"):
        await service.assemble(_command(items))


async def test_two_assemblies_get_distinct_ids() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.package")

    first = await service.assemble(_command([_item()]))
    second = await service.assemble(_command([_item()]))

    assert first.package["package_id"] != second.package["package_id"]


async def test_manifest_includes_count_and_api_version() -> None:
    service, _, _ = _make_service()
    service.grant("e2e.package")

    result = await service.assemble(_command([_item(), _item()]))

    assert result.package["manifest"]["api_version"] == "1.0"
    assert result.package["manifest"]["item_count"] == 2
    assert "package_id" in result.package["manifest"]
