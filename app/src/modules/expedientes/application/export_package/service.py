"""Service implementation for EXP-CAP-037 E2E package + sink (E06)."""

from __future__ import annotations

from typing import Any
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.export_package._evidence import (
    package_event,
)
from app.src.modules.expedientes.application.export_package.command import (
    ExportPackageAuthorizationError,
    ExportPackageCommand,
    ExportPackageError,
    ExportPackageResult,
    ExportPackageValidationError,
)

API_VERSION = "1.0"


class ExportPackageService:
    """Application-level E2E package assembly + sink dispatch (CAP-037).

    The service assembles a deterministic package (manifest + items
    sorted by ordinal) and dispatches it to the injected sink. The
    sink is opaque: production wiring may use S3, SharePoint, an
    internal queue, or a fake. If the sink fails, the service raises
    without emitting an audit event (the rollback is the only signal
    the delivery layer needs).
    """

    def __init__(
        self,
        *,
        audit_log: Any,
        sink: Any,
        permissions: set[str],
    ) -> None:
        self._audit_log = audit_log
        self._sink = sink
        self._permissions = permissions

    def grant(self, permission: str) -> None:
        self._permissions.add(permission)

    async def assemble(self, command: ExportPackageCommand) -> ExportPackageResult:
        self._check_actor(command.actor_id)
        if not command.items:
            raise ExportPackageValidationError("package is empty")

        package_id = uuid4()
        sorted_items = sorted(command.items, key=lambda it: it.ordinal)
        package = {
            "package_id": str(package_id),
            "api_version": API_VERSION,
            "manifest": {
                "package_id": str(package_id),
                "api_version": API_VERSION,
                "item_count": len(sorted_items),
            },
            "items": [
                {
                    "id": str(item.id),
                    "ordinal": item.ordinal,
                    "data": item.data,
                }
                for item in sorted_items
            ],
        }
        try:
            await self._sink.dispatch(package)
        except Exception as exc:
            raise ExportPackageError(f"sink failed for package {package_id!r}: {exc}") from exc
        assert command.actor_id is not None
        await self._audit_log.append(package_event(command.actor_id, package))
        return ExportPackageResult(package=package, manifest_count=len(sorted_items))

    def _check_actor(self, actor_id: UUID | None) -> None:
        if actor_id is None:
            raise ExportPackageAuthorizationError("actor_id is required (deny-by-default)")
        if "e2e.package" not in self._permissions:
            raise ExportPackageAuthorizationError("actor lacks permission 'e2e.package'")


__all__ = ["ExportPackageService"]
