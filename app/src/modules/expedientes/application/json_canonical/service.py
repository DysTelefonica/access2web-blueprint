"""Canonical JSON service for EXP-CAP-034."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.json_canonical._evidence import (
    render_event,
)
from app.src.modules.expedientes.application.json_canonical.command import (
    CANONICAL_API_VERSION,
    CANONICAL_COLLECTIONS,
    JsonCanonicalAuthorizationError,
    JsonCanonicalValidationError,
)
from app.src.modules.expedientes.ports.audit_log import AuditLogPort


def _iso_z(value: Any) -> Any:
    """Return ``value`` serialised as an ISO-Z string, or ``None``."""
    if value is None:
        return None
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.isoformat() + "Z"
        return value.astimezone(UTC).isoformat().replace("+00:00", "Z")
    return value


def _order_key(row: dict[str, Any]) -> tuple[int, str]:
    """Stable ordering by OrdinalE2E then id."""
    ordinal = row.get("ordinal_e2e", 0)
    raw_id = row.get("id", "")
    rid = str(raw_id) if raw_id is not None else ""
    return (ordinal if isinstance(ordinal, int) else 0, rid)


class JsonCanonicalService:
    """Render the canonical E2E JSON envelope.

    Produces ``{meta, data}`` with ``apiVersion 1.0``, nine
    collections, ISO-Z dates, explicit ``None`` for null fields, and
    stable ``OrdinalE2E+id`` ordering. The actual byte-rendering
    stays an adapter concern; the service returns a typed dict that
    mirrors the spec.
    """

    def __init__(
        self,
        *,
        audit_log: AuditLogPort,
        permissions: set[str] | None = None,
        api_version: str = CANONICAL_API_VERSION,
        collections: tuple[str, ...] = CANONICAL_COLLECTIONS,
        now_provider: Callable[[], datetime] | None = None,
    ) -> None:
        self._audit_log = audit_log
        self._permissions: set[str] = set(permissions or ())
        self._api_version = api_version
        self._collections = collections
        self._now = now_provider or (lambda: datetime.now(UTC))

    def grant(self, permission: str) -> None:
        self._permissions.add(permission)

    async def render(
        self, *, collections: dict[str, list[dict[str, Any]]], actor_id: UUID | None = None
    ) -> dict[str, Any]:
        if actor_id is None:
            raise JsonCanonicalAuthorizationError("actor_id is required (deny-by-default)")
        if "e2e.export" not in self._permissions:
            raise JsonCanonicalAuthorizationError("actor lacks permission 'e2e.export'")

        unknown = set(collections) - set(self._collections)
        if unknown:
            raise JsonCanonicalValidationError(f"unknown collection(s): {sorted(unknown)}")

        data: dict[str, list[dict[str, Any]]] = {}
        total_rows = 0
        for name in self._collections:
            rows = list(collections.get(name, ()))
            for _row in rows:
                total_rows += 1
            rows.sort(key=_order_key)
            serialised = [_serialise_row(row) for row in rows]
            data[name] = serialised

        envelope = {
            "meta": {
                "apiVersion": self._api_version,
                "generatedAt": _iso_z(self._now()),
                "totalRows": total_rows,
            },
            "data": data,
        }
        await self._audit_log.append(render_event(actor_id, self._now(), total_rows))
        return envelope


def _serialise_row(row: dict[str, Any]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key, value in row.items():
        if isinstance(value, datetime):
            out[key] = _iso_z(value)
        elif isinstance(value, UUID):
            out[key] = str(value)
        elif hasattr(value, "isoformat") and callable(value.isoformat):
            out[key] = _iso_z(value)
        else:
            out[key] = value
    return out
