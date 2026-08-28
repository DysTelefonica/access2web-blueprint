# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin use case: audit_append (DA-11, D55).

Thin wrapper over ``AuditLogPort.append`` that always injects the
module name ``lanzadera`` and stamps a fresh correlation id. D55:
``payload`` MUST NOT carry any of the forbidden legacy-telemetry
columns (SSID, BSSID, machine name, coordinates, IP) — the legacy-pin
gate rejects any reintroduction. The check is enforced at the
domain level via the ``AuditLogEntry`` value object (the dataclass
does not declare those fields); callers who need telemetry do so via
the structured-logger adapter, NOT through this use case.
"""

from __future__ import annotations

from datetime import datetime
from uuid import UUID, uuid4

from app.src.modules.lanzadera.domain.ports import AuditLog, AuditLogEntry


async def audit_append(
    event_type: str,
    *,
    actor_id: UUID | None,
    target_id: str,
    result: str,
    payload: dict[str, object],
    audit: AuditLog,
    now: datetime,
    correlation_id: UUID | None = None,
    module: str = "lanzadera",
) -> AuditLogEntry:
    """Persist an audit row with the canonical lanzadera shape.

    Returns the freshly-stamped entry so the caller can correlate the
    response with the audit log row. D55: the value object
    constructor enforces ``event_type`` / ``module`` / ``result`` are
    non-empty strings (empty values would slip past validation
    downstream); the caller does not need to repeat the check.
    """
    cid = correlation_id if correlation_id is not None else uuid4()
    entry = AuditLogEntry(
        event_type=event_type,
        actor_id=actor_id,
        target_id=target_id,
        module=module,
        result=result,
        correlation_id=cid,
        payload=payload,
        created_at=now,
    )
    await audit.append(entry)
    return entry


__all__ = ["audit_append"]
