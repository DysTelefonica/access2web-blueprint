# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-11, D27, D55 — pure domain entity; only stdlib imports allowed.
"""Audit-event record (lanzadera-mvp/audit).

DA-11 + D55: the audit row carries NO telemetry fields. SSID, BSSID,
machine name, coordinates and IP are explicitly absent — the legacy
`TbConexiones` SSID column was retired by D55 and any reintroduction
is rejected by `scripts/check_legacy_hashes.py` (DA-13).

`actor_id` is nullable: system-issued events (boot, migrations,
automated lockouts) carry no actor. `payload` is an opaque JSONB dict
the application layer fills per event type.
"""

from __future__ import annotations

from datetime import datetime
from dataclasses import dataclass
from uuid import UUID


@dataclass
class AuditEvent:
    """Audit-event row (mutable entity).

    Mutable on purpose: `payload` may evolve after creation when
    additional context arrives in the same transaction (e.g. the lockout
    event records the final `failed_attempts` count). Other attributes
    are set-once.
    """

    id: UUID
    event_type: str
    actor_id: UUID | None
    target_id: str
    module: str
    result: str
    correlation_id: UUID
    payload: dict[str, object]
    created_at: datetime

    def __post_init__(self) -> None:
        _require_non_empty(self.event_type, "event_type")
        _require_non_empty(self.module, "module")
        _require_non_empty(self.result, "result")


def _require_non_empty(value: str, field_name: str) -> None:
    """Reject empty / whitespace-only strings at the domain boundary."""
    if not value or not value.strip():
        raise ValueError(f"AuditEvent.{field_name} must be a non-empty string")
