"""PostgreSQL adapter for :class:`AuditLogPort`.

D27 + DA-11 + D55: append-only audit log with NO telemetry columns
(SSID, BSSID, machine_name, coordinates, IP — explicitly absent per
D55, pinned by ``scripts/check_legacy_hashes.py``).

**Skeleton** — the bodies raise ``NotImplementedError`` until migration
``0006_seed_audit`` lands. The contract is declared today so every
case-of-use can pin its ``append`` calls against ``typing.Protocol``
structural conformance, and so the same-transaction rollback test
(``tests/lanzadera/audit/test_same_transaction_audit.py``) can wire
against a real adapter when #45 ships.
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import datetime
from uuid import UUID

from app.src.modules.lanzadera.domain.audit_event import AuditEvent
from app.src.modules.lanzadera.domain.ports.audit_log import AuditLogPort


class AuditLogPg(AuditLogPort):
    """SQLAlchemy Core + asyncpg implementation. Skeleton — see module docstring."""

    async def append(self, event: AuditEvent) -> None:
        raise NotImplementedError("AuditLogPg.append awaits migration 0006_seed_audit")

    async def list_for_actor(self, actor_id: UUID, since: datetime) -> Sequence[AuditEvent]:
        raise NotImplementedError("AuditLogPg.list_for_actor awaits migration 0006_seed_audit")
