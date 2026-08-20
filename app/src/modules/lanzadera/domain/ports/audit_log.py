"""AuditLogPort.

D27, D55, DA-11. Append-only boundary — no ``delete`` method, by
contract. Every mutating case-of-use inside the auth/session
boundary MUST invoke ``append`` in the same SQLAlchemy session as the
mutation it audits, so a failed insert propagates and the surrounding
transaction rolls back (DA-11).

D55 strips every telemetry column the legacy ``TbConexiones`` /
``TbAplicacionesAperturas`` carried: SSID, BSSID, machine name,
coordinates, IP. Any reintroduction is rejected by the legacy-pin
gate (``scripts/check_legacy_hashes.py``), so the contract here is
enforced by CI not by adapter code.
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import datetime
from typing import Protocol
from uuid import UUID

from app.src.modules.lanzadera.domain.audit_event import AuditEvent


class AuditLogPort(Protocol):
    """Append-only audit-log boundary."""

    async def append(self, event: AuditEvent) -> None:
        """Persist an audit row.

        Must be called inside the same SQLAlchemy session as the
        mutation it audits — a failed insert here MUST propagate to
        the caller so the surrounding transaction rolls back.
        """
        ...

    async def list_for_actor(self, actor_id: UUID, since: datetime) -> Sequence[AuditEvent]:
        """Return every audit row with ``actor_id == actor_id`` and ``created_at >= since``."""
        ...
