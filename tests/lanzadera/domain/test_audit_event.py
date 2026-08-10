# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, DA-11, D55 — `AuditEvent` is the canonical audit row. **No**
# telemetry fields (SSID, BSSID, machine name, coordinates, IP). All
# identifiers live on the schema; payload is an opaque JSONB dict.
"""Strict TDD — `AuditEvent` entity (Phase 1, task 1.7)."""

from __future__ import annotations

from datetime import datetime, timezone
from uuid import UUID, uuid4

import pytest

from app.src.modules.lanzadera.domain.audit_event import AuditEvent


# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


def _now() -> datetime:
    return datetime(2026, 8, 9, 12, 0, 0, tzinfo=timezone.utc)


def _new_event(
    *,
    event_type: str = "auth.login.success",
    actor_id: UUID | None = None,
    target_id: str = "[email protected]",
    module: str = "lanzadera",
    result: str = "ok",
    correlation_id: UUID | None = None,
    payload: dict | None = None,
) -> AuditEvent:
    return AuditEvent(
        id=uuid4(),
        event_type=event_type,
        actor_id=actor_id,
        target_id=target_id,
        module=module,
        result=result,
        correlation_id=correlation_id or uuid4(),
        payload=payload if payload is not None else {},
        created_at=_now(),
    )


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------


class TestAuditEventConstruction:
    def test_event_carries_required_attributes(self) -> None:
        event_id = uuid4()
        actor_id = uuid4()
        correlation_id = uuid4()
        event = AuditEvent(
            id=event_id,
            event_type="auth.login.success",
            actor_id=actor_id,
            target_id="[email protected]",
            module="lanzadera",
            result="ok",
            correlation_id=correlation_id,
            payload={"reason": "first_login"},
            created_at=_now(),
        )
        assert event.id == event_id
        assert event.event_type == "auth.login.success"
        assert event.actor_id == actor_id
        assert event.target_id == "[email protected]"
        assert event.module == "lanzadera"
        assert event.result == "ok"
        assert event.correlation_id == correlation_id
        assert event.payload == {"reason": "first_login"}

    def test_event_type_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="event_type"):
            _new_event(event_type="")

    def test_module_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="module"):
            _new_event(module="")

    def test_result_must_be_non_empty(self) -> None:
        with pytest.raises(ValueError, match="result"):
            _new_event(result="")


# ---------------------------------------------------------------------------
# DA-11 — actor_id is nullable (system-issued events have no actor)
# ---------------------------------------------------------------------------


class TestAuditEventActorNullable:
    def test_actor_id_can_be_none(self) -> None:
        event = _new_event(actor_id=None)
        assert event.actor_id is None

    def test_system_audit_event_has_no_actor(self) -> None:
        """Boot-time and migration audit rows carry no actor."""
        event = _new_event(event_type="audit.system.boot", actor_id=None)
        assert event.actor_id is None
        assert event.event_type == "audit.system.boot"


# ---------------------------------------------------------------------------
# DA-11 — no telemetry fields
# ---------------------------------------------------------------------------


class TestAuditEventNoTelemetry:
    """DA-11: schema MUST NOT carry SSID, BSSID, machine name, coordinates, IP."""

    def test_event_has_no_telemetry_field(self) -> None:
        """The dataclass declares no field named after any telemetry column."""
        import dataclasses as _dc

        field_names = {field.name for field in _dc.fields(AuditEvent)}
        forbidden = {"ssid", "bssid", "coordinates", "machine_name", "ip_address"}
        assert forbidden.isdisjoint(field_names), (
            f"AuditEvent declares telemetry fields: {forbidden & field_names}"
        )


# ---------------------------------------------------------------------------
# Payload is an opaque JSONB-shaped dict
# ---------------------------------------------------------------------------


class TestAuditEventPayload:
    def test_payload_defaults_to_empty_dict(self) -> None:
        event = _new_event()
        assert event.payload == {}

    def test_payload_accepts_arbitrary_jsonb(self) -> None:
        payload = {"reason": "lockout", "attempts": 5, "force_reset": True}
        event = _new_event(payload=payload)
        assert event.payload == payload
