"""Service implementation for EXP-CAP-051 HPS adapter (H01).

``HpsService`` wraps ``HpsPort`` with deny-by-default, one audit event
per decision, and dedupe by ``idempotency_key`` via the existing
``IdempotencyPort``. ``submit`` and ``query`` are duck-typed to keep
cross-module imports out of the application layer
(``scripts/check_layers.py``).
"""

from __future__ import annotations

import hashlib
import json
from typing import Any
from uuid import UUID

from app.src.modules.expedientes.application.hps._evidence import hps_event
from app.src.modules.expedientes.application.hps.command import (
    HpsAuthorizationError,
    HpsDependencyError,
    HpsQueryCommand,
    HpsResult,
    HpsSubmitCommand,
)
from app.src.modules.expedientes.ports.hps import HpsRecord, HpsSubmission


class HpsService:
    """Application-level HPS integration (CAP-051). Deny-by-default + idempotent."""

    def __init__(self, *, hps: Any, audit_log: Any, idempotency: Any) -> None:
        self._hps = hps
        self._audit_log = audit_log
        self._idempotency = idempotency

    async def submit(self, command: HpsSubmitCommand) -> HpsResult:
        actor_id = _validated_actor(command.actor_id)
        submission = HpsSubmission(
            idempotency_key=command.idempotency_key,
            expediente_id=command.expediente_id,
            payload=command.payload,
        )
        signature = _signature(command.idempotency_key, command.payload)
        existing = await self._idempotency.get(command.idempotency_key)
        if existing is not None and existing.request_signature == signature:
            prior = existing.result
            assert isinstance(prior, HpsRecord)  # typed on the way out
            await self._audit_log.append(
                hps_event(
                    "submit",
                    actor_id=actor_id,
                    target_id=prior.reference,
                    payload={"deduped": True, "key": str(command.idempotency_key)},
                )
            )
            return _result_from(actor_id, prior, deduped=True)
        try:
            record = await self._hps.submit(submission, command.credential)
        except Exception as exc:  # noqa: BLE001 — wrap any driver failure
            raise HpsDependencyError("hps.submit failed; caller UoW must roll back") from exc
        await self._idempotency.record(
            _idempotency_record(command.idempotency_key, signature, record)
        )
        await self._audit_log.append(
            hps_event(
                "submit",
                actor_id=actor_id,
                target_id=record.reference,
                payload={"deduped": False, "key": str(command.idempotency_key)},
            )
        )
        return _result_from(actor_id, record, deduped=False)

    async def query(self, command: HpsQueryCommand) -> HpsResult:
        actor_id = _validated_actor(command.actor_id)
        try:
            record = await self._hps.query(command.reference, command.credential)
        except Exception as exc:  # noqa: BLE001 — wrap any driver failure
            raise HpsDependencyError("hps.query failed; caller UoW must roll back") from exc
        await self._audit_log.append(
            hps_event(
                "query",
                actor_id=actor_id,
                target_id=record.reference,
                payload={"reference": str(command.reference)},
            )
        )
        return _result_from(actor_id, record, deduped=False)


def _validated_actor(actor_id: UUID | None) -> UUID:
    if actor_id is None or actor_id == UUID(int=0):
        raise HpsAuthorizationError("actor_id is required (deny-by-default)")
    return actor_id


def _signature(key: UUID, payload: dict[str, Any]) -> str:
    blob = json.dumps(payload, sort_keys=True, default=str).encode("utf-8")
    return f"{key}|{hashlib.blake2b(blob, digest_size=32).hexdigest()}"


def _result_from(actor_id: UUID, record: HpsRecord, *, deduped: bool) -> HpsResult:
    return HpsResult(
        actor_id=actor_id,
        reference=record.reference,
        status=record.status,
        deduped=deduped,
        payload=record.payload,
    )


def _idempotency_record(key: UUID, signature: str, result: HpsRecord) -> Any:
    """Build an ``IdempotencyRecord``; typed as ``Any`` to avoid importing the
    port dataclass directly (cross-layer gate)."""
    from app.src.modules.expedientes.ports.idempotency import IdempotencyRecord

    return IdempotencyRecord(key=key, request_signature=signature, result=result)


__all__ = ["HpsService"]
