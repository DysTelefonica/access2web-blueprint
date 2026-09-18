"""Strict TDD for EXP-CAP-051 HPS adapter (H01).

Dataclass fakes inline, no MagicMock. Mirrors A04: tests derived from
the spec's 3 escenarios (camino feliz, validación, concurrencia/fallo).
"""

from dataclasses import dataclass, field
from uuid import UUID, uuid4

import pytest

from app.src.modules.expedientes.application.hps.command import (
    HpsAuthorizationError,
    HpsDependencyError,
    HpsQueryCommand,
    HpsResult,
    HpsSubmitCommand,
)
from app.src.modules.expedientes.application.hps.service import HpsService
from app.src.modules.expedientes.ports.hps import HpsRecord, HpsSubmission
from app.src.modules.expedientes.ports.idempotency import IdempotencyRecord


@dataclass
class _Audit:
    events: list[object] = field(default_factory=list)

    async def append(self, event: object) -> None:
        self.events.append(event)


@dataclass
class _FakeHps:
    submitted: list[HpsSubmission] = field(default_factory=list)
    queried: list[UUID] = field(default_factory=list)
    exc: BaseException | None = None

    async def submit(self, submission: HpsSubmission, credential: str) -> HpsRecord:
        if self.exc is not None:
            raise self.exc
        self.submitted.append(submission)
        return HpsRecord(
            reference=submission.expediente_id, status="submitted", payload=submission.payload
        )

    async def query(self, reference: UUID, credential: str) -> HpsRecord:
        if self.exc is not None:
            raise self.exc
        self.queried.append(reference)
        return HpsRecord(reference=reference, status="queried", payload={"ref": str(reference)})


@dataclass
class _Idem:
    store: dict[UUID, IdempotencyRecord] = field(default_factory=dict)

    async def get(self, key: UUID) -> IdempotencyRecord | None:
        return self.store.get(key)

    async def record(self, record: IdempotencyRecord) -> None:
        self.store[record.key] = record


def _service() -> tuple[HpsService, _Audit, _FakeHps, _Idem]:
    hps = _FakeHps()
    audit = _Audit()
    idem = _Idem()
    return (
        HpsService(hps=hps, audit_log=audit, idempotency=idem),
        audit,
        hps,
        idem,
    )


def _submit(
    actor_id: UUID | None = None, *, key: UUID | None = None, payload: dict | None = None
) -> HpsSubmitCommand:
    return HpsSubmitCommand(
        actor_id=actor_id,
        idempotency_key=key or uuid4(),
        expediente_id=uuid4(),
        payload=payload or {"k": "v"},
        credential="cred",
    )


async def test_submit_happy_path_emits_one_event_and_persists_idempotency() -> None:
    actor, key = uuid4(), uuid4()
    svc, audit, hps, idem = _service()
    result = await svc.submit(_submit(actor_id=actor, key=key))

    assert isinstance(result, HpsResult) and result.deduped is False
    assert len(hps.submitted) == 1 and key in idem.store
    assert audit.events[0].event_type == "hps.submit.recorded"  # type: ignore[attr-defined]


async def test_submit_idempotency_dedupes_when_key_and_payload_match() -> None:
    actor, key = uuid4(), uuid4()
    svc, audit, hps, _ = _service()
    first = await svc.submit(_submit(actor_id=actor, key=key))
    second = await svc.submit(_submit(actor_id=actor, key=key))

    assert len(hps.submitted) == 1
    assert second.deduped is True and second.reference == first.reference
    assert audit.events[1].payload["deduped"] is True  # type: ignore[attr-defined]


async def test_submit_different_payload_same_key_bypasses_dedupe() -> None:
    actor, key = uuid4(), uuid4()
    svc, _, hps, _ = _service()
    await svc.submit(
        HpsSubmitCommand(
            actor_id=actor,
            idempotency_key=key,
            expediente_id=uuid4(),
            payload={"a": 1},
            credential="c",
        )
    )
    await svc.submit(
        HpsSubmitCommand(
            actor_id=actor,
            idempotency_key=key,
            expediente_id=uuid4(),
            payload={"a": 2},
            credential="c",
        )
    )
    assert len(hps.submitted) == 2


async def test_missing_actor_raises_authorization_and_skips_audit() -> None:
    svc, audit, hps, _ = _service()
    with pytest.raises(HpsAuthorizationError, match="actor_id"):
        await svc.submit(_submit(actor_id=None))
    assert audit.events == [] and hps.submitted == []


async def test_dependency_failure_wraps_underlying_exception() -> None:
    hps = _FakeHps(exc=RuntimeError("hps down"))
    svc = HpsService(hps=hps, audit_log=_Audit(), idempotency=_Idem())
    with pytest.raises(HpsDependencyError) as excinfo:
        await svc.submit(_submit(actor_id=uuid4()))
    assert isinstance(excinfo.value.__cause__, RuntimeError)


async def test_query_happy_path_emits_one_event_and_returns_record() -> None:
    actor, ref = uuid4(), uuid4()
    svc, audit, hps, _ = _service()
    result = await svc.query(HpsQueryCommand(actor_id=actor, reference=ref, credential="c"))
    assert result.status == "queried" and hps.queried == [ref]
    assert audit.events[0].event_type == "hps.query.recorded"  # type: ignore[attr-defined]


async def test_two_services_over_same_audit_emit_distinct_event_ids() -> None:
    actor, audit = uuid4(), _Audit()
    svc_a = HpsService(hps=_FakeHps(), audit_log=audit, idempotency=_Idem())
    svc_b = HpsService(hps=_FakeHps(), audit_log=audit, idempotency=_Idem())

    await svc_a.query(HpsQueryCommand(actor_id=actor, reference=uuid4(), credential="c"))
    await svc_b.query(HpsQueryCommand(actor_id=actor, reference=uuid4(), credential="c"))

    assert len(audit.events) == 2
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]
