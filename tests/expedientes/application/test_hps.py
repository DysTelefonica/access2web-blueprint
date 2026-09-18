"""Strict TDD for EXP-CAP-051 HPS adapter (H01).

Dataclass fakes inline, no MagicMock. Mirrors the A04 pattern: per-
scenario tests derived from the spec's 3 escenarios (camino feliz,
validación, concurrencia/fallo).
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
    """Records every submit/query; ``exc`` makes it raise."""

    submitted: list[HpsSubmission] = field(default_factory=list)
    queried: list[UUID] = field(default_factory=list)
    record_seq: list[HpsRecord] = field(default_factory=list)
    exc: BaseException | None = None

    async def submit(self, submission: HpsSubmission, credential: str) -> HpsRecord:
        if self.exc is not None:
            raise self.exc
        self.submitted.append(submission)
        rec = HpsRecord(
            reference=submission.expediente_id, status="submitted", payload=submission.payload
        )
        self.record_seq.append(rec)
        return rec

    async def query(self, reference: UUID, credential: str) -> HpsRecord:
        if self.exc is not None:
            raise self.exc
        self.queried.append(reference)
        return HpsRecord(reference=reference, status="queried", payload={"ref": str(reference)})


@dataclass
class _Idempotency:
    store: dict[UUID, IdempotencyRecord] = field(default_factory=dict)

    async def get(self, key: UUID) -> IdempotencyRecord | None:
        return self.store.get(key)

    async def record(self, record: IdempotencyRecord) -> None:
        self.store[record.key] = record


def _service(
    *, hps: _FakeHps | None = None, audit: _Audit | None = None, idem: _Idempotency | None = None
) -> tuple[HpsService, _Audit, _FakeHps, _Idempotency]:
    audit = audit or _Audit()
    hps = hps or _FakeHps()
    idem = idem or _Idempotency()
    return HpsService(hps=hps, audit_log=audit, idempotency=idem), audit, hps, idem


def _submit_cmd(actor_id: UUID | None = None, *, key: UUID | None = None) -> HpsSubmitCommand:
    return HpsSubmitCommand(
        actor_id=actor_id,
        idempotency_key=key or uuid4(),
        expediente_id=uuid4(),
        payload={"k": "v"},
        credential="cred",
    )


async def test_submit_emits_one_event_and_persists_idempotency_record() -> None:
    actor, key = uuid4(), uuid4()
    svc, audit, hps, idem = _service()
    result = await svc.submit(_submit_cmd(actor_id=actor, key=key))

    assert isinstance(result, HpsResult) and result.deduped is False
    assert len(hps.submitted) == 1
    assert key in idem.store and idem.store[key].result.status == "submitted"
    assert audit.events[0].event_type == "hps.submit.recorded"  # type: ignore[attr-defined]


async def test_submit_with_repeated_key_dedupes_and_skips_adapter() -> None:
    actor, key = uuid4(), uuid4()
    svc, audit, hps, idem = _service()
    first = await svc.submit(_submit_cmd(actor_id=actor, key=key))
    second = await svc.submit(_submit_cmd(actor_id=actor, key=key))

    assert len(hps.submitted) == 1
    assert second.deduped is True and second.reference == first.reference
    assert audit.events[1].payload["deduped"] is True  # type: ignore[attr-defined]


async def test_submit_different_payload_for_same_key_bypasses_dedupe() -> None:
    actor, key = uuid4(), uuid4()
    svc, audit, hps, _ = _service()
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
        await svc.submit(_submit_cmd(actor_id=None))
    assert audit.events == [] and hps.submitted == []


async def test_dependency_failure_wraps_underlying_exception() -> None:
    actor = uuid4()
    hps = _FakeHps(exc=RuntimeError("hps down"))
    svc = HpsService(hps=hps, audit_log=_Audit(), idempotency=_Idempotency())
    with pytest.raises(HpsDependencyError) as excinfo:
        await svc.submit(_submit_cmd(actor_id=actor))
    assert isinstance(excinfo.value.__cause__, RuntimeError)


async def test_query_emits_one_event_and_returns_record() -> None:
    actor, ref = uuid4(), uuid4()
    svc, audit, hps, _ = _service()
    result = await svc.query(HpsQueryCommand(actor_id=actor, reference=ref, credential="c"))

    assert result.reference == ref and result.status == "queried"
    assert hps.queried == [ref]
    assert audit.events[0].event_type == "hps.query.recorded"  # type: ignore[attr-defined]


async def test_two_services_over_same_audit_emit_distinct_event_ids() -> None:
    audit = _Audit()
    svc_a = HpsService(hps=_FakeHps(), audit_log=audit, idempotency=_Idempotency())
    svc_b = HpsService(hps=_FakeHps(), audit_log=audit, idempotency=_Idempotency())
    actor = uuid4()

    await svc_a.query(HpsQueryCommand(actor_id=actor, reference=uuid4(), credential="c"))
    await svc_b.query(HpsQueryCommand(actor_id=actor, reference=uuid4(), credential="c"))

    assert len(audit.events) == 2
    assert audit.events[0].id != audit.events[1].id  # type: ignore[attr-defined]
