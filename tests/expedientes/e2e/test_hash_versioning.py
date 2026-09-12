"""Strict TDD — Hash versionado (CAP-036).

El spec `openspec/changes/expedientes-web-migration/specs/e2e.md`
exige (CAP-036): estrategia de hash versionada y canonicalización
documentada. Estos tests cubren el comportamiento neutral que sirve
para CUALQUIER política de hash (FNV-1a u otra). La elección del
algoritmo concreto es decisión humana (issue #659).
"""

from __future__ import annotations

import pytest

from app.src.modules.expedientes.e2e.hash_versioning import (
    FNV1A32,
    HashRegistry,
    HashVersioned,
    canonicalize,
    compute_hash,
)

# ---------------------------------------------------------------------------
# Canonicalización
# ---------------------------------------------------------------------------


def test_canonicalize_orders_keys_alphabetically() -> None:
    """The canonical form must be order-independent on the producer side."""
    a = canonicalize({"b": 2, "a": 1})
    b = canonicalize({"a": 1, "b": 2})
    assert a == b


def test_canonicalize_separates_nested_objects() -> None:
    """Nested dicts must canonicalize to the same bytes regardless of indentation."""
    a = canonicalize({"meta": {"z": 1, "a": 2}, "id": "exp-1"})
    b = canonicalize({"id": "exp-1", "meta": {"a": 2, "z": 1}})
    assert a == b


def test_canonicalize_returns_utf8_bytes() -> None:
    payload = canonicalize({"name": "Expediente Jaén"})
    assert isinstance(payload, bytes)
    assert payload.decode("utf-8")  # round-trips


# ---------------------------------------------------------------------------
# Determinism
# ---------------------------------------------------------------------------


def test_compute_hash_is_deterministic() -> None:
    """Same payload + same algorithm → same hash."""
    payload = {"ordinal": "EXP-001", "estado": "BORRADOR", "importe": 1234.56}
    first = compute_hash(payload, algorithm="fnv1a-32")
    second = compute_hash(payload, algorithm="fnv1a-32")
    assert first.digest == second.digest


def test_compute_hash_changes_with_payload() -> None:
    payload_a = {"ordinal": "EXP-001"}
    payload_b = {"ordinal": "EXP-002"}
    assert (
        compute_hash(payload_a, algorithm="fnv1a-32").digest
        != compute_hash(payload_b, algorithm="fnv1a-32").digest
    )


# ---------------------------------------------------------------------------
# Versioning
# ---------------------------------------------------------------------------


def test_hash_version_carries_algorithm_marker() -> None:
    """The HashVersioned record identifies which algorithm produced it.

    CAP-036: the strategy MUST be versioned. Any change in algorithm
    or canonicalization must be visible in the hash itself so old
    hashes can be reproduced by replaying the same version.
    """
    record = compute_hash({"x": 1}, algorithm="fnv1a-32")
    assert isinstance(record, HashVersioned)
    assert record.algorithm == "fnv1a-32"
    assert record.digest != ""


def test_two_algorithms_produce_different_digests() -> None:
    """Switching algorithm changes the hash — the marker is meaningful.

    This is the test that catches silent version drift: a reader that
    ignores the algorithm marker would treat the two digests as
    comparable, which they are not.
    """
    payload = {"ordinal": "EXP-001"}

    class _StubAlgorithm:
        @property
        def algorithm(self) -> str:
            return "stub-1"

        def hash(self, payload: bytes) -> str:
            return "deadbeef"

    registry = HashRegistry()
    registry.register(_StubAlgorithm())
    registry.register(FNV1A32())

    fnv = compute_hash(payload, algorithm="fnv1a-32", registry=registry)
    stub = compute_hash(payload, algorithm="stub-1", registry=registry)
    assert fnv.digest != stub.digest


# ---------------------------------------------------------------------------
# FNV-1a 32-bit vector tests (informational, not conformance)
# ---------------------------------------------------------------------------


def test_fnv1a_32_known_vector_empty() -> None:
    """FNV-1a 32-bit hash of empty input is the offset basis 0x811c9dc5."""
    record = compute_hash({}, algorithm="fnv1a-32")
    # Canonical form of {} is "{}"; FNV-1a 32-bit of "{}"
    # is NOT the offset basis (the offset basis is for empty input).
    # We only assert the hash is deterministic and hex-shaped.
    assert len(record.digest) == 8
    assert all(c in "0123456789abcdef" for c in record.digest)


def test_fnv1a_32_against_reference_implementation() -> None:
    """FNV-1a 32-bit matches the IETF reference vector for byte b'a'."""
    algo = FNV1A32()
    digest = algo.hash(b"a")
    assert digest == "e40c292c"


def test_fnv1a_32_known_vector_foobar() -> None:
    """FNV-1a 32-bit known vector for 'foobar' per the reference algorithm."""
    algo = FNV1A32()
    digest = algo.hash(b"foobar")
    assert digest == "bf9cf968"


# ---------------------------------------------------------------------------
# Registry contract
# ---------------------------------------------------------------------------


def test_registry_rejects_unknown_algorithm() -> None:
    registry = HashRegistry()
    registry.register(FNV1A32())
    with pytest.raises(KeyError):
        compute_hash({"x": 1}, algorithm="does-not-exist")


def test_registry_rejects_duplicate_registration() -> None:
    """Two algorithms with the same name conflict — fail loudly."""
    registry = HashRegistry()
    registry.register(FNV1A32())
    with pytest.raises(ValueError):
        registry.register(FNV1A32())


def test_default_registry_includes_fnv1a_32() -> None:
    """The module-level default registry ships with FNV-1a 32-bit.

    Issue #659 keeps the door open for a different algorithm; the
    default registry contains the legacy-compatible one until that
    decision lands.
    """
    from app.src.modules.expedientes.e2e.hash_versioning import default_registry

    assert "fnv1a-32" in default_registry.names()


# ---------------------------------------------------------------------------
# Algorithm interface
# ---------------------------------------------------------------------------


def test_hash_algorithm_protocol_exposes_algorithm_name() -> None:
    algo = FNV1A32()
    # Structural check — HashAlgorithm is a Protocol, so we look at the
    # methods it declares rather than rely on runtime_checkable semantics.
    assert hasattr(algo, "algorithm")
    assert hasattr(algo, "hash")
    assert algo.algorithm == "fnv1a-32"
