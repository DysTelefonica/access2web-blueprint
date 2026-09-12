"""Strict TDD — Anexo size limit policy (CAP-010, D-EXP-3).

CAP-010 dice "dentro del límite acordado" — el límite concreto
queda como decisión de producto (issue #253). El módulo entrega
una política configurable con un valor por defecto conservador y un
mecanismo de validación que produce un error diagnosticable.

D-EXP-3 (deny-by-default) implica que un anexo fuera del límite
debe rechazarse ruidosamente, no truncarse ni aceptarse en silencio.
"""

from __future__ import annotations

import pytest

from app.src.modules.expedientes.domain.anexo.size_limit import (
    DEFAULT_MAX_SIZE_BYTES,
    SizeLimitExceeded,
    SizeLimitPolicy,
)


def test_default_max_size_is_25_mib() -> None:
    """The spec leaves 25 MB as an open question; until producto
    decides otherwise, 25 MiB is a safe upper bound that fits most
    office documents and keeps storage costs predictable.

    25 MiB (= 26_214_400 bytes) is 25 * 1024 * 1024 — the binary
    interpretation of "25 MB" common in Spanish IT shops. Once the
    maintainer picks the production figure, the default can move.
    """
    assert DEFAULT_MAX_SIZE_BYTES == 25 * 1024 * 1024


def test_policy_accepts_size_within_limit() -> None:
    policy = SizeLimitPolicy(max_size_bytes=1024)
    # No exception raised.
    policy.ensure_within_limit(size_bytes=1023, context="test")


def test_policy_accepts_size_exactly_at_limit() -> None:
    """The boundary case is inclusive — exactly the limit is OK."""
    policy = SizeLimitPolicy(max_size_bytes=1024)
    policy.ensure_within_limit(size_bytes=1024, context="test")


def test_policy_rejects_size_above_limit() -> None:
    policy = SizeLimitPolicy(max_size_bytes=1024)
    with pytest.raises(SizeLimitExceeded) as exc:
        policy.ensure_within_limit(size_bytes=1025, context="test")
    assert exc.value.size_bytes == 1025
    assert exc.value.limit_bytes == 1024


def test_policy_rejects_size_above_limit_with_strict_greater_than() -> None:
    """The check is strictly ``>`` (not ``>=``): exactly-at-limit is
    accepted. This matches the inclusive boundary test."""
    policy = SizeLimitPolicy(max_size_bytes=1024)
    # 1024 itself is OK (boundary).
    policy.ensure_within_limit(size_bytes=1024, context="test")
    # 1025 raises.
    with pytest.raises(SizeLimitExceeded):
        policy.ensure_within_limit(size_bytes=1025, context="test")


def test_size_limit_error_carries_context() -> None:
    """The exception carries the operation context for audit logs."""
    policy = SizeLimitPolicy(max_size_bytes=10)
    with pytest.raises(SizeLimitExceeded) as exc:
        policy.ensure_within_limit(size_bytes=11, context="upload:doc-42")
    assert exc.value.context == "upload:doc-42"


def test_size_limit_error_message_includes_human_readable_limit() -> None:
    """The error message must include the limit so the operator can
    see at a glance why the upload was rejected."""
    policy = SizeLimitPolicy(max_size_bytes=1024 * 1024)  # 1 MiB
    with pytest.raises(SizeLimitExceeded) as exc:
        policy.ensure_within_limit(size_bytes=2 * 1024 * 1024, context="x")
    message = str(exc.value)
    assert "1 MiB" in message or "1.0 MiB" in message or "1048576" in message


def test_policy_rejects_negative_size() -> None:
    """A negative size is malformed input, not an exceeded limit."""
    policy = SizeLimitPolicy(max_size_bytes=1024)
    with pytest.raises(ValueError):
        policy.ensure_within_limit(size_bytes=-1, context="test")


def test_policy_zero_size_is_invalid_even_when_limit_is_zero() -> None:
    """A zero-size anexo is malformed input, not 'within the limit'."""
    policy = SizeLimitPolicy(max_size_bytes=1024)
    with pytest.raises(ValueError):
        policy.ensure_within_limit(size_bytes=0, context="test")


def test_custom_limit_from_maintainer() -> None:
    """Product decides a different limit (e.g. 100 MB for video
    evidence); the policy accepts any non-negative value."""
    policy = SizeLimitPolicy(max_size_bytes=100 * 1024 * 1024)
    policy.ensure_within_limit(size_bytes=99 * 1024 * 1024, context="video")
    with pytest.raises(SizeLimitExceeded):
        policy.ensure_within_limit(size_bytes=101 * 1024 * 1024, context="video")


def test_policy_rejects_negative_limit_at_construction() -> None:
    """A negative limit is a configuration bug; fail loudly at construction."""
    with pytest.raises(ValueError):
        SizeLimitPolicy(max_size_bytes=-1)


def test_format_mib() -> None:
    """The format helper rounds to one decimal for human display."""
    from app.src.modules.expedientes.domain.anexo.size_limit import _format_mib

    assert _format_mib(1024 * 1024) == "1.0 MiB"
    assert _format_mib(25 * 1024 * 1024) == "25.0 MiB"
