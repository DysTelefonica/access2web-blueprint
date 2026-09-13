"""Strict TDD — Anexo retention policy (CAP-010, D28).

D28 (APROBADO): "Retención por niveles configurable con archivo en
object storage". El módulo entrega:

- `RetentionLevel` enum con tres niveles (operativo, legal, histórico).
- `RetentionPolicy` value object que codifica un nivel + duración.
- Cálculo determinista de expiry (created_at + duration).
- Verificación de "archivo" cuando expira.

Los niveles concretos y las duraciones exactas son decisión de
producto (D28 deja los detalles a la configuración). El módulo
implementa la mecánica; el catálogo de niveles es runtime config.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta

import pytest

from app.src.modules.expedientes.domain.anexo.retention import (
    RetentionExpired,
    RetentionLevel,
    RetentionPolicy,
)

# ---------------------------------------------------------------------------
# Retention levels
# ---------------------------------------------------------------------------


def test_operational_level_short_window() -> None:
    """Operational = in-flight documentation, retained only while
    the expediente is active."""
    level = RetentionLevel.OPERATIONAL
    assert level.value == "operational"


def test_legal_level_long_window() -> None:
    """Legal = regulatory retention (Spanish public contracts default
    to 10 years under LCSP). The exact duration is a config concern;
    the enum names the bucket."""
    level = RetentionLevel.LEGAL
    assert level.value == "legal"


def test_historical_level_indefinite() -> None:
    """Historical = kept indefinitely (archive). Expiry is ``None``."""
    level = RetentionLevel.HISTORICAL
    assert level.value == "historical"


# ---------------------------------------------------------------------------
# RetentionPolicy
# ---------------------------------------------------------------------------


def test_constructs_with_legal_level_and_ten_years() -> None:
    """Default legal retention under Spanish public-sector rules is
    10 years. This is the most common case for the current
    customer; the constructor exposes a sane default."""
    policy = RetentionPolicy.legal_default()
    assert policy.level is RetentionLevel.LEGAL
    assert policy.duration == timedelta(days=10 * 365)


def test_constructs_with_operational_level_and_one_year() -> None:
    """Operational retention defaults to 1 year."""
    policy = RetentionPolicy.operational_default()
    assert policy.level is RetentionLevel.OPERATIONAL
    assert policy.duration == timedelta(days=365)


def test_historical_level_has_no_duration() -> None:
    """Historical = kept indefinitely; duration is ``None``."""
    policy = RetentionPolicy.historical()
    assert policy.level is RetentionLevel.HISTORICAL
    assert policy.duration is None


def test_expires_at_with_duration() -> None:
    """Expire date is created_at + duration."""
    policy = RetentionPolicy(level=RetentionLevel.LEGAL, duration=timedelta(days=10))
    created = datetime(2024, 1, 1, tzinfo=UTC)
    assert policy.expires_at(created) == datetime(2024, 1, 11, tzinfo=UTC)


def test_expires_at_with_historical_is_none() -> None:
    """Historical retention does not expire."""
    policy = RetentionPolicy.historical()
    created = datetime(2024, 1, 1, tzinfo=UTC)
    assert policy.expires_at(created) is None


def test_expires_at_with_naive_datetime_raises() -> None:
    """The created_at argument must be timezone-aware; naive datetimes
    have undefined timezone and would produce a non-deterministic
    expiry."""
    policy = RetentionPolicy.legal_default()
    with pytest.raises(ValueError):
        policy.expires_at(datetime(2024, 1, 1))  # noqa: DTZ001 - intentionally naive


def test_is_expired_with_duration() -> None:
    """``is_expired(at)`` is True iff ``at >= expires_at``."""
    policy = RetentionPolicy(level=RetentionLevel.LEGAL, duration=timedelta(days=30))
    created = datetime(2024, 1, 1, tzinfo=UTC)
    # 1 day before expiry
    assert policy.is_expired(at=created + timedelta(days=29), created_at=created) is False
    # exactly at expiry
    assert policy.is_expired(at=created + timedelta(days=30), created_at=created) is True
    # 1 day after expiry
    assert policy.is_expired(at=created + timedelta(days=31), created_at=created) is True


def test_historical_level_never_expires() -> None:
    """``is_expired`` is always False for historical retention."""
    policy = RetentionPolicy.historical()
    created = datetime(2024, 1, 1, tzinfo=UTC)
    far_future = created + timedelta(days=365 * 100)
    assert policy.is_expired(at=far_future, created_at=created) is False


def test_enforce_raises_when_expired() -> None:
    """``enforce`` is the audit-friendly entry point: raises
    ``RetentionExpired`` with context for the operator."""
    policy = RetentionPolicy(level=RetentionLevel.LEGAL, duration=timedelta(days=10))
    created = datetime(2024, 1, 1, tzinfo=UTC)
    after = created + timedelta(days=11)
    with pytest.raises(RetentionExpired) as exc:
        policy.enforce(at=after, created_at=created, context="exp-42")
    assert exc.value.context == "exp-42"
    assert exc.value.level is RetentionLevel.LEGAL


def test_enforce_passes_when_not_expired() -> None:
    policy = RetentionPolicy.legal_default()
    created = datetime(2024, 1, 1, tzinfo=UTC)
    # 1 day later — within legal default of 10 years
    policy.enforce(at=created + timedelta(days=1), created_at=created, context="exp-42")


def test_enforce_passes_always_for_historical() -> None:
    policy = RetentionPolicy.historical()
    created = datetime(2024, 1, 1, tzinfo=UTC)
    # Far in the future.
    policy.enforce(
        at=created + timedelta(days=365 * 1000),
        created_at=created,
        context="archive",
    )


def test_constructing_operational_or_legal_without_duration_raises() -> None:
    """A policy that has a finite level but ``duration is None`` is
    a configuration bug — caught at construction."""
    with pytest.raises(ValueError):
        RetentionPolicy(level=RetentionLevel.LEGAL, duration=None)


def test_constructing_with_negative_duration_raises() -> None:
    with pytest.raises(ValueError):
        RetentionPolicy(level=RetentionLevel.OPERATIONAL, duration=timedelta(days=-1))
