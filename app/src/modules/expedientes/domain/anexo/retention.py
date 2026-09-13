"""Anexo retention policy (CAP-010, D28).

D28 (APROBADO): "Retención por niveles configurable con archivo en
object storage". El módulo entrega:

- ``RetentionLevel`` enum (operational / legal / historical).
- ``RetentionPolicy`` value object con nivel + duración opcional.
- Cálculo determinista de ``expires_at`` y ``is_expired``.
- Verificación ``enforce`` que produce ``RetentionExpired`` con
  contexto para los logs de auditoría.

Los niveles concretos y las duraciones exactas son decisión de
producto; el módulo implementa la mecánica.

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from enum import StrEnum


class RetentionLevel(StrEnum):
    """Three retention buckets, named for the operational use case.

    - OPERATIONAL: in-flight documentation, alive only while the
      expediente is active.
    - LEGAL: regulatory retention. Default Spanish public-sector
      retention under LCSP is 10 years.
    - HISTORICAL: indefinite archive. The reference lives forever
      (or until manual deletion).

    StrEnum so the level survives JSON serialisation (audit logs,
    API responses) and round-trips through environment variables.
    """

    OPERATIONAL = "operational"
    LEGAL = "legal"
    HISTORICAL = "historical"


class RetentionExpired(Exception):
    """Raised when ``enforce`` is called on an expired anexo.

    Carries the level and the operation context so the audit log
    can explain why the operation was rejected.
    """

    def __init__(self, level: RetentionLevel, context: str) -> None:
        self.level = level
        self.context = context
        super().__init__(f"anexo retention expired (level={level.value}) in {context!r}")


@dataclass(frozen=True)
class RetentionPolicy:
    """Retention rule for an Anexo.

    ``duration`` is the time the anexo lives before expiry. It is
    ``None`` for ``HISTORICAL`` retention (kept indefinitely).
    """

    level: RetentionLevel
    duration: timedelta | None

    def __post_init__(self) -> None:
        if self.level is RetentionLevel.HISTORICAL and self.duration is not None:
            raise ValueError("HISTORICAL retention has indefinite duration; duration must be None")
        if self.level is not RetentionLevel.HISTORICAL and self.duration is None:
            raise ValueError(f"{self.level.value} retention requires a duration")
        if self.duration is not None and self.duration <= timedelta(0):
            raise ValueError(f"duration must be positive, got {self.duration}")

    @classmethod
    def legal_default(cls) -> RetentionPolicy:
        """Spanish public-sector legal retention: 10 years.

        10 years matches Ley 9/2017 (LCSP) artículo 159. The exact
        figure is a configuration concern; this default exists for
        tests and for the early adopters who don't override it.
        """
        return cls(level=RetentionLevel.LEGAL, duration=timedelta(days=10 * 365))

    @classmethod
    def operational_default(cls) -> RetentionPolicy:
        """Operational retention: 1 year — long enough for an active
        expediente to close, short enough to keep storage tidy."""
        return cls(level=RetentionLevel.OPERATIONAL, duration=timedelta(days=365))

    @classmethod
    def historical(cls) -> RetentionPolicy:
        """Historical retention: indefinite archive."""
        return cls(level=RetentionLevel.HISTORICAL, duration=None)

    def expires_at(self, created_at: datetime) -> datetime | None:
        """Return the wall-clock instant at which this anexo expires.

        None for HISTORICAL retention (kept indefinitely).
        """
        self._require_aware(created_at)
        if self.duration is None:
            return None
        return created_at + self.duration

    def is_expired(self, *, at: datetime, created_at: datetime) -> bool:
        """Return True iff the anexo has expired as of ``at``.

        ``at`` is the wall-clock instant being asked; ``created_at``
        is when the anexo entered the system. The comparison is
        inclusive of the exact expiry instant.
        """
        expiry = self.expires_at(created_at)
        if expiry is None:
            return False
        return at >= expiry

    def enforce(self, *, at: datetime, created_at: datetime, context: str) -> None:
        """Audit-friendly entry point: raise ``RetentionExpired`` if the
        anexo is past its retention as of ``at``.

        HISTORICAL retention never raises.
        """
        if self.is_expired(at=at, created_at=created_at):
            raise RetentionExpired(level=self.level, context=context)

    @staticmethod
    def _require_aware(dt: datetime) -> None:
        """Naive datetimes have no timezone; assume UTC and fail loudly.

        A naive expiry would depend on the runtime's local timezone
        and produce non-deterministic results across servers. The
        domain layer is strict on this.
        """
        if dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None:
            raise ValueError(
                f"datetime must be timezone-aware (got naive {dt.isoformat()!r}); "
                "use datetime.now(tz=UTC) or datetime.fromisoformat(..., tz=...)"
            )
