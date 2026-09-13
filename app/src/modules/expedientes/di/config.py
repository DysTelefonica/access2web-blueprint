"""Typed configuration for the Expedientes module (F03, issue #224).

D-EXP-3: "deny/no fake prod" + "readiness falla si se enlazan en producción".
The container's production guard reads its mode from this object.

Production validation rejects empty values for any field the runtime needs
at startup. Development and test modes are permissive — they accept
optional fields as ``None`` so work can proceed without a real database.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum


class AppEnv(StrEnum):
    """Runtime environment marker.

    Drives the deny-by-default guard in :class:`ExpedientesContainer`.
    StrEnum so it round-trips through JSON and environment variables.
    """

    DEVELOPMENT = "development"
    TEST = "test"
    PRODUCTION = "production"


class ConfigValidationError(ValueError):
    """Raised when the typed configuration is missing required values.

    Production-mode configuration must declare every field the runtime
    needs at startup; missing values are rejected eagerly so the process
    fails fast instead of crashing on first use.
    """


class ProductionFakeRefused(RuntimeError):
    """Raised when the container is asked to bind fakes in production.

    D-EXP-3 + issue #224 acceptance: deny/no fake prod. Fakes are
    deterministic test doubles; binding them in production would silently
    accept writes that have no durable backing. The container raises
    this error instead of starting.
    """


def _coerce_app_env(value: object) -> AppEnv:
    """Coerce a string-or-AppEnv into AppEnv, raising ConfigValidationError.

    String inputs are normalised to lowercase before lookup so environment
    variables (``APP_ENV=production``) round-trip without surprises.
    """
    if isinstance(value, AppEnv):
        return value
    try:
        return AppEnv(str(value).lower())
    except ValueError as exc:
        raise ConfigValidationError(
            f"unknown app_env: {value!r} (expected one of {[e.value for e in AppEnv]})"
        ) from exc


@dataclass(frozen=True)
class ExpedientesConfig:
    """Typed configuration for the Expedientes module.

    Defaults to development so first-time callers (tests, REPL) get a
    permissive configuration. Production deployments MUST set ``app_env``
    and ``db_url`` explicitly.
    """

    app_env: AppEnv | str = AppEnv.DEVELOPMENT
    db_url: str | None = None

    def __post_init__(self) -> None:
        coerced = _coerce_app_env(self.app_env)
        if coerced is not self.app_env:
            object.__setattr__(self, "app_env", coerced)

        if self.app_env is AppEnv.PRODUCTION and not self.db_url:
            raise ConfigValidationError(
                "production requires db_url (D-EXP-3: no stub permisivo en producción)"
            )
