"""Strict TDD — ExpedientesConfig typed configuration (F03, issue #224).

The container reads its environment from a typed config object. Production
validation rejects empty values for any field that the runtime needs at
startup (D-EXP-3: deniegue sin stub permisivo en producción).
"""

from __future__ import annotations

import pytest

from app.src.modules.expedientes.di.config import (
    AppEnv,
    ConfigValidationError,
    ExpedientesConfig,
)


def test_default_config_is_development() -> None:
    """No env var, no argument -> development mode."""
    config = ExpedientesConfig()
    assert config.app_env is AppEnv.DEVELOPMENT


def test_db_url_defaults_to_none_in_development() -> None:
    """Dev runs without a real database; the field is optional."""
    config = ExpedientesConfig(app_env=AppEnv.DEVELOPMENT)
    assert config.db_url is None


def test_production_without_db_url_raises() -> None:
    """Production without a database URL is a misconfiguration."""
    with pytest.raises(ConfigValidationError) as exc:
        ExpedientesConfig(app_env=AppEnv.PRODUCTION, db_url=None)
    assert "db_url" in str(exc.value)


def test_production_with_empty_db_url_raises() -> None:
    """An empty string counts as missing."""
    with pytest.raises(ConfigValidationError):
        ExpedientesConfig(app_env=AppEnv.PRODUCTION, db_url="")


def test_production_with_db_url_succeeds() -> None:
    config = ExpedientesConfig(
        app_env=AppEnv.PRODUCTION, db_url="postgresql+asyncpg://user:pass@host/db"
    )
    assert config.db_url.startswith("postgresql")


def test_app_env_accepts_lowercase_string() -> None:
    """Allow ``"production"`` from environment variables (case-insensitive)."""
    config = ExpedientesConfig(
        app_env="production",  # type: ignore[arg-type]
        db_url="postgresql+asyncpg://user:pass@host/db",
    )
    assert config.app_env is AppEnv.PRODUCTION


def test_app_env_rejects_unknown_value() -> None:
    with pytest.raises(ConfigValidationError):
        ExpedientesConfig(app_env="staging")  # type: ignore[arg-type]
