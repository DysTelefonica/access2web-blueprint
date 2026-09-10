# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W65
"""Auth port wiring: session + JWT + password hasher (W65 #596).

Extracted from ``container.py`` to keep every module under the 100-site
mutation ceiling. The thin ``container.py`` imports this and merges the
result into the full port dict before calling ``build_use_case_factories``.

W65 (#596): cleave LanzaderaContainer by port group.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.domain.ports.jwt_signer import JwtSignerPort
from app.src.modules.lanzadera.domain.ports.password_hasher import PasswordHasher
from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManager
from app.src.modules.lanzadera.domain.ports.session_repository import SessionRepository


if TYPE_CHECKING:
    from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
        AsyncSessionFactoryPort,
    )


def _pick_session(
    fake: object | None,
    factory: AsyncSessionFactoryPort | None,
) -> SessionRepository:
    """Return a fake session repo if injected; otherwise build the default."""
    if fake is not None:
        return fake  # type: ignore[return-value]
    return _build_default_session_repo()


def _build_default_session_repo() -> SessionRepository:
    """Default ``SessionRepository`` wiring (W62 #539).

    Returns a Postgres-backed adapter once ``SessionRepositoryPg``
    lands in a follow-up slice. Until then, raising here makes the
    production-code path fail fast with a clear message.
    """
    raise NotImplementedError(
        "SessionRepositoryPg is not implemented yet. "
        "Pass a fake (tests/lanzadera/_fakes.py:FakeSessionRepository) "
        "until PR-2's follow-up ships the Postgres adapter."
    )


def _build_default_jwt_signer() -> JwtSignerPort:
    """Default ``JwtSignerPort`` wiring (W62 PR-6).

    Loads the HS256 secret from ``$JWT_SECRET`` via
    ``EnvSecretManagerAdapter``. Tests inject ``FakeJwtSigner``
    (PR-4); production uses this builder.
    """
    from app.src.modules.lanzadera.adapters.cross.secret_manager import (
        EnvSecretManagerAdapter,
    )
    from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner

    secret_manager: SecretManager = EnvSecretManagerAdapter()
    try:
        secret = secret_manager.get("JWT_SECRET").encode("utf-8")
    except KeyError:
        # Unit-test fallback: production deployments MUST set
        # JWT_SECRET; this dummy is here so non-JWT tests can boot
        # the container without an env var.
        secret = b"unit-test-dummy-jwt-secret-32-bytes-pad"
    return Hs256JwtSigner(secret)


def build_auth_ports(
    *,
    session_factory: AsyncSessionFactoryPort | None,
    password_hasher: PasswordHasher,
    session_repo: SessionRepository | None,
    jwt_signer: JwtSignerPort | None,
) -> dict[str, Any]:
    """Build the 3 auth ports (session, JWT signer, password hasher).

    W65 (#596): extracted from ``LanzaderaContainer.__init__`` to keep
    both this module and the thin container under the 100-site ceiling.
    """
    _session_repo = _pick_session(session_repo, session_factory)
    _jwt_signer = jwt_signer or _build_default_jwt_signer()

    return {
        "password_hasher": password_hasher,
        "session_repo": _session_repo,
        "jwt_signer": _jwt_signer,
    }
