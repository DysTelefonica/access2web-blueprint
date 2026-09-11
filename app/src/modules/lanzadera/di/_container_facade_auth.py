# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W65
"""Auth use case methods + port accessors for LanzaderaContainer (W65 #596).

W65 (#596): cleave LanzaderaContainer by port group.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from uuid import UUID


class LanzaderaFacadeAuth:
    """Auth use cases and read-only auth port accessors.

    ``LanzaderaContainer`` creates one instance per container; callers
    access it via ``container._facade_auth`` (private) or implicitly through
    the ``__getattr__`` delegation on ``LanzaderaContainer``.
    """

    def __init__(
        self,
        use_cases: dict[str, Any],
        clock: Any,
        session_repo: Any,
        jwt_signer: Any,
        password_hasher: Any,
    ) -> None:
        self._use_cases = use_cases
        self._clock = clock
        self._session_repo = session_repo
        self._jwt_signer = jwt_signer
        self._password_hasher = password_hasher

    # -- auth use cases -------------------------------------------------------

    async def login(
        self,
        email: str,
        password: str,
        *,
        actor_id: UUID | None = None,
    ) -> Any:
        return await self._use_cases["login"](
            email=email,
            password=password,
            actor_id=actor_id,
        )

    async def logout(
        self,
        session_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["logout"](
            session_id=session_id,
            actor_id=actor_id,
        )

    async def set_password(
        self,
        email: str,
        *,
        new_password: str,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["set_password"](
            email=email,
            new_password=new_password,
            actor_id=actor_id,
            now=self._clock(),
        )

    # -- auth port accessors -------------------------------------------------

    @property
    def sessions(self) -> Any:
        return self._session_repo

    @property
    def jwt_signer(self) -> Any:
        return self._jwt_signer

    @property
    def password_hasher(self) -> Any:
        return self._password_hasher
