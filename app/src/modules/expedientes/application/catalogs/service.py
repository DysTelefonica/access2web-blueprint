"""Catalogs service for EXP-CAP-015..019."""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID

from app.src.modules.expedientes.application.catalogs._evidence import (
    cpv_change,
    cpv_event,
    query_event,
)
from app.src.modules.expedientes.application.catalogs.command import (
    CatalogQuery,
    CatalogsAuthorizationError,
    CatalogsError,
    CatalogsValidationError,
    CpvDecision,
    CpvQuery,
)
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.catalog_repository import CatalogRepositoryPort

CAPACIDAD_BY_TYPE: dict[str, str] = {
    "comerciales": "EXP-CAP-015",
    "cpv": "EXP-CAP-016",
    "ejercitos": "EXP-CAP-017",
    "suministradores": "EXP-CAP-018",
    "lugares_ejecucion": "EXP-CAP-019",
    "pecal": "EXP-CAP-020",
    "racs": "EXP-CAP-021",
    "grados": "EXP-CAP-022",
    "organos": "EXP-CAP-023",
    "oficinas_programa": "EXP-CAP-024",
}


class CatalogsService:
    """Read-mostly catalogs use case with the CPV decision gate.

    The service is a thin layer over ``CatalogRepositoryPort``: it
    enforces permission, performs the optional text filter, and — for
    CPV — produces a typed decision object with the audit event the
    delivery layer can render.
    """

    def __init__(
        self,
        *,
        catalog_type: str,
        permission_required: str,
        catalog_repo: CatalogRepositoryPort,
        audit_log: AuditLogPort,
        permissions: set[str] | None = None,
    ) -> None:
        if catalog_type not in CAPACIDAD_BY_TYPE:
            raise ValueError(
                f"unknown catalog_type {catalog_type!r}; "
                f"expected one of {sorted(CAPACIDAD_BY_TYPE)}"
            )
        self._catalog_type = catalog_type
        self._permission_required = permission_required
        self._catalog_repo = catalog_repo
        self._audit_log = audit_log
        self._permissions: set[str] = set(permissions or ())

    def grant(self, permission: str) -> None:
        """Test/admin hook to inject a permission."""
        self._permissions.add(permission)

    async def list(self, query: CatalogQuery) -> list[object]:
        self._check_actor(query.actor_id)
        try:
            if query.text and query.fecha is not None:
                rows = await self._catalog_repo.vigentes_search(
                    query.text, query.fecha, limit=query.limit
                )
            elif query.text:
                rows = await self._catalog_repo.search(query.text, limit=query.limit)
            elif query.fecha is not None:
                rows = await self._catalog_repo.list_vigentes_en(query.fecha)
            else:
                rows = await self._catalog_repo.list_all()
        except Exception as exc:
            raise CatalogsError(f"catalogs failed for {self._catalog_type!r}: {exc}") from exc
        await self._audit_log.append(
            query_event(query, datetime.now(UTC), capacidad=self._capacidad())
        )
        if query.limit and query.limit > 0:
            return list(rows)[: query.limit]
        return list(rows)

    async def get(self, *, catalog_id: int, actor_id: UUID) -> object:
        self._check_actor(actor_id)
        try:
            row = await self._catalog_repo.get_by_id(catalog_id)
        except Exception as exc:
            raise CatalogsError(f"catalogs failed for {self._catalog_type!r}: {exc}") from exc
        if row is None:
            raise CatalogsValidationError(
                f"catalog entry {catalog_id!r} not found in {self._catalog_type!r}"
            )
        return row

    async def cpv_decision(self, query: CpvQuery) -> CpvDecision:
        if self._catalog_type != "cpv":
            raise CatalogsValidationError(
                f"cpv_decision only valid for cpv catalog; got {self._catalog_type!r}"
            )
        self._check_actor(query.actor_id)
        digits = query.raw_code.digits
        row: object | None = None
        try:
            row = await self._catalog_repo.get_by_id(int(digits))
        except Exception:
            row = None
        if row is None:
            try:
                row = await self._catalog_repo.get_by_id(digits)  # type: ignore[arg-type]
            except Exception as exc:
                raise CatalogsError(f"catalogs failed for cpv: {exc}") from exc
        decision = CpvDecision(
            raw_code=query.raw_code,
            accepted=row is not None,
            decision="aceptado" if row is not None else "rechazado",
            row=row,
            reason=None if row is not None else f"CPV {digits!r} no aparece en el catálogo maestro",
        )
        at = datetime.now(UTC)
        await self._audit_log.append(cpv_event(query, decision, at))
        await self._audit_log.record_change(cpv_change(query, decision, at))
        return decision

    def _check_actor(self, actor_id: UUID) -> None:
        if actor_id is None:
            raise CatalogsAuthorizationError("actor_id is required (deny-by-default)")
        if self._permission_required not in self._permissions:
            raise CatalogsAuthorizationError(
                f"actor lacks permission {self._permission_required!r} for "
                f"catalog {self._catalog_type!r}"
            )

    def _capacidad(self) -> str:
        return CAPACIDAD_BY_TYPE[self._catalog_type]
