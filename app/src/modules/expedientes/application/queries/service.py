"""Query services use case for EXP-CAP-025..029."""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.expedientes.application.queries._evidence import (
    bandeja_event,
    export_event,
    tareas_event,
)
from app.src.modules.expedientes.application.queries.command import (
    BandejaPage,
    BandejaQuery,
    ExportacionExcelCommand,
    ExportacionExcelJob,
    QueriesAuthorizationError,
    QueriesError,
    TareasCalculadasQuery,
    TareasResult,
)
from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.ports.audit_log import AuditLogPort
from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort

TAREA_BUCKETS: tuple[str, ...] = (
    "desconocido",
    "recepcion_hito",
    "adjudicado_sin_contrato",
    "tsol_sin_s4h",
    "oferta_prolongada",
    "otro",
)


class QueriesService:
    """Read-mostly query services for the expedientes lane.

    Covers CAP-025 (bandeja paginada), CAP-026 (búsqueda avanzada),
    CAP-027 (búsqueda técnica — reuses the bandeja path with a different
    permission), CAP-028 (exportación Excel) and CAP-029 (tareas
    calculadas).
    """

    def __init__(
        self,
        *,
        expediente_repo: ExpedienteRepositoryPort,
        audit_log: AuditLogPort,
        permissions: set[str] | None = None,
        expedientes_provider: Callable[[], list[Expediente]] | None = None,
    ) -> None:
        self._expediente_repo = expediente_repo
        self._audit_log = audit_log
        self._permissions: set[str] = set(permissions or ())
        self._expedientes_provider: Callable[[], list[Expediente]] = (
            expedientes_provider if expedientes_provider is not None else (lambda: [])
        )

    def grant(self, permission: str) -> None:
        """Test/admin hook to inject a permission."""
        self._permissions.add(permission)

    def bind_expedientes(self, provider: Callable[[], list[Expediente]]) -> None:
        """Bind the read model used by the worklist calculation."""
        self._expedientes_provider = provider

    async def bandeja(self, query: BandejaQuery) -> BandejaPage:
        self._check_actor(query.actor_id, "bandeja.read")
        try:
            rows, total = await self._expediente_repo.search(
                estado=query.estado,
                codigo=query.codigo,
                responsable_id=query.responsable_id,
                juridica_id=query.juridica_id,
                suministrador_id=query.suministrador_id,
                limit=query.limit,
                offset=query.offset,
            )
        except Exception as exc:
            raise QueriesError(f"queries failed for bandeja: {exc}") from exc
        await self._audit_log.append(bandeja_event(query, datetime.now(UTC), total=total))
        return BandejaPage(
            rows=list(rows),
            total=total,
            limit=query.limit,
            offset=query.offset,
        )

    async def busqueda_avanzada(self, query: BandejaQuery) -> BandejaPage:
        return await self.bandeja(query)

    async def busqueda_tecnica(self, query: BandejaQuery) -> BandejaPage:
        self._check_actor(query.actor_id, "bandeja.technical")
        return await self.bandeja(query)

    async def exportar_excel(self, cmd: ExportacionExcelCommand) -> ExportacionExcelJob:
        self._check_actor(cmd.actor_id, "export.excel")
        job = ExportacionExcelJob(
            job_id=uuid4(),
            filtros=dict(cmd.filtros),
            formato=cmd.formato,
            requested_at=datetime.now(UTC),
        )
        await self._audit_log.append(export_event(cmd, job))
        return job

    async def tareas_calculadas(self, query: TareasCalculadasQuery) -> TareasResult:
        self._check_actor(query.actor_id, "tareas.read")
        contadores: dict[str, int] = {bucket: 0 for bucket in TAREA_BUCKETS}
        for expediente in self._expedientes_provider():
            bucket = self._classify(expediente, query.now)
            contadores[bucket] += 1
        result = TareasResult(contadores=contadores, total=sum(contadores.values()))
        await self._audit_log.append(tareas_event(query, result, datetime.now(UTC)))
        return result

    def _check_actor(self, actor_id: UUID, permission: str) -> None:
        if actor_id is None:
            raise QueriesAuthorizationError("actor_id is required (deny-by-default)")
        if permission not in self._permissions:
            raise QueriesAuthorizationError(f"actor lacks permission {permission!r}")

    @staticmethod
    def _classify(expediente: Expediente, now: datetime) -> str:
        if expediente.estado is ExpedienteEstado.BORRADOR:
            return "desconocido"
        if expediente.estado is ExpedienteEstado.ADJUDICADO:
            return "adjudicado_sin_contrato"
        return "otro"
