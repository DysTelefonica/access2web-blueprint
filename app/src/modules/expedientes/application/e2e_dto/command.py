"""Commands and outcomes for EXP-CAP-033 stable DTO."""

from dataclasses import dataclass
from datetime import date
from typing import Any

from app.src.modules.expedientes.domain.expediente import Expediente
from app.src.modules.expedientes.domain.expediente_estado import ExpedienteEstado
from app.src.modules.expedientes.domain.expediente_tipo import ExpedienteTipo

SUPPORTED_API_VERSION = "1.0"


class E2EDtoError(Exception):
    """Base error for the DTO exchange use case."""


class E2EDtoValidationError(E2EDtoError):
    """The DTO or the inbound payload is malformed or has the wrong version."""


class E2EDtoAuthorizationError(E2EDtoError):
    """The caller has no authenticated actor or lacks permission."""


@dataclass(frozen=True)
class ExpedienteDTO:
    """Stable, versioned DTO for the E2E channel (CAP-033).

    No technical IDs (``id``, ``version``, ``id_expediente_padre``) leak
    through the wire. Dates are ISO-8601 strings, optional fields are
    ``None``. The contract carries a version stamp so E2E partners can
    detect mismatched producers.
    """

    api_version: str
    tipo: str
    estado: str
    fecha_inicio_contrato: str | None
    fecha_fin_contrato: str | None

    def to_payload(self) -> dict[str, Any]:
        return {
            "apiVersion": self.api_version,
            "tipo": self.tipo,
            "estado": self.estado,
            "fechaInicioContrato": self.fecha_inicio_contrato,
            "fechaFinContrato": self.fecha_fin_contrato,
        }

    def to_aggregate(self) -> "Expediente":
        try:
            inicio = (
                date.fromisoformat(self.fecha_inicio_contrato)
                if self.fecha_inicio_contrato
                else None
            )
            fin = date.fromisoformat(self.fecha_fin_contrato) if self.fecha_fin_contrato else None
        except (TypeError, ValueError) as exc:
            raise E2EDtoValidationError(f"fecha ISO-8601 inválida en DTO: {exc}") from exc
        # DTOs do not carry technical IDs; reconstruct with a fresh
        # UUID and a synthetic parent for the LOTE/BASED types so the
        # aggregate invariant (padre required) is satisfied. The
        # delivery layer is responsible for remapping the technical
        # ID when the partner sends one back.
        from uuid import uuid4

        # AM/SUBM/...: no parent required. LOTE/BASED: synthetic parent
        # to keep the aggregate valid; the delivery layer remaps the
        # partner-supplied ID on top.
        padre = uuid4() if self.tipo in {"LOTE", "BASED"} else None
        try:
            return Expediente(
                id=uuid4(),
                tipo=ExpedienteTipo(self.tipo),
                estado=ExpedienteEstado(self.estado),
                version=1,
                id_expediente_padre=padre,
                fecha_inicio_contrato=inicio,
                fecha_fin_contrato=fin,
            )
        except ValueError as exc:
            raise E2EDtoError(f"DTO cannot be reconstructed: {exc}") from exc


def from_aggregate(
    expediente: Expediente, *, api_version: str = SUPPORTED_API_VERSION
) -> ExpedienteDTO:
    return ExpedienteDTO(
        api_version=api_version,
        tipo=expediente.tipo.value,
        estado=expediente.estado.value,
        fecha_inicio_contrato=(
            expediente.fecha_inicio_contrato.isoformat()
            if expediente.fecha_inicio_contrato
            else None
        ),
        fecha_fin_contrato=(
            expediente.fecha_fin_contrato.isoformat() if expediente.fecha_fin_contrato else None
        ),
    )


@dataclass(frozen=True)
class E2EDtoCommand:
    expediente: Expediente
    actor_id: Any
    api_version: str = SUPPORTED_API_VERSION


@dataclass(frozen=True)
class E2EDtoResult:
    dto: ExpedienteDTO
    reconstructed: Expediente
