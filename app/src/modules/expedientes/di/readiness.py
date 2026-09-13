"""``/readyz`` endpoint for the Expedientes module (F03, issue #224).

CAP-047: "Camino feliz: sistema arrancado y dependencias (Lanzadera,
Riesgos, NC) disponibles → contrato ``/readyz`` verde".

D-EXP-3: when production is reached without a working ``ReadinessPort``,
the endpoint MUST return 503. The endpoint does not know about the
production guard (the container refuses to build in that case) — the 503
is a defence-in-depth signal that the host can rely on regardless of how
the container was constructed.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from app.src.modules.expedientes.di.container import ExpedientesContainer
    from app.src.modules.expedientes.ports.readiness import ReadinessCheck, ReadinessResult


@dataclass(frozen=True)
class ReadinessResponse:
    """The wire shape of the ``/readyz`` response.

    Status 200 when every check passes; 503 otherwise. The body is JSON-
    serialisable for the FastAPI layer (F04+) to wrap directly.
    """

    status_code: int
    body: dict[str, object] = field(default_factory=dict)


def _check_to_dict(check: ReadinessCheck) -> dict[str, object]:
    return {"name": check.name, "ok": check.ok, "message": check.message}


def _result_to_body(result: ReadinessResult) -> dict[str, object]:
    return {
        "ready": result.ready,
        "checks": [_check_to_dict(c) for c in result.checks],
    }


async def readiness_response(container: ExpedientesContainer) -> ReadinessResponse:
    """Resolve the ``/readyz`` response from the container's readiness port.

    The container must hold a wired ``ReadinessPort``. When the check
    fails (or the port returns ``None`` because nothing was configured),
    the response is 503 with the failed check listed.
    """
    result = await container.readiness.check()
    if result is None:
        return ReadinessResponse(
            status_code=503,
            body={"ready": False, "checks": [], "reason": "no readiness check configured"},
        )
    body = _result_to_body(result)
    return ReadinessResponse(
        status_code=200 if result.ready else 503,
        body=body,
    )


__all__ = ["ReadinessResponse", "readiness_response"]
