"""FastAPI application composition root.

Phase 0 ships only a stub with a single `get_current_admin` dependency
placeholder. Real routes, handlers, and middleware arrive with the delivery
phase (PR 5 — Phase 5).

Architectural note: this module is the **driving adapter** boundary. The
hexagonal `di/container.py` (Phase 4) will wire ports to adapters; the
delivery routes in `app/src/modules/lanzadera/delivery/http/app.py`
(Phase 5) will import from this module. Phase 0 keeps the wiring minimal so
`uvicorn app.src.main:app` can boot without half-implemented code.
"""

from __future__ import annotations

from typing import Annotated

from fastapi import Depends, FastAPI, status

# ---------------------------------------------------------------------------
# Application factory
# ---------------------------------------------------------------------------

app = FastAPI(
    title="Platform — Lanzadera MVP",
    version="0.0.0+phase0",
    description=(
        "Hexagonal Python platform for the Lanzadera admin module. "
        "Phase 0 wires quality gates only; domain, ports, application, "
        "adapters, delivery and CLI arrive in PR 2..PR 8."
    ),
)


# ---------------------------------------------------------------------------
# Phase 0 placeholder dependency — replaced by a real `is_global_admin` gate
# once `ports/global_admin_repository.py` exists (Phase 1, PR 4).
# ---------------------------------------------------------------------------


async def get_current_admin() -> dict[str, str]:
    """Stub: returns a placeholder admin payload.

    Replaced in PR 5 by a real `Annotated[UUID, Depends(get_current_admin)]`
    extraction that decodes the session cookie and checks
    `GlobalAdminRepositoryPort.is_global_admin`.
    """
    return {"email": "[email protected]", "user_id": "00000000-0000-0000-0000-000000000000"}


CurrentAdmin = Annotated[dict[str, str], Depends(get_current_admin)]


# ---------------------------------------------------------------------------
# Healthcheck — used by docker-compose health probes and CI smoke tests.
# ---------------------------------------------------------------------------


@app.get("/health", status_code=status.HTTP_200_OK, tags=["health"])
async def healthcheck() -> dict[str, str]:
    """Liveness probe — does NOT touch the database.

    The `health` endpoint is intentionally side-effect free so the orchestrator
    can hit it without booting Postgres. The `ready` endpoint (Phase 5) will
    validate database + cache + secret-manager connectivity.
    """
    return {"status": "ok", "phase": "0", "service": "app-lanzadera-mvp"}
