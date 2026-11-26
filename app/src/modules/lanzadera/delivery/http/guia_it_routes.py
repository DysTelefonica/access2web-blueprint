"""Admin-only IT handover guide for the Citrix and Linux rehearsal.

The page exposes internal paths and infrastructure flows, so it uses the
same ``require_global_admin`` gate as the other admin pages and is never
cached by intermediaries.
"""

from __future__ import annotations

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.delivery.http import admin as _admin
from app.src.modules.lanzadera.delivery.http.guia_it_content import GUIDE_CONTEXT


def register_guia_it_routes(router: APIRouter, *, templates: Jinja2Templates) -> None:
    """Attach ``GET /guia-it-citrix-linux`` to the admin router."""

    @router.get("/guia-it-citrix-linux", response_class=HTMLResponse)
    async def guia_it(request: Request) -> HTMLResponse:
        await _admin.require_global_admin(request)
        response = templates.TemplateResponse(request, "admin/guia_it.html", GUIDE_CONTEXT)
        response.headers["Cache-Control"] = "no-store"
        return response


__all__ = ["register_guia_it_routes"]
