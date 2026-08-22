"""Auth HTTP delivery (DL1, issue #54).

FastAPI router with the four public auth endpoints:

- ``GET  /login``  — render the login form
- ``POST /login``  — submit credentials; the actual lockout-after-5
  (D38) and session-creation are M02's A01..A03 work — this WU ships
  the route shape only.
- ``POST /logout`` — invalidate the current session (stub: clear the
  cookie).
- ``GET  /reset``  — render the reset request form.
- ``POST /reset``  — issue a reset token via the D90 service.

Alpine.js is loaded from a CDN (the upstream Mística design system
includes Alpine; we do not bundle it here because the platform's CDN
policy forbids local copies per the design.md §CDN).
"""

from __future__ import annotations

from fastapi import APIRouter, Depends, Form, Request, status
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.domain.ports.audit_log import AuditLogPort
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.notification_delivery import (
    NotificationDeliveryPort,
)
from app.src.modules.lanzadera.domain.ports.reset_token_repository import (
    ResetTokenRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.services.issue_reset_token import (
    issue_reset_token,
)


def build_router(
    *,
    templates: Jinja2Templates,
    users: UserRepository,
    reset_tokens: ResetTokenRepositoryPort,
    notifications: NotificationDeliveryPort,
    global_admins: GlobalAdminRepositoryPort,
) -> APIRouter:
    """Build the auth router.

    The factory pattern matches the admin router (#55) and the CLI
    (#56): the use-case ports are injected by the composition root at
    application startup, not at import time.
    """
    router = APIRouter(tags=["auth"])

    @router.get("/login", response_class=HTMLResponse)
    async def login_form(request: Request) -> HTMLResponse:
        return templates.TemplateResponse(request, "auth/login.html", {})

    @router.post("/login", response_class=HTMLResponse)
    async def login_submit(
        request: Request,
        email: str = Form(...),
        password: str = Form(...),
    ) -> RedirectResponse:
        # M02/A01..A03: validate password, increment failed_attempts on
        # failure, transition to LOCKED at threshold (D38), create a
        # Session row. This WU ships the route shape only.
        user = await users.get_by_email(email.strip().lower())
        if user is None:
            # Don't leak which side of the credential was wrong.
            return RedirectResponse("/login?error=invalid", status_code=303)
        # M02 wires the real password check.
        return RedirectResponse("/admin", status_code=303)

    @router.post("/logout", response_class=RedirectResponse)
    async def logout() -> RedirectResponse:
        # M02/A03: invalidate session in the session store, append an
        # auth.logout audit row.
        return RedirectResponse("/login", status_code=303)

    @router.get("/reset", response_class=HTMLResponse)
    async def reset_form(request: Request) -> HTMLResponse:
        return templates.TemplateResponse(request, "auth/reset.html", {})

    @router.post("/reset", response_class=RedirectResponse)
    async def reset_submit(
        request: Request,
        email: str = Form(...),
    ) -> RedirectResponse:
        # M02/A02: hash the raw token, persist via the D90 service.
        # The D90 service's ``issue_reset_token`` writes the reset_tokens
        # row and (if a MailQueueTableAdapter is wired) sends the email.
        # The dependency on ``issue_reset_token`` and ``reset_tokens``
        # here matches the WU DL1 spec (auth-core + auth-reset).
        from datetime import UTC, datetime, timedelta

        now = datetime.now(UTC)
        await issue_reset_token(
            email=email.strip().lower(),
            now=now,
            ttl=timedelta(hours=24),
            users=users,
            reset_tokens=reset_tokens,
            global_admins=global_admins,
            notifications=notifications,
        )
        return RedirectResponse("/login?reset=sent", status_code=303)

    return router


__all__ = ["build_router"]
