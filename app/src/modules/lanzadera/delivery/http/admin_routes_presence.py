# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — admin HTTP routes for the real-time presence slice.
"""Admin HTTP route handlers for the W60 (#522) real-time presence slice.

Three ``/admin/presence/...`` endpoints:

- ``POST /admin/presence/heartbeat`` — record a heartbeat for the calling
  user. The X-User-ID header stands in for ``request.state.user_id``
  until the auth middleware ships with W62; the W60 contract pins the
  behaviour behind that header so the slice is exercisable end-to-end
  before the real auth wiring lands.
- ``GET /admin/presence`` — one-shot JSON snapshot of the currently
  connected set.
- ``GET /admin/presence/stream`` — Server-Sent Events emitter that polls
  the connected set every 5 seconds and re-broadcasts it as
  ``data: <json>\\n\\n``.

The router is mounted directly in ``app/src/main.py`` (W60 ships its
own ``/admin`` prefix), unlike the rest of the admin routes which
register through ``admin_routes.py``. The container is resolved via
``request.state.container`` (set by the FastAPI app's lifespan or by a
small middleware that mirrors ``app.state.container`` onto every
request) so the SSE emitter does not need a closure binding — the
stream is a long-lived per-request coroutine that would not survive the
outer closure's lifetime if the router were registered at startup.
"""

from __future__ import annotations

import asyncio
import json
from collections.abc import AsyncIterator
from uuid import UUID

from fastapi import APIRouter, HTTPException, Request, status
from fastapi.responses import StreamingResponse

from app.src.modules.lanzadera.domain.presence import ConnectedUser

router = APIRouter()

# W60 (#522): heartbeat interval for the SSE emitter. The spec calls for
# a 5-second poll cadence (admin console refresh feels live, the DB query
# stays cheap). If the heartbeat / disconnect TTL lands in a later wave,
# this constant survives — the emitter continues to drive its own clock
# regardless of how stale a row can become.
_SSE_POLL_INTERVAL_SECONDS: float = 5.0


def _serialise_user(user: ConnectedUser) -> dict[str, object]:
    """Render a ``ConnectedUser`` as the JSON shape the SSE emitter broadcasts.

    Keys are stable (``user_id`` / ``email`` / ``connected_at`` /
    ``last_seen``); values are JSON-safe. ``user_id`` is stringified
    because ``json.dumps`` cannot serialise a ``UUID`` natively and the
    client expects strings across the network.
    """
    return {
        "user_id": str(user.user_id),
        "email": user.email,
        "connected_at": user.connected_at.isoformat(),
        "last_seen": user.last_seen.isoformat(),
    }


def _container(request: Request) -> object:
    """Return the container attached to the request (or 503 if absent).

    FastAPI's ``app.state.container`` is set by the lifespan; the
    ``request.state.container`` it forwards to is populated by a small
    middleware that mirrors the app state onto every request. The
    presence routes never reach for ``app.state.container`` directly
    so the SSE loop stays correct even after ``app.state`` is recreated
    (the middleware reads ``app.state.container`` at request time).
    """
    container = getattr(request.state, "container", None)
    if container is None:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="presence routes require the container (lifespan not ready)",
        )
    return container


@router.post("/presence/heartbeat")
async def heartbeat(request: Request) -> dict[str, bool]:
    """Record a heartbeat for the calling user (X-User-ID header for W60).

    The W62 auth middleware will replace the header lookup with
    ``request.state.user_id``; until then ``X-User-ID`` lets the slice
    be exercised end-to-end. A missing header is treated as a no-op so
    the client can probe whether the endpoint is reachable without
    forcing an identity.
    """
    container = _container(request)
    raw_user_id = request.headers.get("X-User-ID")
    if raw_user_id:
        try:
            user_id = UUID(raw_user_id)
        except ValueError as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"X-User-ID is not a valid UUID: {raw_user_id!r}",
            ) from exc
        use_cases = container.use_cases  # type: ignore[attr-defined]
        await use_cases["track_presence"](user_id=user_id)
    return {"ok": True}


@router.get("/presence")
async def list_connected(request: Request, limit: int = 100) -> list[dict[str, object]]:
    """Return a JSON snapshot of the currently connected set.

    ``limit`` defaults to 100 rows — the same default the SSE emitter
    uses. The endpoint renders the ``ConnectedUser`` list as a JSON
    array via ``_serialise_user``; consumers that need a strict
    envelope can wrap the array themselves.
    """
    container = _container(request)
    use_cases = container.use_cases  # type: ignore[attr-defined]
    users = await use_cases["get_connected_users"](limit=limit)
    return [_serialise_user(u) for u in users]


@router.get("/presence/stream")
async def presence_stream(request: Request) -> StreamingResponse:
    """Stream the connected set as Server-Sent Events.

    The emitter polls ``get_connected_users`` every 5 seconds and writes
    the snapshot as ``data: <json>\\n\\n``. FastAPI's ``StreamingResponse``
    keeps the coroutine alive across the response lifetime; the loop
    exits when the client disconnects (``request.is_disconnected()``)
    so the server does not leak generators.

    Cancellation safety: the inner ``try/except asyncio.CancelledError``
    makes the loop shut down cleanly when the ASGI server closes the
    connection. FastAPI's ``TestClient`` does not exercise the
    long-running path, so the integration test covers only the
    ``Content-Type`` header — the live behaviour lives behind a runtime
    load test (out of scope for W60).
    """
    container = _container(request)
    use_cases = container.use_cases  # type: ignore[attr-defined]

    async def event_generator() -> AsyncIterator[str]:
        try:
            while True:
                if await request.is_disconnected():
                    break
                users = await use_cases["get_connected_users"](limit=100)
                payload = json.dumps([_serialise_user(u) for u in users])
                yield f"data: {payload}\n\n"
                await asyncio.sleep(_SSE_POLL_INTERVAL_SECONDS)
        except asyncio.CancelledError:  # pragma: no cover - server-side only
            raise

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive"},
    )


__all__ = ["router"]
