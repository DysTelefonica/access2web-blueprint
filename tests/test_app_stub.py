# HARNESS-PROVENANCE: lanzadera-mvp Phase 0 — healthcheck wiring.
"""Smoke test for the FastAPI composition root.

Phase 0 ships only the application factory (`app.src.main.app`) and a single
`/health` route. This test exercises both via FastAPI's in-process `TestClient`
so the coverage gate sees a non-empty statement surface — without any real
application logic, the global coverage would be 0% and the `--cov-fail-under=85`
floor would fail.

We use `fastapi.testclient.TestClient` (Starlette's sync wrapper around httpx)
instead of `httpx.ASGITransport` because httpx 0.28's ASGITransport is
async-only and would require `pytest-asyncio`. The TestClient gives us a
sync API with no extra deps; the same dependency already powers the FastAPI
delivery layer in Phase 5.
"""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient


@pytest.fixture(scope="module")
def client():
    from app.src.main import app

    with TestClient(app) as session:
        yield session


def test_healthcheck_returns_ok(client: TestClient) -> None:
    response = client.get("/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["phase"] == "1"
    assert payload["service"] == "lanzadera"


def test_app_factory_exposes_expected_metadata() -> None:
    """The FastAPI app carries the Phase 0 contract in its title and version."""
    from app.src.main import app

    assert app.title == "Platform — Lanzadera MVP"
    # `version` is pinned in pyproject.toml — `0.0.0+phase0` is the Phase 0
    # marker the delivery layer uses to gate preview builds.
    assert "phase0" in app.version.lower()
