# Spec: Stub replacement + auth_bypass retirement (PR-7)

> Capability del change `w62-auth-flow`. Cierra el último eslabón: las rutas admin leen `request.state.user_id` directamente, y `auth_bypass` se retira de los tests.

## Purpose

Las rutas admin (`admin_routes_apps.py`, `admin_routes_users.py`, `admin_routes_misc.py`) tienen un helper `_actor_id(request)` que retorna `None` como stub. PR-7 reemplaza este helper con lectura real desde `request.state.user_id` (poblado por el middleware de PR-5).

Además, los tests de integración usan `auth_bypass` fixture (PR-W-TEST #519) que monkeypatchea `require_global_admin`. Con la implementación real de PR-6, el bypass ya no es necesario — los tests deben usar una sesión real (admin global) para ejercitar las rutas destructivas.

## Requirements

### Requirement: _actor_id reads from request.state

The system SHALL replace the ``_actor_id(request)`` stub in each
admin route file with a call that reads
``getattr(request.state, "user_id", None)``.

#### Scenario: Authenticated admin request

- GIVEN a request with ``request.state.user_id = <admin_uuid>``
- WHEN the route calls ``_actor_id(request)``
- THEN it returns ``<admin_uuid>``

#### Scenario: Unauthenticated request

- GIVEN a request with ``request.state.user_id = None``
- WHEN the route calls ``_actor_id(request)``
- THEN it returns ``None`` (the route already gated by ``require_global_admin`` before this call)

### Requirement: auth_bypass retirement

The system SHALL remove the ``auth_bypass`` fixture from
``tests/lanzadera/conftest.py`` and replace its uses in the admin
integration tests with an ``auth_session`` fixture that injects a
real session into the request state.

#### Scenario: auth_session fixture

- GIVEN the ``auth_session`` fixture is active (auto-applied to admin integration tests)
- WHEN an admin route is invoked
- THEN ``request.state.user_id`` is set to a known global admin UUID

The fixture replaces the test-side stubbing: instead of
``monkeypatch.setattr(_admin, "require_global_admin", lambda: None)``, the
fixture seeds a global admin row + a session row in the in-memory
fakes, then injects a minimal FastAPI app whose middleware populates
``request.state.user_id`` from a stub JWT.

### Requirement: Test domain remains unit/integration

The system SHALL NOT introduce a Postgres dependency in the test
path. Tests still run against ``FakeUserRepository`` /
``FakeSessionRepository`` / ``FakeGlobalAdminRepository`` /
``FakeAuditLog``.

## File-surface contract

| File | Action | Notes |
|---|---|---|
| `app/src/modules/lanzadera/delivery/http/admin_routes_apps.py` | MODIFY | `_actor_id` reads `request.state.user_id`. |
| `app/src/modules/lanzadera/delivery/http/admin_routes_users.py` | MODIFY | same. |
| `app/src/modules/lanzadera/delivery/http/admin_routes_misc.py` | MODIFY | same. |
| `tests/lanzadera/conftest.py` | MODIFY | remove `auth_bypass` fixture definition. Add `auth_session` fixture that pre-seeds a global admin and a session in the in-memory fakes. |
| `tests/lanzadera/delivery/test_admin_routes_integration.py` | MODIFY | remove `@pytest.mark.usefixtures("auth_bypass")` from each test; add `@pytest.mark.usefixtures("auth_session")`. |
| `tests/lanzadera/di/test_lanzadera_container.py` | MODIFY | remove the explicit `auth_bypass` fixture invocation in any test that uses it (none currently — confirmed during PR-7 implementation). |
| `docs/03-aplicaciones/lanzadera/epic.md` | UPDATE | mark W62 as completed. |

## Implementation reference — auth_session fixture

```python
# tests/lanzadera/conftest.py
@pytest.fixture
def auth_session(fake_fixtures: FakeFixtures) -> Iterator[None]:
    """Seed a global admin + session so admin routes see a real user.

    The admin integration tests mount a minimal FastAPI app that:
    1. populates ``request.state.user_id`` from a session id resolved
       out of a stub Bearer token;
    2. calls ``require_global_admin`` (PR-6 implementation) which
       checks ``request.state.user_id`` against ``FakeGlobalAdminRepository``.

    This replaces the old ``auth_bypass`` fixture that monkeypatched
    the placeholder gate.
    """
    admin_id = uuid4()
    fake_fixtures.global_admins.grant(admin_id)
    session_id = uuid4()
    fake_fixtures.users.add(
        User(
            id=admin_id,
            email="admin@enterprise.test",
            name="Admin",
            dni_encrypted=b"enc:admin",
            password_hash="fake:irrelevant",
            status=UserStatus.ACTIVE,
            failed_attempts=0,
            last_login_at=None,
            created_at=datetime.now(UTC),
            updated_at=datetime.now(UTC),
        )
    )
    fake_fixtures.sessions.sessions[session_id] = Session(
        id=session_id,
        user_id=admin_id,
        created_at=datetime.now(UTC),
        expires_at=datetime.now(UTC) + timedelta(hours=1),
    )
    # The integration tests read the session id from a stub token set
    # by the test (not via real JWT — that's PR-5's job).
    yield {"admin_id": admin_id, "session_id": session_id}
```

Tests then use the fixture like:

```python
@pytest.mark.usefixtures("auth_session")
async def test_create_user_endpoint(container_with_app) -> None:
    response = await client.post("/admin/users", json={...})
    assert response.status_code == 201
```

The fixture is opt-in (not autouse) so non-auth integration tests can skip it.

## Decisiones cerradas

- **auth_session is opt-in** (not autouse): only admin integration tests opt in via `@pytest.mark.usefixtures("auth_session")`. Other tests that don't need auth stay untouched.
- **No test-side stubbing**: the `auth_bypass` fixture's `monkeypatch.setattr(_admin, "require_global_admin", lambda: None)` is fully gone.
- **`_actor_id` becomes a 1-line function**: read state, return value. Keep it as a helper for readability.
- **JWT in tests**: the `auth_session` fixture does NOT sign a real JWT. The integration tests mount the routes WITHOUT the `AuthMiddleware` (or with a test-only middleware that reads `X-Test-Session-Id` and populates `request.state`). This avoids pulling PR-5's JWT signing into every admin integration test.

  Wait — actually, simpler approach: the test mounts the routes WITHOUT the AuthMiddleware, and instead sets `request.state.user_id` via a test-only middleware. Or just have the test directly call `require_global_admin` after the admin route, with the state pre-populated by a fixture.

  Decision: use a **test-only middleware** that reads `X-Test-Session-Id` and populates state. This keeps the admin integration tests pure (no JWT signing required) while exercising the real `require_global_admin` (which is the contract we want to verify).

## Tests updates

The PR modifies existing tests but does NOT add new tests in PR-7 — the existing tests should pass without changes (modulo the auth_session fixture injection).

If any test was relying on `auth_bypass` and now needs `auth_session`, the diff catches it. Expected count: ~8 tests in `test_admin_routes_integration.py` plus 1 in `test_lanzadera_container.py`.

## Verification

```bash
# Verify auth_bypass is gone
grep -r "auth_bypass" tests/ --include="*.py"
# → (no results)

# Suite completa
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app --cov=app --cov-report=term
# → 733 + ~0 = ~733 passed (no new tests; existing ones pass via auth_session)

# Gates
ruff format --check --config app/pyproject.toml .
ruff check --config app/pyproject.toml .
/usr/bin/python3.12 -m mypy --explicit-package-bases app/
python3 scripts/check_complexity.py --root .
python3 scripts/check_mutation_sites.py --root .
python3 scripts/check_test_classification.py --root tests/lanzadera
python3 scripts/check_check_classification.py --strict
# → all pass

# Smoke: el epic.md se actualiza con W62 done
grep -c "W62" docs/03-aplicaciones/lanzadera/epic.md
# → ≥ 1
```

## Cierre del change

Tras mergear PR-7:
- Cerrar issue #537 (umbrella W62).
- Archivar el change `w62-auth-flow` con `openspec archive w62-auth-flow` (si el comando existe; si no, mover manualmente el directorio a `openspec/changes/archive/`).
- CHANGELOG: añadir W62 como bloque.
