# Spec: Logout use case (PR-3)

> Capability del change `w62-auth-flow`. Sub-capacidad `auth-core` extendida con terminación de sesión.

## Purpose

Inversa de `login`: termina la sesión activa del usuario revocando la fila `Session` asociada. La fila no se borra — `expires_at` se acorta a `now` para preservar la trazabilidad del audit (D-W62-2 + DA-11).

## Requirements

### Requirement: Revoke by session id

The system SHALL revoke the session identified by ``session_id`` via
``SessionRepository.revoke(session_id, now)``, which shortens
``expires_at`` to ``now`` (idempotent on the second call).

#### Scenario: Successful logout

- GIVEN a session row exists for ``session_id`` with ``expires_at > now``
- WHEN ``logout(session_id, *, now, sessions, audit, actor_id)`` is invoked
- THEN ``sessions.revoke(session_id, now)`` is called
- AND ``sessions.sessions[session_id].expires_at == now``
- AND ``audit.append`` records ``event_type="auth.logout.success"``, ``result="success"``, ``target_id=str(session_id)``, ``payload={"user_id": <user_id>}``
- AND no exception is raised

#### Scenario: Logout of unknown session id

- GIVEN no session row exists for ``session_id``
- WHEN ``logout(session_id, *, now, sessions, audit, actor_id)`` is invoked
- THEN ``SessionNotFoundError`` is raised
- AND ``audit.append`` records ``event_type="auth.logout.failure"``, ``result="failure"``, ``target_id=str(session_id)``, ``payload={"reason": "unknown_session"}``
- AND ``sessions.revoke_calls`` is empty

#### Scenario: Idempotent logout (already-revoked session)

- GIVEN a session row exists for ``session_id`` and ``expires_at < now``
- WHEN ``logout(session_id, *, now2, sessions, audit, actor_id)`` is invoked with ``now2 > now``
- THEN ``sessions.sessions[session_id].expires_at`` stays at the earlier timestamp (``now``)
- AND ``audit.entries`` has two success rows
- AND no exception is raised

### Requirement: Validation

The system SHALL raise ``ValueError`` when ``session_id`` is ``None``.

#### Scenario: None session id

- GIVEN ``session_id is None``
- WHEN ``logout(None, *, now, sessions, audit)`` is invoked
- THEN ``ValueError`` is raised

### Requirement: Exception surface

The system SHALL raise ``SessionNotFoundError`` (defined in
``domain/errors.py``) on unknown session id. The caller (delivery,
PR-5) maps this to HTTP 401 to avoid leaking which sessions were once
valid.

### Requirement: Audit row on every path

The system SHALL emit exactly one audit row per call regardless of
success or failure (DA-11).

## File-surface contract

| File | Action | Notes |
|---|---|---|
| `app/src/modules/lanzadera/application/logout.py` | NEW | async function `logout(session_id, *, now, sessions, audit, actor_id) -> None`. |
| `app/src/modules/lanzadera/domain/errors.py` | MODIFY | add `class SessionNotFoundError(Exception)`. |
| `app/src/modules/lanzadera/di/container.py` | MODIFY | add public `async def logout(self, session_id, *, actor_id=None) -> None` calling `self._use_cases["logout"](...)`. |
| `app/src/modules/lanzadera/di/use_cases.py` | MODIFY | add `from app.src.modules.lanzadera.application.logout import logout`. Add `"logout": functools.partial(logout, sessions=session_repo, audit=audit)` to the factory dict. |
| `tests/lanzadera/application/test_logout.py` | NEW | Categoría 2 — 7 test cases (see below). |

## Tests Cat 2 — `tests/lanzadera/application/test_logout.py`

All tests use `FakeSessionRepository`, `FakeAuditLog` from `tests/lanzadera._fakes`. NO mocks.

1. `test_logout_revokes_session` — happy path: session expires_at becomes `now`.
2. `test_logout_appends_audit_success_row` — single audit row with `event_type="auth.logout.success"`, `target_id=str(session.id)`, `payload={"user_id": str(session.user_id)}`.
3. `test_logout_records_exactly_one_revoke_call` — `sessions.revoke_calls == [session.id]`.
4. `test_logout_raises_session_not_found_for_unknown_session_id` — `pytest.raises(SessionNotFoundError)`.
5. `test_logout_appends_audit_failure_row_for_unknown_session` — payload has `reason="unknown_session"`.
6. `test_logout_is_idempotent_on_second_call` — second revoke doesn't raise, `expires_at` stays at first timestamp.
7. `test_logout_raises_value_error_for_none_session_id` — `pytest.raises(ValueError)`.

Helper `_seed_session(sessions, ...)` MUST `await sessions.create(session)` (async fake — learned the hard way in PR-2 worktree).

## Container integration

```python
# di/use_cases.py (factory dict entry)
"logout": functools.partial(
    logout,
    sessions=session_repo,
    audit=audit,
),
```

`session_repo` is already wired in PR-2; no container refactor needed.

## Decisión sobre BASELINE updates

`application/logout.py` debe quedar bajo 100 mutation sites (PR-2 baseline shows login.py at 167 → may exceed ceiling). Si excede, añadir entry al `BASELINE` de `scripts/check_mutation_sites.py` con `target_date="2027-02-13"`.

`container.py` no crece en este PR (sólo añadimos un método público que llama `self._use_cases["logout"]`). BASELINE ya cubierta.

## Verificación

```bash
# Sólo los nuevos tests
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/lanzadera/application/test_logout.py -v
# → 7 passed

# Suite completa
/usr/bin/python3.12 -m pytest -c app/pyproject.toml --rootdir=app --cov=app --cov-report=term
# → 706 + 7 = 713 passed, coverage ≥ 82%

# Gates
ruff format --check --config app/pyproject.toml .
ruff check --config app/pyproject.toml .
/usr/bin/python3.12 -m mypy --explicit-package-bases app/
python3 scripts/check_complexity.py --root .
python3 scripts/check_mutation_sites.py --root .
python3 scripts/check_test_classification.py --root tests/lanzadera
# → all pass
```
