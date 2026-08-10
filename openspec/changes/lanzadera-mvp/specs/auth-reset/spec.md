# Lanzadera · Auth reset flow specification

> Capability del change `lanzadera-mvp`. Reset de密码 con tokens one-time de 24 h, emisión vía notification port (cola por tabla como adapter v1), consumo atómico. **No** incluye el bootstrap del primer admin global (ver `auth-bootstrap`); **no** incluye el almacenamiento de credenciales (ver `auth-core`).

## Purpose

Esta specification describe el reset flow del dominio `auth`: `issue_reset_token(user_id) -> str` y `consume_reset_token(token, new_password) -> bool`. Cubre la emisión, supersession, expiración de 24 h, consumo atómico y el guard que impide emitir tokens antes de que exista un global admin.

Decisiones de origen: D90.

## Requirements

### Requirement: Reset token issuance and supersession

The system SHALL expose `issue_reset_token(user_id) -> str` that returns a one-time token with a 24-hour expiry. The raw value SHALL be emitted via the notification port only; the server SHALL store the hash.

#### Scenario: Issued token is single-use and time-bounded

- GIVEN a user with `status = 'password_reset_required'`
- WHEN `issue_reset_token(user_id)` runs
- THEN the persistence layer stores `token_hash`, `expires_at = now + 24h`, `consumed_at = NULL`, `superseded_at = NULL`
- AND the notification adapter is invoked with the raw token

#### Scenario: Re-issuing supersedes the previous token

- GIVEN a user has an unconsumed, unexpired token
- WHEN `issue_reset_token(user_id)` runs again
- THEN the previous token is marked `superseded_at = now` and the new token is persisted

### Requirement: Reset token consumption

The system SHALL expose `consume_reset_token(token, new_password) -> bool` that atomically validates the token, marks it consumed, and updates the user's password hash. The helper SHALL be `false` if the token is unknown, expired, superseded, or already consumed.

#### Scenario: Valid token updates the password

- GIVEN a user has a valid unconsumed token
- WHEN `consume_reset_token(token, new_password)` runs
- THEN the helper returns `true`, `users.password_hash = hash_password(new_password)` (Argon2id), `users.status = 'active'`, and the token is marked `consumed_at = now`

#### Scenario: Unknown, expired, superseded, or already-consumed token is rejected

- GIVEN a token whose hash is unknown, expired, superseded, or already consumed
- WHEN `consume_reset_token(token, new_password)` runs
- THEN the helper returns `false` and `password_hash` is unchanged

### Requirement: Reset-by-email disabled until a global admin exists

The system SHALL refuse to issue reset tokens until at least one global admin exists. This guarantees that the first admin cannot bootstrap themselves via email (ver `auth-bootstrap` para el camino CLI exclusivo).

#### Scenario: issue_reset_token rejected with no_global_admin

- GIVEN no global admin exists
- WHEN any flow tries to call `issue_reset_token` for a non-admin user
- THEN the system rejects the call with a `no_global_admin` error

## Cross-references

D90.

## Decisiones pendientes

`##ABIERTO##` Política exacta del canal de notificación al usuario cuando se emite un token (P20, SMTP corporativo real). El MVP usa la cola por tabla como adapter v1; el contrato del `NotificationDeliveryPort` lo concreta `sdd-design`.
