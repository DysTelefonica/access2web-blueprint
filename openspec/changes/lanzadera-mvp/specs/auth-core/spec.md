# Lanzadera · Auth core specification

> Capability del change `lanzadera-mvp`. Credential storage con Argon2id, login, lockout configurable. **No** incluye compat con hashes legacy: D36+D37 están OBSOLETOS; D88+D89 los sustituyen. La columna `legacy_hash` no existe en el schema.

## Purpose

Esta specification describe el núcleo del dominio `auth`: hashing Argon2id, verificación, lockout configurable y persistencia de credenciales sin legacy. Cubre `CredentialHasherArgon2id.hash`, `.verify`, lockout y la regla «no legacy verification path exists». No cubre reset de contraseña (ver `auth-reset`), bootstrap CLI, identidad ni sesiones web persistentes.

Decisiones de origen: D9, D38, D39, D40, D42, D88, D89.

## Requirements

### Requirement: Argon2id-only credential storage

The system SHALL hash every password with Argon2id via `argon2-cffi==25.1.0` using the `RFC_9106_LOW_MEMORY` profile (Argon2id, 64 MiB memory, 3 iterations, 4 threads). The schema SHALL NOT contain `legacy_hash`, `password_legacy`, or any pre-Argon2id field.

#### Scenario: Hash uses Argon2id with the pinned profile

- GIVEN `await CredentialHasherArgon2id.hash(plain) -> str` is invoked
- WHEN the helper computes the hash
- THEN the produced string is the Argon2id PHC string with embedded parameters m=65536, t=3, p=4

#### Scenario: Verify accepts matching plaintext and rejects wrong

- GIVEN `password_hash = await CredentialHasherArgon2id.hash('correct horse battery staple')`
- WHEN `await CredentialHasherArgon2id.verify('correct horse battery staple', password_hash)` runs
- THEN the helper returns `true`
- AND `await CredentialHasherArgon2id.verify('wrong', password_hash)` returns `false`

#### Scenario: Schema forbids legacy hash column

- GIVEN the migration `0004_seed_users` runs
- WHEN the schema is asserted
- THEN the `users` table has no `legacy_hash`, `password_legacy`, or `pass_hash_v1` column

### Requirement: Seeded users start with reset required

The system SHALL persist every migrated user with `password_hash = NULL` and `status = 'password_reset_required'`. The user SHALL NOT authenticate until `consume_reset_token` succeeds (ver `auth-reset`).

#### Scenario: Empty password attempt is rejected

- GIVEN a user has `password_hash = NULL`, `status = 'password_reset_required'`
- WHEN the login endpoint is invoked with any password
- THEN the system rejects the login and no Argon2id computation runs against an empty hash

### Requirement: Lockout after repeated failures

The system SHALL lock a user after `lockout_threshold` (default 5) consecutive failed logins. The lock SHALL last `lockout_duration` (default 1 hour) unless a global admin unlocks the user.

#### Scenario: Fifth failure locks the account

- GIVEN a user has 4 consecutive failed attempts
- WHEN the 5th attempt fails
- THEN the user transitions to `locked` and an audit record of type `auth.lockout` is appended

#### Scenario: Global admin unlock resets the counter

- GIVEN a user is in `locked`
- WHEN the global admin invokes the unlock action
- THEN the user transitions to `password_reset_required` and the failed-attempt counter is reset to 0

### Requirement: No legacy compat layer

The system SHALL NOT contain any code path that reads or verifies a legacy hash. The auth module SHALL be modelled for Argon2id only.

#### Scenario: No legacy verification path exists

- GIVEN the auth module is searched at the source level
- WHEN static analysis looks for `legacy_hash`, `verify_legacy`, `sha256`, `old_password`, `migrate_password`
- THEN no source file in `auth` references those symbols
- AND a pinning test asserts the absence

## Cross-references

D9, D38, D39, D40, D42, D88, D89, QC-5.

## Decisiones pendientes

`##ABIERTO##` Política de expiración de contraseña configurable (D41). El MVP implementa `status = 'password_reset_required'` sin caducidad periódica; el reset flow cubre la recuperación.
