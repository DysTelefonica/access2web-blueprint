# Lanzadera · Users specification

> Capability del change `lanzadera-mvp`. Ciclo de vida de usuarios: alta, baja, perfil y consulta. No cubre auth (ver `auth`) ni autorización de plataforma (ver `global_admins`).

## Purpose

Esta specification describe el dominio `users` del módulo `lanzadera`. Cubre el almacenamiento en PostgreSQL, las invariantes de los atributos personales y la separación de scopes. No cubre autenticación, credentials storage, ni emisión de tokens de reset.

Decisiones de origen: D5, D7, D9, D42, D56, D58, D89.

## Requirements

### Requirement: User identity record

The system SHALL persist a canonical user record per person authorised to access the platform. The record SHALL include corporate email (unique, normalised lowercase), display name, encrypted national ID, internal status, and timestamps. The record SHALL NOT store any plaintext credential.

#### Scenario: Create user with minimum attributes

- GIVEN a global admin submits an email, a display name and a national ID
- WHEN the create-user use case runs
- THEN the system stores `email = lower(input)`, `name = input.name`, `dni_encrypted = encrypt(input.dni)`, `status = 'password_reset_required'`, `password_hash = NULL`
- AND no plaintext copy of the national ID is persisted

#### Scenario: Duplicate email rejected

- GIVEN a user with the same corporate email already exists
- WHEN the create-user use case runs with that email
- THEN the system rejects the operation with a duplicate-email error
- AND no row is inserted

#### Scenario: National ID stored encrypted at rest

- GIVEN a user has been created with a national ID
- WHEN a row is read directly from the database
- THEN the `dni_encrypted` column carries ciphertext only
- AND the encryption key is fetched from the secret manager, never from the repository

### Requirement: User lifecycle states

The system SHALL keep the user record in one of four states: `active`, `disabled`, `password_reset_required`, `locked`. Transitions SHALL be triggered only by the events the spec authorises; any unauthorised transition SHALL be rejected.

#### Scenario: Global admin disables a user

- GIVEN an active user exists
- WHEN a global admin issues the disable command
- THEN the user transitions to `disabled`
- AND future login attempts SHALL be rejected by `auth`
- AND an audit record is appended

#### Scenario: User is unlocked only by global admin

- GIVEN a user is in `locked`
- WHEN the unlock action is invoked
- THEN the system accepts it only if the actor is a global admin; otherwise the action is rejected

### Requirement: Profile attributes redaction

The system SHALL expose a profile view that returns display name, email, status, and last successful login. The view SHALL NOT include `password_hash`, `dni_encrypted`, or any token value.

#### Scenario: Profile view redacts sensitive fields

- GIVEN a user record exists
- WHEN any read endpoint or use case returns the profile
- THEN the response contains `email`, `name`, `status`, `last_login_at`
- AND does NOT contain `password_hash`, `dni_encrypted`, or any `reset_token_*` field

### Requirement: National ID encryption key separation

The system SHALL source the national ID encryption key from a secret manager injected at composition root. The key value SHALL NOT appear in environment variables committed to the repository, in logs, or in CLI arguments.

#### Scenario: Key sourced from secret manager

- GIVEN the composition root is built
- WHEN the `users` module is initialised
- THEN the encryption key is obtained from the secret manager port
- AND no plaintext key is read from `os.environ` or argv

## Cross-references

D5, D7, D9, D42, D56, D58, D89; QC-2, QC-5, QC-9.

## Decisiones pendientes

`##ABIERTO##` Política exacta de normalización del email (lowercase completo vs `local-part` lower + `domain` lower, IDN de segundo nivel).
