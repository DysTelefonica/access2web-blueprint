# Lanzadera · Global admins specification

> Capability del change `lanzadera-mvp`. Siembra y operación de administradores globales de plataforma. No cubre administradores de aplicación (D22) ni el bootstrap del primer admin (ver `auth`).

## Purpose

Esta specification describe el dominio `global_admins` del módulo `lanzadera`. Cubre la siembra desde `GLOBAL_ADMIN_EMAILS`, la idempotencia al arranque del proceso y la gestión de altas y bajas por CLI. No cubre el bootstrap del primer admin (ver `auth`); no cubre administradores de aplicación (D22 los define por módulo).

Decisiones de origen: D7, D21, D23, D25, D26, D42, D48, D91.

## Requirements

### Requirement: Global admin identity

The system SHALL persist the set of global admin emails in a `global_admins` table joined to `users`. The join SHALL be unique on `user_id`. The system SHALL NOT materialise a global admin role via a column on `users`.

#### Scenario: Global admin row is unique per user

- GIVEN a user is a global admin
- WHEN the row is written
- THEN exactly one `global_admins` row references `user_id`
- AND a second attempt with the same `user_id` is rejected

### Requirement: Bootstrap from `GLOBAL_ADMIN_EMAILS`

The system SHALL expose a `BootstrapAdapter` that reads `GLOBAL_ADMIN_EMAILS` (semicolon-separated) on process start and ensures each email is a global admin. The adapter SHALL be idempotent across restarts.

#### Scenario: First run creates the users

- GIVEN `GLOBAL_ADMIN_EMAILS = '[email protected];[email protected]'`
- AND neither user exists
- WHEN the bootstrap adapter runs
- THEN two `users` rows are created with `status = 'password_reset_required'`, two `global_admins` rows reference them, and an audit record of type `global_admins.bootstrap` is appended

#### Scenario: Subsequent run is idempotent and missing env var is empty

- GIVEN the previous scenario already ran
- WHEN the bootstrap adapter runs again
- THEN no new rows are created and existing rows are not mutated
- AND if `GLOBAL_ADMIN_EMAILS` becomes unset, the adapter returns without error and no user is created

### Requirement: Grant and revoke via CLI

The system SHALL expose `gentle-ai platform user grant-global-admin <email>` and `revoke-global-admin <email>` as the only paths to mutate `global_admins`. The CLI SHALL refuse to revoke the last global admin.

#### Scenario: Grant and revoke round-trip

- GIVEN a user exists with `status = 'active'`
- WHEN a global admin invokes `grant-global-admin <email>` then `revoke-global-admin <email>`
- THEN the user becomes a global admin and is removed; an audit record is appended for each mutation

#### Scenario: Cannot revoke the last global admin

- GIVEN exactly one global admin exists
- WHEN the CLI is invoked with `revoke-global-admin <self>`
- THEN the CLI rejects the action and the row is unchanged

### Requirement: Global admin required for destructive actions

The system SHALL require global admin membership for any destructive or scope-changing action. The CLI SHALL enforce the check; the web layer SHALL enforce the same check via the same use case.

#### Scenario: Non-admin cannot invoke destructive CLI

- GIVEN a non-admin actor invokes any `gentle-ai platform user ...` subcommand
- WHEN the CLI checks authorization
- THEN the command is rejected and no mutation runs

#### Scenario: Destructive CLI requires explicit confirmation

- GIVEN a global admin invokes `revoke-global-admin`
- WHEN the CLI parses the call
- THEN the CLI prompts for an explicit confirmation and the action runs only after confirmation

### Requirement: Audit emission for global admin mutations

The system SHALL append an audit record for every bootstrap, grant, and revoke operation. The record SHALL include actor, target email, action type, and timestamp.

#### Scenario: Audit record is emitted for every mutation

- GIVEN a global admin mutation runs
- WHEN the use case completes
- THEN an audit record of type `global_admins.<action>` is appended with actor, target, timestamp, and correlation ID

## Cross-references

D7, D21, D23, D25, D26, D42, D48, D91. ABIERTO: canal de notificación al usuario cuando recibe o pierde el rol (D43 lo hereda).
