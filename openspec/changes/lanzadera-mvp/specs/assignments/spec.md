# Lanzadera · Assignments specification

> Capability del change `lanzadera-mvp`. Matriz efectiva de permisos user × app × profile y regla de mapeo legacy → moderno. No cubre el catálogo de apps (ver `apps`) ni los profiles (ver `profiles`).

## Purpose

Esta specification describe el dominio `user_app_assignments` del módulo `lanzadera`. Cubre la matriz efectiva de permisos, la regla de mapeo desde `TbUsuariosAplicacionesPermisos` y la regla de exclusividad `SinAcceso`. No cubre la identidad del usuario (ver `users`) ni los profiles abstractos (ver `profiles`).

Decisiones de origen: D7, D22, D23, D42, D43, D56, D85, D102, D110.

## Requirements

### Requirement: Assignment matrix

The system SHALL persist the triple `(user_id, app_id, profile_id)` as the only source of truth for effective permissions. Each row SHALL be unique on the triple. The schema SHALL NOT mirror the legacy 7-flag column structure.

#### Scenario: Single assignment creates one row

- GIVEN a user has `profile = ADMIN` on an app
- WHEN the assignment use case runs
- THEN exactly one row joins the three IDs
- AND no row carries a column named `Administrador`, `Calidad`, `Técnico`, `Economía`, `Secretaría`, `CalidadAvisos` or `SinAcceso`

#### Scenario: Compound assignment creates multiple rows

- GIVEN a legacy user has `Calidad = 'Sí'` AND `CalidadAvisos = 'Sí'` on the same app
- WHEN the mapping rule applies
- THEN two rows are created: `CALIDAD` and `CALIDAD_AVISOS`

### Requirement: `SinAcceso` exclusivity rule

The system SHALL treat `SinAcceso = 'Sí'` as exclusive. Mixed legacy rows with `SinAcceso` plus other `Sí` flags collapse to a single `SIN_ACCESO` row.

#### Scenario: Compound row with `SinAcceso` collapses to single SIN_ACCESO

- GIVEN a legacy row has `SinAcceso = 'Sí'`, `Administrador = 'Sí'`, `Técnico = 'Sí'`
- WHEN the mapping rule applies
- THEN exactly one row is created with `profile_id = SIN_ACCESO`
- AND no row with `profile_id = ADMIN` or `TECNICO` is created for that (user, app)

#### Scenario: Compound row without `SinAcceso` follows the normal rule

- GIVEN a legacy row has `SinAcceso = NULL` and `Administrador = 'Sí'`, `Técnico = 'Sí'`
- WHEN the mapping rule applies
- THEN two rows are created: `ADMIN` and `TECNICO`; no `SIN_ACCESO` row is created

### Requirement: All-None defaults to DEFAULT

The system SHALL treat `all_No / NULL` legacy rows as `DEFAULT`. The mapping rule SHALL emit at least one row per (user, app) when the legacy source had a row at all.

#### Scenario: All-No legacy row maps to DEFAULT

- GIVEN a legacy row has every flag set to `'No'` or NULL
- WHEN the mapping rule applies
- THEN one row is created with `profile_id = DEFAULT` and no other row is created for that (user, app)

### Requirement: Assignment write is admin-gated

The system SHALL accept assignment mutations only from actors authorised by `global_admins`. App-level admins MAY assign users to profiles within their app scope.

#### Scenario: Non-admin actor is rejected

- GIVEN a non-admin actor invokes the create-assignment use case
- WHEN the boundary checks the actor
- THEN the operation is rejected and no row is inserted

#### Scenario: Global admin assigns a profile

- GIVEN a global admin invokes the create-assignment use case
- WHEN the boundary checks the actor
- THEN the operation is accepted, a row is inserted, and an audit record is appended

### Requirement: Assignment read is filtered by viewer scope

The system SHALL return only the rows the requesting user is allowed to see. The list endpoint SHALL NOT expose assignments for users outside the actor's scope.

#### Scenario: Listing is scope-filtered

- GIVEN a global admin requests the assignments for a target user
- WHEN the list endpoint runs
- THEN the response contains every row for that user
- AND a per-app admin viewing the same endpoint sees only the rows for that app

## Cross-references

D7, D22, D23, D42, D43, D56, D85, D102, D110. ABIERTO: peso de `SinAcceso` en navegación menú global.
