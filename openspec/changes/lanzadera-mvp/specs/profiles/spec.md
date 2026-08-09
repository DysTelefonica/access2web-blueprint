# Lanzadera · Profiles specification

> Capability del change `lanzadera-mvp`. Profiles de aplicación declarados en JSONB y la relación profile-app. No cubre la asignación a usuarios (ver `assignments`) ni las capabilities por módulo (D45).

## Purpose

Esta specification describe el dominio `profiles` del módulo `lanzadera`. Cubre la fila de profile por app, las capabilities declaradas en JSONB y el conjunto de profiles disponibles para asignación. No cubre la asignación efectiva a usuarios (ver `assignments`); no reemplaza la capability declaration por módulo (D45).

Decisiones de origen: D5, D7, D22, D45, D46, D56, D58, D110.

## Requirements

### Requirement: Profile per application

The system SHALL persist one or more profiles per app. Each profile SHALL carry a stable code, a display name, a `capabilities` JSONB map, and an `active` flag.

#### Scenario: Standard profile is provisioned at seed

- GIVEN the seed migration `0003_seed_profiles` runs
- WHEN the migration completes
- THEN each of the 20 apps has at least one profile with `code = 'default'`
- AND each default profile's `capabilities` is a non-empty JSON object

#### Scenario: Per-app profiles map legacy roles

- GIVEN the legacy `TbUsuariosAplicacionesPermisos` flags for a given app include `Administrador`, `Calidad`, `Técnico`
- WHEN the seed maps those flags to profiles
- THEN at least one profile per flag is created under that app
- AND the `capabilities` map reflects the legacy role semantics

### Requirement: Capabilities JSONB contract

The system SHALL store `capabilities` as a JSONB map of stable string keys to JSON-typed values. The map SHALL NOT contain anything other than `{string: string|number|boolean}`.

#### Scenario: Capability map is structurally validated

- GIVEN a profile is created or updated
- WHEN the boundary receives the payload
- THEN the JSONB map is validated against the schema
- AND any non-conforming payload is rejected before persistence

#### Scenario: Legacy `F3-F9` dynamic fields are not migrated

- GIVEN a legacy `TbPermisos` row carries `F3..F9` text columns
- WHEN the migration decision is applied
- THEN no row in the new schema mirrors that anti-pattern
- AND any data preserved from those columns is stored as a JSONB entry with explicit names

### Requirement: Profile deactivation is non-destructive

The system SHALL allow a global admin to set `active = false` on a profile. The profile row SHALL remain available for historical assignment lookups; assigning a deactivated profile SHALL be rejected.

#### Scenario: Deactivated profile is not assignable

- GIVEN a profile exists with `active = false`
- WHEN any use case attempts to assign it
- THEN the use case rejects the operation
- AND existing assignments remain readable for audit

#### Scenario: Reactivation is allowed

- GIVEN a profile with `active = false`
- WHEN the global admin reactivates it
- THEN the row stores `active = true` and assignments are again accepted

### Requirement: Profiles are version-aware per app

The system SHOULD expose a profile list per app that excludes profiles whose app has been retired. The list endpoint SHALL be backed by the cache port only when the profile set is stable for the cache TTL.

#### Scenario: Listing profiles for a retired app returns empty

- GIVEN an app is in `registration_status = 'retired'`
- WHEN the profile list for that app is requested
- THEN an empty list is returned
- AND no `404` is raised when the app is reachable through history

## Cross-references

D5, D7, D22, D45, D46, D56, D58, D110.

## Decisiones pendientes

`##ABIERTO##` Set canónico de capabilities por app para los 20 IDs en alcance. El spec se mantiene agnóstico; el catálogo concreto lo cierra `sdd-design`.
