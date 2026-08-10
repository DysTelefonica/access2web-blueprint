# Lanzadera · Apps specification

> Capability del change `lanzadera-mvp`. Catálogo de aplicaciones registradas en la plataforma: metadatos, topología de despliegue y requisitos de presencia física. No cubre el lanzamiento legacy (retirado por D52) ni capabilities por app (ver `profiles`).

## Purpose

Esta specification describe el dominio `apps` del módulo `lanzadera`. Cubre la fila de catálogo de cada aplicación, los atributos de despliegue y la separación entre registro técnico y activación global. No cubre el lanzamiento runtime (D52 lo retira); no cubre capabilities por app (ver `profiles`).

Decisiones de origen: D5, D7, D52, D53, D56, D58, D75, D85.

## Requirements

### Requirement: Application catalog entry

The system SHALL persist one record per application registered in the platform. The record SHALL include a stable identifier, display name, short code, deployment topology, and office-presence requirement. The record SHALL NOT include any user data.

#### Scenario: Seed 20 apps from `TbAplicaciones`

- GIVEN the seed migration `0002_seed_apps` runs against an empty schema
- WHEN the migration completes
- THEN the `apps` table holds exactly 20 rows
- AND each row carries `id`, `name`, `short_code`, `deployment_topology`, `requires_office_presence`, `registration_status = 'active'`
- AND no `command`, `shell`, `unc_path`, or launcher attribute is migrated

#### Scenario: Application without office dependency is marked central

- GIVEN an app whose legacy `EjecucionEnOficina = 'No'` or NULL
- WHEN the migration reads that row
- THEN the persisted record stores `deployment_topology = 'central'`, `requires_office_presence = false`

### Requirement: Deployment topology flag

The system SHALL classify each app as `central` or `office-nas`. The classification SHALL be immutable for a given app version; changes require a new versioned registration.

#### Scenario: Office-NAS app flagged for office presence

- GIVEN an app whose legacy `EjecucionEnOficina = 'Sí'`
- WHEN the migration reads that row
- THEN the persisted record stores `deployment_topology = 'office-nas'`, `requires_office_presence = true`

#### Scenario: Topology change requires new version

- GIVEN an app is registered with `deployment_topology = 'central'`
- WHEN a global admin attempts to flip the topology in place
- THEN the system rejects the in-place change
- AND requires a new app version or a new registration row

### Requirement: Registration lifecycle

The system SHALL keep each app in one of three registration states: `pending`, `active`, `retired`. The deployment technical path creates `pending` rows; the global admin transitions to `active`. Retired is terminal.

#### Scenario: Deployment creates pending row

- GIVEN a new service is deployed
- WHEN the registration hook runs
- THEN a row is inserted with `registration_status = 'pending'`
- AND the app is not visible to any non-admin user

#### Scenario: Global admin activates pending app

- GIVEN an app is in `pending`
- WHEN the global admin confirms the activation
- THEN the row transitions to `active`, an audit record is appended, and the app becomes visible according to user assignments

### Requirement: Application listing is permission-filtered

The system SHALL return only the apps the requesting user is authorised to see. The list endpoint SHALL NOT expose global IDs of hidden apps.

#### Scenario: User sees only assigned apps

- GIVEN a user has 3 assignments across 20 apps
- WHEN the user requests the apps list
- THEN the response contains exactly those 3 apps
- AND no other `id` is included in the payload

## Cross-references

D5, D7, D52, D53, D56, D58, D75, D85.

## Decisiones pendientes

`##ABIERTO##` Lista de campos de catálogo legacy que NO migran (`Pass`, `Comando`, `URLDIrectorioIconoAplicacion`, etc.). Requiere tabla de mapeo en `sdd-design`.
