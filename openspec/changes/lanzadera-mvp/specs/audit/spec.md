# Lanzadera · Audit specification

> Capability del change `lanzadera-mvp`. Log canónico de eventos de autenticación y apertura de aplicación. **No** incluye telemetría de SSID, ubicación o coordenadas (D55 la retira).

## Purpose

Esta specification describe el dominio `audit` del módulo `lanzadera`. Cubre eventos de autenticación y apertura de app. No cubre la operación de la cola de notificación (D11, D13); no cubre auditoría de operación plataforma fuera del scope Lanzadera.

Decisiones de origen: D11, D13, D27, D28, D29, D55, D56, D112, P-21.

## Requirements

### Requirement: Audit event types

The system SHALL emit at minimum: `auth.login.success`, `auth.login.failure`, `auth.lockout`, `auth.reset.issued`, `auth.reset.consumed`, `auth.bootstrap.set_password`, `global_admins.grant`, `global_admins.revoke`, `app.open`. Every event SHALL carry `actor`, `target`, `module`, `result`, `timestamp`, `correlation_id`.

#### Scenario: Login success and failure emit one event each

- GIVEN a user authenticates
- WHEN the login flow completes
- THEN an event of type `auth.login.success` (or `auth.login.failure`) is appended with actor, target, module, `result`, timestamp, correlation_id
- AND the failure case payload does NOT contain the attempted password or any hash

### Requirement: No legacy telemetry fields

The system SHALL NOT persist or emit SSID, BSSID, geolocation coordinates, or machine name. The schema SHALL NOT contain columns for those fields.

#### Scenario: Schema forbids telemetry columns

- GIVEN the audit table is created
- WHEN the schema is asserted
- THEN no `ssid`, `bssid`, `coordinates`, `machine_name` column exists
- AND a pinning test asserts the absence

### Requirement: Seed audit from legacy without telemetry

The system SHALL seed the audit table from legacy `TbConexiones` and `TbAplicacionesAperturas` during migration `0006`. The seed SHALL preserve authentication and app-open events; it SHALL NOT preserve SSID, location, or coordinates.

#### Scenario: Legacy connections map to audit events

- GIVEN legacy rows in `TbConexiones` carry email, last connection, success flag
- WHEN the seed runs
- THEN each row produces one historical event of type `auth.login.success` (or `failure` depending on `Exitoso`)
- AND the source SSID, office, and coordinates are skipped

#### Scenario: Legacy app openings map to audit events

- GIVEN legacy rows in `TbAplicacionesAperturas` carry `IDAplicacion`, `NombreUsuario`, `FechaApertura`, `EnOficina`
- WHEN the seed runs
- THEN each row produces one event of type `app.open` and the `EnOficina` flag is dropped

### Requirement: Audit emission is mandatory

The system SHALL emit audit events as a contractual side effect of every auth or app-open mutation. The audit write SHALL be attempted in the same transaction; the mutation SHALL roll back if the audit write fails.

#### Scenario: Auth commit rolls back if audit is not written

- GIVEN a login mutation and the audit write happen in the same transaction
- WHEN the audit write fails
- THEN the login mutation is rolled back and the caller sees a typed error

### Requirement: Retention levels

The system SHALL support configurable retention. The MVP SHALL keep 90 days hot and 1 year total. The retention mechanism SHALL be implemented behind a port to allow archive to object storage.

#### Scenario: Retention policy is configurable

- GIVEN the audit retention service is initialised
- WHEN the configuration is read
- THEN the system applies the configured hot and total windows
- AND any event older than the total window is purged

### Requirement: Sync to NAS office

The system SHALL expose a manual sync procedure that emits the audit subset required by office-only apps. The sync SHALL NOT be automatic; it SHALL be triggered by the documented runbook.

#### Scenario: Manual sync runs on demand

- GIVEN a global admin invokes the sync procedure
- WHEN the procedure runs
- THEN the audit subset is exported, an audit record of type `audit.sync.office` is appended, and no automatic scheduling is configured

## Cross-references

D11, D13, D27, D28, D29, D55, D56, D112, H4. ABIERTO: periodos definitivos de retención por cumplimiento normativo (P10); MVP usa 90 d hot + 1 año total.
