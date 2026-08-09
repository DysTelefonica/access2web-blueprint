# Lanzadera · Auth bootstrap specification

> Capability del change `lanzadera-mvp`. Aprovisionamiento del primer global admin vía CLI exclusivo (D25). Antes de que exista un global admin, ningún reset por email es posible (ver `auth-reset`).

## Purpose

Esta specification describe el bootstrap del primer global admin: comando CLI `gentle-ai platform user set-password <email>`. Es el único camino para crear el primer admin global; el reset por email queda bloqueado hasta que este exista (ver `auth-reset`).

Decisiones de origen: D25, D91.

## Requirements

### Requirement: First global admin bootstrap via CLI

The system SHALL expose `gentle-ai platform user set-password <email>` as the exclusive path to provision the first global admin. The platform SHALL refuse to issue reset tokens by email until at least one global admin exists.

#### Scenario: CLI sets a global admin password

- GIVEN the platform is freshly installed and no global admin exists
- WHEN the CLI is invoked with `set-password <email>` and a new password
- THEN the user is created or upgraded to global admin
- AND `password_hash = hash_password(new_password)` (Argon2id)
- AND the user's `status = 'active'`

#### Scenario: CLI is the exclusive path before any global admin exists

- GIVEN no global admin exists
- WHEN any other path (UI form, REST endpoint, scheduled job) attempts to create a global admin
- THEN the system rejects the attempt and logs a security audit event

## Cross-references

D25, D91.

## Decisiones pendientes

`##ABIERTO##` Política de auditoría para intentos fallidos de crear admin global (qué severidad, qué notificación a SOC). Diferido del MVP; suficiente con que el evento quede en el log canónico de auditoría.
