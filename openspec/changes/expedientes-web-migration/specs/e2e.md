# Especificación delta — e2e

Contrato de capacidad agnóstico de plataforma; no replica pantallas ni mecanismos Access.

## REQUISITOS AÑADIDOS

### Requirement: EXP-CAP-033 — DTO estable
El sistema MUST intercambiar DTO de dominio versionado sin IDs técnicos. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN DTO de dominio conforme al contrato versionado
- WHEN intercambia el DTO
- THEN se acepta y reconstruye el agregado con los campos de dominio estables, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita dto estable
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante dto estable
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-034 — JSON determinista
El sistema MUST emitir `{meta,data}`, apiVersion 1.0, nueve colecciones, ISO-Z/null y orden OrdinalE2E+ID. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN nueve colecciones completas, fechas ISO-Z y null explícito
- WHEN genera el JSON canónico
- THEN se obtiene `{meta,data}`, apiVersion 1.0 y orden OrdinalE2E+ID reproducible, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita json determinista
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante json determinista
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-035 — Batch E2E
El sistema MUST gestionar ciclo de batch y detalle con atomicidad explícita. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN selección manual válida y sesión E2E del usuario
- WHEN crea y ejecuta un batch
- THEN se registra ciclo de vida, detalle y resultado según la política de atomicidad aprobada, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita batch e2e
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante batch e2e
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-036 — Hash versionado
El sistema MUST aplicar estrategia de hash versionada y canonicalización. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN payload canónico y versión de hash aprobada para el cutover
- WHEN calcula el hash de contenido
- THEN se persiste hash y versión de canonicalización de forma reproducible, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita hash versionado
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante hash versionado
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-037 — Paquete E2E
El sistema MUST generar paquete completo con manifest y trazabilidad. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN expedientes seleccionados y dependencias disponibles
- WHEN genera el paquete E2E
- THEN se publica paquete completo con manifest y referencias trazables, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita paquete e2e
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante paquete e2e
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-038 — Trazabilidad E2E
El sistema MUST relacionar expediente, batch, artefacto y resultado. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN expediente, batch y artefacto existentes
- WHEN consulta la trazabilidad
- THEN se devuelve la relación completa entre expediente, batch, artefacto y resultado, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita trazabilidad e2e
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante trazabilidad e2e
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-039 — Selección manual
El sistema MUST aislar selección temporal por sesión y usuario. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN usuario autenticado y expedientes seleccionables de su ámbito
- WHEN selecciona expedientes manualmente
- THEN se guarda la selección en la sesión aislada del usuario, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita selección manual
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante selección manual
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-040 — Destino por usuario
El sistema MUST resolver destino por usuario, nunca ruta local contractual. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN usuario autenticado y destino configurado por su cuenta
- WHEN resuelve el destino de exportación
- THEN se utiliza el destino del usuario sin convertir rutas locales en contrato, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita destino por usuario
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante destino por usuario
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-041 — Sesión E2E
El sistema MUST reanudar/cerrar sesión sin mezclar usuarios. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN sesión E2E activa del usuario
- WHEN reanuda o cierra la sesión
- THEN se conserva el progreso y la auditoría sin mezclar usuarios, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita sesión e2e
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante sesión e2e
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-042 — Ordinal E2E
El sistema MUST mantener ordinal único y consistente dominio/trazabilidad. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN familia acíclica y datos con ordinales estables
- WHEN expande la familia y calcula el ordinal E2E
- THEN se obtiene un ordinal único y consistente en dominio y trazabilidad, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita ordinal e2e
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante ordinal e2e
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

## Política de hash y cutover

La estrategia de hash MUST estar versionada y su canonicalización documentada. La continuidad FNV queda OPEN QUESTION hasta disponer de golden evidence. Antes del cutover debe aprobarse una de estas políticas: (A) continuidad FNV-1a demostrada contra golden; o (B) nuevo algoritmo/versionado con reconciliación y aceptación explícitas.
