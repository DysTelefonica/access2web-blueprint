# Especificación delta — write-resilience

Contrato de capacidad agnóstico de plataforma; no replica pantallas ni mecanismos Access.

## REQUISITOS AÑADIDOS

### Requirement: EXP-CAP-030 — Autosave generales y fechas
El sistema MUST guardar cambios generales y fechas con transacción. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN agregado existente, payload válido y transacción disponible
- WHEN ejecuta autosave de datos generales y fechas
- THEN se confirma una única transacción con los cambios guardados, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita autosave generales y fechas
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante autosave generales y fechas
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-031 — Autosave relacionados
El sistema MUST guardar hitos, modificados y secciones editables sin perder relaciones. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN secciones editables y relaciones válidas
- WHEN ejecuta autosave de hitos, modificados y restantes
- THEN se confirman cabecera y relacionados de forma atómica, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita autosave relacionados
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante autosave relacionados
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-032 — Idempotencia y feedback
El sistema MUST usar clave idempotente, bloquear doble envío y exponer progreso aria-busy/error. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN clave de idempotencia nueva y controles de interfaz activos
- WHEN envía una mutación editable
- THEN se procesa una sola mutación y se muestra progreso accesible mediante aria-busy, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita idempotencia y feedback
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante idempotencia y feedback
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.
