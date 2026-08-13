# Especificación delta — query-and-tasks

Contrato de capacidad agnóstico de plataforma; no replica pantallas ni mecanismos Access.

## REQUISITOS AÑADIDOS

### Requirement: EXP-CAP-025 — Bandeja paginada
El sistema MUST filtrar por estado/código/JP/jurídica/suministrador y mostrar detalle solo lectura. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN actor autorizado y filtros válidos de estado, código o responsable
- WHEN consulta la bandeja paginada
- THEN se devuelve una página estable con detalle de solo lectura, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita bandeja paginada
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante bandeja paginada
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-026 — Búsqueda avanzada
El sistema MUST preservar semántica combinada de filtros y resultados deterministas. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN actor autorizado y combinación válida de filtros avanzados
- WHEN ejecuta una búsqueda avanzada
- THEN se devuelven resultados deterministas con la semántica de filtros preservada, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita búsqueda avanzada
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante búsqueda avanzada
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-027 — Búsqueda técnica
El sistema MUST permitir lectura técnica solo con permiso efectivo. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN actor con permiso técnico efectivo y criterios válidos
- WHEN ejecuta una búsqueda técnica
- THEN se devuelve información técnica en modo solo lectura, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita búsqueda técnica
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante búsqueda técnica
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-028 — Exportación Excel
El sistema MUST generar exportación reproducible autorizada con filtros aplicados. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN actor autorizado, filtros válidos y formato Excel solicitado
- WHEN exporta los resultados
- THEN se genera una exportación Excel reproducible con los filtros aplicados, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita exportación excel
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante exportación excel
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-029 — Tareas calculadas
El sistema MUST calcular seis worklists: desconocido, recepción/hito, adjudicado sin contrato, TSOL sin S4H y oferta prolongada. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN expedientes con reglas de tarea evaluables
- WHEN calcula las worklists operativas
- THEN se generan las seis categorías de tareas y sus contadores, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita tareas calculadas
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante tareas calculadas
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.
