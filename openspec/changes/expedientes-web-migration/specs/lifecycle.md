# Especificación delta — lifecycle

Contrato de capacidad agnóstico de plataforma; no replica pantallas ni mecanismos Access.

## REQUISITOS AÑADIDOS

### Requirement: EXP-CAP-001 — Alta transaccional
El sistema MUST crear cabecera e hijos; origen manual o HPS autorizado. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN actor autenticado con permiso de escritura y datos de expediente completos
- WHEN registra un alta manual o autorizada por HPS
- THEN se persisten cabecera, hijos, read-model y último cambio en una transacción confirmada, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita alta transaccional
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante alta transaccional
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-002 — Edición concurrente
El sistema MUST distinguir null de vacío y conservar invariantes. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN agregado existente, versión vigente y campos editables válidos
- WHEN guarda una edición del agregado
- THEN se actualizan los campos manteniendo invariantes y control de concurrencia, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita edición concurrente
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante edición concurrente
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-003 — Baja condicionada
El sistema MUST impedir pérdida de hijos, relaciones y auditoría. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN expediente con relaciones permitidas para eliminación y actor autorizado
- WHEN solicita la baja condicionada
- THEN se elimina el agregado de forma auditable sin afectar relaciones no incluidas, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita baja condicionada
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante baja condicionada
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-004 — Cambio de tipo
El sistema MUST validar elegibilidad y padre para AM/Lote/Basado. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN expediente elegible y padre válido cuando el tipo lo requiere
- WHEN cambia el tipo de expediente
- THEN se aplica el nuevo tipo y sus efectos derivados de forma consistente, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita cambio de tipo
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante cambio de tipo
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-005 — Estado y garantía
El sistema MUST calcular estados y garantía desde fechas, flags y motivo; conservar fuente y calculado. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN fechas, flags y estado manual coherentes
- WHEN consulta o recalcula estado y garantía
- THEN se devuelve estado calculado, fechas, garantía y motivo junto con sus valores fuente, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita estado y garantía
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante estado y garantía
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-006 — Jerarquía acuerdos-lotes-basados
El sistema MUST mantener padre válido e integridad autorreferente. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN acuerdo marco o lote padre válido y sin ciclos
- WHEN crea o consulta la jerarquía de acuerdos, lotes y basados
- THEN se mantiene la relación padre-hijo y la integridad de la jerarquía, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita jerarquía acuerdos-lotes-basados
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante jerarquía acuerdos-lotes-basados
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-007 — Ordinal funcional
El sistema MUST derivar ordinal estable dentro de jerarquía. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN jerarquía válida con identificadores estables
- WHEN deriva el ordinal funcional
- THEN se asigna un ordinal único y reproducible dentro de la jerarquía, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita ordinal funcional
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante ordinal funcional
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.
