# Especificación delta — related-data

Contrato de capacidad agnóstico de plataforma; no replica pantallas ni mecanismos Access.

## REQUISITOS AÑADIDOS

### Requirement: EXP-CAP-008 — Hitos temporales
El sistema MUST validar fecha de hito, garantía e importe. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN expediente existente, fecha de hito e importe válidos
- WHEN registra un hito
- THEN se guarda descripción, fecha, garantía e importe vinculados al expediente, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita hitos temporales
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante hitos temporales
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-009 — Modificados e historial
El sistema MUST registrar número, fechas y cambios de campo/último cambio. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN modificado con número y fechas coherentes y actor autorizado
- WHEN registra un modificado o cambio de campo
- THEN se conserva el historial reciente y el último cambio con usuario y fecha, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita modificados e historial
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante modificados e historial
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-010 — Anexos y referencias
El sistema MUST gestionar referencia opaca y eliminación condicionada por fichero/registro. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN referencia documental opaca autorizada y dentro del límite acordado
- WHEN crea o elimina una referencia de anexo
- THEN se registra la referencia y su retención sin exponer rutas locales, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita anexos y referencias
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante anexos y referencias
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-011 — Anualidades
El sistema MUST conservar año, periodo e importes IVA/IPSI/IGIC/exento. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN expediente existente, año y desglose fiscal monetario
- WHEN guarda una anualidad
- THEN se conservan año, periodo e importes IVA, IPSI, IGIC y exento, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita anualidades
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante anualidades
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-012 — Responsables
El sistema MUST asignar usuario por rol, JP y flags. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN usuario vigente y rol permitido para el expediente
- WHEN asigna un responsable
- THEN se registra la relación por rol, incluyendo JP y flags aplicables, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita responsables
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante responsables
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-013 — Entidades y jurídicas
El sistema MUST mantener relaciones y cadena de entidades. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN entidades existentes y cadena acíclica
- WHEN relaciona entidades y jurídicas
- THEN se conservan las relaciones y la cadena sin huérfanos, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita entidades y jurídicas
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante entidades y jurídicas
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-014 — Suministradores y UTE
El sistema MUST aplicar jerarquía «el árbol manda», contratista/subcontratista y UTE. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN suministradores existentes y árbol acíclico
- WHEN asigna contratista, subcontratista, UTE o jerarquía
- THEN se aplica la regla «el árbol manda» y se conserva el árbol, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita suministradores y ute
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante suministradores y ute
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.
