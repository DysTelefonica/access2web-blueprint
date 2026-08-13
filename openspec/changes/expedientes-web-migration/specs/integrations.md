# Especificación delta — integrations

Contrato de capacidad agnóstico de plataforma; no replica pantallas ni mecanismos Access.

## REQUISITOS AÑADIDOS

### Requirement: EXP-CAP-051 — HPS
El sistema MUST aceptar alta/consulta autenticada e idempotente. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN credencial y contrato HPS válidos con clave idempotente
- WHEN solicita alta o consulta HPS
- THEN se crea o consulta el expediente sin duplicación, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita hps
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante hps
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-052 — AGEDYS
El sistema MUST intercambiar datos con ownership y errores explícitos. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN contrato AGEDYS vigente y ownership aprobado
- WHEN intercambia datos con AGEDYS
- THEN se confirma el intercambio y se registra ownership y resultado, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita agedys
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante agedys
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-053 — Gestión de Riesgos
El sistema MUST enlazar por identificador estable. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN identificador estable de expediente y servicio de Riesgos disponible
- WHEN consulta el enlace de Riesgos
- THEN se devuelve el proyecto de Riesgos asociado, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita gestión de riesgos
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante gestión de riesgos
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-054 — No Conformidades
El sistema MUST consultar estado NC por expediente/código S4H. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN código de expediente o S4H y servicio de NC disponible
- WHEN consulta No Conformidades
- THEN se devuelve el estado NC asociado al expediente, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita no conformidades
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante no conformidades
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-055 — Correo
El sistema MUST emitir notificación mediante puerto y trazabilidad. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN puerto de correo disponible y evento auditable
- WHEN emite una notificación
- THEN se envía el aviso y se registra la trazabilidad, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita correo
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante correo
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

### Requirement: EXP-CAP-056 — Documentos
El sistema MUST usar referencia opaca a SharePoint/object storage con retención y autorización. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

#### Scenario: camino feliz
- GIVEN referencia opaca autorizada y almacenamiento disponible
- WHEN consulta o vincula un documento
- THEN se devuelve el documento según retención y autorización, con salida determinista y auditable.
#### Scenario: validación y autorización
- GIVEN datos inválidos, permiso ausente o dependencia no resoluble
- WHEN solicita documentos
- THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.

#### Scenario: concurrencia o fallo
- GIVEN dos solicitudes concurrentes o un fallo de dependencia durante documentos
- WHEN se procesa la operación
- THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.
