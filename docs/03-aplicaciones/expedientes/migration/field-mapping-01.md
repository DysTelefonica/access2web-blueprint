# Field Mapping Tranche 01 — Tablas 01–15

Mapeo campo-a-campo de las primeras 15 tablas del backend legacy `Expedientes_datos.accdb` hacia las capacidades del spec `expedientes-web-migration`.

**Fuente legacy**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (extraído con Jackcess)
**Spec destino**: `openspec/changes/expedientes-web-migration/specs/`
**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9)

## Resumen de tablas

| # | Tabla legacy | Tipo | Capacidad destino | Estado |
|---|---|---|---|---|
| 01 | Copia de TbExpedientes | Histórico/backup | No migrar (backup legacy) | Rechazar |
| 02 | Copia de TbExpedientesConEntidades | Histórico/backup | No migrar (backup legacy) | Rechazar |
| 03 | ListaPrevia | Histórico | No migrar (lista preliminar legacy) | Rechazar |
| 04 | TbAusExpPostAGEDO | Auxiliar AGEDO | EXP-CAP-051 (AGEDYS) | Cuarentena |
| 05 | TbAuxEstadosMartina | Auxiliar | No migrar (auxiliar legacy) | Rechazar |
| 06 | TbAuxNemotecnico | Auxiliar | EXP-CAP-007 (Ordinal) | Transformar |
| 07 | TbCambios | Auditoría | EXP-CAP-046 (Auditoría) | Migrar |
| 08 | TbComerciales | Catálogo | EXP-CAP-015 (Comerciales) | Migrar |
| 09 | TbConfMostrarEstado | Config | EXP-CAP-005 (Estado) | Transformar |
| 10 | TbCPV | Catálogo | EXP-CAP-016 (CPV) | Migrar |
| 11 | TbDatosEconomicosExpedientes | Datos expediente | EXP-CAP-001..003 (Ciclo) | Migrar |
| 12 | TbE2EExportBatch | E2E | EXP-CAP-035 (Batch E2E) | Migrar |
| 13 | TbE2EExportBatchDetalle | E2E | EXP-CAP-035 (Batch E2E) | Migrar |
| 14 | TbE2EExportSeleccionTemp | E2E | EXP-CAP-039 (Selección manual) | Migrar |
| 15 | TbE2EJsonDestinationUserConfig | E2E | EXP-CAP-040 (Destino por usuario) | Migrar |

## Detalle por tabla

### 01. Copia de TbExpedientes → RECHAZAR

**Razón**: Tabla de backup legacy (`Copia de`). No contiene datos activos. El diseño D-EXP-9 establece que el legacy permanece intacto durante la transición; las copias de seguridad no se migran.

**Campos**: 55 columnas (idénticas a TbExpedientes)
**Acción**: No migrar. Documentar como backup histórico.

### 02. Copia de TbExpedientesConEntidades → RECHAZAR

**Razón**: Tabla de backup legacy (`Copia de`). Mismo criterio que #01.

**Campos**: 23 columnas (idénticas a TbExpedientesConEntidades)
**Acción**: No migrar. Documentar como backup histórico.

### 03. ListaPrevia → RECHAZAR

**Razón**: Tabla de lista preliminar legacy. No tiene capacidad asociada en el spec. Probablemente usada para reportes internos legacy.

**Campos**: TBD (verificar en schema.sql)
**Acción**: No migrar. Documentar como tabla legacy sin capacidad asociada.

### 04. TbAusExpPostAGEDO → CUARENTENA

**Razón**: Tabla auxiliar de AGEDO (sistema de gestión de expedientes antiguo). Relacionada con EXP-CAP-051 (integración AGEDYS). Requiere análisis adicional para determinar si los datos son relevantes para la migración.

**Campos**: TBD (verificar en schema.sql)
**Acción**: Cuarentena. Analizar contenido antes de decidir migración.

### 05. TbAuxEstadosMartina → RECHAZAR

**Razón**: Tabla auxiliar legacy (`TbAux`). No tiene capacidad asociada en el spec. Probablemente usada para cálculos internos legacy.

**Campos**: TBD (verificar en schema.sql)
**Acción**: No migrar. Documentar como tabla auxiliar legacy.

### 06. TbAuxNemotecnico → TRANSFORMAR

**Razón**: Tabla auxiliar de nemotécnicos (códigos cortos de identificación). Relacionada con EXP-CAP-007 (Ordinal funcional). El campo `Nemotecnico` en TbExpedientes es un identificador corto usado para búsqueda rápida.

**Campos**: TBD (verificar en schema.sql)
**Acción**: Transformar. El nemotécnico se mantiene como campo en la tabla `expedientes` del nuevo schema.

### 07. TbCambios → MIGRAR

**Razón**: Tabla de auditoría de cambios. Relacionada con EXP-CAP-046 (Auditoría). Contiene el historial de cambios de cada expediente.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `FechaCambio` (SHORT_DATE_TIME) → `audit_log.created_at`
- `IDUsuarioCambio` (TEXT) → `audit_log.user_id` (transformar a FK numérica)
- `CampoCambiado` (TEXT) → `audit_log.field_name`
- `ValorAnterior` (MEMO) → `audit_log.old_value`
- `ValorNuevo` (MEMO) → `audit_log.new_value`

**Acción**: Migrar a `audit_log` del nuevo schema. Transformar `IDUsuarioCambio` de TEXT a FK numérica.

### 08. TbComerciales → MIGRAR

**Razón**: Catálogo de comerciales. Relacionada con EXP-CAP-015 (Comerciales). Tabla maestra simple.

**Campos clave**:
- `IDComercial` (LONG) → `comerciales.id`
- `Comercial` (TEXT) → `comerciales.nombre`
- `Descripcion` (MEMO) → `comerciales.descripcion`

**Acción**: Migrar a `comerciales` del nuevo schema. Sin transformaciones.

### 09. TbConfMostrarEstado → TRANSFORMAR

**Razón**: Configuración de visualización de estados. Relacionada con EXP-CAP-005 (Estado y garantía). Contiene reglas de visualización de estados en la UI.

**Campos**: TBD (verificar en schema.sql)
**Acción**: Transformar. Las reglas de visualización se mueven a la capa de presentación (Jinja2 templates) del nuevo schema.

### 10. TbCPV → MIGRAR

**Razón**: Catálogo CPV (Common Procurement Vocabulary). Relacionada con EXP-CAP-016 (CPV). Tabla maestra simple.

**Campos clave**:
- `IDCPV` (LONG) → `cpv.id`
- `CPV` (TEXT) → `cpv.codigo`
- `Descripcion` (MEMO) → `cpv.descripcion`

**Acción**: Migrar a `cpv` del nuevo schema. Sin transformaciones.

### 11. TbDatosEconomicosExpedientes → MIGRAR

**Razón**: Datos económicos de expedientes. Relacionada con EXP-CAP-001..003 (Ciclo de vida). Contiene importes y datos financieros.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `ImporteLicitacion` (DOUBLE) → `expedientes.importe_licitacion`
- `ImporteContratacion` (DOUBLE) → `expedientes.importe_contratacion`

**Acción**: Migrar. Los importes se consolidan en la tabla `expedientes` del nuevo schema.

### 12. TbE2EExportBatch → MIGRAR

**Razón**: Batch de exportación E2E. Relacionada con EXP-CAP-035 (Batch E2E). Contiene los batches de exportación.

**Campos clave**:
- `IDBatch` (LONG) → `e2e_batches.id`
- `FechaCreacion` (SHORT_DATE_TIME) → `e2e_batches.created_at`
- `Estado` (TEXT) → `e2e_batches.status`

**Acción**: Migrar a `e2e_batches` del nuevo schema.

### 13. TbE2EExportBatchDetalle → MIGRAR

**Razón**: Detalle de batch de exportación E2E. Relacionada con EXP-CAP-035 (Batch E2E). Contiene los expedientes incluidos en cada batch.

**Campos clave**:
- `IDBatch` (LONG) → FK a `e2e_batches.id`
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `OrdinalE2E` (TEXT) → `e2e_batch_items.ordinal`

**Acción**: Migrar a `e2e_batch_items` del nuevo schema.

### 14. TbE2EExportSeleccionTemp → MIGRAR

**Razón**: Selección temporal de exportación E2E. Relacionada con EXP-CAP-039 (Selección manual). Contiene la selección temporal de expedientes para exportación.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDUsuario` (TEXT) → `e2e_selections.user_id` (transformar a FK numérica)
- `FechaSeleccion` (SHORT_DATE_TIME) → `e2e_selections.selected_at`

**Acción**: Migrar a `e2e_selections` del nuevo schema. Transformar `IDUsuario` de TEXT a FK numérica.

### 15. TbE2EJsonDestinationUserConfig → MIGRAR

**Razón**: Configuración de destino JSON por usuario. Relacionada con EXP-CAP-040 (Destino por usuario). Contiene la configuración de destino de exportación por usuario.

**Campos clave**:
- `IDUsuario` (TEXT) → `e2e_user_configs.user_id` (transformar a FK numérica)
- `DestinoJson` (MEMO) → `e2e_user_configs.destination_path`
- `Activo` (TEXT) → `e2e_user_configs.active` (transformar a BOOLEAN)

**Acción**: Migrar a `e2e_user_configs` del nuevo schema. Transformar `IDUsuario` de TEXT a FK numérica y `Activo` de TEXT a BOOLEAN.

## Transformaciones globales

| Transformación | Aplica a | Razón |
|---|---|---|
| `IDUsuario*` TEXT → FK numérica | TbCambios, TbE2EExportSeleccionTemp, TbE2EJsonDestinationUserConfig | El legacy usa TEXT para IDs de usuario; el nuevo schema usa FK numéricas |
| `Activo` TEXT(2) → BOOLEAN | TbE2EJsonDestinationUserConfig | D102: booleanos Text(2) → BOOLEAN |
| `EsAM/EsLote/EsBasado` TEXT(2) → BOOLEAN | TbExpedientes (tabla principal) | D102: booleanos Text(2) → BOOLEAN |
| Nemotécnico → campo en `expedientes` | TbAuxNemotecnico | El nemotécnico se consolida en la tabla principal |

## Cuarentena

| Tabla | Razón | Acción requerida |
|---|---|---|
| TbAusExpPostAGEDO | Auxiliar AGEDO, contenido desconocido | Analizar contenido antes de decidir migración |

## Rechazos

| Tabla | Razón |
|---|---|
| Copia de TbExpedientes | Backup legacy, no migrar |
| Copia de TbExpedientesConEntidades | Backup legacy, no migrar |
| ListaPrevia | Lista preliminar legacy, sin capacidad asociada |
| TbAuxEstadosMartina | Auxiliar legacy, sin capacidad asociada |
