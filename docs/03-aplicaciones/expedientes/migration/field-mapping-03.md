# Field Mapping Tranche 03 — Tablas 31–49 + Rechazos

Mapeo campo-a-campo de las tablas 31-49 del backend legacy `Expedientes_datos.accdb` hacia las capacidades del spec `expedientes-web-migration`. Este tranche completa el mapping de las 49 tablas locales del backend.

**Fuente legacy**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (extraído con Jackcess)
**Spec destino**: `openspec/changes/expedientes-web-migration/specs/`
**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9)

## Resumen de tablas

| # | Tabla legacy | Tipo | Capacidad destino | Estado |
|---|---|---|---|---|
| 31 | TbExpedientesJefaturas | Relación | EXP-CAP-012 (Responsables) | Migrar |
| 32 | TbExpedientesJuridicas | Relación | EXP-CAP-013 (Entidades) | Migrar |
| 33 | TbExpedientesLugaresEjecucion | Relación | EXP-CAP-019 (Lugares) | Migrar |
| 34 | TbExpedientesModificados | Modificados | EXP-CAP-009 (Modificados) | Migrar |
| 35 | TbExpedientesPECAL | Relación | EXP-CAP-020 (PECAL) | Migrar |
| 36 | TbExpedientesRACS | Relación | EXP-CAP-021 (RAC) | Migrar |
| 37 | TbExpedientesResponsables | Relación | EXP-CAP-012 (Responsables) | Migrar |
| 38 | TbExpedientesSuministradores | Relación | EXP-CAP-014 (Suministradores) | Migrar |
| 39 | TbGradosClasificacion | Catálogo | EXP-CAP-022 (Grados) | Migrar |
| 40 | TbJefaturas | Catálogo | EXP-CAP-012 (Responsables) | Migrar |
| 41 | TbJuridicas | Catálogo | EXP-CAP-013 (Entidades) | Migrar |
| 42 | TbLugaresEjecucion | Catálogo | EXP-CAP-019 (Lugares) | Migrar |
| 43 | TbOficinasPrograma | Catálogo | EXP-CAP-024 (Oficinas) | Migrar |
| 44 | TbOrganosContratacion | Catálogo | EXP-CAP-023 (Órganos) | Migrar |
| 45 | TbPECAL | Catálogo | EXP-CAP-020 (PECAL) | Migrar |
| 46 | TbRACS | Catálogo | EXP-CAP-021 (RAC) | Migrar |
| 47 | TbResponsablesPorRol | Catálogo | EXP-CAP-012 (Responsables) | Migrar |
| 48 | TbSuministradores | Catálogo | EXP-CAP-014 (Suministradores) | Migrar |
| 49 | TbUltimoCambio | Auditoría | EXP-CAP-046 (Auditoría) | Migrar |

## Detalle por tabla

### 31. TbExpedientesJefaturas → MIGRAR

**Razón**: Relación expediente-jefatura. Relacionada con EXP-CAP-012 (Responsables). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDJefatura` (LONG) → FK a `jefaturas.id`

**Acción**: Migrar a `expediente_jefaturas` del nuevo schema. Sin transformaciones.

### 32. TbExpedientesJuridicas → MIGRAR

**Razón**: Relación expediente-jurídica. Relacionada con EXP-CAP-013 (Entidades y jurídicas). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDJuridica` (LONG) → FK a `juridicas.id`

**Acción**: Migrar a `expediente_juridicas` del nuevo schema. Sin transformaciones.

### 33. TbExpedientesLugaresEjecucion → MIGRAR

**Razón**: Relación expediente-lugar de ejecución. Relacionada con EXP-CAP-019 (Lugares de ejecución). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDLugarEjecucion` (LONG) → FK a `lugares_ejecucion.id`

**Acción**: Migrar a `expediente_lugares_ejecucion` del nuevo schema. Sin transformaciones.

### 34. TbExpedientesModificados → MIGRAR

**Razón**: Modificados de expedientes. Relacionada con EXP-CAP-009 (Modificados e historial). Contiene el historial de modificaciones.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDModificado` (LONG) → `expediente_modificados.id`
- `FechaModificado` (SHORT_DATE_TIME) → `expediente_modificados.fecha`
- `Descripcion` (MEMO) → `expediente_modificados.descripcion`

**Acción**: Migrar a `expediente_modificados` del nuevo schema. Sin transformaciones.

### 35. TbExpedientesPECAL → MIGRAR

**Razón**: Relación expediente-PECAL. Relacionada con EXP-CAP-020 (PECAL). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDPECAL` (LONG) → FK a `pecal.id`

**Acción**: Migrar a `expediente_pecal` del nuevo schema. Sin transformaciones.

### 36. TbExpedientesRACS → MIGRAR

**Razón**: Relación expediente-RAC. Relacionada con EXP-CAP-021 (RAC). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDRAC` (LONG) → FK a `racs.id`

**Acción**: Migrar a `expediente_racs` del nuevo schema. Sin transformaciones.

### 37. TbExpedientesResponsables → MIGRAR

**Razón**: Relación expediente-responsable. Relacionada con EXP-CAP-012 (Responsables). Tabla de relación many-to-many con rol.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDResponsable` (LONG) → FK a `responsables.id`
- `IDRol` (LONG) → FK a `roles.id`

**Acción**: Migrar a `expediente_responsables` del nuevo schema. Sin transformaciones.

### 38. TbExpedientesSuministradores → MIGRAR

**Razón**: Relación expediente-suministrador. Relacionada con EXP-CAP-014 (Suministradores y UTE). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDSuministrador` (LONG) → FK a `suministradores.id`
- `Tipo` (TEXT) → `expediente_suministradores.tipo` (contratista/subcontratista/UTE)

**Acción**: Migrar a `expediente_suministradores` del nuevo schema. Sin transformaciones.

### 39. TbGradosClasificacion → MIGRAR

**Razón**: Catálogo de grados de clasificación. Relacionada con EXP-CAP-022 (Grados de clasificación). Tabla maestra simple.

**Campos clave**:
- `IDGradoClasificacion` (LONG) → `grados_clasificacion.id`
- `Grado` (TEXT) → `grados_clasificacion.nombre`
- `Descripcion` (MEMO) → `grados_clasificacion.descripcion`

**Acción**: Migrar a `grados_clasificacion` del nuevo schema. Sin transformaciones.

### 40. TbJefaturas → MIGRAR

**Razón**: Catálogo de jefaturas. Relacionada con EXP-CAP-012 (Responsables). Tabla maestra simple.

**Campos clave**:
- `IDJefatura` (LONG) → `jefaturas.id`
- `Jefatura` (TEXT) → `jefaturas.nombre`
- `Descripcion` (MEMO) → `jefaturas.descripcion`

**Acción**: Migrar a `jefaturas` del nuevo schema. Sin transformaciones.

### 41. TbJuridicas → MIGRAR

**Razón**: Catálogo de jurídicas. Relacionada con EXP-CAP-013 (Entidades y jurídicas). Tabla maestra simple.

**Campos clave**:
- `IDJuridica` (LONG) → `juridicas.id`
- `Juridica` (TEXT) → `juridicas.nombre`
- `Descripcion` (MEMO) → `juridicas.descripcion`

**Acción**: Migrar a `juridicas` del nuevo schema. Sin transformaciones.

### 42. TbLugaresEjecucion → MIGRAR

**Razón**: Catálogo de lugares de ejecución. Relacionada con EXP-CAP-019 (Lugares de ejecución). Tabla maestra simple.

**Campos clave**:
- `IDLugarEjecucion` (LONG) → `lugares_ejecucion.id`
- `Lugar` (TEXT) → `lugares_ejecucion.nombre`
- `Descripcion` (MEMO) → `lugares_ejecucion.descripcion`

**Acción**: Migrar a `lugares_ejecucion` del nuevo schema. Sin transformaciones.

### 43. TbOficinasPrograma → MIGRAR

**Razón**: Catálogo de oficinas de programa. Relacionada con EXP-CAP-024 (Oficinas de programa). Tabla maestra simple.

**Campos clave**:
- `IDOficinaPrograma` (LONG) → `oficinas_programa.id`
- `Oficina` (TEXT) → `oficinas_programa.nombre`
- `Descripcion` (MEMO) → `oficinas_programa.descripcion`

**Acción**: Migrar a `oficinas_programa` del nuevo schema. Sin transformaciones.

### 44. TbOrganosContratacion → MIGRAR

**Razón**: Catálogo de órganos de contratación. Relacionada con EXP-CAP-023 (Órganos de contratación). Tabla maestra simple.

**Campos clave**:
- `IDOrganoContratacion` (LONG) → `organos_contratacion.id`
- `Organo` (TEXT) → `organos_contratacion.nombre`
- `Descripcion` (MEMO) → `organos_contratacion.descripcion`

**Acción**: Migrar a `organos_contratacion` del nuevo schema. Sin transformaciones.

### 45. TbPECAL → MIGRAR

**Razón**: Catálogo PECAL. Relacionada con EXP-CAP-020 (PECAL). Tabla maestra simple.

**Campos clave**:
- `IDPECAL` (LONG) → `pecal.id`
- `PECAL` (TEXT) → `pecal.nombre`
- `Descripcion` (MEMO) → `pecal.descripcion`

**Acción**: Migrar a `pecal` del nuevo schema. Sin transformaciones.

### 46. TbRACS → MIGRAR

**Razón**: Catálogo RACS. Relacionada con EXP-CAP-021 (RAC). Tabla maestra simple.

**Campos clave**:
- `IDRAC` (LONG) → `racs.id`
- `RAC` (TEXT) → `racs.nombre`
- `Descripcion` (MEMO) → `racs.descripcion`

**Acción**: Migrar a `racs` del nuevo schema. Sin transformaciones.

### 47. TbResponsablesPorRol → MIGRAR

**Razón**: Catálogo de responsables por rol. Relacionada con EXP-CAP-012 (Responsables). Tabla maestra que define los roles de responsable.

**Campos clave**:
- `IDRol` (LONG) → `roles.id`
- `Rol` (TEXT) → `roles.nombre`
- `Descripcion` (MEMO) → `roles.descripcion`

**Acción**: Migrar a `roles` del nuevo schema. Sin transformaciones.

### 48. TbSuministradores → MIGRAR

**Razón**: Catálogo de suministradores. Relacionada con EXP-CAP-014 (Suministradores y UTE). Tabla maestra simple.

**Campos clave**:
- `IDSuministrador` (LONG) → `suministradores.id`
- `Suministrador` (TEXT) → `suministradores.nombre`
- `Descripcion` (MEMO) → `suministradores.descripcion`

**Acción**: Migrar a `suministradores` del nuevo schema. Sin transformaciones.

### 49. TbUltimoCambio → MIGRAR

**Razón**: Último cambio de cada expediente. Relacionada con EXP-CAP-046 (Auditoría). Contiene el último cambio registrado.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `FechaCambio` (SHORT_DATE_TIME) → `expediente_ultimo_cambio.fecha`
- `IDUsuarioCambio` (TEXT) → `expediente_ultimo_cambio.usuario` (transformar a FK numérica)
- `CampoCambiado` (TEXT) → `expediente_ultimo_cambio.campo`

**Acción**: Migrar a `expediente_ultimo_cambio` del nuevo schema. Transformar `IDUsuarioCambio` de TEXT a FK numérica.

## Transformaciones globales

| Transformación | Aplica a | Razón |
|---|---|---|
| `IDUsuarioCambio` TEXT → FK numérica | TbUltimoCambio | El legacy usa TEXT para IDs de usuario; el nuevo schema usa FK numéricas |

## Resumen completo del mapping (49 tablas)

| Tranche | Tablas | Migrar | Transformar | Cuarentena | Rechazar |
|---|---|---|---|---|---|
| 01 (01-15) | 15 | 8 | 2 | 1 | 4 |
| 02 (16-30) | 15 | 12 | 0 | 2 | 1 |
| 03 (31-49) | 19 | 19 | 0 | 0 | 0 |
| **Total** | **49** | **39** | **2** | **3** | **5** |

## Rechazos acumulados (todos los tranches)

| Tabla | Tranche | Razón |
|---|---|---|
| Copia de TbExpedientes | 01 | Backup legacy |
| Copia de TbExpedientesConEntidades | 01 | Backup legacy |
| ListaPrevia | 01 | Lista preliminar legacy |
| TbAuxEstadosMartina | 01 | Auxiliar legacy |
| TbExpedientes_antes | 02 | Histórico legacy |

## Cuarentena acumulada (todos los tranches)

| Tabla | Tranche | Razón | Acción requerida |
|---|---|---|---|
| TbAusExpPostAGEDO | 01 | Auxiliar AGEDO | Analizar contenido |
| TbExpAgedys | 02 | Integración AGEDYS | Analizar contenido |
| TbExpAGEDYS1 | 02 | Integración AGEDYS (variante) | Analizar contenido |
