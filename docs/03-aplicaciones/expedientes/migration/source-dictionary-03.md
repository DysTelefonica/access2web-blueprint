# Source Dictionary Tranche 03 — Tablas 21–30

Diccionario físico de las tablas ordinales 21-30 del backend legacy `Expedientes_datos.accdb` hacia la migración web del Expediente. Todas las tablas de esta tranche tienen descripción detallada en `docs/03-aplicaciones/expedientes/data-model.md`.

**Fuente legacy**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (extraído con Jackcess el 2026-08-18, PR #393).
**Fuente de columnas**: `docs/03-aplicaciones/expedientes/data-model.md` (todas las tablas en esta tranche tienen §detallada).
**Mapeo a capacidades**: `docs/03-aplicaciones/expedientes/migration/field-mapping-03.md`.
**Spec destino**: `openspec/changes/expedientes-web-migration/specs/`.
**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9).
**Autoridad de ordinales**: las posiciones 21–30 siguen el ordinal de `data-model.md` (§Tabla de mapeo ordinal → nombre de tabla).

## Resumen de tablas

| # | Tabla legacy | Columnas | Filas | Capacidad destino | Estado |
|---|---|---:|---:|---|---|
| 21 | `TbExpedientesCPVs` | 3 | 429 | EXP-CAP-013 (CPVs por expediente) | Migrar |
| 22 | `TbExpedientesE2E` | 7 | — | EXP-CAP-050 (Sincronización E2E) | Migrar |
| 23 | `TbExpedientesHitos` | 6 | 46 | EXP-CAP-014 (Hitos) | Migrar |
| 24 | `TbExpedientesJefaturas` | 3 | — | EXP-CAP-006 (Jefaturas por expediente) | Migrar |
| 25 | `TbExpedientesJuridicas` | 6 | 417 | EXP-CAP-007 (Jurídicas por expediente) | Migrar |
| 26 | `TbExpedientesLugaresEjecucion` | 3 | 194 | EXP-CAP-006 (Lugares de ejecución por expediente) | Migrar |
| 27 | `TbExpedientesModificados` | 6 | 37 | EXP-CAP-046 (Auditoría / Modificaciones) | Migrar |
| 28 | `TbExpedientesPECAL` | 3 | 366 | EXP-CAP-018 (PECAL por expediente) | Migrar |
| 29 | `TbExpedientesRACS` | 3 | — | EXP-CAP-018 (RACS por expediente) | Migrar |
| 30 | `TbExpedientesResponsables` | 6 | — | EXP-CAP-019 (Responsables por expediente) | Migrar |

## Detalle por tabla

### 21. TbExpedientesCPVs → MIGRAR

**Razón**:join entre expedientes y códigos CPV. Mapea contra `EXP-CAP-013` (CPVs por expediente). Tabla dejoin con PK compuesta.

**Columnas** (de `data-model.md` §TbExpedientesCPVs):

PK compuesta: `IDCPVExpediente`, `IDCPV`, `IDExpediente`.

**Ownership**: equipo de expediente / catálogos.

**Volumen**: 429 filas cosechadas.

**Acción**: migrar a `expediente_cpvs` del nuevo schema. La PK compuesta se migra como `BIGSERIAL` + UK únicos.

### 22. TbExpedientesE2E → MIGRAR

**Razón**: sincronización E2E con hash del payload. Mapea contra `EXP-CAP-050` (Sincronización E2E). **D114**: consolidar con `TbExpedientes.HashActual` y `TbExpedientes.HashUltimaExportacion` en una sola tabla `expediente_sincronizacion_e2e`.

**Columnas** (de `data-model.md` §TbExpedientesE2E):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` |
| `HashPayload` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — hash del payload E2E |
| `Estado` | 10 (Text) | 20 | false | `VARCHAR(20) NULL` |
| `FechaCreacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaModificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `UsuarioCreacion` | 10 (Text) | 100 | false | `VARCHAR(100) NULL` |
| `UsuarioModificacion` | 10 (Text) | 100 | false | `VARCHAR(100) NULL` |

**Ownership**: equipo de integraciones E2E.

⚠️ **D114**: consolidar con `TbExpedientes.HashActual` y `TbExpedientes.HashUltimaExportacion` en una sola tabla `expediente_sincronizacion_e2e`. Requiere decisión de diseño antes de migrar.

**Acción**: migrar a `expediente_sincronizacion_e2e` del nuevo schema. Aplicar D114 una vez cerrada.

### 23. TbExpedientesHitos → MIGRAR

**Razón**: hitos con importes por expediente. Mapea contra `EXP-CAP-014` (Hitos). Tabla de hitos contractuales.

**Columnas** (de `data-model.md` §TbExpedientesHitos):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDHitoExpediente` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `Descripcion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaHito` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaGarantiaHito` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` — fecha de garantía |
| `Importe` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` |

**Ownership**: equipo de expediente / jurídica.

**Volumen**: 46 filas cosechadas.

**Acción**: migrar a `expediente_hitos` del nuevo schema.

### 24. TbExpedientesJefaturas → MIGRAR

**Razón**:join entre expedientes y jefaturas. Mapea contra `EXP-CAP-006` (Jefaturas por expediente). Tabla dejoin con PK compuesta.

**Columnas** (de `data-model.md` §TbExpedientesJefaturas):

PK compuesta: `IDJefaturaExpediente`, `IDJefatura`, `IDExpediente`.

**Ownership**: equipo de expediente / organización.

**Acción**: migrar a `expediente_jefaturas` del nuevo schema. La PK compuesta se migra como `BIGSERIAL` + UK únicos.

### 25. TbExpedientesJuridicas → MIGRAR

**Razón**: entidades jurídicas por expediente (contratista principal + subcontratista). Mapea contra `EXP-CAP-007` (Jurídicas por expediente). Tabla dejoin con atributos.

**Columnas** (de `data-model.md` §TbExpedientesJuridicas):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteJuridica` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `IDJuridica` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `IDSuministrador` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `ContratistaPrincipal` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `SubContratista` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

**Ownership**: equipo de expediente / jurídica.

**Volumen**: 417 filas cosechadas.

**Acción**: migrar a `expediente_juridicas` del nuevo schema. Consolidar `VARCHAR(2)` → `BOOLEAN` (D102).

### 26. TbExpedientesLugaresEjecucion → MIGRAR

**Razón**:join entre expedientes y lugares de ejecución. Mapea contra `EXP-CAP-006` (Lugares de ejecución por expediente). Tabla dejoin con PK compuesta.

**Columnas** (de `data-model.md` §TbExpedientesLugaresEjecucion):

PK compuesta: `IDExpedienteLugarEjecucion`, `IDExpediente`, `IDLugarEjecucion`.

**Ownership**: equipo de expediente / organización.

**Volumen**: 194 filas cosechadas.

**Acción**: migrar a `expediente_lugares_ejecucion` del nuevo schema. La PK compuesta se migra como `BIGSERIAL` + UK únicos.

### 27. TbExpedientesModificados → MIGRAR

**Razón**: modificaciones al expediente (modificados contractuales). Mapea contra `EXP-CAP-046` (Auditoría / Modificaciones). Tabla de histórico de modificaciones.

**Columnas** (de `data-model.md` §TbExpedientesModificados):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteModificado` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `NModificado` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — número de modificado |
| `FechaFirmaModificado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaFinModificado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `Descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |

**Ownership**: equipo de expediente / jurídica.

**Volumen**: 37 filas cosechadas.

**Acción**: migrar a `expediente_modificados` del nuevo schema.

### 28. TbExpedientesPECAL → MIGRAR

**Razón**:join entre expedientes y códigos PECAL. Mapea contra `EXP-CAP-018` (PECAL por expediente). Tabla dejoin con PK compuesta.

**Columnas** (de `data-model.md` §TbExpedientesPECAL):

PK compuesta: `IDPECALExpediente`, `IDExpediente`, `IDPECAL`.

**Ownership**: equipo de expediente / catálogo.

**Volumen**: 366 filas cosechadas.

**Acción**: migrar a `expediente_pecal` del nuevo schema. La PK compuesta se migra como `BIGSERIAL` + UK únicos.

### 29. TbExpedientesRACS → MIGRAR

**Razón**:join entre expedientes y RACS. Mapea contra `EXP-CAP-018` (RACS por expediente). Tabla dejoin con PK compuesta.

**Columnas** (de `data-model.md` §TbExpedientesRACS):

PK compuesta: `IDRacExpediente`, `IDExpediente`, `IDRAC`.

**Ownership**: equipo de expediente / catálogo.

**Acción**: migrar a `expediente_racs` del nuevo schema. La PK compuesta se migra como `BIGSERIAL` + UK únicos.

### 30. TbExpedientesResponsables → MIGRAR

**Razón**: responsables por expediente con rol. Mapea contra `EXP-CAP-019` (Responsables por expediente). Tabla dejoin con atributos extendido.

**Columnas** (de `data-model.md` §TbExpedientesResponsables):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteResponsable` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IdExpediente` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `IdUsuario` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `CorreoSiempre` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsJefeProyecto` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `esPreventa` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

**Ownership**: equipo de expediente / gestión de usuarios.

**Acción**: migrar a `expediente_responsables` del nuevo schema. Consolidar `VARCHAR(2)` → `BOOLEAN` (D102).

## Transformaciones globales aplicables

| Transformación | Aplica a | Razón |
|---|---|---|
| `Sí/No` TEXT(2) → BOOLEAN | 25, 30 | D102 cross-cutting: booleanos como texto pasan a `BOOLEAN` en PostgreSQL |
| FK LONG → INTEGER | Todas las tablas | El nuevo schema usa `INTEGER` para FKs |
| PK compuesta → BIGSERIAL + UK | 21, 24, 26, 28, 29 | Las PK compuestas Access se migran como surrogate + UK únicos |
| `CURRENCY` → NUMERIC(15,2) | 23 (Importe) | Precisión decimal sin pérdida |
| D114: consolidar E2E | 22 | TbExpedientesE2E + TbExpedientes.HashActual + HashUltimaExportacion |

## Pendientes explícitos

| Tabla | Pendiente | Resolución esperada |
|---|---|---|
| 22 (TbExpedientesE2E) | D114: consolidar hash E2E en una tabla | Decisión de diseño antes de migrar |

## Bloqueos por evidencia runtime

Ninguna tabla en esta tranche tiene bloqueos: todas tienen descripción detallada en `data-model.md`.

## Cuarentena y rechazos

No hay tablas en cuarentena o rechazo en esta tranche.

## Navegación

Tranche anterior: [Source Dictionary Tranche 02 — Tablas 11–20](source-dictionary-02.md).
Tranche siguiente: [Source Dictionary Tranche 04 — Tablas 31–40](source-dictionary-04.md).
Mapping relacionado: [Field Mapping Tranche 03 — Tablas 21–30](field-mapping-03.md).
[← Back to Expedientes migration](README.md).
