# Source Dictionary Tranche 04 — Tablas 31–40

Diccionario físico de las tablas ordinales 31-40 del backend legacy `Expedientes_datos.accdb`. Estas tablas no están detalladas en `data-model.md`; sus columnas provienen de `schema.sql` y requieren evidencia runtime mediante `dysflow get_schema` contra la fuente autoritativa para asignar PK, FK, relaciones y reglas de validación.

**Fuente legacy**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (extraído con Jackcess el 2026-08-18, PR #393).
**Fuente de columnas**: `docs/03-aplicaciones/expedientes/data-model.md` — **no detalla ninguna tabla de esta tranche**.
**Spec destino**: `openspec/changes/expedientes-web-migration/specs/`.
**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9).
**Autoridad de ordinales**: las posiciones 31–40 siguen el ordinal de `data-model.md` (§Tabla de mapeo ordinal → nombre de tabla).

## Resumen de tablas

| # | Tabla legacy | Columnas | Filas | Capacidad destino | Estado |
|---|---|---:|---:|---|---|
| 31 | `TbExpedientesSuministradores` | 7 | 72 | EXP-CAP-015 (Suministradores por expediente) | Migrar — detallado en data-model.md |
| 32 | `TbCambios` | 10 | — | EXP-CAP-046 (Auditoría) | Bloqueado — necesita evidencia runtime |
| 33 | `TbUltimoCambio` | 4 | — | EXP-CAP-046 (Auditoría) | CUARENTENA |
| 34 | `TbConfMostrarEstado` | 3 | 23 | EXP-CAP-005 (Estado y garantía) | CUARENTENA — preferencia UI |
| 35 | `TbE2EExportBatch` | 9 | — | EXP-CAP-050 (Sincronización E2E) | CUARENTENA — operación E2E |
| 36 | `TbE2EExportBatchDetalle` | 8 | — | EXP-CAP-050 (Sincronización E2E) | CUARENTENA — operación E2E |
| 37 | `TbE2EExportSeleccionTemp` | 5 | — | EXP-CAP-050 (Sincronización E2E) | RECHAZAR — tabla temporal |
| 38 | `TbE2EJsonDestinationUserConfig` | 5 | — | EXP-CAP-050 (Sincronización E2E) | CUARENTENA — config E2E |
| 39 | `TbExpedientesE2E` | — | — | — | RE-ENTRY (ordinal 22 ya detallado) |
| 40 | `TbDatosEconomicosExpedientes` | 7 | 91 | EXP-CAP-001 (Expediente principal) | Bloqueado — necesita evidencia runtime |

## Detalle por tabla

### 31. TbExpedientesSuministradores → MIGRAR

**Razón**:join entre expedientes y suministradores (jerárquico). Mapea contra `EXP-CAP-015` (Suministradores por expediente). Tabla dejoin con atributos extendido.

**Columnas** (de `data-model.md` §TbExpedientesSuministradores):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteSuministrador` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `IDSuministrador` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `IDPadre` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — padre (jerarquía) |
| `Descripcon` | 12 (Memo) | 0 | false | `TEXT NULL` — ⚠️ typo: debería ser "Descripcion" |
| `ContratistaPrincipal` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `SubContratista` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

**Ownership**: equipo de expediente / jurídica.

**Volumen**: 72 filas cosechadas.

⚠️ **Typo en nombre de columna**: `Descripcon` (falta la "i") — verificar si es error de Access o se migra con el typo.

**Acción**: migrar a `expediente_suministradores` del nuevo schema. Consolidar `VARCHAR(2)` → `BOOLEAN` (D102). Corregir typo en nombre de columna en PostgreSQL.

### 32. TbCambios → BLOQUEADO — necesita evidencia runtime

**Razón**: historial de cambios de auditoría. Mapea contra `EXP-CAP-046` (Auditoría). `data-model.md` la lista como ordinal 32 pero no la detalla.

**Columnas** (de `schema.sql`; `data-model.md` no detalla):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDCambio` | 4 (LongInteger) | 4 | — | BIGSERIAL PK — por confirmar |
| `NombreTabla` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `NombreCampoID` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `ValorCampoID` | 4 (LongInteger) | 4 | — | `INTEGER` — por confirmar |
| `NombreCampo` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `ValorInicial` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `ValorFinal` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `FechaCambio` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |
| `IDUsuarioCambio` | 4 (LongInteger) | 4 | — | `INTEGER` — por confirmar |
| `Accion` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |

**Ownership**: equipo de auditoría.

**Volumen**: no disponible (0 en la extracción original; requiere nueva lectura runtime).

**Bloqueos**: `data-model.md` no detalla esta tabla; requiere `dysflow get_schema` para confirmar la estructura completa.

**Acción**: migrar a `audit_log` del nuevo schema pending evidencia runtime. Aplicar D102 para campos `Sí/No` si los hay.

### 33. TbUltimoCambio → CUARENTENA

**Razón**: tabla de último cambio por expediente (cache de auditoría). Mapea contra `EXP-CAP-046` (Auditoría). Tabla auxiliar cuyo contenido puede reconstruirse desde `TbCambios`.

**Columnas** (de `schema.sql`):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `ID` | 4 (LongInteger) | 4 | — | BIGSERIAL PK — por confirmar |
| `IDExpediente` | 4 (LongInteger) | 4 | — | `INTEGER` — FK — por confirmar |
| `FechaCambio` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |
| `IDUsuarioCambio` | 4 (LongInteger) | 4 | — | `INTEGER` — por confirmar |

**Ownership**: equipo de auditoría.

**Acción**: cuarentena. Analizar si el contenido se puede reconstruir desde `TbCambios` (ordinal 32) en cuyo caso se descarta; si es irreemplazable, migrar.

### 34. TbConfMostrarEstado → CUARENTENA

**Razón**: configuración por usuario de qué estados mostrar en la UI. Mapea contra `EXP-CAP-005` (Estado y garantía). Tabla de preferencias de presentación.

**Columnas** (de `schema.sql`):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `ID` | 4 (LongInteger) | 4 | — | BIGSERIAL PK — por confirmar |
| `UsuarioRed` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `MostrarEstado` | 10 (Text) | 2 | — | `VARCHAR(2)` — ⚠️ por confirmar |

**Ownership**: equipo de presentación.

**Volumen**: 23 filas cosechadas.

**Acción**: cuarentena. Analizar si las reglas de visualización pasan a `Jinja2` templates del nuevo schema o se conservan en `preferencias_usuario`.

### 35. TbE2EExportBatch → CUARENTENA

**Razón**: batch de exportación E2E. Mapea contra `EXP-CAP-050` (Sincronización E2E). Tabla operacional del sistema E2E.

**Columnas** (de `schema.sql`; parcial):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDBatch` | 4 (LongInteger) | 4 | — | BIGSERIAL PK — por confirmar |
| `SessionId` | 10 (Text) | 100 | — | `VARCHAR(100)` — por confirmar |
| `UsuarioConectado` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `Estado` | 10 (Text) | 50 | — | `VARCHAR(50)` — por confirmar |
| `CreatedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |
| `StartedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |
| `CompletedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |
| `TotalSeleccionados` | 4 (LongInteger) | 4 | — | `INTEGER` — por confirmar |
| `TotalExportados` | 4 (LongInteger) | 4 | — | `INTEGER` — por confirmar |
| `ErrorMessage` | 12 (Memo) | 0 | — | `TEXT` — por confirmar |

**Ownership**: equipo de integraciones E2E.

**Acción**: cuarentena. Analizar si estos datos se migran o se descartan (batch histórico).

### 36. TbE2EExportBatchDetalle → CUARENTENA

**Razón**: detalle de batch de exportación E2E. Mapea contra `EXP-CAP-050`. Tabla operacional del sistema E2E.

**Columnas** (de `schema.sql`; parcial):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDBatchDetalle` | 4 (LongInteger) | 4 | — | BIGSERIAL PK — por confirmar |
| `IDBatch` | 4 (LongInteger) | 4 | — | `INTEGER` — FK — por confirmar |
| `IDExpediente` | 4 (LongInteger) | 4 | — | `INTEGER` — FK — por confirmar |
| `OrdinalSeleccion` | 4 (LongInteger) | 4 | — | `INTEGER` — por confirmar |
| `HashExportado` | 10 (Text) | 64 | — | `VARCHAR(64)` — hash de exportación |
| `Estado` | 10 (Text) | 50 | — | `VARCHAR(50)` — por confirmar |
| `CreatedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |
| `ExportedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |

**Ownership**: equipo de integraciones E2E.

**Acción**: cuarentena. Analizar si estos datos se migran o se descartan.

### 37. TbE2EExportSeleccionTemp → RECHAZAR

**Razón**: tabla temporal de selección de expedientes para exportación E2E. Mapea contra `EXP-CAP-050`. Tabla efímera de sesión.

**Columnas** (de `schema.sql`):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDTemp` | 4 (LongInteger) | 4 | — | BIGSERIAL PK — por confirmar |
| `UsuarioConectado` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `SessionId` | 10 (Text) | 100 | — | `VARCHAR(100)` — por confirmar |
| `IDExpediente` | 4 (LongInteger) | 4 | — | `INTEGER` — FK — por confirmar |
| `CreatedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |

**Ownership**: equipo de integraciones E2E.

**Acción**: no migrar. Tabla temporal de sesión; el nuevo sistema recrea la selección en memoria o en caché Redis.

### 38. TbE2EJsonDestinationUserConfig → CUARENTENA

**Razón**: configuración de destino JSON por usuario para exportación E2E. Mapea contra `EXP-CAP-050`. Tabla de configuración.

**Columnas** (de `schema.sql`):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDConfig` | 4 (LongInteger) | 4 | — | BIGSERIAL PK — por confirmar |
| `UsuarioRed` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |
| `RutaDestino` | 12 (Memo) | 0 | — | `TEXT` — path de destino |
| `CreatedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |
| `UpdatedAt` | 8 (DateTime) | 8 | — | `TIMESTAMP` — por confirmar |

**Ownership**: equipo de integraciones E2E.

**Acción**: cuarentena. Analizar si la configuración de destino se migra a `preferencias_usuario` del nuevo schema.

### 39. TbExpedientesE2E → RE-ENTRY (ordinal 22 ya detallado)

Esta entrada ordinal apunta a `TbExpedientesE2E`, que ya está detallada en la entrada ordinal 22 de `data-model.md` y en el source-dictionary-03.md. No es una tabla independiente; se documenta como entrada ordinal duplicada para completar el mapeo de posiciones 01–49.

Ver entrada ordinal 22 → [Source Dictionary Tranche 03](source-dictionary-03.md).

### 40. TbDatosEconomicosExpedientes → BLOQUEADO — necesita evidencia runtime

**Razón**: datos económicos por expediente. Mapea contra `EXP-CAP-001` (Expediente principal). Tabla de join económico.

**Columnas** (de `schema.sql`):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpediente` | 4 (LongInteger) | 4 | — | `INTEGER` — FK — por confirmar |
| `MontoTotal` | 7 (Currency) | 8 | — | `NUMERIC(15,2)` — por confirmar |
| `CosteExternoPrevisto` | 7 (Currency) | 8 | — | `NUMERIC(15,2)` — por confirmar |
| `CosteInternoPrevisto` | 7 (Currency) | 8 | — | `NUMERIC(15,2)` — por confirmar |
| `GastoMaximoPrevisto` | 7 (Currency) | 8 | — | `NUMERIC(15,2)` — por confirmar |
| `InversionPrevista` | 7 (Currency) | 8 | — | `NUMERIC(15,2)` — por confirmar |
| `Observaciones` | 10 (Text) | 255 | — | `VARCHAR(255)` — por confirmar |

**Ownership**: equipo de expediente / financiación.

**Volumen**: 91 filas cosechadas.

**Bloqueos**: `data-model.md` no detalla esta tabla; requiere `dysflow get_schema` para confirmar FK, cardinalidad y posibles duplicaciones con columnas de `TbExpedientes` (ordinal 14).

**Acción**: migrar a `expediente_datos_economicos` del nuevo schema pending evidencia runtime.

## Transformaciones globales aplicables

| Transformación | Aplica a | Razón |
|---|---|---|
| `Sí/No` TEXT(2) → BOOLEAN | 31 (ContratistaPrincipal, SubContratista) | D102 cross-cutting |
| FK LONG → INTEGER | Todas las tablas | El nuevo schema usa `INTEGER` para FKs |
| `CURRENCY` → NUMERIC(15,2) | 40 (datos económicos) | Precisión decimal sin pérdida |
| Descartar temporales | 37 (TbE2EExportSeleccionTemp) | Tabla efímera de sesión; no migrar |

## Pendientes explícitos

| Tabla | Pendiente | Resolución esperada |
|---|---|---|
| 32 (TbCambios) | dysflow get_schema para confirmar columnas y cardinalidad | Ejecutar dysflow get_schema sobre TbCambios |
| 33 (TbUltimoCambio) | Analizar si es redundante con TbCambios | Comparar contenido; si redundante, rechazar |
| 34 (TbConfMostrarEstado) | Decidir si preferencias UI migran a preferencias_usuario | Análisis UX; puede descartarse si preferencia no se usa |
| 35, 36, 38 (E2E export) | Analizar si datos históricos E2E migran | Decisión de negocio: migrar historial o descartar |
| 40 (TbDatosEconomicosExpedientes) | dysflow get_schema para confirmar FK y duplicación con ordinal 14 | Ejecutar dysflow get_schema; comparar con columnas de TbExpedientes |

## Bloqueos por evidencia runtime

Todas las tablas de esta tranche (excepto la 31, detallada en data-model.md) requieren `dysflow get_schema` contra `C:\00repos\datos\Expedientes_datos.accdb` para confirmar columnas, tipos, FKs, cardinalidad yPK antes de definir el mapeo PostgreSQL.

## Cuarentena y rechazos

| Categoría | Tablas | Acción |
|---|---|---|
| Cuarentena | 33, 34, 35, 36, 38 | Analizar contenido antes de decidir entre migrar o rechazar |
| Rechazar | 37 | Tabla temporal; no migrar |

## Navegación

Tranche anterior: [Source Dictionary Tranche 03 — Tablas 21–30](source-dictionary-03.md).
Tranche siguiente: [Source Dictionary Tranche 05 — Tablas 41–49 (re-entries)](source-dictionary-05.md).
[← Back to Expedientes migration](README.md).
