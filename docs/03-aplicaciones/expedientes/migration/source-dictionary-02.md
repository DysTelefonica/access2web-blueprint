# Source Dictionary Tranche 02 — Tablas 11–20

Diccionario físico de las tablas ordinales 11-20 del backend legacy `Expedientes_datos.accdb` hacia la migración web del Expediente. Cada entrada declara nombre, columnas, volumen cosechado, capacidad destino del spec `expedientes-web-migration`, y estado de migración. Las tablas sin descripción detallada en `docs/03-aplicaciones/expedientes/data-model.md` quedan marcadas como bloqueadas a la espera de evidencia runtime.

**Fuente legacy**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (extraído con Jackcess el 2026-08-18, PR #393).
**Fuente de columnas**: `docs/03-aplicaciones/expedientes/data-model.md` (cuando cubre la tabla; ver ordinal → tabla en §Tabla de mapeo ordinal → nombre de tabla).
**Mapeo a capacidades**: `docs/03-aplicaciones/expedientes/migration/field-mapping-02.md` / `field-mapping-03.md`.
**Spec destino**: `openspec/changes/expedientes-web-migration/specs/`.
**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9).
**Autoridad de ordinales**: las posiciones 11–20 siguen el ordinal de `data-model.md` (§Tabla de mapeo ordinal → nombre de tabla).

## Resumen de tablas

| # | Tabla legacy | Columnas | Filas | Capacidad destino | Estado |
|---|---|---:|---:|---|---|
| 11 | `TbRACS` | 4 | 37 | EXP-CAP-018 (RACS) | Migrar |
| 12 | `TbResponsablesPorRol` | 3 | 11 | EXP-CAP-019 (Responsables) | Migrar |
| 13 | `TbSuministradores` | 10 | 72 | EXP-CAP-015 (Comerciales/Suministradores) | Migrar |
| 14 | `TbExpedientes` | 75 | 453 | EXP-CAP-001 (Expediente principal) | Migrar — ⚠️ necesita evidencia runtime |
| 15 | `TbExpedientesConEntidades` | 23 | 451 | EXP-CAP-001 (Vista desnormalizada) | VIEW (no migrar tabla) |
| 16 | `TbExpedientesAnualidades` | 11 | 174 | EXP-CAP-008 (Anualidades) | Migrar |
| 17 | `TbExpedientesAnexos` | 3 | 712 | EXP-CAP-009 (Anexos) | Migrar |
| 18 | `TbExpedientesCadenaContratacion` | 9 | — | EXP-CAP-010 (Cadena de contratación) | Migrar |
| 19 | `TbExpedientesCodigoCompras` | 3 | — | EXP-CAP-011 (Códigos de compra) | Migrar |
| 20 | `TbExpedientesComerciales` | 3 | 333 | EXP-CAP-012 (Comerciales por expediente) | Migrar |

## Detalle por tabla

### 11. TbRACS → MIGRAR

**Razón**: catálogo de responsables de actuación contractual (RACS). Mapea contra `EXP-CAP-018` (RACS). Tabla maestra simple.

**Columnas** (de `data-model.md` §TbRACS):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDRAC` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `RAC` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `CORREO` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — correo del RAC |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |

**Ownership**: equipo de catálogos.

**Volumen**: 37 filas cosechadas.

**Acción**: migrar a `racs` del nuevo schema. Sin transformaciones de tipo.

### 12. TbResponsablesPorRol → MIGRAR (parcial)

**Razón**: catálogo de responsables por rol. Mapea contra `EXP-CAP-019` (Responsables). Tabla simple.

**Columnas** (de `schema.sql`; `data-model.md` no detalla columnas):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDResponsablePorRol` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDUsuario` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK a `usuarios` |
| `Rol` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` |

**Ownership**: equipo de gestión de usuarios.

**Volumen**: 11 filas cosechadas.

**Bloqueos**: descripción detallada pendiente — `data-model.md` no cubre esta tabla; las columnas provienen de `schema.sql` y requieren verificación runtime para confirmar cardinalidad y FK semánticas.

**Acción**: migrar a `responsables_por_rol` del nuevo schema. Aplicar decisión de mapeo de `IDUsuario` una vez cerrada la evidencia runtime.

### 13. TbSuministradores → MIGRAR (parcial)

**Razón**: catálogo de suministradores. Mapea contra `EXP-CAP-015` (Comerciales/Suministradores). Tabla maestra central.

**Columnas** (de `schema.sql`; `data-model.md` detalla 4 columnas pero no las 10 completas):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDSuministrador` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Nombre` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CIF` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` — UNIQUE |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `TramitadoraHPS` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `Nemotecnico` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

**Ownership**: equipo de catálogos.

**Volumen**: 72 filas cosechadas.

**Bloqueos**: 2 columnas adicionales (`TramitadoraHPS`, `Nemotecnico`) provienen de `schema.sql` y no están detalladas en `data-model.md`; requieren verificación runtime para confirmar uso y posible consolidación.

**Acción**: migrar a `suministradores` del nuevo schema. Las columnas `TramitadoraHPS` y `Nemotecnico` se migran como `VARCHAR(2)` y `VARCHAR(255)` pending runtime; si no tienen uso activo, se descartan.

### 14. TbExpedientes → MIGRAR (crítico, necesita evidencia runtime)

**Razón**: tabla central del expediente. Mapea contra `EXP-CAP-001` (Expediente principal). 75 columnas en `schema.sql`; `data-model.md` la lista como ordinal 14 pero no la detalla.

**Columnas** (de `schema.sql`; `data-model.md` no detalla; 75 columnas en total):

Fragmento de columnas conocidas:

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpediente` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpedientePadre` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK autorreferente |
| `Nemotecnico` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Titulo` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `ImporteLicitacion` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` |
| `ImporteContratacion` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` |
| `CodExp` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CodExpLargo` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CodS4H` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaInicioContrato` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaFinContrato` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaFinGarantia` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `EsAM` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `Ordinal` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaCreacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaUltimoCambio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `HashActual` | 10 (Text) | 64 | false | `VARCHAR(64) NULL` — hash E2E |
| *(61 columnas restantes)* | — | — | — | Ver `schema.sql` y evidencia runtime |

**Ownership**: equipo de expediente (crítico).

**Volumen**: 453 filas cosechadas.

**Bloqueos**: `data-model.md` no detalla esta tabla; las 75 columnas provienen de `schema.sql` y de `Estructura_Datos.md`. Se requiere `dysflow get_schema` contra la fuente autoritativa para confirmar PK, FK, y uso real de cada columna antes de definir el mapeo PostgreSQL.

**Acción**: migrar a `expedientes` del nuevo schema. El ordinal 14 es la tabla pivot central; todas las demás tablas Expedientes referencian su `IDExpediente`. La migración de esta tabla condiciona el resto.

### 15. TbExpedientesConEntidades → VIEW (no migrar tabla)

**Razón**: vista desnormalizada del expediente con entidades relacionadas en cadenas separadas por comas. Mapea contra `EXP-CAP-001`. **D113 propuesta**: migrar como `VIEW` en PostgreSQL, no como tabla.

**Columnas** (de `data-model.md` §TbExpedientesConEntidades, 23 columnas):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpediente` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` PK |
| `Clasificacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `OrganoContratacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `OficinaPrograma` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Ejercito` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Estado` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `ResponsableCalidad` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `ResponsableSeguridad` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CadenaPecal` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `Pecal` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `CadenaContratistas` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaSubContratistas` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaSuministradores` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaComerciales` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaJPs` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaRACs` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaCorreoRACs` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaHitos` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `TipoParaLista` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CadenaLugares` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| `CadenaJuridicas` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — cadena |
| *(3 columnas restantes)* | — | — | — | Ver `schema.sql` |

⚠️ **D113 propuesta**: esta tabla es una **vista materializada** que se debe regenerar como `VIEW` en PostgreSQL. Las 13 `Cadena*` son strings separados por comas que **no escalan**. La migración es mantener las tablas normalizadas y definir la `VIEW` como query que une `expedientes` con sus tablas dejoin.

**Ownership**: equipo de expediente.

**Volumen**: 451 filas.

**Acción**: no migrar como tabla. Crear `VIEW expediente_entidades` en PostgreSQL que replique el resultado de la query Access.

### 16. TbExpedientesAnualidades → MIGRAR

**Razón**: anualidades con importes por tipo de impuesto. Mapea contra `EXP-CAP-008` (Anualidades). 11 columnas, 174 filas.

**Columnas** (de `data-model.md` §TbExpedientesAnualidades):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDAnualidad` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `Año` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `BIIVA` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` — Base imponible IVA |
| `BIIPSI` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` — Base imponible IPSI |
| `BIIGIC` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` — Base imponible IGIC |
| `BIEXENTA` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` — Base imponible exenta |
| `IVA` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` — cuota IVA |
| `IPSI` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` — cuota IPSI |
| `IGIC` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` — cuota IGIC |
| `PeriodoFacturacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

⚠️ **Sistema multi-impuestos** (IVA, IPSI, IGIC) — migrar con `NUMERIC(15,2)` en PostgreSQL. **CRÍTICO para auditoría regulatoria**.

**Ownership**: equipo de expediente / financiación.

**Volumen**: 174 filas cosechadas.

**Acción**: migrar a `expediente_anualidades` del nuevo schema.

### 17. TbExpedientesAnexos → MIGRAR

**Razón**: documentos anexos al expediente. Mapea contra `EXP-CAP-009` (Anexos). Tabla de documentos vinculados.

**Columnas** (de `data-model.md` §TbExpedientesAnexos):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDDocumento` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `NombreDocumento` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

**Ownership**: equipo de expediente.

**Volumen**: 712 filas cosechadas.

**Acción**: migrar a `expediente_anexos` del nuevo schema. Los documentos reales se almacenan en object storage (S3-compatible); esta tabla registra los metadatos.

### 18. TbExpedientesCadenaContratacion → MIGRAR

**Razón**: cadena de subcontratación con reglas. Mapea contra `EXP-CAP-010` (Cadena de contratación). Tabla de relaciones jerárquicas de contratistas.

**Columnas** (de `data-model.md` §TbExpedientesCadenaContratacion):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `ID` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDPadre` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — padre (jerarquía) |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `IDSuministrador` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `AplicaCalidad` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `AplicaRiesgos` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `AplicaContratosClasificados` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `AplicaHPS` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — ⚠️ anomalía: Text 255 vs Text 2 |
| `Observaciones` | 12 (Memo) | 0 | false | `TEXT NULL` |

**Ownership**: equipo de expediente / jurídica.

**Acción**: migrar a `expediente_cadena_contratacion` del nuevo schema. Consolidar los `VARCHAR(2)` a `BOOLEAN` (D102).

### 19. TbExpedientesCodigoCompras → MIGRAR

**Razón**: códigos de compra por expediente. Mapea contra `EXP-CAP-011` (Códigos de compra). Tabla dejoin simple.

**Columnas** (de `data-model.md` §TbExpedientesCodigoCompras):

PK + `IDExpediente` (FK) + `CodCompras` (Text 255).

**Ownership**: equipo de expediente / compras.

**Acción**: migrar a `expediente_codigos_compras` del nuevo schema.

### 20. TbExpedientesComerciales → MIGRAR

**Razón**: comerciales asignados al expediente. Mapea contra `EXP-CAP-012` (Comerciales por expediente). Tabla dejoin con PK compuesta.

**Columnas** (de `data-model.md` §TbExpedientesComerciales):

PK compuesta: `IDComercialExpediente`, `IDComercial`, `IDExpediente`.

**Ownership**: equipo comercial.

**Volumen**: 333 filas cosechadas.

**Acción**: migrar a `expediente_comerciales` del nuevo schema. La PK compuesta se migra como `BIGSERIAL` + UK únicos.

## Transformaciones globales aplicables

| Transformación | Aplica a | Razón |
|---|---|---|
| `Sí/No` TEXT(2) → BOOLEAN | TbExpedientesCadenaContratacion (AplicaCalidad/Riesgos/ContratosClasificados), TbExpedientes (EsAM, etc.) | D102 cross-cutting: booleanos como texto pasan a `BOOLEAN` en PostgreSQL |
| FK LONG → INTEGER | Todas las tablas Expedientes | El nuevo schema usa `INTEGER` para FKs |
| `CURRENCY` → NUMERIC(15,2) | TbExpedientes (ImporteLicitacion, ImporteContratacion) | Precisión decimal sin pérdida |
| VIEW en lugar de tabla | TbExpedientesConEntidades | D113: generar VIEW desde tablas normalizadas en PostgreSQL |

## Pendientes explícitos

| Tabla | Pendiente | Resolución esperada |
|---|---|---|
| 12 (TbResponsablesPorRol) | Confirmar cardinalidad y FK semánticas de IDUsuario | dysflow get_schema contra fuente autoritativa |
| 13 (TbSuministradores) | Verificar uso activo de TramitadoraHPS y Nemotecnico | dysflow get_schema; posible descarte si sin uso |
| 14 (TbExpedientes) | Detallar las 75 columnas; confirmar PK, FK, uso real | dysflow get_schema; decisión de split (campos HPS, E2E) |
| 15 (TbExpedientesConEntidades) | D113: definir query de VIEW PostgreSQL | Spec EXP-CAP-001 + diseño D113 |

## Bloqueos por evidencia runtime

Las tablas 12, 13 y 14 carecen de descripción detallada en `data-model.md`. La política de esta tranche es marcarlas como bloqueadas hasta contar con `dysflow get_schema` o equivalente contra `C:\00repos\datos\Expedientes_datos.accdb`. Las entradas de `schema.sql` cubren únicamente nombre y tipo básico, insuficiente para asignar PK, FK y reglas de validación en el mapeo a PostgreSQL.

La tabla 14 (TbExpedientes) es **crítica**: condiciona el mapeo de todas las demás tablas Expedientes que referencian `IDExpediente`.

## Cuarentena y rechazos

| Categoría | Tablas | Acción |
|---|---|---|
| VIEW (no tabla) | 15 | Generar VIEW PostgreSQL desde tablas normalizadas |
| Migra con evidencia runtime | 12, 13, 14 | dysflow get_schema antes de definir mapeo PostgreSQL |

## Navegación

Tranche anterior: [Source Dictionary Tranche 01 — Tablas 01–10](source-dictionary-01.md).
Tranche siguiente: [Source Dictionary Tranche 03 — Tablas 21–30](source-dictionary-03.md).
Mapping relacionado: [Field Mapping Tranche 02 — Tablas 16–30](field-mapping-02.md), [Field Mapping Tranche 03](field-mapping-03.md).
[← Back to Expedientes migration](README.md).
