# Source Dictionary Tranche 01 — Tablas 01–10

Diccionario físico de las primeras diez tablas del backend legacy `Expedientes_datos.accdb` hacia la migración web del Expediente. Cada entrada declara nombre, columnas, volumen cosechado, capacidad destino del spec `expedientes-web-migration`, y estado de migración. Las tablas sin descripción detallada en `docs/03-aplicaciones/expedientes/data-model.md` quedan marcadas como bloqueadas a la espera de evidencia runtime.

**Fuente legacy**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (extraído con Jackcess el 2026-08-18, PR #393).
**Fuente de columnas**: `docs/03-aplicaciones/expedientes/data-model.md` (cuando cubre la tabla).
**Mapeo a capacidades**: `docs/03-aplicaciones/expedientes/migration/field-mapping-01.md` (PR #395).
**Spec destino**: `openspec/changes/expedientes-web-migration/specs/`.
**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9).
**Autoridad de ordinales**: las posiciones 01–49 siguen el orden de aparición en `schema.sql`.

## Resumen de tablas

| # | Tabla legacy | Columnas | Filas | Capacidad destino | Estado |
|---|---|---:|---:|---|---|
| 01 | Copia de TbExpedientes | 55 | 0 | No migrar (backup legacy) | Rechazar |
| 02 | Copia de TbExpedientesConEntidades | 23 | 365 | No migrar (backup legacy) | Rechazar |
| 03 | ListaPrevia | 52 | 99 | No migrar (lista preliminar legacy) | Rechazar |
| 04 | TbAusExpPostAGEDO | 1 | 203 | EXP-CAP-051 (AGEDYS) | Cuarentena |
| 05 | TbAuxEstadosMartina | 3 | 143 | No migrar (auxiliar legacy) | Rechazar |
| 06 | TbAuxNemotecnico | 1 | 0 | EXP-CAP-007 (Ordinal) | Transformar |
| 07 | TbCambios | 10 | 0 | EXP-CAP-046 (Auditoría) | Migrar |
| 08 | TbComerciales | 3 | 26 | EXP-CAP-015 (Comerciales) | Migrar |
| 09 | TbConfMostrarEstado | 3 | 23 | EXP-CAP-005 (Estado y garantía) | Transformar |
| 10 | TbCPV | 3 | 73 | EXP-CAP-016 (CPV) | Migrar |

## Detalle por tabla

### 01. Copia de TbExpedientes → RECHAZAR

**Razón**: tabla de backup legacy (`Copia de`). No contiene datos activos en el momento de la extracción (`Rows: 0`).

**Columnas**: 55 columnas, idénticas a `TbExpedientes` salvo autonuméricos y agregados. El detalle vive en `field-mapping-01.md` §01 y en el spec de ciclo de vida (`lifecycle`).

**Ownership**: legacy, sin equipo asignado. No se requiere decisión de negocio previa al rechazo.

**Volumen**: 0 filas cosechadas.

**Acción**: no migrar. Documentar como backup histórico en la página de la ausencia del cambio.

### 02. Copia de TbExpedientesConEntidades → RECHAZAR

**Razón**: tabla de backup legacy (`Copia de`). Réplica desnormalizada de `TbExpedientesConEntidades` con cadenas separadas por comas. Mismo criterio que la entrada #01.

**Columnas**: 23 columnas (13 de ellas `Cadena*`, cadenas separadas por comas; ver D-EXP-1 propuesta de reemplazo por `VIEW`).

**Ownership**: legacy.

**Volumen**: 365 filas cosechadas.

**Acción**: no migrar. La tabla original `TbExpedientesConEntidades` (entrada #26) es la que pasa a PostgreSQL como `VIEW`.

### 03. ListaPrevia → RECHAZAR

**Razón**: tabla preliminar usada en flujos legacy de carga. No figura en el spec del change `expedientes-web-migration` porque ningún subsistema la referencia.

**Columnas**: 52 columnas, mezcla de datos del expediente y campos auxiliares (`FPreOferta`, `FOferta`, `Jefe Proyecto`, `Racs`, `Hitos`, etc.).

**Ownership**: legacy.

**Volumen**: 99 filas cosechadas.

**Acción**: no migrar. Documentar como tabla preliminar sin capacidad asociada.

### 04. TbAusExpPostAGEDO → CUARENTENA

**Razón**: tabla auxiliar de AGEDO (sistema de gestión de expedientes anterior). Apunta al integrador AGEDYS como `EXP-CAP-051`, pero el contenido no se ha revisado.

**Columnas**: 1 columna (`ID` LONG). El esquema cosechado no incluye detalle de FKs ni claves de cruce.

**Ownership**: equipo de integraciones (a confirmar).

**Volumen**: 203 filas cosechadas.

**Bloqueos**: descripción detallada pendiente — `data-model.md` no cubre esta tabla; la columna única `ID` no basta para asignar mapeo de campos sin contraste contra `dysflow get_schema` sobre la fuente autoritativa.

**Acción**: cuarentena. Analizar contenido en una iteración posterior antes de decidir entre migrar, transformar o rechazar.

### 05. TbAuxEstadosMartina → RECHAZAR

**Razón**: tabla auxiliar legacy (`TbAux`). No aparece en el spec.

**Columnas**: 3 columnas (`ID` LONG, `EstadoMartina` TEXT, `Igual` TEXT(2)). El campo `Igual` es `Sí/No` como texto (ver D102 cross-cutting).

**Ownership**: legacy.

**Volumen**: 143 filas cosechadas.

**Acción**: no migrar. Documentar como auxiliar legacy sin capacidad asociada.

### 06. TbAuxNemotecnico → TRANSFORMAR

**Razón**: tabla auxiliar de nemotécnicos (códigos cortos de identificación de expediente). El nemotécnico se mantiene como campo en `expedientes` del nuevo schema; la tabla auxiliar se descarta.

**Columnas**: 1 columna (`Nemotecnico` TEXT). Tabla vacía al momento de la extracción.

**Ownership**: legacy.

**Volumen**: 0 filas cosechadas.

**Acción**: transformar. Consolidar `Nemotecnico` como columna dentro de `expedientes.nemotecnico` y descartar la tabla.

### 07. TbCambios → MIGRAR

**Razón**: tabla de auditoría de cambios. Mapea contra `EXP-CAP-046` (Auditoría). Historial de cambios por expediente, tabla y campo.

**Columnas** (de `schema.sql`): `IDCambio` LONG, `NombreTabla` TEXT(255), `NombreCampoID` TEXT(255), `ValorCampoID` LONG, `NombreCampo` TEXT(255), `ValorInicial` TEXT(255), `ValorFinal` TEXT(255), `FechaCambio` SHORT_DATE_TIME, `IDUsuarioCambio` LONG, `Accion` TEXT(255).

**Ownership**: equipo de auditoría (a confirmar).

**Volumen**: 0 filas cosechadas al momento de la extracción. El volumen real en producción requiere nueva lectura runtime.

**Acción**: migrar a `audit_log` del nuevo schema. Transformar `IDUsuarioCambio` LONG a FK numérica.

### 08. TbComerciales → MIGRAR

**Razón**: catálogo de comerciales. Mapea contra `EXP-CAP-015` (Comerciales). Tabla maestra simple.

**Columnas** (de `data-model.md` §`TbComerciales`):

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDComercial` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Comercial` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `Descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |

**Ownership**: equipo de catálogos.

**Volumen**: 26 filas cosechadas.

**Acción**: migrar a `comerciales` del nuevo schema. Sin transformaciones de tipo.

### 09. TbConfMostrarEstado → TRANSFORMAR

**Razón**: configuración por usuario de qué estados mostrar en la UI. Mapea contra `EXP-CAP-005` (Estado y garantía). Las reglas de visualización se mueven a la capa de presentación.

**Columnas** (de `schema.sql`): `ID` LONG, `UsuarioRed` TEXT(255), `MostrarEstado` TEXT(2).

**Ownership**: equipo de presentación.

**Volumen**: 23 filas cosechadas.

**Bloqueos**: descripción detallada pendiente — `data-model.md` no cubre esta tabla; las reglas de mapeo (`MostrarEstado` como Sí/No texto) requieren verificación contra el formulario legacy.

**Acción**: transformar. Las reglas de visualización pasan a `Jinja2` templates del nuevo schema; el registro por usuario se conserva como `preferencias_usuario`.

### 10. TbCPV → MIGRAR

**Razón**: catálogo CPV (Common Procurement Vocabulary). Mapea contra `EXP-CAP-016` (CPV). Tabla maestra simple.

**Columnas** (de `data-model.md` §`TbCPV`): PK + `CPV` + `DESCRIPCION`.

**Ownership**: equipo de catálogos.

**Volumen**: 73 filas cosechadas.

**Bloqueos**: la decisión de dígitos CPV (8 vs 9 dígitos, ver `D-EXP-3` pregunta abierta) afecta al tipo PostgreSQL final.

**Acción**: migrar a `cpv` del nuevo schema. Sin transformaciones de tipo. Aplicar la decisión CPV una vez quede cerrada.

## Transformaciones globales aplicables

| Transformación | Aplica a | Razón |
|---|---|---|
| `IDUsuario*` TEXT → FK numérica | TbCambios (columna `IDUsuarioCambio`) | El legacy usa TEXT o LONG para IDs de usuario; el nuevo schema usa FK numérica |
| `Sí/No` TEXT(2) → BOOLEAN | TbAuxEstadosMartina, TbConfMostrarEstado | D102 cross-cutting: booleanos como texto pasan a `BOOLEAN` en PostgreSQL |
| Nemotécnico → columna en `expedientes` | TbAuxNemotecnico | Consolidación en la tabla principal del expediente |
| `Copia de...` → descarte directo | 01, 02 | El legacy permanece intacto durante la transición; las copias no se migran |

## Pendientes explícitos

| Tabla | Pendiente | Resolución esperada |
|---|---|---|
| 04 (TbAusExpPostAGEDO) | Revisar contenido antes de decidir | Análisis en iteración posterior; capacidad destino AGEDYS confirmada |
| 09 (TbConfMostrarEstado) | Detallar reglas de mapeo `MostrarEstado` | Verificación contra formulario legacy; conversión a `preferencias_usuario` |
| 10 (TbCPV) | Dígitos CPV (8 vs 9) | Decisión `D-EXP-3` |

## Bloqueos por evidencia runtime

Las tablas 04, 05 y 09 carecen de descripción detallada en `data-model.md`. La política de esta tranche es marcarlas como bloqueadas hasta contar con `dysflow get_schema` o equivalente contra `C:\00repos\datos\Expedientes_datos.accdb`. Las entradas de `schema.sql` cubren únicamente nombre y tipo básico, insuficiente para asignar PK, FK y reglas de validación en el mapeo a PostgreSQL.

## Cuarentena y rechazos

| Categoría | Tablas | Acción |
|---|---|---|
| Cuarentena | 04 | Analizar contenido antes de decidir |
| Rechazo | 01, 02, 03, 05 | Backup, preliminar o auxiliar legacy sin capacidad |

## Navegación

Tranche anterior: — (es el primero).
Tranche siguiente: [Source Dictionary Tranche 02 — Tablas 11–20](source-dictionary-02.md).
Mapping relacionado: [Field Mapping Tranche 01 — Tablas 01–15](field-mapping-01.md).
[← Back to Expedientes migration](README.md).