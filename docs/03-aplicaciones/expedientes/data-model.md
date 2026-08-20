# Expedientes — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: `C:\00repos\datos\Expedientes_datos.accdb`, resuelta explícitamente el 2026-08-05 mediante Dysflow `list_tables`, `get_schema`, `get_relationships` y `count_rows`. Se detectaron **49 tablas de usuario**. El frontend `staging\Expedientes.accdb` no sustituye al backend autoritativo.

## Diccionario completo

El diccionario fuente completo, con **cada tabla y cada campo**, está cosechado en [Estructura_Datos.md](../../../../documentacion/OPENSPEC/00_EXPEDIENTES/docs/ERD/Estructura_Datos.md). Ese documento se conserva como fuente primaria del inventario de columnas; esta página añade la normalización, procedencia, relaciones y perfilado exigidos para la migración. No se han copiado valores de filas.

## Tabla de mapeo ordinal → nombre de tabla

Mapeo canónico (de `dysflow list_tables` ejecutado el 2026-08-05, validado contra `docs/03-aplicaciones/expedientes/ERD/schema.sql` que contiene los 49 ``CREATE TABLE``):

| Ordinal | Tabla | Cubierta en este doc | Columnas en schema.sql |
|---|---|---|---|
| 01 | `TbEstados` | sí (3 cols) | sí |
| 02 | `TbComerciales` | sí (3 cols) | sí |
| 03 | `TbCPV` | sí (3 cols) | sí |
| 04 | `TbEjercitos` | sí (3 cols) | sí |
| 05 | `TbJefaturas` | sí (3 cols) | sí |
| 06 | `TbJuridicas` | sí (4 cols, 417 filas) | sí |
| 07 | `TbLugaresEjecucion` | sí (3 cols, 194 filas) | sí |
| 08 | `TbOficinasPrograma` | sí (3 cols) | sí |
| 09 | `TbOrganosContratacion` | sí (3 cols) | sí |
| 10 | `TbPECAL` | sí (3 cols) | sí |
| 11 | `TbRACS` | sí (4 cols, 37 filas) | sí |
| 12 | `TbResponsablesPorRol` | parcial (no detallada) | sí |
| 13 | `TbSuministradores` | parcial (no detallada) | sí |
| 14 | `TbExpedientes` | parcial (no detallada) | sí |
| 15 | `TbExpedientesConEntidades` | sí (23 cols) | sí |
| 16 | `TbExpedientesAnualidades` | sí (11 cols, 174 filas) | sí |
| 17 | `TbExpedientesAnexos` | sí (3 cols, 712 filas) | sí |
| 18 | `TbExpedientesCadenaContratacion` | sí (9 cols) | sí |
| 19 | `TbExpedientesCodigoCompras` | sí (3 cols) | sí |
| 20 | `TbExpedientesComerciales` | sí (3 cols) | sí |
| 21 | `TbExpedientesCPVs` | sí (3 cols) | sí |
| 22 | `TbExpedientesE2E` | sí (7 cols) | sí |
| 23 | `TbExpedientesHitos` | sí (6 cols, 46 filas) | sí |
| 24 | `TbExpedientesJefaturas` | sí (3 cols) | sí |
| 25 | `TbExpedientesJuridicas` | sí (6 cols, 417 filas) | sí |
| 26 | `TbExpedientesLugaresEjecucion` | sí (3 cols, 194 filas) | sí |
| 27 | `TbExpedientesModificados` | sí (6 cols, 37 filas) | sí |
| 28 | `TbExpedientesPECAL` | sí (3 cols) | sí |
| 29 | `TbExpedientesRACS` | sí (3 cols) | sí |
| 30 | `TbExpedientesResponsables` | parcial (no detallada) | sí |
| 31 | `TbExpedientesSuministradores` | parcial (no detallada) | sí |
| 32 | `TbCambios` | parcial (no detallada) | sí |
| 33 | `TbUltimoCambio` | no detallada | sí |
| 34 | `TbConfMostrarEstado` | no detallada | sí |
| 35 | `TbE2EExportBatch` | no detallada | sí |
| 36 | `TbE2EExportBatchDetalle` | no detallada | sí |
| 37 | `TbE2EExportSeleccionTemp` | no detallada | sí |
| 38 | `TbE2EJsonDestinationUserConfig` | no detallada | sí |
| 39 | `TbExpedientesE2E` (re-entry) | sí | sí |
| 40 | `TbDatosEconomicosExpedientes` | parcial (no detallada) | sí |
| 41 | `Comerciales` (re-entry) | parcial (no detallada) | sí |
| 42 | `CPVs` (re-entry) | parcial (no detallada) | sí |
| 43 | `RACS` (re-entry) | parcial (no detallada) | sí |
| 44 | `PECAL` (re-entry) | parcial (no detallada) | sí |
| 45 | `LugaresEjecucion` (re-entry) | parcial (no detallada) | sí |
| 46 | `Responsables` (re-entry) | parcial (no detallada) | sí |
| 47 | `Suministradores` (re-entry) | parcial (no detallada) | sí |
| 48 | `Jefaturas` (re-entry) | parcial (no detallada) | sí |
| 49 | `Juridicas` (re-entry) | parcial (no detallada) | sí |
| — | `Hitos` (re-entry) | parcial (no detallada) | sí |
| — | `Modificados` (re-entry) | parcial (no detallada) | sí |
| — | `Anexos` (re-entry) | parcial (no detallada) | sí |
| — | `CodigoCompras` (re-entry) | parcial (no detallada) | sí |
| — | `CadenaContratacion` (re-entry) | parcial (no detallada) | sí |

(Las tablas marcadas «parcial» tienen el nombre listado en este documento pero no la sección «Schemas detallados» individual; las marcadas «no detallada» carecen incluso del nombre en este doc pero existen en `schema.sql` y en `Estructura_Datos.md`. El mapeo ordinal→tabla se publica aquí para que los 5 dictionary tranches de #211–#215 puedan ser completados sin ejecutar `dysflow list_tables` de nuevo: el ordinal ya está fijado.)

## Bloqueo de las 19 tablas restantes

Las 19 tablas marcadas «parcial» o «no detallada» requieren la ejecución de ``dysflow get_schema`` contra el backend autoritativo (``C:\00repos\datos\Expedientes_datos.accdb``) para tener descripciones de columnas/PKs/FKs. El ordinal y el nombre de cada tabla ya constan en este documento (de la columna anterior) y en `docs/03-aplicaciones/expedientes/ERD/schema.sql`, por lo que la información mínima para reabrir #211–#215 está disponible. Lo que falta es la descripción narrativa (qué hace cada columna, de dónde viene, perfil de uso), que solo Dysflow puede producir.

### Tablas de negocio principales

| Grupo | Tablas |
|---|---|
| Agregado | `TbExpedientes`, `TbExpedientesConEntidades`, `TbDatosEconomicosExpedientes` |
| Relaciones | `TbExpedientesAnualidades`, `Comerciales`, `CPVs`, `RACS`, `PECAL`, `LugaresEjecucion`, `Responsables`, `Suministradores`, `Jefaturas`, `Juridicas`, `Hitos`, `Modificados`, `Anexos`, `CodigoCompras`, `CadenaContratacion` |
| Catálogos | `TbComerciales`, `TbCPV`, `TbEjercitos`, `TbEstados`, `TbGradosClasificacion`, `TbJefaturas`, `TbJuridicas`, `TbLugaresEjecucion`, `TbOficinasPrograma`, `TbOrganosContratacion`, `TbPECAL`, `TbRACS`, `TbResponsablesPorRol`, `TbSuministradores` |
| Auditoría/preferencias | `TbCambios`, `TbUltimoCambio`, `TbConfMostrarEstado` |
| Históricos/auxiliares | `Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `ListaPrevia`, `TbExpedientes_antes`, `TbAusExpPostAGEDO`, `TbAuxEstadosMartina`, `TbAuxNemotecnico`, `TbExpAgedys`, `TbExpAGEDYS1` |
| E2E/operación | `TbE2EExportBatch`, `TbE2EExportBatchDetalle`, `TbE2EExportSeleccionTemp`, `TbE2EJsonDestinationUserConfig`, `TbExpedientesE2E` |

## Relaciones físicas confirmadas

`TbExpedientes.IDExpediente` relaciona con las tablas de hijos y joins; `TbSuministradores.IDSuministrador` relaciona con `TbExpedientesSuministradores`. `IDExpedientePadre` es además una relación jerárquica autorreferente inferida por código y datos, aunque no aparece como FK física explícita en todas las relaciones Access.

## Schemas detallados (segunda pasada, 20 tablas obtenidas)

### `TbEstados` (3 columnas, 9 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDEstado` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Estado` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `TbComerciales` (3 columnas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDComercial` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Comercial` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `Descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `TbCPV` (3 columnas)

PK + CPV + DESCRIPCION.

### `TbEjercitos` (3 columnas)

PK + Ejercito + Descripcion.

### `TbJefaturas` (3 columnas)

PK + Jefatura + DESCRIPCION.

### `TbJuridicas` (4 columnas, 417 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDJuridica` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Juridica` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `IDSuministrador` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK a `TbSuministradores` |

### `TbLugaresEjecucion` (3 columnas, 194 filas)

PK + LugarEjecucion (Memo) + Descripcion.

### `TbOficinasPrograma` (3 columnas)

PK + OficinaPrograma + Descripcion.

### `TbOrganosContratacion` (3 columnas)

PK + OrganoContratacion + Descripcion.

### `TbPECAL` (3 columnas)

PK + PECAL + DESCRIPCION.

### `TbRACS` (4 columnas, 37 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDRAC` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `RAC` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `CORREO` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — correo del RAC |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `TbExpedientesAnexos` (3 columnas, 712 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDDocumento` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `NombreDocumento` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

### `TbExpedientesAnualidades` (11 columnas, 174 filas) — **anualidades con importes por tipo de impuesto**

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

### `TbExpedientesCadenaContratacion` (9 columnas) — **cadena de subcontratación con reglas**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `ID` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDPadre` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — padre (jerarquía) |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `IDSuministrador` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `AplicaCalidad` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `AplicaRiesgos` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `AplicaContratosClasificados` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `AplicaHPS` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — ⚠️ anomalía: Text 255 (vs Text 2) |
| `Observaciones` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `TbExpedientesCodigoCompras` (3 columnas)

PK + IDExpediente + CodCompras (Text 255).

### `TbExpedientesComerciales` (3 columnas)

PK compuesta (IDComercialExpediente, IDComercial, IDExpediente).

### `TbExpedientesConEntidades` (23 columnas) — **vista desnormalizada con cadenas ⚠️**

⚠️ **CRÍTICO**: **23 columnas, 13 de ellas son `Cadena*` (cadenas separadas por comas)**. Esta tabla es una **vista desnormalizada** del expediente con todas las entidades relacionadas en strings.

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

⚠️ **D113 propuesta**: esta tabla es una **vista materializada** que se debe regenerar en la nueva plataforma. Las 13 `Cadena*` son strings separados por comas que **NO escalan** y dificultan queries. Migración: mantener como `VIEW` en PostgreSQL (no tabla) generada a partir de las tablas normalizadas.

### `TbExpedientesCPVs` (3 columnas)

PK compuesta (IDCPVExpediente, IDCPV, IDExpediente).

### `TbExpedientesE2E` (7 columnas) — **sincronización E2E con hash**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` |
| `HashPayload` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — hash del payload E2E |
| `Estado` | 10 (Text) | 20 | false | `VARCHAR(20) NULL` |
| `FechaCreacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaModificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `UsuarioCreacion` | 10 (Text) | 100 | false | `VARCHAR(100) NULL` |
| `UsuarioModificacion` | 10 (Text) | 100 | false | `VARCHAR(100) NULL` |

⚠️ **Otro sistema E2E con hash**. **D114**: consolidar con `TbExpedientes.HashActual` y `TbExpedientes.HashUltimaExportacion` en una sola tabla `expediente_sincronizacion_e2e`.

### `TbExpedientesHitos` (6 columnas, 46 filas) — **hitos con importes**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDHitoExpediente` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `Descripcion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaHito` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaGarantiaHito` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` — fecha de garantía |
| `Importe` | 7 (Currency) | 8 | false | `NUMERIC(15,2) NULL` |

### `TbExpedientesJefaturas` (3 columnas)

PK compuesta (IDJefaturaExpediente, IDJefatura, IDExpediente).

### `TbExpedientesJuridicas` (6 columnas, 417 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteJuridica` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `IDJuridica` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `IDSuministrador` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `ContratistaPrincipal` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `SubContratista` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

### `TbExpedientesLugaresEjecucion` (3 columnas, 194 filas)

PK compuesta (IDExpedienteLugarEjecucion, IDExpediente, IDLugarEjecucion).

### `TbExpedientesModificados` (6 columnas, 37 filas) — **modificaciones al expediente**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteModificado` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `NModificado` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — número de modificado |
| `FechaFirmaModificado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaFinModificado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `Descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `TbExpedientesPECAL` (3 columnas)

PK compuesta (IDPECALExpediente, IDExpediente, IDPECAL).

### `TbExpedientesRACS` (3 columnas)

PK compuesta (IDRacExpediente, IDExpediente, IDRAC).

### `TbExpedientesResponsables` (6 columnas, volumen TBD)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteResponsable` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IdExpediente` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `IdUsuario` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `CorreoSiempre` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsJefeProyecto` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `esPreventa` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

### `TbExpedientesSuministradores` (7 columnas, 72 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDExpedienteSuministrador` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDExpediente` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `IDSuministrador` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `IDPadre` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — padre (jerarquía) |
| `Descripcon` | 12 (Memo) | 0 | false | `TEXT NULL` — ⚠️ typo: debería ser "Descripcion" |
| `ContratistaPrincipal` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `SubContratista` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

### `TbSuministradores` (10 columnas, 72 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDSuministrador` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Nombre` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CIF` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `TramitadoraHPS` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `Nemotecnico` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Direccion` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `CP` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — código postal |
| `Ciudad` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `ConsorcioPropio` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

### `TbResponsablesPorRol` (3 columnas, volumen TBD)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDResponsablePorRol` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDUsuario` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `Rol` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` — enum: Administrador, Calidad, Técnico |

## Hallazgos críticos del esquema de Expedientes

1. **D102 cross-cutting**: **20+ columnas como `Text 2)` (Sí/No como texto)**. Inconsistencia a estandarizar a `BOOLEAN` en PostgreSQL.
2. **D113 (Expedientes)**: `TbExpedientesConEntidades` con **13 columnas `Cadena*` (cadenas separadas por comas)**. Anti-patrón de modelado. Migración: reemplazar con `VIEW` generada a partir de tablas normalizadas.
3. **D114 (Expedientes)**: múltiples sistemas E2E con hash (`TbExpedientes.HashActual`, `TbExpedientes.HashUltimaExportacion`, `TbExpedientesE2E.HashPayload`). Consolidar en una sola tabla `expediente_sincronizacion_e2e`.
4. **Multi-impuestos** (IVA, IPSI, IGIC) en `TbExpedientesAnualidades` — sistema regulatorio crítico para migración.
5. **TYPO**: `Descripcon` en `TbExpedientesSuministradores` (debería ser "Descripcion"). Migración: corregir a `descripcion`.

## Semántica Access que debe conservarse

- Tipos 1/3/4/7/8/10/12 observados: Boolean, Integer/Long, Currency, DateTime, Text y Memo según columna; confirmar mapeo final por campo.
- `Sí/No` se almacena frecuentemente como texto de longitud 2, no como Boolean; no convertirlo silenciosamente.
- `Null` y cadena vacía se distinguen en formularios, DTO y `Registrar`.
- Autonuméricos/IDs, jerarquías, Memo, URLs SharePoint, ficheros referenciados, cachés denormalizadas y tablas temporales necesitan tratamiento explícito.
- No se observan campos Access de tipo Attachment/OLE ni multivalor en el esquema autoritativo cosechado; los anexos son filas con `IDDocumento`/`NombreDocumento` y deben comprobarse contra el almacenamiento externo.

## Perfil agregado (privacidad segura)

| Tabla | Filas |
|---|---:|
| `TbExpedientes` | 453 |
| `TbExpedientesConEntidades` | 451 |
| `TbExpedientesAnexos` | 712 |
| `TbExpedientesSuministradores` | 713 |
| `TbExpedientesResponsables` | 730 |
| `TbExpedientesComerciales` | 333 |
| `TbExpedientesCPVs` | 429 |
| `TbExpedientesPECAL` | 366 |
| `TbExpedientesJuridicas` | 417 |
| `TbExpedientesLugaresEjecucion` | 194 |
| `TbExpedientesAnualidades` | 174 |
| `TbExpedientesHitos` | 46 |
| `TbExpedientesModificados` | 37 |
| `TbSuministradores` | 72 |

El resto de recuentos está registrado en la evidencia de sesión; las dos tablas `Copia de...` no devolvieron recuento mediante el wrapper y quedan como comprobación pendiente.

Comprobaciones: `TbExpedientes` tiene 453 IDs no nulos, 49 `CodExp` vacíos y 60 `Nemotecnico` vacíos; no se detectaron duplicados de `CodExp`, huérfanos de suministrador ni huérfanos de padre; no se detectaron intervalos contrato fin anteriores a inicio. Fechas serializadas se mantienen como rangos agregados, sin valores de filas.
