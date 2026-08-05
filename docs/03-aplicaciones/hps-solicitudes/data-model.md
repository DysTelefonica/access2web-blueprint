# HPS_Solicitudes — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: **`C:\00repos\datos\Solicitudes_HPS_datos.accdb`** (13 MB) — backend autoritativo. **Inventario real obtenido vía Dysflow MCP el 2026-08-05** con `accessPath` + `backendPath` absolutos. Confirmado que el `Solicitudes_HPS_datos.accdb` local en la raíz del repo tiene la misma estructura y volumen que el autoritativo (el repo lee del autoritativo vía `TbConfiguracionBackends`).

## Inventario real Dysflow (2026-08-05, backend autoritativo)

| Categoría | Resultado |
|---|---|
| **Tablas totales** | **11** en `Solicitudes_HPS_datos.accdb` |
| **Filas en `TbSolicitudes`** | **245** (en uso activo) |
| **Filas en `TbResponsables`** | **27** |
| **Filas en `TbJustificaciones`** | **7** |
| **Filas en `TbLogs`** | **0** (vacía; presumiblemente en desuso, reemplazada por `TbLogsGeneral`) |
| **Filas en `TbLogsGeneral`** | **2058** (log general, alto volumen) |
| **Filas en `TbCorreosEnviados`** | **9** |
| **Filas en `TbHPSGrado`** | **13** (catálogo) |
| **Filas en `TbSolicitudesFechas`** | **245** (1:1 con `TbSolicitudes`) |
| **Filas en `TbConfiguracion`** | (no contado; presumido 1 fila) |
| **Filas en `TbUltimoCambio`** | (no contado; presumido 245 filas) |
| **Filas en `Copia de TbExpedientes`** | (legacy; no contado) |
| **FK relationships** (user tables) | **3** entre tablas de usuario |

⚠️ **Hallazgo del volumen**: HPS_Solicitudes **está en uso activo** con **245 solicitudes** vs 1 de Condor. El backend NO está casi vacío — es un sistema en producción.

⚠️ **Convención de nombres**: las tablas usan `Tb` mayúscula + CamelCase (`TbSolicitudes`, `TbResponsables`, etc.). Consistente con la mayoría de las apps del ecosistema.

### Lista completa de las 11 tablas

```
Copia de TbExpedientes (legacy copy)
TbConfiguracion
TbCorreosEnviados
TbHPSGrado
TbJustificaciones
TbLogs (vacía)
TbLogsGeneral
TbResponsables
TbSolicitudes
TbSolicitudesFechas
TbUltimoCambio
```

## Schemas reales (Dysflow)

### `TbSolicitudes` (28 columnas, 245 filas) — **cabecera con datos personales**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDSolicitud` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Gestor` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `DNI` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal (D98)** |
| `Nombre` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Apellido1` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Apellido2` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `FNacimiento` | 8 (DateTime) | 8 | false | `DATE NULL` — **dato personal** |
| `LugarNacimiento` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `email` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Telefono` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a Lanzadera `TbExpedientes` |
| `IDEmpresaUsuario` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual |
| `IDEmpresaTramitadora` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual |
| `URLAdjunto` | 12 (Memo) | 0 | false | `TEXT NULL` — ruta a fichero externo |
| `FechaCreacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaUltimoCambio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `UsuarioCreacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `UsuarioUltimoCambio` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `emailResponsable` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — ⚠️ **FK por email a `TbResponsables.Correo`** (no por ID) |
| `Estado` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Observaciones` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `Tipo` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — tipo de solicitud |
| `CorreosAutomaticos` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — Sí/No (booleano como texto ⚠️) |
| `Motivo_HPS` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `IDUsuarioHPS` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a HPS `TbUsuarios` |
| `URLAdjuntoEnvioONS` | 12 (Memo) | 0 | false | `TEXT NULL` — ruta a fichero para ONS |
| `Nemotecnico` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `idjustificacion` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `TbJustificaciones` ⚠️ |

⚠️ **Datos personales (D98)**: `DNI`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `Telefono` requieren la misma política que HPS (D92). **245 filas** con datos personales completos.

### `TbResponsables` (presumido 27 filas)

⚠️ Schema no inspeccionado en esta pasada. Pendiente para una segunda iteración.

### `TbJustificaciones` (7 filas)

⚠️ Schema no inspeccionado. Volumen bajo (7 filas).

### `TbHPSGrado` (13 filas) — **catálogo de grados HPS**

⚠️ Schema no inspeccionado. Catálogo con 13 grados.

### `TbConfiguracion` (volumen presumido 1)

⚠️ Schema no inspeccionado.

### `TbLogs` (0 filas) — **vacía, presumiblemente en desuso**

⚠️ Schema no inspeccionado. **0 filas** sugiere que esta tabla está en desuso y fue reemplazada por `TbLogsGeneral` (2058 filas).

### `TbLogsGeneral` (2058 filas) — **log general activo**

⚠️ Schema no inspeccionado. Volumen alto (2058 filas) confirma que es el log principal.

### `TbCorreosEnviados` (9 filas) — **registro de correos**

⚠️ Schema no inspeccionado.

### `TbSolicitudesFechas` (245 filas, 1:1 con TbSolicitudes)

⚠️ Schema no inspeccionado.

### `TbUltimoCambio` (volumen presumido 245)

⚠️ Schema no inspeccionado. Tracking de último cambio por solicitud.

### `Copia de TbExpedientes` (legacy copy)

⚠️ Pendiente para una segunda iteración. Patrón legacy de copia antes de cambios masivos.

## Relaciones físicas reales (3 FK entre user tables)

| Origen | Columna FK | Destino | Columna FK |
|---|---|---|---|
| `TbJustificaciones` | `idjustificacion` | `TbSolicitudes` | `idjustificacion` ⚠️ (dirección unusual) |
| `TbResponsables` | `Correo` | `TbSolicitudes` | `emailResponsable` ⚠️ **FK por email (texto), no por ID** |
| `TbSolicitudes` | `IDSolicitud` | `TbSolicitudesFechas` | `IDSolicitud` (1:N clásico) |

⚠️ **Inconsistencias detectadas en FKs**:

1. **`TbResponsables.Correo → TbSolicitudes.emailResponsable`**: FK por **texto email**, no por ID. Si un email cambia en `TbResponsables`, la FK lógica se rompe. Data integrity gap. **Migración**: en PostgreSQL, agregar columna `idResponsable` (FK numérica) y poblar la FK por email. Mantener el email como campo independiente.

2. **`TbJustificaciones.idjustificacion → TbSolicitudes.idjustificacion`**: la FK va de Justificaciones hacia Solicitudes, no al revés. La Solicitud tiene su propio `idjustificacion` y la Justificación la referencia de vuelta. **Confuso**, probablemente un error de diseño. Migración: formalizar como FK con la dirección correcta (`TbSolicitudes.idjustificacion → TbJustificaciones.id`).

3. **`TbSolicitudes.idExpediente`**: FK conceptual a Lanzadera `TbExpedientes`. Sin constraint. Cross-app.

4. **`TbSolicitudes.IDUsuarioHPS`**: FK conceptual a HPS `TbUsuarios`. Sin constraint. Cross-app.

5. **`TbSolicitudes.IDEmpresaUsuario` / `IDEmpresaTramitadora`**: FKs conceptuales. Sin constraint.

6. **NO hay FK** entre `TbSolicitudes` y `TbConfiguracion`, `TbHPSGrado`, `TbLogs`, `TbLogsGeneral`, `TbCorreosEnviados`, `TbUltimoCambio`. Tablas sueltas.

### Pendientes de discovery (segunda pasada)

Schemas de: `TbLogs` (vacía, no enumerable), `TbLogsGeneral` (falló en get_schema por colección vacía), `Copia de TbExpedientes` (legacy, presumida). Se obtendrán en una iteración posterior. **Los 7 schemas principales restantes ya están documentados** (ver a continuación).

## Schemas detallados (segunda pasada, 7 tablas)

### `TbResponsables` (4 columnas, 27 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDResponsable` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Nombre` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `Correo` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `Descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |

⚠️ **`Correo` es el campo clave de la FK conceptual** (D99: `TbResponsables.Correo → TbSolicitudes.emailResponsable`). Migración: agregar `idResponsable` numérica y mantener `Correo` para búsqueda.

### `TbJustificaciones` (4 columnas, 7 filas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idjustificacion` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `titulo` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `activa` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT TRUE` |

### `TbHPSGrado` (2 columnas, 13 filas) — **catálogo con clave compuesta**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `TipoHPS` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` — parte 1 de la clave compuesta |
| `Grado` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` — parte 2 de la clave compuesta |

PK compuesta `(TipoHPS, Grado)`. Migración: `PRIMARY KEY (tipo_hps, grado)` en PostgreSQL.

### `TbSolicitudesFechas` (21 columnas, 245 filas) — **workflow regulatorio de HPS con 20 fechas**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDSolicitud` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — PK + FK a `TbSolicitudes` |
| `FechaEnvioExcel` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaRecepcionExcel` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaTramitacionAltaMarga` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaEnvioDPS` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaEnvioONS` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaRegistroEnHPS` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCorreoRecordatorioExcel1` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCorreoRecordatorioExcel2` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCorreoRecordatorioRellenoMarga1` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCorreoRecordatorioRellenoMarga2` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaDesestimado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaAutocancelacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaPrevistaCorreoRecordatorioExcel1` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaPrevistaCorreoRecordatorioExcel2` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaPrevistaCorreoRecordatorioRellenoMarga1` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaPrevistaCorreoRecordatorioRellenoMarga2` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaPrevistaAutocancelacionPreMarga` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaPrevistaAutocancelacionMarga` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCancelado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaEnvioTraspasoONS` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |

⚠️ **CRÍTICO**: **20 fechas en 21 columnas**. Esto es un **workflow regulatorio completo de HPS** con eventos: envío/recepción de Excel, recordatorios, desestimación, autocancelación, traspasos a ONS. **Migración crítica**: cada fecha es un evento de workflow que debe preservarse como `TIMESTAMP NULL` en PostgreSQL. Decidir si se normaliza (consolidar fechas relacionadas) o se preserva tal cual.

### `TbConfiguracion` (15 columnas, 1 fila presumida)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `ID` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `DiasParaRecordatorioExcel1` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `DiasParaRecordatorioExcel2` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `DiasCancelacionPreMARGA` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `DiasParaRecordatorioRellenoMarga1` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `DiasParaRecordatorioRellenoMarga2` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `DiasCancelacionMARGA` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `BuzonSeguridad` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — buzón de seguridad |
| `EmailDirectorSeguridad` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CorreodeEnvio` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `AutocancelacionPreMARGA` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (inconsistencia D102) |
| `AutocancelacionMARGA` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (inconsistencia D102) |
| `CorreosAutomaticos` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (inconsistencia D102) |
| `VersionPlantillasHTML` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — versionado de plantillas HTML |
| `VersionPlantillasExcel` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — versionado de plantillas Excel |

⚠️ **CRÍTICO**: 3 campos de booleanos como `Text(2)` (inconsistencia D102). Migración: `BOOLEAN` en PostgreSQL.

⚠️ **CRÍTICO**: **6 campos de días** (`DiasParaRecordatorio*` y `DiasCancelacion*`) son config del workflow. Migración: a config del módulo + tabla de workflow declarativo (similar a D96 de Condor).

### `TbUltimoCambio` (4 columnas, 245 filas presumidas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `ID` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDSolicitud` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK a `TbSolicitudes` |
| `FechaCambio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `IDUsuarioCambio` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `TbUsuarios` (Lanzadera) |

### `TbCorreosEnviados` (21 columnas, 9 filas) — **sistema completo de correos**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDCorreo` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `URLAdjunto` | 12 (Memo) | 0 | false | `TEXT NULL` — ruta a fichero adjunto |
| `Aplicacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — aplicación origen |
| `Destinatarios` | 12 (Memo) | 0 | false | `TEXT NULL` — destinatarios (TO) |
| `DestinatariosConCopia` | 12 (Memo) | 0 | false | `TEXT NULL` — destinatarios (CC) |
| `DestinatariosConCopiaOculta` | 12 (Memo) | 0 | false | `TEXT NULL` — destinatarios (BCC) |
| `Asunto` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaEnvio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaOrdenEnvio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaGrabacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `NombrePlantilla` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `VersionPlantilla` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CadenaRecursos` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `IDSolicitud` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK a `TbSolicitudes` |
| `Accion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `DesencadenadoPor` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Programado` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (inconsistencia D102) |
| `TipoCorreo` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a un catálogo de tipos |
| `Observaciones` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `Intentos` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — número de intentos de envío |
| `FechaProceso` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |

⚠️ **CRÍTICO**: sistema completo de correos con adjuntos, destinatarios múltiples (TO/CC/BCC), plantillas, versionado, intentos de envío, programación. Migración: traducir a un **servicio de correo** server-side con la misma semántica. El campo `Programado` (Text 2) es inconsistencia D102.