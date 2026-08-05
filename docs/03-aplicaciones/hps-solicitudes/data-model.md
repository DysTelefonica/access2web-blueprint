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

Schemas de: `TbResponsables`, `TbJustificaciones`, `TbHPSGrado`, `TbConfiguracion`, `TbLogs`, `TbLogsGeneral`, `TbCorreosEnviados`, `TbSolicitudesFechas`, `TbUltimoCambio`, `Copia de TbExpedientes`. Se obtendrán en una iteración posterior.