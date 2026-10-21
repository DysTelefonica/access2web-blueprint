# Matriz campo a campo — tablas definidas solo en `src`

> Parte de la [matriz campo a campo](../matriz-campos.md). Commit fijado `ec6a8c3`.

Estas tablas no están en el `ERD/schema.sql` del snapshot: `M_E2EMigration.bas` las crea por DAO. A diferencia del resto de la matriz, aquí `src` sí declara nulabilidad, clave primaria e índices, de modo que la fila es evidencia directa del DDL (`:línea`). No se declaran claves foráneas: el comentario deja su creación al test o fixture y recuerda que las tablas pueden estar en backends distintos (`M_E2EMigration.bas:92-94, 130`); `EnsureRelation` existe (`:327-356`) pero ningún módulo de producción la invoca.

## TbE2EEnlaces

Historial de enlaces E2E por expediente: un episodio por alta y baja (`M_E2EMigration.bas:14-16`).

| Origen | Tipo origen (DAO) | Admite nulos | Evidencia (DDL en `src`) | Destino | Tipo objetivo |
|---|---|---|---|---|---|
| `TbE2EEnlaces.IDEnlace` | Auto  · PK | No | `M_E2EMigration.bas:76` | `expedientes.TbE2EEnlaces.IDEnlace` | BIGINT (identidad) |
| `TbE2EEnlaces.IDExpediente` | Long | No | `M_E2EMigration.bas:77` | `expedientes.TbE2EEnlaces.IDExpediente` | BIGINT |
| `TbE2EEnlaces.OrdinalE2E` | Long | Sí | `M_E2EMigration.bas:78` | `expedientes.TbE2EEnlaces.OrdinalE2E` | BIGINT |
| `TbE2EEnlaces.FechaAlta` | Date | No | `M_E2EMigration.bas:79` | `expedientes.TbE2EEnlaces.FechaAlta` | TIMESTAMP |
| `TbE2EEnlaces.FechaBaja` | Date | Sí | `M_E2EMigration.bas:80` | `expedientes.TbE2EEnlaces.FechaBaja` | TIMESTAMP |
| `TbE2EEnlaces.IDUsuarioAlta` | Long | Sí | `M_E2EMigration.bas:81` | `expedientes.TbE2EEnlaces.IDUsuarioAlta` | BIGINT |
| `TbE2EEnlaces.IDUsuarioBaja` | Long | Sí | `M_E2EMigration.bas:82` | `expedientes.TbE2EEnlaces.IDUsuarioBaja` | BIGINT |
| `TbE2EEnlaces.MotivoBaja` | Text(255) | Sí | `M_E2EMigration.bas:83` | `expedientes.TbE2EEnlaces.MotivoBaja` | VARCHAR(255) |

Índices: clave primaria sobre `IDEnlace` y secundarios sobre `IDExpediente`, `OrdinalE2E` y `FechaAlta` (`M_E2EMigration.bas:87-90`).

## TbE2ECola

Cola de cambios pendientes que consulta un script del servidor (`M_E2EMigration.bas:17-21`). Las tres últimas columnas se añaden después, en `ProvisionarTbE2EColaVerificacion`.

| Origen | Tipo origen (DAO) | Admite nulos | Evidencia (DDL en `src`) | Destino | Tipo objetivo |
|---|---|---|---|---|---|
| `TbE2ECola.IDCola` | Auto  · PK | No | `M_E2EMigration.bas:112` | `expedientes.TbE2ECola.IDCola` | BIGINT (identidad) |
| `TbE2ECola.IDExpediente` | Long | No | `M_E2EMigration.bas:113` | `expedientes.TbE2ECola.IDExpediente` | BIGINT |
| `TbE2ECola.OrdinalE2E` | Long  · índice único | No | `M_E2EMigration.bas:114` | `expedientes.TbE2ECola.OrdinalE2E` | BIGINT |
| `TbE2ECola.HashDetectado` | Text(64) | Sí | `M_E2EMigration.bas:115` | `expedientes.TbE2ECola.HashDetectado` | VARCHAR(64) |
| `TbE2ECola.FechaDeteccion` | Date | No | `M_E2EMigration.bas:116` | `expedientes.TbE2ECola.FechaDeteccion` | TIMESTAMP |
| `TbE2ECola.TipoOrigen` | Text(20) | No | `M_E2EMigration.bas:117` | `expedientes.TbE2ECola.TipoOrigen` | VARCHAR(20) |
| `TbE2ECola.Estado` | Text(20) | No | `M_E2EMigration.bas:118` | `expedientes.TbE2ECola.Estado` | VARCHAR(20) |
| `TbE2ECola.FechaProcesado` | Date | Sí | `M_E2EMigration.bas:119` | `expedientes.TbE2ECola.FechaProcesado` | TIMESTAMP |
| `TbE2ECola.IDUsuarioOrigen` | Long | Sí | `M_E2EMigration.bas:120` | `expedientes.TbE2ECola.IDUsuarioOrigen` | BIGINT |
| `TbE2ECola.JSONGenerado` | Memo | Sí | `M_E2EMigration.bas:121` | `expedientes.TbE2ECola.JSONGenerado` | TEXT |
| `TbE2ECola.EstadoImportacion` | Text(20) | No | `M_E2EMigration.bas:406, 452` | `expedientes.TbE2ECola.EstadoImportacion` | VARCHAR(20) |
| `TbE2ECola.FechaVerificacion` | Date | Sí | `M_E2EMigration.bas:413` | `expedientes.TbE2ECola.FechaVerificacion` | TIMESTAMP |
| `TbE2ECola.MotivoFallo` | Text(255) | Sí | `M_E2EMigration.bas:420` | `expedientes.TbE2ECola.MotivoFallo` | VARCHAR(255) |

Índices: clave primaria sobre `IDCola`; secundarios sobre `Estado`, `IDExpediente` y `FechaDeteccion` (`M_E2EMigration.bas:125-128`); índice único `UX_TbE2ECola_OrdinalE2E` sobre `OrdinalE2E` (`:269-273`). `OrdinalE2E` se rellena con `IDCola` en las filas previas (`:204-206`) y `EstadoImportacion` con «Pendiente» (`:430-432`) porque DAO no admite valores por defecto.

## Columnas de RC1 que el snapshot no tiene

| Origen | Evidencia | Observación |
|---|---|---|
| `TbExpedientesConEntidades.OrdinalE2E` | `ExpedienteEntidadOperaciones.cls:133` | `Fields("OrdinalE2E")` escrito; el `schema.sql` del snapshot solo lo tiene en `TbExpedientes` |
| `TbExpedientesConEntidades.EnlazadoE2E` | `ExpedienteEntidadOperaciones.cls:137` | Se escribe «Sí» o «No» según `E2EEnlacesService.EstaVinculado`; no existe en el snapshot |
| `TbConfiguracionBackends` | `ModFixtureF45CacheE2E.bas:43` | Tabla local del frontend (`BackendSandbox`, `PasswordBackend`) citada solo en un módulo de fixtures de prueba; no es dato de negocio ni se migra |

## Tablas vinculadas a otros backends

El `schema.sql` registra 24 tablas vinculadas que no se extrajeron porque su fichero no existía en la máquina de extracción. Se resumen por backend de destino; el detalle de columnas pertenece al estudio de cada aplicación.

| Backend | Tablas vinculadas | Aplicación propietaria |
|---|---|---|
| `AGEDYS_DATOS.accdb` | 14: `TbComunicados`, `TbDpDInformeCondicionamiento`, `TbExpedientesAnualidades1`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbNPedido`, `TbProyectos_AGEDYS`, `TbResponsablesExpedientes`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresSAP`, `TbVisadoFacturas_Nueva`, `TbVisadosGenerales` | AGEDYS (fuera de las ocho aplicaciones) |
| `Correos_datos.accdb` | 1: `TbCorreosEnviados` | correo compartido |
| `Expedientes_datos.accdb` | 1: `TbExpedientes1` | el propio backend de Expedientes |
| `Gestion_Riesgos_Datos.accdb` | 3: `TbGestionRiesgos`, `TbProyectosEdiciones_Riesgos`, `TbProyectos_Riesgos` | Gestión de Riesgos |
| `HPST.accdb` | 2: `TbUsuariosHistoricos_HPS`, `TbUsuarios_HPS` | HPS |
| `Lanzadera_Datos.accdb` | 2: `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | Lanzadera |
| `NoConformidades_Datos.accdb` | 1: `TbNoConformidades` | No Conformidades |

La propiedad es una inferencia por el nombre del fichero; no se ha comprobado. La tabla vinculada al propio `Expedientes_datos.accdb` indica un vínculo circular o heredado que el ERD de origen debe aclarar.
