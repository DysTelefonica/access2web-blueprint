# Matriz campo a campo — HPS Solicitudes

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Issue:** #1077 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Cada columna de origen de las 11 tablas locales del backend de HPS Solicitudes, con la evidencia de su uso en `src` y un candidato de tipo para la migración web. Ninguna fila está aprobada: el ERD de origen queda pendiente del propietario (D-HPS-1 del README), porque `dysflow` no está disponible en el entorno de trabajo y el backend del clon (`Solicitudes_HPS_datos.accdb`) no es legible en local. Esta primera entrega cubre 6 tablas core; la matriz-2 (próxima PR) cierra las 5 restantes y el anexo de las 9 tablas linked.

## Fuente y límites

- **Columnas de origen:** [`ERD/schema.sql`](../ERD/schema.sql), volcado del backend `Solicitudes_HPS_datos.accdb` del snapshot. 11 tablas locales, 122 columnas, volúmenes de fila por tabla (snapshot 2026-08-06): 70/1/9/13/7/0/2058/27/245/245/5. **No está revalidada contra el backend del commit fijado.**
- **Lo que falta en esa fuente:** solo columna, tipo Jet/ACE y `Rows`. No hay claves primarias, índices, relaciones, nulabilidad, valores por defecto ni FKs. Las linked tables (TbExpedientes, TbHPS, TbMotivoHPS, TbUsuarios, etc.) están marcadas como `LINKED/SKIPPED` en el volcado porque su fichero de datos enlazado no estaba presente; estas se llevan a la matriz-2.
- **Evidencia en `src`:** se busca cada columna en los ficheros de producción que citan la tabla (excluidos los `Test*`); lectura estática. Las tablas en `modulo.Metodo` se cuentan como evidencia de la capa que ejecuta SQL; las clases de dominio y operaciones confirman que la columna se materializa como propiedad.
- **Sin valores:** la matriz no contiene datos de ninguna tabla, solo estructura. `TbConfiguracion.Rows = 1` indica que el backend guarda una única fila (configuración global); `TbLogsGeneral.Rows = 2058` indica el volumen de auditoría y merece un análisis de retención antes de la migración.
- **Cadenas concatenadas de persona/datos:** `TbSolicitudes.Gestor`, `emailResponsable`, `UsuarioCreacion` y `TbLogsGeneral.DNI` cargan valores reales en el `.form.txt`; en `src/` solo se tratan como `String` y se filtran en las fichas. Migración: anonimizar o depurar antes del corte (decisión abierta, ver D-HPS-3 del README).

## Reglas de mapeo de tipos (propuesta)

| Tipo de origen | Tipo objetivo propuesto | Decisión abierta |
|---|---|---|
| `LONG` | `INTEGER` | Identidad: el backend actual usa `Long(4)` autoincremental con `Max+1`; la web usa otra convención (UUID o `BIGINT`) — la correspondencia de claves entre Access y web no está definida |
| `INT` | `SMALLINT` | — |
| `TEXT(255)` | `VARCHAR(255)` | — |
| `TEXT(2)` | `BOOLEAN` (con CHECK) | `Sí`/`No` y otros dos-Estados en `TbConfiguracion`, `TbSolicitudes`, `TbLogsGeneral` |
| `MEMO` | `TEXT` | — |
| `BOOLEAN` | `BOOLEAN` | `activa` en `TbJustificaciones` |
| `SHORT_DATE_TIME` | `TIMESTAMP` (con `WITH TIME ZONE`) | Zona horaria y `Date` Access vs `TIMESTAMP` PostgreSQL |

## Tablas cubiertas en esta entrega (6)

- [TbConfiguracion](#tbconfiguracion)
- [TbCorreosEnviados](#tbcorrosenviados)
- [TbResponsables](#tbresponsables)
- [TbSolicitudes](#tbsolicitudes)
- [TbSolicitudesFechas](#tbsolicitudesfechas)
- [TbJustificaciones](#tbjustificaciones)

## Tablas pendientes para la matriz-2 (5 + anexo linked)

TbCopia de TbExpedientes, TbHPSGrado, TbLogs, TbLogsGeneral, TbUltimoCambio + anexo de 9 tablas linked (TbExpedientes, TbExpedientesLugaresEjecucion, TbHPS, TbHPSEquivalencia, TbHPSGrado1, TbMotivoHPS, TbUsuarios, TbUsuariosEntidades, TbUsuariosHistoricos). D-HPS-1: pendiente de extracción de solo lectura por el workflow `dysflow-erd` (#1144) desde la máquina Windows del propietario.

---

<a id="tbconfiguracion"></a>

### TbConfiguracion

Una única fila (`Rows: 1`). Configuración global de la aplicación: ventanas de cancelación automática de MARGA, recordatorios Excel/MARGA, parámetros de plantillas HTML/Excel y datos de seguridad. Cargada por `Entorno.cls` (composition root) y leída por toda la aplicación a través de globales (`m_ObjEntorno.Configuracion`).

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `ID` | `LONG(4)` | `INTEGER PK` | `classes/ConfiguracionOperaciones.cls:Registrar` (asigna `id`); `modules/constructor.bas:getConfiguracion` |
| `DiasParaRecordatorioExcel1` | `INT(2)` | `SMALLINT` | `Form_FormConfiguracion.cls:74-77`; `classes/SolicitudFechasOperaciones.cls:1-50`; `Form_FormSolicitudesAltaFechas.cls:1600, 1625-1629` |
| `DiasParaRecordatorioExcel2` | `INT(2)` | `SMALLINT` | `Form_FormConfiguracion.cls:78-81`; `Form_FormSolicitudesAltaFechas.cls:1598` |
| `DiasCancelacionPreMARGA` | `INT(2)` | `SMALLINT` | `Form_FormConfiguracion.cls:82-85`; `Form_FormSolicitudesAltaFechas.cls:1592-1596` |
| `DiasParaRecordatorioRellenoMarga1` | `INT(2)` | `SMALLINT` | `Form_FormConfiguracion.cls:86-89` |
| `DiasParaRecordatorioRellenoMarga2` | `INT(2)` | `SMALLINT` | `Form_FormConfiguracion.cls:90-93` |
| `DiasCancelacionMARGA` | `INT(2)` | `SMALLINT` | `Form_FormConfiguracion.cls:94-97` |
| `BuzonSeguridad` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormConfiguracion.cls:99-100`; `Form_FormEnvioCorreo.cls:179-180` |
| `EmailDirectorSeguridad` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormConfiguracion.cls:101-104` |
| `CorreodeEnvio` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormConfiguracion.cls:105-108` (no referenciado en el subform de alta: candidato a omisión) |
| `AutocancelacionPreMARGA` | `TEXT(2)` | `BOOLEAN` (CHECK `IN ('Sí','No')`) | `Form_FormConfiguracion.cls:115-118` |
| `AutocancelacionMARGA` | `TEXT(2)` | `BOOLEAN` (CHECK `IN ('Sí','No')`) | `Form_FormConfiguracion.cls:119-122` |
| `CorreosAutomaticos` | `TEXT(2)` | `BOOLEAN` | `Form_FormConfiguracion.cls:127-130`; `Form_Form0BDOpciones.cls:230-256` (rampa de menú comentada) |
| `VersionPlantillasHTML` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormConfiguracion.cls:131-135`; `Form_FormCorreo.cls:296-302`; `Form_FormEnvioCorreo.cls:177-180` |
| `VersionPlantillasEXCEL` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormConfiguracion.cls:136-139`; `Form_FormEnvioCorreo.cls:223-225` |

Sin escrituras en clases/operaciones; lectura centralizada en `m_ObjEntorno` y escritura exclusiva desde `Form_FormConfiguracion` y `ConfiguracionOperaciones.Registrar`. Migración: una sola fila con claves de caché (`m_ObjEntorno.Configuracion = Nothing`) tras guardar; en la web se sustituye por `GET /configuracion` y `PUT /configuracion/{id}`.

---

<a id="tbcorrosenviados"></a>

### TbCorreosEnviados

Cola de correos (`Rows: 9` en el snapshot, dinámico). Persistida por `CorreoOperaciones.Registrar` (vía `Form_FormCorreo.ComandoRegistrar_Click`, `Form_FormEnvioCorreo.Registrar`, `Form_FormCorreoPruebas.cmdEnviarPrueba_Click`).

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `IDCorreo` | `LONG(4)` | `INTEGER PK` | `Form_FormCorreo.cls:241-249`; `classes/Correo.cls:1-50` |
| `URLAdjunto` | `MEMO` | `TEXT` | `Form_FormCorreo.cls:163`; `Form_FormEnvioCorreo.cls:223-225` |
| `Aplicacion` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormCorreo.cls:160`; `modules/constructor.bas:EnviarCorreoGenerico` |
| `Destinatarios` | `MEMO` | `TEXT` | `Form_FormCorreo.cls:88`; `Form_FormEnvioCorreo.cls:179` |
| `DestinatariosConCopia` | `MEMO` | `TEXT` | `Form_FormCorreo.cls:90`; `Form_FormEnvioCorreo.cls:180` |
| `DestinatariosConCopiaOculta` | `MEMO` | `TEXT` | `Form_FormCorreo.cls:158` (sin columna de UI en el `.form.txt` actual) |
| `Asunto` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormCorreo.cls:85`; `Form_FormEnvioCorreo.cls:178` |
| `FechaEnvio` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormCorreo.cls:100`; `Form_FormCorreosGestion.cls:476` (IsDate gate) |
| `FechaOrdenEnvio` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormCorreo.cls:91-99` |
| `FechaGrabacion` | `SHORT_DATE_TIME` | `TIMESTAMP NOT NULL DEFAULT NOW()` | (no referenciada en `src/`; la fija `Registrar` o la BD) |
| `NombrePlantilla` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormCorreo.cls:278-294` |
| `VersionPlantilla` | `TEXT(255)` | `VARCHAR(255)` | (no referenciada en `src/`; ver `VersionPlantillasHTML` de `TbConfiguracion`) |
| `CadenaRecursos` | `MEMO` | `TEXT` | `Form_FormCorreo.cls:165`; `Form_FormCorreoPruebas.cls:215-220` |
| `IDSolicitud` | `LONG(4)` | `INTEGER FK` | `Form_FormCorreo.cls:305-309`; `Form_FormEnvioCorreo.cls:175-176` (FK implícita) |
| `Accion` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormCorreo.cls:251` (cargada, no persistida por este formulario) |
| `DesencadenadoPor` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormCorreo.cls:170` (clon por reenvío) |
| `Programado` | `TEXT(2)` | `BOOLEAN` | `Form_FormCorreo.cls:91-99` (set `Sí`/`No`) |
| `TipoCorreo` | `LONG(4)` | `INTEGER FK` (o `SMALLINT` con CHECK) | `Form_FormEnvioCorreo.cls:186-200` (mapeo a `EnumTipoEnvioCorreo`) |
| `Observaciones` | `MEMO` | `TEXT` | (no referenciada en `src/`) |
| `Intentos` | `LONG(4)` | `INTEGER DEFAULT 0` | (no referenciada en `src/`) |
| `FechaProceso` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | (no referenciada en `src/`) |

Migración: `GET /correos?estado=...&solicitud=...` paginado; `Destinatarios*` se mueven a una tabla `correos_enviados_destinatarios` (1-N); `TipoCorreo` se convierte en `SMALLINT` con un dominio fijo de 7 valores; `IDSolicitud` admite `NULL` para la fila de prueba (`CONST_IDSOLICITUD_PRUEBA = -1`).

---

<a id="tbresponsables"></a>

### TbResponsables

Catálogo de responsables (27 filas en el snapshot). CRUD a través de `Form_FormResponsablesGestion` y `Form_FormResponsable`. Relación con `TbSolicitudes` por email: `TbSolicitudes.emailResponsable` se vincula a `TbResponsables.Correo` (a verificar la FK física).

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `IDResponsable` | `LONG(4)` | `INTEGER PK` | `classes/Responsable.cls:1-50`; `ResponsableOperaciones.cls:Registrar` |
| `Nombre` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormResponsablesGestion.cls:128-148`; `Form_FormResponsable.cls:18-44` |
| `Correo` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormResponsablesGestion.cls:128-148`; `modules/constructor.bas:getSolicitudesConResponsable(p_Correo)` (vínculo por email) |
| `Descripcion` | `MEMO` | `TEXT` | `Form_FormResponsablesGestion.cls:128-148` |

Migración: la FK por `email` (no por `IDResponsable`) en `TbSolicitudes` se sustituye por FK numérica real a `responsables.id`. La columna `Descripcion` se conserva; la web la puede mostrar como descripción breve de la persona responsable.

---

<a id="tbsolicitudes"></a>

### TbSolicitudes

Tabla núcleo (245 filas en el snapshot). 28 columnas con datos personales y de estado. Cargada por `Solicitud.cls` y mutada por `SolicitudOperaciones.cls` desde `Form_FormSolicitudAlta`, `Form_FormSolicitudesAltaFechas`, `Form_FormSolicitudRenovacion`, `Form_FormSolicitudRenovacionDatos`, `Form_FormSolicitudesAltaTraspasoDatos`, `Form_FormSolicitudesTraspasoFechas`, `Form_FormSolicitudesAltaRenovacionFechas`. Relación 1-N con `TbSolicitudesFechas` por `IDSolicitud` (FK implícita). Vinculada a `TbExpedientes` por `IDExpediente` (a verificar la FK física) y a `TbHPS`/`TbUsuariosHPS` (no modelada en local; referencia a `D-HPS-1`).

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `IDSolicitud` | `LONG(4)` | `INTEGER PK` | `Form_FormSolicitudesGestion.cls:438-481`; `Solicitud.cls:1-50` |
| `Gestor` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormTareasTramitadorPendientes.cls:117-156` (dato personal embebido en `.form.txt`) |
| `DNI` | `TEXT(255)` | `VARCHAR(255)` (anonimizar en producción) | `TbLogsGeneral.DNI` (auditoría); `Form_FormAltaDatosSolicitante.cls:218-220, 292-296` (DNI/NIE sintético) |
| `Nombre` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormSolicitudAlta.cls:79-101`; `Form_FormExpedientesBusqueda.cls:233-238` |
| `Apellido1` | `TEXT(255)` | `VARCHAR(255)` | id. |
| `Apellido2` | `TEXT(255)` | `VARCHAR(255)` | id. |
| `FNacimiento` | `SHORT_DATE_TIME` | `DATE` | `Form_FormAltaDatosSolicitante.cls:218-220` |
| `LugarNacimiento` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormAltaDatosSolicitante.cls:218-220` |
| `email` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormSolicitudAlta.cls:79-101`; `Form_FormExpedientesBusqueda.cls:233-238` |
| `Telefono` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormAltaDatosSolicitante.cls:218-220` |
| `IDExpediente` | `LONG(4)` | `INTEGER FK` | `Form_FormExpedientesBusqueda.cls:233-238`; `Form_FormSolicitudesAltaTraspasoDatos.cls:169-176` |
| `IDEmpresaUsuario` | `LONG(4)` | `INTEGER FK` | `Form_FormAltaDatosSolicitante.cls:218-220` (vínculo empresa solicitante) |
| `IDEmpresaTramitadora` | `LONG(4)` | `INTEGER FK` | `Form_FormAltaDatosSolicitante.cls:218-220` |
| `URLAdjunto` | `MEMO` | `TEXT` | `Form_FormSolicitudRenovacionDatos.cls:340-350`; `Form_FormSolicitudesAltaFechas.cls:1598` |
| `FechaCreacion` | `SHORT_DATE_TIME` | `TIMESTAMP NOT NULL DEFAULT NOW()` | `SolicitudOperaciones.cls:1-50` (en INSERT) |
| `FechaUltimoCambio` | `SHORT_DATE_TIME` | `TIMESTAMP NOT NULL` | `SolicitudOperaciones.cls:1-50`; `Form_FormTareasTramitadorPendientes.cls:494-499` (literal fecha+hora) |
| `UsuarioCreacion` | `TEXT(255)` | `VARCHAR(255)` (FK a `users.id`) | `Form_FormAltaDatosSolicitante.cls:218-220`; `SolicitudOperaciones.cls:1-50` |
| `UsuarioUltimoCambio` | `TEXT(255)` | `VARCHAR(255)` (FK a `users.id`) | `SolicitudOperaciones.cls:1-50` |
| `emailResponsable` | `TEXT(255)` | `VARCHAR(255)` FK a `responsables.correo` | `Form_FormSolicitudRenovacionDatos.cls:340-350`; `SolicitudOperaciones.cls:1-50` |
| `Estado` | `TEXT(255)` | `VARCHAR(255)` (con CHECK sobre dominio) | `Form_FormSolicitudesGestion.cls:208-216`; `Form_FormExpedientesBusqueda.cls:233-238` |
| `Observaciones` | `MEMO` | `TEXT` | `Form_FormSolicitudRenovacionDatos.cls:340-350` |
| `Tipo` | `TEXT(255)` | `VARCHAR(255)` (Alta / Traspaso / Renovación) | `Form_FormSolicitudesGestion.cls:113-119`; `Form_FormEnvioCorreo.cls:113-119` (conmutación padre) |
| `CorreosAutomaticos` | `TEXT(2)` | `BOOLEAN` | `Form_FormAltaDatosSolicitante.cls:218-220` |
| `Motivo_HPS` | `MEMO` | `TEXT` | `Form_FormMotivoHPS.cls:22-26`; `Form_FormMotivosHPS.cls:22-26` |
| `IDUsuarioHPS` | `LONG(4)` | `INTEGER FK` (a `TbUsuarios` linked) | `Form_FormSolicitudRenovacionDatos.cls:640-643`; `Form_FormEnvioCorreo.cls:175-176` |
| `URLAdjuntoEnvioONS` | `MEMO` | `TEXT` | `Form_FormSolicitudesTraspasoFechas.cls:678-681`; `SolicitudOperaciones.RegistroEnvioTraspasoONS` (UPDATE) |
| `Nemotecnico` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormExpedientesBusqueda.cls:233-238`; `Form_FormExpedienteDetalle.cls:54-69` |
| `idjustificacion` | `LONG(4)` | `INTEGER FK` | `Form_FormAltaDatosSolicitante.cls:218-220`; `Form_FormJustificacion.cls:22-26` |

Migración: varias columnas de datos personales (DNI, email, teléfono) requieren decisión de anonimización o seudonimización en la web; `emailResponsable` se convierte en FK numérica real a `responsables.id`; `IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `IDUsuarioHPS`, `idjustificacion` se confirman como FKs (D-HPS-1 las define formalmente).

---

<a id="tbsolicitudesfechas"></a>

### TbSolicitudesFechas

Relación 1-1 con `TbSolicitudes` por `IDSolicitud` (245 filas en el snapshot, una por solicitud). 20 columnas de fechas para los hitos del workflow HPS (envío Excel, recepción, MARGA, DPS, ONS, registro HPS, recordatorios, desestimación, autocancelación, traspaso ONS). Mutada por `SolicitudFechasOperaciones.cls` desde los 7 hitos de `Form_FormEnvioCorreo` y desde los handlers `ComandoAdjuntarExcel*Click`.

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `IDSolicitud` | `LONG(4)` | `INTEGER PK FK` | `Form_FormEnvioCorreo.cls:186-200`; `SolicitudFechasOperaciones.cls:1-50` |
| `FechaEnvioExcel` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200`; `SolicitudOperaciones.RegistroEnvioExcel` |
| `FechaRecepcionExcel` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200`; `SolicitudOperaciones.RegistroEnvioExcel` |
| `FechaTramitacionAltaMarga` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200`; `SolicitudOperaciones.RegistroEnvioAutocancelaconMARGA` |
| `FechaEnvioDPS` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200`; `SolicitudOperaciones.RegistroEnvioAutocancelaconPREMARGA` |
| `FechaEnvioONS` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200`; `SolicitudOperaciones.RegistroEnvioAutocancelaconPREMARGA` |
| `FechaRegistroEnHPS` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200`; `SolicitudOperaciones.RegistroEnvioAutocancelaconMARGA` |
| `FechaCorreoRecordatorioExcel1` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200` |
| `FechaCorreoRecordatorioExcel2` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaCorreoRecordatorioRellenoMarga1` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaCorreoRecordatorioRellenoMarga2` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaDesestimado` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaAutocancelacion` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaPrevistaCorreoRecordatorioExcel1` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaPrevistaCorreoRecordatorioExcel2` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaPrevistaCorreoRecordatorioRellenoMarga1` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaPrevistaCorreoRecordatorioRellenoMarga2` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaPrevistaAutocancelacionPreMarga` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaPrevistaAutocancelacionMarga` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | id. |
| `FechaCancelado` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `Form_FormEnvioCorreo.cls:186-200` (seleccionada por el `Select Case`) |
| `FechaEnvioTraspasoONS` | `SHORT_DATE_TIME` | `TIMESTAMP NULL` | `SolicitudOperaciones.RegistroEnvioTraspasoONS` |

Migración: tabla con 20 columnas `TIMESTAMP NULL`; el modelo se sustituye por una tabla `solicitud_hitos` con filas (`solicitud_id`, `tipo_hito`, `fecha_real`, `fecha_prevista`) que normaliza los 20 pares a dos columnas. La decisión está abierta (D-HPS-2 del README).

---

<a id="tbjustificaciones"></a>

### TbJustificaciones

Catálogo de motivos de justificación (7 filas en el snapshot). CRUD a través de `Form_FormJustificacion` y `Form_FormBuscarJustificacion`. Vinculada a `TbSolicitudes` por `idjustificacion`.

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `idjustificacion` | `LONG(4)` | `INTEGER PK` | `modules/JustificacionRepositorio.bas:1-50`; `Solicitud.cls:idjustificacion` (propiedad) |
| `titulo` | `TEXT(255)` | `VARCHAR(255)` | `Form_frmJustificacion.cls:7-13`; `Form_frmBuscarJustificacion.cls:13-17` |
| `descripcion` | `MEMO` | `TEXT` | `Form_frmJustificacion.cls:7-13` |
| `activa` | `BOOLEAN` | `BOOLEAN` | `Form_frmBuscarJustificacion.cls:13-17` (filtro) |

Migración: `GET /justificaciones?activa=...&q=...`; índice `GIN` sobre `titulo` para búsqueda rápida. Sin secretos.

## Pendientes

- Matriz-2 (próxima PR): 5 tablas locales restantes (`Copia de TbExpedientes`, `TbHPSGrado`, `TbLogs`, `TbLogsGeneral`, `TbUltimoCambio`) + anexo de 9 tablas linked (D-HPS-1: pendiente del workflow `dysflow-erd` en Windows del propietario).
- Confirmar claves primarias, índices, nulabilidad y FKs físicas en cada columna (D-HPS-1).
- Decisión de `Sí`/`No` en `TEXT(2)` (4 columnas en `TbConfiguracion`/`TbSolicitudes`/`TbLogsGeneral`): `BOOLEAN` con CHECK `IN ('Sí','No')` o dominio `VARCHAR(8)` con constraint; la elección define el contrato de la API web.
- Decisión de fechas de `TbSolicitudesFechas`: 20 columnas aplanadas frente a tabla `solicitud_hitos` (D-HPS-2).
- Política de anonimización/seudonimización para `DNI`, `email`, `emailResponsable`, `Gestor` (datos personales en `TbSolicitudes`) y de auditoría en `TbLogsGeneral.DNI`/`Usuario` (D-HPS-3).
- Cobertura de `dbOpenSnapshot`, `dbOpenDynamic`, `dbPessimistic`, `dbOptimistic`: la coexistencia debe documentar la diferencia con la semántica de PostgreSQL `READ COMMITTED` (D-HPS-4; ver `coexistencia.md`).
- Mapeo de `IDExpediente` ↔ tabla linked `TbExpedientes`: una `TbExpedientes_datos.accdb` aparte, sin `dysflow` local.
- Estructura de las tablas `TbHPSEquivalencia` y `TbHPSGrado1`: volcadas como `LINKED/SKIPPED`, contenido y FK pendientes.
- Reemplazo de las 9 linked: ¿se migran como linked en la web o se consolidan en una sola BD?
