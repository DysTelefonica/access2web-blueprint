# Capacidades de HPS Solicitudes frente al estudio previo

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Issue:** #1077 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Clasifica las 18 capacidades del [estudio previo](../capabilities.md) como **añadida**, **modificada**, **retirada** o **sin cambios** respecto al snapshot del 2026-08-06, con la evidencia del `src` y el enlace a las fichas de formulario. Las capacidades del estudio no tienen identificador; aquí se numeran como `C-HPS-S-1` a `C-HPS-S-18` solo para referirse a ellas.

## Método y límites

- La clasificación compara cada objeto de lógica con el mismo fichero del snapshot (`data/staging/hps-solicitudes/src`) tras quitar `Attribute VB_`, comentarios, líneas en blanco y diferencias de espacios y mayúsculas. **Resultado: los 74 ficheros de código (34 formularios, 14 módulos, 26 clases) son idénticos.** Por eso todas las capacidades salen «sin cambios en el código». Coincide con el estudio previo: 0 diferencias entre `staging` y `main` desde el commit fijado.
- «Sin cambios» es igualdad de texto: no prueba que la capacidad funcione, ni que el estudio previo la describiera bien, ni que el binario no versionado coincida. Las fichas contrastan el comportamiento y sus hallazgos se señalan en la última columna.
- La **disposición aprobada** es la del estudio previo (que este estudio no modifica).
- Las métricas de uso («Arranque e identidad», «Flags de operación», «Roles», etc.) son del estudio previo y no se han revalidado.

## Clasificación

| Capacidad | Objetos de lógica | Estado frente al snapshot | Disposición aprobada | Fichas | Hallazgos |
|---|---|---|---|---|---|
| C-HPS-S-1 Arranque e identidad | `Variables Globales.bas:EVE` (`:231`), `constructor.bas:getUsuario` (`:309`), `Form_frmSplash` (encadenado a `Form_Form0BDOpciones`) | sin cambios | Preservar y modernizar; integrar con la identidad única de la web (D-HPS-8) | [Form_frmSplash](formularios/Form_frmSplash.md), [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md) | H-HPS-S-1, 2, 3, 4 |
| C-HPS-S-2 Configuración de backend | `constructor.bas:getConfiguracion` (`:270`), `Configuracion.cls`, `ConfiguracionOperaciones.cls`, `Form_FormConfiguracion` | sin cambios | Preservar `TbConfiguracion` en backend (singleton, decisión del propietario); sustituir las TempVars (`EnDesarrollo`, `ActivadoCorreoAutomatico`, etc.) por `GET /configuracion` | [Form_FormConfiguracion](formularios/Form_FormConfiguracion.md) | H-HPS-S-9, 10, 11, 12, 13 |
| C-HPS-S-3 Flags de operación | `Variables Globales.bas:263-270` (`ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado`) | sin cambios | Mantener como flags de aplicación durante la coexistencia; candidato a tipar como `BOOLEAN` en la web | (cubierto por C-HPS-S-2) | — |
| C-HPS-S-4 Roles | `Variables Globales.bas:317-336` (`EsUsuarioTecnicoCalculado`); 2 niveles (Administrador / Técnico) | sin cambios | Preservar y modernizar; el modelo de 2 niveles se traduce a `role: 'admin' \| 'tech'` y se valida en cada endpoint | [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md) | H-HPS-S-6 |
| C-HPS-S-5 Solicitudes (alta, edición, renovación, cambio de tipo, baja) | `Solicitud.cls`, `SolicitudOperaciones.cls`, `SolicitudFechas.cls`, `SolicitudFechasOperaciones.cls` | sin cambios | Preservar y modernizar; la API web cubrirá las 5 operaciones con `POST /solicitudes`, `PATCH /solicitudes/{id}`, etc. | [Form_FormSolicitudAlta](formularios/Form_FormSolicitudAlta.md), [Form_FormSolicitudRenovacion](formularios/Form_FormSolicitudRenovacion.md), [Form_FormSolicitudRenovacionDatos](formularios/Form_FormSolicitudRenovacionDatos.md), [Form_FormSolicitudRenovacionFechas](formularios/Form_FormSolicitudRenovacionFechas.md), [Form_FormSolicitudesAltaRenovacionFechas](formularios/Form_FormSolicitudesAltaRenovacionFechas.md), [Form_FormCambioTipo](formularios/Form_FormCambioTipo.md) | H-HPS-S-14 a 22, 31 a 42, 50 a 58 |
| C-HPS-S-6 Traspasos a ONS | `Form_FormAdjuntaTraspasoONS`, `Form_FormSolicitudAltaTraspaso`, `Form_FormSolicitudesAltaTraspasoDatos`, `Form_FormSolicitudesTraspasoFechas`, `SolicitudOperaciones.RegistroEnvioTraspasoONS` (sin transacción) | sin cambios | Preservar y modernizar; la web cubre el traspaso como `POST /solicitudes/{id}/traspaso-ons`; transacción atómica obligatoria (D-HPS-11) | [Form_FormSolicitudesAltaTraspasoDatos](formularios/Form_FormSolicitudesAltaTraspasoDatos.md), [Form_FormSolicitudesTraspasoFechas](formularios/Form_FormSolicitudesTraspasoFechas.md), [Form_FormAdjuntaTraspasoONS](formularios/Form_FormAdjuntaTraspasoONS.md) | H-HPS-S-45, 46, 47, 48 |
| C-HPS-S-7 Adjuntos Excel | `Form_FormAdjuntarExcelSolicitante` (sin cablear: D-HPS-10), `SolicitudOperaciones.RegistroEnvioExcel` | sin cambios | Preservar y modernizar; la web cubre el adjunto con `POST /solicitudes/{id}/adjuntos` y validación de tipo MIME/tamaño | [Form_FormAdjuntarExcelSolicitante](formularios/Form_FormAdjuntarExcelSolicitante.md) | H-HPS-S-18, 19, 23 |
| C-HPS-S-8 Plantillas HTML | `Form_FormPlantillasHTML`, `Form_FormWeb`, `Form_FormCorreo.ComandoVerHTML_Click`, `Form_FormEnvioCorreo.EstablecerDatos`, `Form_FormCorreoPruebas.cmdVerHTML_Click` | sin cambios | Preservar y modernizar; las plantillas se sirven firmadas con `ContentVersion`; el render local se sustituye por el navegador del usuario | [Form_FormCorreo](formularios/Form_FormCorreo.md), [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md), [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) | H-HPS-S-33, 35 |
| C-HPS-S-9 Renovación | `Form_FormSolicitudRenovacion`, `Form_FormSolicitudRenovacionDatos`, `Form_FormSolicitudRenovacionFechas`, `Form_FormSolicitudesAltaRenovacionFechas` | sin cambios | Preservar y modernizar; la API web cubre la renovación como `POST /solicitudes/{id}/renovacion` | [Form_FormSolicitudRenovacion](formularios/Form_FormSolicitudRenovacion.md), [Form_FormSolicitudRenovacionDatos](formularios/Form_FormSolicitudRenovacionDatos.md), [Form_FormSolicitudRenovacionFechas](formularios/Form_FormSolicitudRenovacionFechas.md), [Form_FormSolicitudesAltaRenovacionFechas](formularios/Form_FormSolicitudesAltaRenovacionFechas.md) | H-HPS-S-31 a 42 |
| C-HPS-S-10 Cambio de tipo | `Form_FormCambioTipo` (sin cablear: H-HPS-S-29) | sin cambios | Preservar y modernizar; la API web cubre el cambio como `PATCH /solicitudes/{id}/tipo` con validación de rol | [Form_FormCambioTipo](formularios/Form_FormCambioTipo.md) | H-HPS-S-29 |
| C-HPS-S-11 Justificaciones | `Justificacion.cls`, `JustificacionServicio.cls`, `Form_frmJustificacion`, `Form_frmBuscarJustificacion` | sin cambios | Preservar y modernizar; `GET /justificaciones?activa=...&q=...` con índice `GIN` sobre `titulo` | [Form_frmJustificacion](formularios/Form_frmJustificacion.md), [Form_frmBuscarJustificacion](formularios/Form_frmBuscarJustificacion.md) | — |
| C-HPS-S-12 Responsables | `Responsable.cls`, `ResponsableOperaciones.cls`, `Form_FormResponsable`, `Form_FormResponsablesGestion`; relación por `email` con `TbResponsables.Correo` | sin cambios | Preservar y modernizar; FK por `email` se sustituye por FK numérica a `responsables.id` (decisión P-EXP-3) | [Form_FormResponsable](formularios/Form_FormResponsable.md), [Form_FormResponsablesGestion](formularios/Form_FormResponsablesGestion.md) | H-HPS-S-49, 51, 52, 53, 54, 55 |
| C-HPS-S-13 Motivos HPS | `MotivoHPS.cls`, `Form_FormMotivoHPS`, `Form_FormMotivosHPS` (ambos sin cablear: H-HPS-S-23, S-29) | sin cambios | Preservar y modernizar; `GET /motivos-hps?tipo=...` con índice en `(TipoHPS, Grado)` | [Form_FormMotivoHPS](formularios/Form_FormMotivoHPS.md), [Form_FormMotivosHPS](formularios/Form_FormMotivosHPS.md) | H-HPS-S-23, 28, 29 |
| C-HPS-S-14 Entidades / expedientes | `Form_FormEntidades` (sin cablear: ver hallazgos), `Form_FormExpedientesBusqueda`, `Form_FormExpedienteDetalle` | sin cambios | Preservar y modernizar; la búsqueda de expediente se traduce a `GET /expedientes?palabra=...&page=...` con `ILIKE '%…%'` y paginación servidor | [Form_FormEntidades](formularios/Form_FormEntidades.md), [Form_FormExpedientesBusqueda](formularios/Form_FormExpedientesBusqueda.md), [Form_FormExpedienteDetalle](formularios/Form_FormExpedienteDetalle.md) | H-HPS-S-7, 8, 11, 12, 39, 40, 41 |
| C-HPS-S-15 Tareas pendientes (tramitador) | `Form_FormTareasTramitadorPendientes` (11 estados) | sin cambios | Preservar y modernizar; `GET /tareas?estado=...&tramitador=...&page=...`; reemplazar las 11 consultas casi idénticas de `constructor.bas:668-1513` por una vista materializada | [Form_FormTareasTramitadorPendientes](formularios/Form_FormTareasTramitadorPendientes.md) | H-HPS-S-49, 50, 56, 58 |
| C-HPS-S-16 Log general | `LogGeneral.cls`, `LogGeneralOperaciones.cls` → `TbLogsGeneral` (2058 filas) | sin cambios | Preservar y modernizar; partición por `Fecha` (mensual/semestral); seudonimización de `DNI` y `Usuario` (D-HPS-3) | (sin ficha de formulario propia; consumir `Form_FormCorreoPruebas` y `Form_FormCorreo` indirectamente) | H-HPS-S-45 |
| C-HPS-S-17 Correos | `Correo.cls`, `CorreoOperaciones.cls`, `CorreoServicio.cls`, `Form_FormCorreo`, `Form_FormCorreosGestion`, `Form_FormEnvioCorreo`, `Form_FormCorreoPruebas` → `TbCorreosEnviados` (con `IDSolicitud`) | sin cambios | Preservar y modernizar; la web cubre envío de correo como `POST /correos` y `POST /correos/preview`; SMTP autenticado con TLS (D-HPS-12) | [Form_FormCorreo](formularios/Form_FormCorreo.md), [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md), [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md), [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) | H-HPS-S-31 a 49 |
| C-HPS-S-18 Registro en HPS | `TbSolicitudes.IDUsuarioHPS`, flag `RegistroEnHPS = "Sí"`; la integración con la web de HPS es vía API, no vía linked | sin cambios | Preservar y modernizar; la integración se hace vía endpoint `POST /hps/registro` con autenticación federada (decisión abierta P-EXP-3) | (cubierto por las fichas de Solicitud y SolicitudRenovacionDatos) | H-HPS-S-26, 34 |

**Resumen:** 18 de 18 capacidades «sin cambios en el código»; 0 añadidas, 0 modificadas, 0 retiradas del código. La aprobación se mantiene en la cobertura de cada capacidad (preservar y modernizar) con la salvedad de las decisiones D-HPS-8 (identidad), D-HPS-11 (transacciones) y D-HPS-12 (SMTP) que requieren acción del mantenedor.

## Contraste entre el estudio previo y las fichas

Lo que las fichas añaden y el estudio no recogía, en términos de defectos:

- **Defectos de `Err.Numbr`** (H-HPS-S-25, S-37): cinco formas de la cadena (`AltaDatosSolicitante`, `SolicitudRenovacionDatos`, `CambioTipo`, `EnvioCorreo`, `SolicitudesAltaFechas`) referencian un miembro inexistente; el mensaje reventa al construirse. El estudio documentaba la intención pero no este typo.
- **Generador de DNI/NIE sintético** (H-HPS-S-26): `ComandoGenerarDNIValido` asigna un DNI/NIE sin validar letra; combinado con `ComandoRegistrar` puede persistirlo en producción si `EnPruebas = "Sí"`. Hallazgo sensible fuera del lote.
- **`UPDATE` en 4 tablas sin transacción** (H-HPS-S-28, S-44): `MotivoHPS.EditarMotivo` y `SolicitudOperaciones.RegistroEnvioTraspasoONS` no usan `BeginTrans`; caída intermedia deja inconsistencia detectable.
- **Cuatro formularios sin cablear** (H-HPS-S-23, S-29, S-38, S-26): `Form_FormAltaDatosSolicitante`, `Form_FormMotivosHPS`, `Form_FormCambioTipo` y `ComandoGenerarDNIValido` no tienen `DoCmd.OpenForm` ni `SourceObject` que los invoque en `src/`; son residuo de un refactor previo o los lanza una macro del binario.
- **Apertura con `OpenArgs` como `String` libre** (H-HPS-S-36): `Form_FormCorreo` valida `"MODO_GESTION"` como literal; cualquier llamante puede desactivar la edición.
- **SQL por concatenación** (H-HPS-S-18, S-33, S-39): `ComandoAdjuntarExcel_Click` interpola nombre de fichero y fecha en `SQL`; `Constructor.bas:3888-3905` concatena en `InStr` para el filtro de expedientes.

## Capacidades de código sin entrada propia en el estudio

Candidatas a incorporar o a confirmar como subsumidas; son propuestas, no decisiones:

| Candidata | Evidencia | Entrada más cercana |
|---|---|---|
| Generación de PDF de solicitud (no en el `.cls`) | (no en `src/`) | C-HPS-S-5 |
| Importación desde archivos Excel masivos | (no en `src/`) | C-HPS-S-7 |
| Cálculo de plazos de cancelación según `TbConfiguracion.DiasCancelacionPreMARGA` | `SolicitudFechasOperaciones.cls` | C-HPS-S-5 |
| Auditoría de accesos por usuario (vinculado a `TbLogsGeneral`) | `LogGeneral.cls` | C-HPS-S-16 |
| Reenvío de correos sin reescribir plantilla | `Form_FormCorreo.ComandoReenviar_Click` (clona `cadenaRecursos` y `URLAdjunto` antiguos: H-HPS-S-31, S-33) | C-HPS-S-17 |
| Distinción entre `IDSolicitud` real y `CONST_IDSOLICITUD_PRUEBA` en la cola de envío | `Form_FormCorreoPruebas.cmdEnviarPrueba_Click` con `m_ObjEntorno.ModoNuevoEnvioCorreo` (H-HPS-S-42) | C-HPS-S-17 |

## Pendientes

- Confirmar la asociación capacidad-objeto de la tabla con quien conozca HPS Solicitudes.
- Contrastar semánticamente, y no solo por texto, las capacidades de identidad, permisos y registro en HPS.
- Decidir si las linked `TbHPS`/`TbUsuarios` se consolidan con Lanzadera o se federan (D-HPS-8).
- Decidir si `C-HPS-S-15` (tareas pendientes) se queda como `GET /tareas` o se consolida en la cola de Lanzadera.
- Confirmar la equivalencia entre el flag `RegistroEnHPS = "Sí"` y la API real de la web de HPS: a verificar el endpoint `POST /hps/registro` y su contrato.
- Cobertura de pruebas de las 7 ramas de `SolicitudOperaciones.Registro*` y de los 11 estados de `Form_FormTareasTramitadorPendientes`: ¿hay test para cada una?
- Decisión de `ExpedienteUnificado = "No"` (desactivado): si se reactiva en la web, qué endpoint lo cubre; si se descarta, se deja como flag obsoleto.
