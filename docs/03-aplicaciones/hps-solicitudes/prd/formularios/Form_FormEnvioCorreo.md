# Form_FormEnvioCorreo

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormEnvioCorreo.cls` (427 líneas) y `.form.txt` (14 controles `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Pantalla modal de envío de uno de los siete hitos de correo predefinidos (`EnvioExcel`, `RecordatorioExcel1/2`, `CancelacionPreMarga`, `RecordatorioMarga1/2`, `CancelacionMarga`). Se abre con `OpenArgs = CLng(tipoEnvio)` desde varios formularios de Solicitud y desde el menú; pre-rellena los textos con la misma lógica de negocio que `Form_FormCorreo`, valida la fecha contra la fecha de edición, y dispara la rama correspondiente en `SolicitudOperaciones` para registrar el hito, persistir en `TbSolicitudesFechas` y enviar el correo. Su diseño permite alternar entre «Registrar y enviar» (default) y «Reenviar correo» según el estado histórico de la solicitud. Las citas son `Form_FormEnvioCorreo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudAltaTraspasoDatos` | `DoCmd.OpenForm "Form_FormEnvioCorreo"` con `OpenArgs = CLng(tipoEnvio)` cuando se pulsa un botón de hito | `Form_FormSolicitudesAltaTraspasoDatos.cls:181-186` (en `ComandoEnvio*_Click`) |
| `Form_FormSolicitudRenovacionDatos` | id. | `Form_FormSolicitudRenovacionDatos.cls:174-178` |
| `Form_FormSolicitudesAltaRenovacionFechas` | `Form_FormEnvioCorreo.EstablecerDatos` tras pulsar un botón de hito | ver `ComandoEnvio*_Click` (en subagent output, líneas no leídas) |
| `Form_FormSolicitudesTraspasoFechas` | id. | id. |

Estado compartido: globales `m_ObjSolicitudActiva`, `m_ObjUsuarioConectado`, `m_ObjEntorno`, y `m_TipoEnvioCorreo` (campo privado del formulario, alimentado por `OpenArgs`).

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesAltaFechas` / `Form_FormSolicitudesTraspasoFechas` / `Form_FormSolicitudRenovacionFechas` | `Form_FormSolicitudesAltaFechas.EstablecerDatos` o análogo, llamado en `ComandoEnviar_Click` al terminar | `:113-119` |
| `SolicitudOperaciones.Registro*` | 7 ramas en un `Select Case` sobre `m_TipoEnvioCorreo` | `:240-260` |
| `m_ObjSolicitudActiva.SolicitudFechas` | `Set m_ObjSolicitudActiva = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error)` (refresco) | `:101-104, 162-165` |
| `Form_FormCorreosGestion` (no directo) | Los eventos `Registro*` del formulario no se propagan a `Form_FormCorreosGestion`; el refresco se delega al padre. | (no en este `.cls`) |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblBloqueo` | etiqueta | `ComandoEnviar_Click` la hace visible y cambia su `Caption` a «Procesando envío...» | `:62-64` |
| `lblTitulo` | etiqueta | `EstablecerDatos` pinta el título del hito (`Solicitud.TituloParaFormulario`) | `:99, 174-176` |
| `FECHA` | fecha | `ValidarFechaEdicion` exige fecha válida y no futura; `ComandoEnviar_Click` la pasa a `p_Nz(Me.FECHA, "")` | `:265-280, 350-352` |
| `chkActualizarFecha` | checkbox | `EstablecerDatos` lo marca y deshabilita si el hito no existe, lo desmarca y habilita si ya existe (modo reenvío) | `:209-221` |
| `Asunto` | texto | `EstablecerDatos` lo pinta con `Solicitud.AsuntoCalculado(tipoEnvio)` | `:100, 178` |
| `DESTINATARIOS` | texto | `EstablecerDatos` lo pinta con `Solicitud.DestinatariosCalculados` | `:101, 179` |
| `DestinatariosConCopia` | texto | `EstablecerDatos` lo pinta con `Solicitud.DestinatariosConCopiaCalculados(tipoEnvio)` | `:102, 180` |
| `ComandoEnviar` | botón | `ComandoEnviar_Click` ejecuta la rama correspondiente en `SolicitudOperaciones` | `:36-141` |
| `ComandoGrabarFecha` | botón | `ComandoGrabarFecha_Click` graba solo fecha sin enviar correo | `:310-352` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` abre la ayuda | `:11-33` |
| `cmdSalir` | botón | `cmdSalir_Click` cierra sin guardar | `:6-8` |
| `Programado` | texto bloqueado | `EstablecerDatos` lo pone a «No» (envío inmediato) | `:182` |

## Flujos

### Apertura (`Form_Open`, `:359-393`)

1. `Ajustar Me`; `If Not IsNumeric(Nz(Me.OpenArgs(), ""))` → `m_Error = "El formulario se ha abierto con parámetros insuficientes"` y `Err.Raise 1000` (`:363-369`).
2. `m_TipoEnvioCorreo = CLng(Me.OpenArgs())` (`:371`).
3. `Me.AllowEdits = False`; `Me.Caption = m_TituloFormulario` (`:373-375`).
4. `EstablecerDatos m_Error`; si falla, `Err.Raise 1000` (`:377-381`).
5. `AvanceCerrar`; `Me.AllowEdits = True`; `HabilitarEdicionFecha` (`:383-389`).

### Pintado (`EstablecerDatos`, `:168-225`)

1. `If m_ObjSolicitudActiva Is Nothing` → `Err.Raise 1000` (`:171-174`).
2. Carga `m_titulo`, `m_Asunto`, `m_Destinatarios`, `m_DestinatariosEnCopia` desde `m_ObjSolicitudActiva.TituloParaFormulario` / `AsuntoCalculado` / `DestinatariosCalculados` / `DestinatariosConCopiaCalculados` (`:175-180`).
3. Pinta los controles `lblTitulo`, `Asunto`, `DESTINATARIOS`, `DestinatariosConCopia`, `Programado = "No"` (`:181-183`).
4. Calcula `fechaYaRegistrada` con un `Select Case` sobre `m_TipoEnvioCorreo` que consulta la columna correspondiente de `m_ObjSolicitudActiva.SolicitudFechas` (`:186-200`).
5. Si `IsDate(fechaYaRegistrada)`: `chkActualizarFecha.value = False`, `.Enabled = True`, `ComandoEnviar.Caption = "Reenviar Correo"`, `Me.FECHA = fechaYaRegistrada` (modo reenvío, permite corregir fecha) (`:209-216`).
6. Si no: `chkActualizarFecha.value = True`, `.Enabled = False`, `ComandoEnviar.Caption = "Registrar y Enviar"`, `Me.FECHA = Date` (modo registro obligatorio) (`:218-222`).
7. `HabilitarEdicionFecha` y `ComandoEnviar.Enabled = True` (`:223-224`).

### Validación de fecha (`ValidarFechaEdicion`, `:407-422`)

1. `IsDate(Nz(Me.FECHA, ""))` → si falso, `p_Error = "Debe indicar una fecha válida."` y devuelve `False` (`:411-414`).
2. `d = CDate(Me.FECHA)`; `If d > Date` → `p_Error = "No se permite una fecha posterior a la fecha de edición (hoy)."` y devuelve `False` (`:416-421`).

### Desbloqueo de fecha (`HabilitarEdicionFecha`, `:424-441`)

1. `Me.AllowEdits = True`; `With Me.FECHA: .Enabled = True, .Locked = False, .TabStop = True` (`:426-432`).
2. Si `Me.FECHA.ControlSource` empieza por `=`, contiene `(`, `+` o `&` (expresión calculada), vacía `ControlSource` para desenlazar y permitir edición (`:434-440`).

### Envío (`ComandoEnviar_Click`, `:36-141`)

1. Compone `msg` según `Me.chkActualizarFecha.value`: «REGISTRAR y ENVIAR» o «REENVÍO sin modificar fecha» (`:38-44`).
2. `pregunta = MsgBox(msg, vbYesNo + vbDefaultButton2)`. Sale si no es `vbYes` (`:46-50`).
3. Si `chkActualizarFecha.value = True` y `Not ValidarFechaEdicion`, `Err.Raise 1000` (`:52-55`).
4. `Me.lblBloqueo.Visible = True; .Caption = "Procesando envío..."` (`:62-64`).
5. `Registrar EnumSiNo.Sí, m_Error` (`:66-69`).
6. `Set m_ObjSolicitudActiva = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error)` (`:101-104`).
7. Refresca el padre según el `TIPO` de la solicitud: `Alta` → `Form_FormSolicitudesAltaFechas.EstablecerDatos`; `Traspaso` → `Form_FormSolicitudesTraspasoFechas.EstablecerDatos`; resto → `Form_FormSolicitudesAltaRenovacionFechas.EstablecerDatos` (`:113-119`).
8. Mensaje final de éxito y `DoCmd.Close` (`:124-141`).

### Persistencia (`Registrar`, `:227-292`)

1. Decide `blnSoloEnvio = IIf(Me.chkActualizarFecha.value = True, False, True)`. Si `p_EnviandoCorreo = EnumSiNo.no`, fuerza `blnSoloEnvio = False` (botón «Grabar fecha» que solo graba) (`:232-241`).
2. `Set m_SolicitudOp = New SolicitudOperaciones`; `Set .solicitud = m_ObjSolicitudActiva` (`:244-246`).
3. `Select Case m_TipoEnvioCorreo`: 7 ramas a `m_SolicitudOp.Registro*Excel/Recordatorio*/Autocancelacion*` con `p_EnviandoCorreo:=p_EnviandoCorreo, p_Fecha:=Nz(Me.FECHA, ""), p_Destinatarios, p_DestinatariosCopia, p_Asunto, p_blnSoloEnvio, p_Error` (`:247-274`).
4. Si `m_TipoEnvioCorreo` no encaja en ninguna rama, `p_Error = "Acción no codificada"` y `Err.Raise 1000` (`:275-277`).
5. `PintarIndicadores` y refresca el padre según `m_ObjSolicitudActiva.TIPO` (`:283-290`).

### Grabación de fecha sin envío (`ComandoGrabarFecha_Click`, `:310-352`)

1. Confirma con `MsgBox("Va a grabar sólo la fecha, NO SE VA A MANDAR CORREO ELECTRÓNICO", vbYesNo)` (`:312-314`).
2. `ValidarFechaEdicion` (`:317-319`).
3. `Registrar EnumSiNo.no, m_Error` (`:325-327`).
4. Refresca `m_ObjSolicitudActiva` y cierra (`:330-340`).

### Errores

`Form_Open`, `ComandoEnviar_Click`, `ComandoGrabarFecha_Click` y `ComandoAyuda_Click` comparten el patrón de la casa. `Form_Open` aborta con `DoCmd.Close`; `ComandoEnviar_Click` hace `Me.lblBloqueo.Visible = False` antes de mostrar el `MsgBox` para que la UI no quede bloqueada. `EstablecerDatos` y `ValidarFechaEdicion` no muestran `MsgBox` ellos mismos: su `p_Error` lo muestra el llamador. `cmdSalir_Click` usa `On Error Resume Next`.

## Reglas, validaciones y cálculos

- **Siete tipos de correo predefinidos:** `EnvioExcel`, `RecordatorioExcel1`, `RecordatorioExcel2`, `CancelacionPreMarga`, `RecordatorioMarga1`, `RecordatorioMarga2`, `CancelacionMarga`. Las 7 ramas en `Registrar` conmutan a las 7 firmas de `SolicitudOperaciones`. A verificar que el `EnumTipoEnvioCorreo` del `VG.bas` no haya añadido tipos nuevos.
- **Modo «Registrar y enviar» vs «Reenviar»:** determinado por `EstablecerDatos` consultando `m_ObjSolicitudActiva.SolicitudFechas.{Columna}` para el tipo. Si existe fecha, modo reenvío con `chkActualizarFecha` opcional (el usuario puede corregir la fecha y reenviar con la nueva). Si no existe, registro obligatorio y checkbox bloqueado.
- **Fecha no futura:** la regla es post-edición, no histórica. Una fecha anterior a hoy se acepta como reenvío (corrección).
- **Conmutación entre padre:** tras enviar, refresca el formulario padre según `m_ObjSolicitudActiva.TIPO` (`Alta`, `Traspaso`, otros). El refresco no cubre la pantalla de `Form_FormSolicitudesAltaRenovacionFechas` cuando `TIPO = "Renovación"` (cae en el «resto»). A verificar.
- **Concurrencia:** no hay `BeginTrans` ni versión de fila; dos pantallas abiertas sobre la misma solicitud podrían pisar fechas. `m_ObjSolicitudActiva` se refresca tras `Registrar`, pero la lectura posterior no compara contra una versión leída originalmente.
- **Validación de fecha opcional:** `ValidarFechaEdicion` se llama solo cuando `chkActualizarFecha = True` en `ComandoEnviar_Click` (`:52-55`); `ComandoGrabarFecha_Click` la llama siempre (`:317-319`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga inicial del formulario | `m_ObjSolicitudActiva` (en memoria) | `Solicitud.TituloParaFormulario` / `AsuntoCalculado` / `DestinatariosCalculados` / `DestinatariosConCopiaCalculados` | `:175-180` |
| Consulta del estado histórico | `m_ObjSolicitudActiva.SolicitudFechas` (en memoria) | `FechaEnvioExcel` / `FechaCorreoRecordatorioExcel1/2` / `FechaCancelado` / `FechaCorreoRecordatorioRellenoMarga1/2` | `:186-200` |
| Persistencia del hito | `TbSolicitudesFechas` (vía `SolicitudOperaciones.Registro*`) | 7 métodos `Registro*Excel/Recordatorio*/Autocancelacion*` con `p_EnviandoCorreo`, `p_Fecha`, `p_Destinatarios`, `p_DestinatariosCopia`, `p_Asunto`, `p_blnSoloEnvio` | `:247-274` |
| Envío del correo | `TbCorreosEnviados` (vía `EnviarCorreoGenerico` en `Automiatizacion.EnviarCorreoGenerico` llamado desde `SolicitudOperaciones.Registro*`) | DAO centralizado | (no en este `.cls`) |
| Refresco del padre | `Form_FormSolicitudesAltaFechas.EstablecerDatos` / análogo | `DoCmd.FormRefresh` indirecto | `:113-119, 287-290` |
| Log de auditoría (vía padre) | `TbLogsGeneral` (vía `PintarIndicadores`) | DAO centralizado | (no en este `.cls`) |

Sin escrituras en catálogos ni en `TbUsuarios`. No lee directamente tablas; toda la información llega por la `Solicitud` en memoria y sus agregados.

## Coexistencia

`TbSolicitudesFechas` migra a `lanzadera.hps_solicitudes.solicitudes_fechas` con una columna `tipo_hito` `ENUM` con los 7 valores actuales. `m_TipoEnvioCorreo` global se sustituye por un discriminador en la URL (`/solicitudes/{id}/hitos/{tipo}`). `Set m_ObjSolicitudActiva = constructor.getSolicitud(...)` se traduce a un endpoint `GET /solicitudes/{id}` que devuelve un grafo con `solicitud_fechas` anidadas. El refresco del padre pasa a una `Suscripción` por websocket o una invalidación de caché local. El `Select Case` de las 7 ramas se condensa en una única función `SolicitudFechas.registrar(tipo, fecha, ...)` que delega a una política configurable. `chkActualizarFecha` se vuelve un toggle del lado del cliente que el backend respeta como `if-match` opcional. `ValidarFechaEdicion` se implementa como un guard de FastAPI con el reloj del servidor como referencia.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-50.** `ComandoEnviar_Click` refresca el padre según el `TIPO` de la solicitud (`:113-119`): cuando `TIPO = "Renovación"`, el bloque «resto» ejecuta `Form_FormSolicitudesAltaRenovacionFechas.EstablecerDatos`, pero `Form_FormSolicitudesAltaFechas.EstablecerDatos` para `Alta` no cubre la pantalla específica de renovación. Si en la misma sesión hay una pantalla `Form_FormSolicitudesAltaRenovacionFechas` y otra `Form_FormSolicitudesAltaFechas` abiertas, solo se refresca una. `Form_FormEnvioCorreo.cls:113-119`
- **H-HPS-S-51.** `Registrar` aborta con «Acción no codificada» ante un `m_TipoEnvioCorreo` desconocido: el usuario pierde el trabajo; debería mantener el formulario abierto con foco en el campo que corresponda. `Form_FormEnvioCorreo.cls:275-277`
- **H-HPS-S-52.** `HabilitarEdicionFecha` desenlaza el control cuando el `ControlSource` parece una expresión calculada (`:434-440`). Si el `ControlSource` cambia entre llamadas, el desenlace ocurre en cada apertura; con `Form_Current` (`:395-399`) re-ejecutándose en cada movimiento de registro, el comportamiento es difícil de seguir. A verificar si `Form_Current` debe seguir llamándolo.
- **H-HPS-S-53.** `ComandoEnviar_Click` usa `On Error Resume Next` solo en el bloque de refresco del padre (`:108-119`); los errores al refrescar quedan silenciados. Si la pantalla padre está cerrada o no responde, el `Form_FormEnvioCorreo` cierra y el usuario asume éxito sin saber que el padre no se refrescó.
- **H-HPS-S-54.** `Registrar` no actualiza `m_ObjSolicitudActiva.SolicitudFechas.{columna}` antes de volver: la versión en memoria queda obsoleta hasta que `ComandoEnviar_Click` la refresca al final. Si entre la escritura y el refresco una lectura muestra datos obsoletos, el diagnóstico se confunde.
- **H-HPS-S-55.** `Me.Caption = m_TituloFormulario` en `Form_Open` (`:375`) y luego `ComandoEnviar_Click` lo sobreescribe con un literal en el `MsgBox vbInformation` final (`:124-130`): el caption se pinta con el título calculado en `EstablecerDatos` (`:174-176`), pero `m_TituloFormulario` lo sobreescribe después. A verificar el orden de pintado en `Form_Open`.
- **H-HPS-S-56.** `chkActualizarFecha.Enabled = False` cuando el hito no existe (`:220`): el usuario no puede desactivar la actualización, pero el caption del botón dice «Registrar y Enviar» (`:221`) y `ComandoGrabarFecha_Click` tiene su propio caption sin restricción. La documentación sugiere que el botón «Grabar fecha» existe para grabar sin correo, pero su caption no es explícito.
- **H-HPS-S-57.** `ComandoGrabarFecha_Click` no refresca el padre tras grabar (sólo `Set m_ObjSolicitudActiva` y `DoCmd.Close`): la pantalla de `Form_FormSolicitudesAltaFechas` abierta detrás no se actualiza con la nueva fecha hasta que el usuario la refresca manualmente.
- **H-HPS-S-58.** `ValidarFechaEdicion` rechaza fechas futuras (`:420`), pero `EstablecerDatos` inicializa con `Me.FECHA = Date` (`:222`): la fecha de hoy siempre es válida. Una fecha anterior a hoy también es válida: el rechazo es solo de futuras.

**Defectos de seguridad sensibles (no publicados con detalle, fuera del lote):** `Me.FECHA` se desenlaza en `HabilitarEdicionFecha` y se acepta cualquier valor que el usuario teclee; el backend no recibe un timestamp firmado ni una prueba de que la fecha es legítima. Los `p_Destinatarios` y `p_DestinatariosCopia` se concatenan tal cual a `SolicitudOperaciones.Registro*` que persiste en `TbSolicitudesFechas` y luego encola el correo. La validación de formato de correo electrónico no se aplica en el formulario: cualquier string se acepta como destinatario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pasar a un endpoint `POST /solicitudes/{id}/hitos/{tipo}` con payload JSON; el backend valida la fecha contra el reloj del servidor y rechaza futuras o muy pasadas; la actualización de pantallas padres se hace con un store reactivo o notificaciones por websocket; el `Select Case` se sustituye por una política de `SOLICITUD_HITOS` configurable; el `MsgBox` de error se sustituye por un toast con detalle accesible desde la consola; el botón «Grabar fecha» se elimina y se unifica con el botón principal mediante un toggle explícito de "actualizar fecha sí/no".

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir el formulario con un `tipoEnvio` válido y una solicitud activa, se pintan título, asunto y destinatarios calculados.
2. Si el hito existe, el caption dice «Reenviar Correo» y el checkbox está habilitado; si no, dice «Registrar y Enviar» y el checkbox está bloqueado en true.
3. «Enviar» valida la fecha (no futura) y dispara la rama correspondiente en `SolicitudOperaciones`.
4. «Grabar fecha» persiste la fecha sin enviar correo.
5. Tras una operación, el padre (`AltaFechas` / `TraspasoFechas` / `AltaRenovacionFechas`) se refresca.
6. Si el `tipoEnvio` no es válido, el formulario aborta con «El formulario se ha abierto con parámetros insuficientes».

## Pendientes

- `EnumTipoEnvioCorreo`: valor numérico de cada miembro, posibles valores no contemplados.
- Forma de `SolicitudOperaciones.Registro*`: cómo se construye la `INSERT/UPDATE` en `TbSolicitudesFechas`, y si se concatena el `IDSolicitud` (riesgo de inyección si no se parametriza).
- Comportamiento de `PintarIndicadores` y `Form_FormSolicitudesAltaFechas.EstablecerDatos`: idempotencia, manejo de errores.
- Si `ComandoGrabarFecha` debe seguir existiendo una vez que se separa el endpoint (la separación cliente-servidor lo hace redundante).
- Si el `chkActualizarFecha` es accesible por teclado: el evento `Enter` en un checkbox no es estándar; a verificar la navegación por teclado.
- Política de refresco del padre: ¿debería refrescarse también `Form_FormCorreosGestion` si está abierto desde un proceso de envío?
- Cobertura de pruebas de las 7 ramas de `Registrar`: ¿hay test para cada una?
