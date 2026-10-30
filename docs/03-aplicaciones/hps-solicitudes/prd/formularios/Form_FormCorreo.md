# Form_FormCorreo

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormCorreo.cls` (403 líneas) y `.form.txt` (14 controles `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Pantalla de detalle/edición de un correo de la aplicación. Carga la entidad `Correo` por `constructor.getCorreo(p_IDCorreo:=m_ObjCorreoActivo.IDCORREO, p_Error)`, expone los campos editables, valida cambios con `HaHabidoCambios`, persiste con `CorreoOperaciones.Registrar` y emite los eventos públicos `Alta` o `Editado` para refrescar la lista padre. Cuando se abre desde `Form_FormCorreosGestion` con `OpenArgs = "MODO_GESTION"` desactiva el botón de registrar y, si el correo ya fue enviado, bloquea todos los campos y muestra el botón «Reenviar» para clonar el correo. Las citas son `Form_FormCorreo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCorreosGestion` | `DoCmd.OpenForm "FormCorreo", OpenArgs:="MODO_GESTION"` con `m_ObjCorreoActivo` ya fijado desde la selección de la lista | `Form_FormCorreosGestion.cls:285, 305` |
| `WithEvents` | `Private WithEvents m_Form As Form_FormCorreo` en `Form_FormCorreosGestion`; los eventos `Alta` y `Editado` se re-enganchan para refrescar la lista | `Form_FormCorreosGestion.cls:5, 397-403` |
| Otras vías (Solicitud, Tareas, etc.) | `m_ObjCorreoActivo` se rellena antes de abrir el formulario | (no en este `.cls`) |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormCorreosGestion` | `RaiseEvent Alta` o `RaiseEvent Editado`; `DoCmd.Close acForm, Me.Name, acSaveNo` | `:135-138, 207-209` |
| `CorreoOperaciones.Registrar` y `CambiarFechaOrdenEnvio` | Persistencia de la entidad y refresco de la fecha programada | `:120, 122-126` |

No abre otros formularios.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Asunto` | texto bloqueado o habilitado según modo | `ComandoRegistrar_Click` lo vuelca a `m_ObjCorreoActivo.Asunto` | `.form.txt:control Asunto`; `:85` |
| `DESTINATARIOS` | texto | `ComandoRegistrar_Click` lo vuelca a `m_ObjCorreoActivo.DESTINATARIOS` | `:88` |
| `DestinatariosConCopia` | texto | volcado análogo a `m_ObjCorreoActivo.DestinatariosConCopia` | `:90` |
| `Accion` | texto bloqueado | `EstablecerDatos` lo carga; no se persiste por este formulario | `:251` |
| `FechaOrdenEnvio` | fecha bloqueable | `ComandoRegistrar_Click` decide `Programado = "Sí"` si es fecha válida, o `"No"` y vacía `FechaOrdenEnvio` | `:91-99` |
| `FechaEnvio` | fecha bloqueable | volcado a `m_ObjCorreoActivo.FechaEnvio` | `:100` |
| `ComandoRegistrar` | botón | `ComandoRegistrar_Click` | `:51-138` |
| `ComandoReenviar` | botón | `ComandoReenviar_Click` clona el correo y lo da de alta como nuevo | `:142-186` |
| `ComandoVerHTML` | botón | `ComandoVerHTML_Click` resuelve la plantilla con la URL del entorno y la renderiza local | `:266-319` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` | `:18-49` |
| `cmdSalir` | botón | `cmdSalir_Click` cierra sin guardar | `:11-14` |

## Flujos

### Apertura (`Form_Open`, `:325-369`)

1. `Ajustar Me`; `Me.ComandoRegistrar.Enabled = False` por defecto; `Me.AllowEdits = True` para los modos que lo requieran; `Me.Caption = m_TituloFormulario` (`:332-336`).
2. `EsTecnico <> EnumSiNo.Sí` habilita `ComandoRegistrar` (`:339-341`).
3. `EstablecerDatos m_Error` carga la entidad desde `constructor.getCorreo`; si devuelve error, aborta con `Err.Raise 1000` (`:344-347`).
4. `AvanceCerrar` y desbloquea la barra de progreso (`:350-353`).

### Pintado (`EstablecerDatos`, `:238-263`)

1. Recupera `m_CorreoAlInicio = constructor.getCorreo(p_IDCorreo:=m_ObjCorreoActivo.IDCORREO, p_Error)` y aborta si viene `Nothing` (`:241-249`).
2. Copia selectivamente los campos cuyo valor original no está vacío (Asunto, Acción, fechas, destinatarios): un valor vacío no sobrescribe lo que el control ya tuviera (`:250-258`).
3. `AplicarReglasNegocioVisuales` decide qué controles quedan editables (`:261`).

### Reglas visuales (`AplicarReglasNegocioVisuales`, `:276-303`)

- Si `OpenArgs = "MODO_GESTION"`, `ComandoRegistrar.Enabled = False` y `ComandoReenviar.Visible = False` (`:280-282`).
- Si además `IsDate(m_CorreoAlInicio.FechaEnvio)` (ya enviado): `BloquearControles True` y vuelve visible el botón «Reenviar», que sigue permitiendo generar un clon (`:285-290`).
- Si no está enviado: `FechaOrdenEnvio.Enabled = False`, pero el registro inicial sigue siendo posible desde otras vías (`:292-296`).
- En modo normal (sin `MODO_GESTION`): «Reenviar» permanece oculto y el botón de registrar depende de `EsTecnico` (`:298-302`).

### Bloqueo físico (`BloquearControles`, `:305-318`)

Fija `Locked = True` sobre Asunto, DESTINATARIOS, DestinatariosConCopia, Acción, FechaEnvio, FechaOrdenEnvio bajo `On Error Resume Next`. No afecta al botón.

### Guardado (`ComandoRegistrar_Click`, `:51-138`)

1. Si `OpenArgs = "MODO_GESTION"` y el correo ya está enviado, aborta con `MsgBox` informativo (`:60-65`).
2. `HaHabidoCambios` con `m_Error`; si no hay cambios o hay error, `Err.Raise 1000` (`:67-75`).
3. Volcado del formulario a `m_ObjCorreoActivo.Asunto/DESTINATARIOS/.../FechaOrdenEnvio`/`FechaEnvio` con gestión inteligente: si `IsDate(Me.FechaOrdenEnvio)` entonces `.Programado = "Sí"`, en caso contrario `.Programado = "No"` y la fecha vacía (`:79-101`).
4. `Set m_CorreoOp = New CorreoOperaciones` y `Registrar p_CorreoAlInicio:=m_CorreoAlInicio`; si falla, `Err.Raise 1000` (`:103-119`).
5. `CambiarFechaOrdenEnvio` para sincronizar la fecha programada explícitamente (`:120-126`).
6. `RaiseEvent Alta` si era nuevo, `RaiseEvent Editado` si era edición; cierra el formulario (`:130-138`).

### Reenvío (`ComandoReenviar_Click`, `:142-186`)

1. Pide confirmación (`MsgBox vbYesNo`); sale si el usuario cancela (`:146-148`).
2. Crea un `New Correo` clonando los datos relevantes (Asunto, destinatarios, `nombrePlantilla`, `cadenaRecursos`, `URLAdjunto`, `tipoCorreo`, `Accion`, `IDSolicitud`) y resetea fechas (`FechaOrdenEnvio = ""`, `FechaEnvio = ""`, `Programado = "No"`), marcando `DesencadenadoPor` con el usuario de red y el sufijo «(Reenvío)» (`:151-174`).
3. `Registrar` y, si va bien, muestra `MsgBox` informativo y `RaiseEvent Alta`; cierra (`:176-186`).

### Previsualización HTML (`ComandoVerHTML_Click`, `:266-319`)

1. Resuelve `m_NombrePlantilla = m_CorreoAlInicio.nombrePlantilla`; como red de seguridad, si viniera vacío deduce desde `m_CorreoAlInicio.solicitud.NombrePlantillaParaCorreoCalculada` y aborta si no se puede (`:278-294`).
2. Construye `m_URLPlantilla = m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida & m_NombrePlantilla`; aborta si el fichero no existe (`:296-302`).
3. `ObtenerHTMLEnLocal p_URLPlantillaHTML:=m_URLPlantilla, p_IDSolicitud:=m_CorreoAlInicio.IDSolicitud, p_Error:=m_Error` (`:305-309`).

### Errores

Patrón de la casa: `Err.Raise 1000` para errores internos ya redactados; `CorreoAlAdministrador m_Error` solo cuando `Err.Number <> 1000`; `DoCmd.Hourglass False`; `MsgBox m_Error vbCritical`. Aplica a todos los handlers excepto `cmdSalir_Click` (que usa `On Error Resume Next`).

## Reglas, validaciones y cálculos

- **Detección de cambios (`:187-235`):** 5 comparaciones manuales entre `m_CorreoAlInicio` y los controles (`Asunto`, `FechaOrdenEnvio` con `Nz`, `FechaEnvio`, `DESTINATARIOS`, `DestinatariosConCopia`). Vacío inicial → `True`. Sin detección automática de nuevas propiedades: defecto silencioso si se añade un campo.
- **Fechas programadas:** `IsDate` se aplica a `Me.FechaOrdenEnvio` antes de fijar `Programado = "Sí"`. Vacío lo deja en `"No"` sin fecha, comportamiento de «envío manual inmediato o pendiente».
- **Trazabilidad de reenvío:** `DesencadenadoPor` se compone como `m_ObjUsuarioConectado.UsuarioRed & " (Reenvío)"` para que la auditoría distinga el reenvío de la primera generación.
- **Concurrencia:** no hay `BeginTrans` ni `Version`/`LastModified` en la escritura; dos ediciones simultáneas se pisan sin aviso.
- **Plantilla dinámica:** si el correo no tiene `nombrePlantilla` fijado, la previsualización cae en `solicitud.NombrePlantillaParaCorreoCalculada(tipoCorreo)`. A verificar si la deducción cubre todos los tipos de correo.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga de la entidad `Correo` | `TbCorreosEnviados` | DAO (constructor centralizado) | `classes/Correo.cls:1-50`; `:241-249` |
| Persistencia del correo | `TbCorreosEnviados` (insert/update + `CambiarFechaOrdenEnvio`) | `CorreoOperaciones.Registrar` con `p_CorreoAlInicio` | `:103-126` |
| Previsualización HTML | fichero local generado desde `URLDirectorioPlantillasHTMLVersionEstablecida` | `ObtenerHTMLEnLocal` con imágenes copiadas | `:296-309`; `FU.bas:ObtenerHTMLEnLocal` |
| Plantilla ausente | `m_CorreoAlInicio.solicitud.NombrePlantillaParaCorreoCalculada` (memoria) | cálculo de clase de dominio | `:283-292` |

Sin escrituras en otras tablas. No toca `TbSolicitudes` directamente.

## Coexistencia

`TbCorreosEnviados` migra a `lanzadera.hps_solicitudes.correos_enviados` (candidata por nombre; id, asunto, fechas, destinatarios, `IDSolicitud`, `URLAdjunto`, `cadenaRecursos`, `nombrePlantilla`, `VersionPlantilla`, `Estado`, `Accion`, `Aplicacion`, `tipoCorreo`, `Programado`). `m_ObjCorreoActivo` global se sustituye por un endpoint `GET /correos/{id}`. `m_CorreoServicio.ObtenerCorreosGestion` (en `Form_FormCorreosGestion`) se traduce a un `GET /correos?estado=...&solicitud=...` paginado y filtrado en servidor. `ObtenerHTMLEnLocal` se sustituye por un servicio de previsualización con imágenes versionadas por URL firmada. `m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida` deja de ser una ruta de red: el web sirve la plantilla con `ContentVersion` por entorno.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-31.** `ComandoReenviar_Click` solo clona el correo cuando el usuario confirma: el botón «Reenviar» está oculto por defecto y solo aparece si el correo ya está enviado. El clon copia `cadenaRecursos` y `URLAdjunto` antiguos: si entre tanto la plantilla o el adjunto cambiaron, el clon usa los valores obsoletos. `Form_FormCorreo.cls:151-174`
- **H-HPS-S-32.** `ComandoAyuda_Click` y todos los `Comando*_Click` de este formulario llevan el mismo patrón de «literal de error copy-paste» que `Form_FormConfiguracion.cls:42, 84, 218, 280`; p. ej. aquí `:42` y `:132-137` repiten la construcción del mensaje. A verificar el origen y si hay un helper compartido.
- **H-HPS-S-33.** `ComandoReenviar_Click` no recalcula `cadenaRecursos` desde la solicitud origen: la cadena queda con los `||parametros||` del correo original aunque la solicitud haya cambiado. `Form_FormCorreo.cls:165`
- **H-HPS-S-34.** `HaHabidoCambios` solo compara 5 campos y devuelve `True` con la primera diferencia: si la edición solo cambia `Accion` (no persistido por este formulario) o un futuro campo, la primera diferencia cubre la edición accidental y la operación procede.
- **H-HPS-S-35.** `BloquearControles True` se ejecuta con `On Error Resume Next`; controles no presentes en el `.form.txt` actual (p. ej. `DestinatariosConCopiaOculta`, que sí existe en `Reenviar_Click`) no se bloquean por omisión silenciosa.
- **H-HPS-S-36.** `Me.OpenArgs` se valida como `Nz("", "") = "MODO_GESTION"`: el modo se transmite como string literal, sin validación de origen (cualquier parte del código puede abrir con ese `OpenArgs` y desactivar la edición del correo).

**Defectos de interfaz heredados (a verificar):** `Caption` del diseño con valor «Edición de Suministrador» (`.form.txt:39`), `lblTitulo` lleva rótulo genérico, `ComandoAyuda` y `ComandoReenviar` con `Caption` heredado de la plantilla; `ComandoRegistrar` con `Caption` por defecto «Editar» (preserva lo del diseño anterior). `m_TituloFormulario` se compone en `VG.bas:388-401` con el título concreto de la aplicación y la versión: la UI muestra ese caption y, sobreescribiéndolo, el `Me.Caption = m_TituloFormulario` en `Form_Open` queda invisible.

**Defectos de seguridad sensibles (no publicados con detalle, fuera del lote):** el envío se hace a través de `CorreoServicio.Enviar` (clase helper) que persiste en `TbCorreosEnviados` y delega en `CDO.Message`; `m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida` se evalúa sin allowlist y se sirve al `ObtenerHTMLEnLocal` para escribir HTML temporal. La traza de quién envía qué se conserva en `DesencadenadoPor` (`:170`) pero el HTML de la plantilla no se sella ni se firma: una plantilla modificada en disco se enviaría sin huella detectable en la fila de `TbCorreosEnviados`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: distinguir alta de edición en `m_ObjCorreoActivo` con un `EsNuevo` en lugar de `m_CorreoAlInicio Is Nothing`; trazabilidad de versiones de plantilla con hash del fichero renderizado; helper único para mensajes de error de la casa (eliminar el literal copy-paste); `OpenArgs` validado contra tipo enumerado, no contra string libre; `p_CorreoAlInicio` en `Registrar` se sustituye por un `Id` y un `ExpectedVersion` para concurrencia optimista.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el formulario con un `Correo` existente pinta los campos y bloquea los no editables según `MODO_GESTION` y estado de envío.
2. Sin cambios, «Registrar» aborta con «No hay cambios que guardar».
3. Con cambios, «Registrar» persiste, emite `Alta` o `Editado` y cierra.
4. Un correo ya enviado abierto desde gestión solo permite «Reenviar», que clona y registra.
5. La previsualización HTML muestra la plantilla con los datos de la solicitud.

## Pendientes

- Forma exacta de `m_TituloFormulario` cuando se entra desde `MODO_GESTION` (`.lblTitulo` se ignora).
- Si `m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida` resuelve a un UNC con autenticación integrada o por clave.
- Comportamiento de `ObtenerHTMLEnLocal` ante ficheros muy grandes (no se ve límite explícito).
- Estructura de la fila en `TbCorreosEnviados` (no leída) y si el orden de columnas del INSERT en `CorreoOperaciones.Registrar` es estable.
- Reglas de visibilidad de `ComandoReenviar` cuando el correo está «Programado» y aún no enviado (no se contempla explícitamente: cae en el caso B de `AplicarReglasNegocioVisuales` y queda oculto).
