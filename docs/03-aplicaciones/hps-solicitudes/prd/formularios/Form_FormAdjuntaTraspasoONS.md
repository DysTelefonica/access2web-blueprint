# Form_FormAdjuntaTraspasoONS

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormAdjuntaTraspasoONS.cls` (234 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo modal para anexar el documento de traspaso a la ONS (sistema externo) sobre la solicitud activa. Es un formulario **no enlazado**: muestra la URL local cacheada y la fecha del envío en `SolicitudFechas.FechaEnvioONS`; al pulsar «Registrar», copia el fichero seleccionado al directorio `URLDirectorioAnexo` y persiste dos `UPDATE` consecutivos (`TbSolicitudes.URLAdjuntoEnvioONS` y `TbSolicitudesFechas.FechaEnvioONS`) sin transacción. Emite el evento público `AdjuntadoEnvio` para que `Form_FormSolicitudesTraspasoFechas` repinte sus fechas. Las citas son `Form_FormAdjuntaTraspasoONS.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesTraspasoFechas` | `DoCmd.OpenForm "FormAdjuntaTraspasoONS"` (si ya estaba abierto, lo cierra antes) | `Form_FormSolicitudesTraspasoFechas.cls:295-302` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesTraspasoFechas` | `RaiseEvent AdjuntadoEnvio` consumido por `m_FormAnexoTraspaso_AdjuntadoEnvio` que llama `EstablecerDatos` y `PintarIndicadores` | `:88`, `Form_FormSolicitudesTraspasoFechas.cls:1131-1166` |

Estado compartido: `m_ObjSolicitudActiva` (la solicitud activa en `VG.bas`), `m_ObjEntorno`, `m_TituloFormulario`, `fso` (todas globales, no declaradas en este `cls`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Formulario` | — | `OnOpen` → `Form_Open` | `Ajustar Me`, `Me.AllowEdits = False`, `EstablecerDatos`, `AvanceCerrar`, `Me.AllowEdits = True`; aborta cerrando el form | `:166-202` |
| `lblTitulo` | etiqueta de cabecera | — | `Caption = "ENVÍO DOC PARA LA ONS"` (literal) | `:117` |
| `URL` | texto deshabilitado | — | Muestra «..\<fichero>» (solo nombre, no ruta completa) | `:121-126, :165-170` |
| `FECHA` | texto fecha corta | — | `FechaEnvioONS` o `Date` por defecto | `:131-141` |
| `ComandoSeleccionar` | botón carpeta | `OnClick` → `ComandoSeleccionar_Click` | `Seleccionar(True, …)` (file dialog), guarda ruta en `m_URLLocal` y refresca `URL`; `AvanceCerrar` | `:152-184` |
| `ComandoAyuda` | botón con icono ayuda | `OnClick` → `ComandoAyuda_Click` | `AbrirAyuda m_Error` | `:13-37` |
| `ImagenURL` | imagen clip (visible si hay adjunto) | `OnClick` → `ImagenURL_Click` | `Ejecutar(Me.hwnd, "open", m_URLLocal, …)` si existe; si no, `AbrirEnLocal m_ObjSolicitudActiva.URLAdjuntoEnvioONSCalculado` | `:217-234` |
| `ComandoRegistrar` | botón «Editar» (`ControlTipText = "Registrar"`) | `OnClick` → `ComandoRegistrar_Click` | valida `m_URLLocal <> ""`; llama `SolicitudOperaciones.RegistroEnvioTraspasoONS`; `RaiseEvent AdjuntadoEnvio`; `PintarIndicadores`; cierra | `:56-104` |
| `cmdSalir` | botón Salir | `OnClick` → `cmdSalir_Click` | `DoCmd.Close acForm, Me.Name, acSaveNo` | `:8-11` |

## Flujos

### Apertura (`Form_Open`, `:166-202`)

1. `DoCmd.Hourglass True`; `Ajustar Me`; `Me.AllowEdits = False` (`:174-182`).
2. `EstablecerDatos m_Error`; aborta con `Err.Raise 1000` si devuelve error (`:184-188`).
3. `AvanceCerrar`; `DoCmd.Hourglass False`; `Me.AllowEdits = True` (`:189-197`).
4. La rama `errores:` avisa por correo, muestra `MsgBox` y cierra el formulario (no deja el diálogo abierto a medio pintar) (`:198-202`).

### Pintado (`EstablecerDatos`, `:108-150`)

1. Si `m_ObjSolicitudActiva` es `Nothing`, devuelve «No hay datos previos de la solicitud» y aborta (`:114-117`).
2. Fija `lblTitulo.Caption = "ENVÍO DOC PARA LA ONS"` literal — la opción recepción no tiene título dedicado (`:117`).
3. Si `URLAdjuntoEnvioONS <> ""`, muestra el icono clip y pinta `URL = "..\" & fso.GetFileName(.URLAdjuntoEnvioONSCalculado)`; si no, oculta el icono y deja `URL = Null` (`:119-127`).
4. Si `SolicitudFechas` no es `Nothing`, usa `FechaEnvioONS` si está definida, o `Date` por defecto (`:129-140`).
5. `Me.Caption = m_TituloFormulario` (`:147`).

### Selección de archivo (`ComandoSeleccionar_Click`, `:152-184`)

`m_URLLocal = Seleccionar(True, "Seleccione el archivo", m_Error)`; si el diálogo devuelve vacío, `Me.URL = Null`; si no, `Me.URL = "..\" & fso.GetFileName(m_URLLocal)`; termina con `AvanceCerrar`.

### Apertura del adjunto desde el icono (`ImagenURL_Click`, `:217-234`)

Si `fso.FileExists(m_URLLocal)` (el usuario ya ha seleccionado en la sesión actual), lanza `Ejecutar(Me.hwnd, "open", m_URLLocal, "", "", 1)` (ShellExecute); si no, recurre a `AbrirEnLocal m_ObjSolicitudActiva.URLAdjuntoEnvioONSCalculado`.

### Registro (`ComandoRegistrar_Click`, `:56-104`)

1. Si `m_URLLocal = ""`, aborta con «Debe seleccionar el archivo a registrar» (`:62-66`).
2. Crea `SolicitudOperaciones`, le inyecta `.solicitud = m_ObjSolicitudActiva` y llama `RegistroEnvioTraspasoONS m_URLLocal, Nz(Me.FECHA, ""), m_Error` (`:68-74`).
3. Si devuelve `m_Error`, aborta con `Err.Raise 1000` (`:76-78`).
4. `RaiseEvent AdjuntadoEnvio` (`:88`); el padre repinta.
5. `PintarIndicadores` (`:90`); `DoCmd.Close acForm, Me.Name, acSaveNo` (`:97`).

`RegistroEnvioTraspasoONS` (`SolicitudOperaciones.cls:1619-1702`): valida que `IDSolicitud <> ""`, que `FechaRecepcionExcel` sea fecha y sea ≤ `FechaEnvioONS`, calcula `m_URLFinal = getURLFinal(p_URLLocal, .IDSolicitud, p_Error)`, copia el fichero a `URLDirectorioAnexo` y ejecuta los dos `UPDATE`.

## Reglas, validaciones y cálculos

- **Concurrencia cero:** `RegistroEnvioTraspasoONS` (`SolicitudOperaciones.cls:1680-1690`) hace dos `UPDATE` por `Execute` directa sin `BeginTrans`/`CommitTrans`/`Rollback`. Un fallo entre ambos deja el fichero copiado en disco y el estado de fecha/cons URL parcialmente persistido.
- **`m_URLLocal` no se inicializa en `Form_Load`:** su valor declarado `Private m_URLLocal As String` (`:9`) queda en la vida del módulo; un error de flujo que relance la apertura puede arrastrar la selección anterior (defecto menor, ver §Defectos).
- **`m_Resultado` declarado y no usado:** la línea `Dim m_Resultado As String` (`:60`) coexiste con un `m_Resultado` global que `Form_FormSolicitudesTraspasoFechas` usa tras `RegistrarUsuarioHPSEnHPS`. En este formulario nunca se le asigna valor: queda `Empty`.
- **`FECHA` validada por `IsDate`:** `RegistroEnvioTraspasoONS` convierte fechas con `CDate(... FechaRecepcionExcel) > CDate(... FechaEnvioONS)` (`SolicitudOperaciones.cls:1651`); un `FECHA` futuro muy lejano rompe al comparar, pero no hay validación previa en este `cls`.
- **Visibilidad del icono clip:** `ImagenURL.Visible = True` solo si `URLAdjuntoEnvioONS <> ""` (`:121-126`); no se desactiva explícitamente cuando el adjunto se borra, pero el botón «Editar» cierra el diálogo, así que la siguiente apertura lo recalcula.
- **`m_URLLocal` pasado a `Ejecutar(Me.hwnd, "open", …, 1)`** sin allowlist de extensiones: admite cualquier `.exe`, `.bat`, `.lnk` presente en el sistema si el usuario lo elige a través del diálogo «Seleccionar»; ver §Aparte seguridad.
- **Caption del botón ComandoAyuda:** en el `.form.txt` el `Caption = "Salir"` y `ControlTipText = "Ayuda"` (`.form.txt:595`); el handler actual invoca `AbrirAyuda`. El rótulo del botón no concuerda con su acción real; defecto cosmetic claro.

## Lecturas y escrituras

| Operación | Tabla o recurso | Tipo | Evidencia |
|---|---|---|---|
| Pintar adjunto | `m_ObjSolicitudActiva.URLAdjuntoEnvioONS`, `URLAdjuntoEnvioONSCalculado` | memoria | `:121-126` |
| Pintar fecha | `m_ObjSolicitudActiva.SolicitudFechas.FechaEnvioONS` | memoria | `:131-140` |
| `RegistroEnvioTraspasoONS` (`SolicitudOperaciones.cls:1619-…`) | `TbSolicitudes.URLAdjuntoEnvioONS` y `TbSolicitudesFechas.FechaEnvioONS`; `TbSolicitudes.URLAdjuntoEnvioONS=Null`/`Null` en borrado posterior | `UPDATE` por `Execute` + `fso.CopyFile`/`fso.DeleteFile` | `SolicitudOperaciones.cls:1679-1690, :1736-1742` |
| `ObtenerCarpetaSiNoExiste` (vía `RegistroEnvioTraspasoONS`) | sistema de archivos: `URLDirectorioAnexo` | `fso.CreateFolder` | `SolicitudOperaciones.cls:1670-1672` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la pareja de `UPDATE` directos (`SolicitudOperaciones.cls:1680-1690`) sin transacción es candidata directa a sustitución por un `PUT /solicitudes/<id>/envio-ons` atómico. La copia por `fso.CopyFile` (`:1679`) pide reemplazo por blob storage + URL prefirmada; el `ShellExecute` «abrir con aplicación asociada» (`:225`) depende de la asociación de extensiones del sistema operativo del usuario, sin equivalente directo en web.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El título `lblTitulo.Caption = "ENVÍO DOC PARA LA ONS"` (`:117`) está hardcodeado, pese a que el comentario de cabecera (`:3-4`) dice «ha de valer para envío y para recepción»: no hay variación para el modo recepción.
- `m_URLLocal` no se borra al abrir el formulario (`:9`): si la cancelación por error cierra y vuelve a abrir, podría arrastrarse la selección anterior hasta que el usuario elija otra o la borre — defecto de ciclo de vida.
- `Dim m_Resultado As String` (`:60`) sobra en este procedimiento: ni se asigna ni se usa; quedó del copy-paste.
- El `Caption` del `ComandoAyuda` no concuerda con su acción: `.form.txt:595`.
- En `ComandoSeleccionar_Click`, `AvanceCerrar` se llama tras construir `m_URLLocal`, pero antes de `Me.AllowEdits = …` — a verificar que el ciclo `AllowEdits` del formulario padre no lo deje en estado inválido.
- `ImagenURL_Click` no avisa si `m_URLLocal` apunta a un fichero recién movido/borrado: cae en la rama `AbrirEnLocal` solo si el `fso.FileExists` falla, pero sin mensaje claro en ese caso (`:224-231`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: subida autenticada con drag-and-drop y MIME allowlist (pdf/doc/docx); fecha siempre editable con validación `max = today`; emisión del evento «AdjuntadoEnvio» como evento del backend (webhook) en lugar de `RaiseEvent` VBA.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el diálogo con la solicitud activa y con `URLAdjuntoEnvioONS = ""` deja `ImagenURL.Visible = False` y `URL = Null`.
2. Abrirlo con adjunto previo pinta `URL` con `..\<fichero>` y muestra el icono clip.
3. Pulsar el icono clip con `m_URLLocal` aún en sesión abre el fichero local; si no, abre `URLAdjuntoEnvioONSCalculado`.
4. Pulsar «Registrar» sin seleccionar fichero aborta con mensaje «Debe seleccionar el archivo a registrar».
5. Tras registrar con éxito: el padre repinta (`AdjuntadoEnvio`) y el diálogo se cierra con `acSaveNo`.
6. La fecha por defecto es `Date` cuando `SolicitudFechas.FechaEnvioONS` está vacío.
7. Un error no controlado en `RegistroEnvioTraspasoONS` se envía al administrador y muestra `MsgBox` antes de cerrar.

## Pendientes

- Definición exacta de `Seleccionar` y `AvanceCerrar` (`FU.bas:41-…`, `FU.bas:555-…`).
- Procedencia real del título `m_TituloFormulario` (versión + entorno de `VG.bas:388-399`).
- Sentido exacto de `m_Resultado` global en `Form_FormSolicitudesTraspasoFechas` y si choca con la local sin asignar de este `cls`.
- Tratamiento real de la fase «DOC DE LA ONS» (recepción): el comentario `ha de valer para envío y para recepción` (`:3-4`) sugiere otro flujo aún no modelado aquí.

### Aparte — defectos de seguridad sensibles (a verificar)

- **`ImagenURL_Click` invoca `Ejecutar(Me.hwnd, "open", m_URLLocal, "", "", 1)`** sobre una ruta local tomada del diálogo del usuario (`:225`) sin allowlist de extensión ni validación de MIME. Un usuario (o un anexo previamente aceptado) puede llevar a `.exe`/`.lnk`/`.bat` y disparar `ShellExecute` con la asociación por defecto del sistema. Reemplazo propuesto: validar `LCase(Right(m_URLLocal, 4))` ∈ `{.pdf, .doc, .docx}` y abrir como blob en navegador.
- **`m_URLLocal` se persiste durante la vida del módulo `cls`** (`:9`): no se borra en `Form_Load`. Si la aplicación reusa la variable tras un ciclo de error, podría persistir una ruta que ya no existe; combinado con `ImagenURL_Click` que usa primero `m_URLLocal`, la rama de «abrir local» puede quedar apuntando a un fichero previamente seleccionado por otro usuario en otra sesión de la misma instancia.
