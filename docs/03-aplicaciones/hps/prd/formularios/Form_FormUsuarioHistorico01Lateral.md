# Form_FormUsuarioHistorico01Lateral

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuarioHistorico01Lateral.cls` (130 líneas) y `.form.txt` (934 líneas) · **PRD:** [README](../README.md)

Subformulario de navegación del histórico: aloja el icono de usuario, tres etiquetas informativas de entorno (Versión, Datos en, Modo Escritura) y dos etiquetas-botón —«Gral.» y «Anexos»— que cambian el `SourceObject` del `SubFormCentral` del formulario padre para mostrar `FormUsuarioHistorico02Edicion` o `FormUsuarioHistorico02Anexos`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormUsuarioHistorico01Lateral.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuarioHistorico00Principal` | Se carga como `SubFormLateral.SourceObject = "FormUsuarioHistorico01Lateral"` (inicial y defensivo) | `Form_FormUsuarioHistorico00Principal.cls:78-80`, `Form_FormUsuarioHistorico00Principal.form.txt:851` |
| `Form_FormUsuarioHistorico00Principal` → `SubFormCentral` | El lateral cambia el `SourceObject` del subformulario central según la pestaña elegida | `:47, 63`, `Form_FormUsuarioHistorico00Principal.form.txt:890` |
| `m_ObjUsuarioHistoricoActivo` (global) | Origen del estado compartido; si es `Nothing`, `SeleccionarEnlace ""` sale en silencio | `VariablesEntorno.bas:227`, `:95-97` |

No abre formularios por sí solo: las dos pestañas se limitan a conmutar el `SourceObject` del padre (`:47, 63`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Pinta las etiquetas de entorno, vacía los tags, marca la pestaña General por defecto | `:7-29` |
| (formulario) | — | `OnUnload` implícito | Sin `Form_Unload` propio; el estado global sobrevive | — |
| `Imagen7` | imagen | — | Icono de usuario (`UsuarioIcono.png`) | `.form.txt:244-245` |
| `lblDatosGenerales` («Gral.») | etiqueta-botón, `Tag = "Elegibles;Seleccionado"` | `OnClick` → `lblDatosGenerales_Click` | Si el central no es ya edición, fuerza `FormUsuarioHistorico02Edicion`, reordena tags y aplica estilo | `:57-84`, `.form.txt:677-693` |
| `lblAnexos` («Anexos») | etiqueta-botón, `Tag = "Elegibles;NoSeleccionado"` | `OnClick` → `lblAnexos_Click` | Si ya es Anexos, sale sin tocar nada; si no, fija `FormUsuarioHistorico02Anexos`, reordena tags y aplica estilo | `:32-55`, `.form.txt:705-721` |
| `lblVersion` | etiqueta informativa | — | Recibe `"Versión: " & VersionAplicacion` | `:13`, `.form.txt:733-746` |
| `lblDatosEn` | etiqueta informativa | — | Recibe `"Datos en: " & DatosEn` | `:14`, `.form.txt:758-771` |
| `lblModoEscritura` | etiqueta informativa | — | Recibe `"Modo Escritura: " & ModoEscritura` | `:15`, `.form.txt:783-796` |
| `Detalle` | sección | — | Fondo `BackColor = 8210719`; sin lógica | `.form.txt:232, 230` |

## Flujos

### Apertura (`Form_Load`, `:7-29`)

1. Entra con `On Error GoTo errores` y vacía `m_Error` (`:9-10`).
2. Pinta `lblVersion.Caption` con `m_ObjEntorno.VersionAplicacion` (`:13`; `Entorno.cls:1526-1561`).
3. Pinta `lblDatosEn.Caption` con `m_ObjEntorno.DatosEn` («Oficina», «Local» o «#ERR»; `:14`; `Entorno.cls:2219-2241`).
4. Pinta `lblModoEscritura.Caption` con `m_ObjEntorno.ModoEscritura` («Sí», «No» o «#ERR»; `:15`; `Entorno.cls:2242-2253`).
5. `SeleccionarEnlace , m_Error` se llama con el argumento omitido, lo que entra en la rama `p_TipoHPS = ""` y, si `m_ObjUsuarioHistoricoActivo` está asignado, marca como `Elegibles;NoSeleccionado` todas las etiquetas que tengan «;» en el `Tag` (`:16`; `:94-107`).
6. Si `m_Error` volvió lleno, eleva error interno 1000 (`:17-19`).
7. Llama directamente a `lblDatosGenerales_Click` para fijar la pestaña General de inicio (`:20`).
8. En error: compone mensaje, envía correo al administrador, `MsgBox` crítico y **cierra el formulario sin guardar** (`:22-28`).

### Cambio a Datos Generales (`lblDatosGenerales_Click`, `:57-84`)

1. Entra con `On Error GoTo errores` y vacía `m_Error` (`:58-60`).
2. Si el `SourceObject` del `SubFormCentral` del padre **ya es** `FormUsuarioHistorico02Edicion`, omite la asignación (`:62-64`).
3. `SeleccionarEnlace "", m_Error` aplica la rama «limpiar» del helper (`:66`; `:94-107`). Si devuelve error, eleva 1000 (`:67-69`).
4. Marca `lblDatosGenerales.Tag = "Elegibles;Seleccionado"` (`:71`).
5. `EstablecerControlesformlateral Me, m_Error` repinta los colores del lateral con la paleta del `Entorno` (`:72`; `FG.bas:963-1004`). Si falla, eleva 1000 (`:73-75`).
6. En error: mensaje, correo y `MsgBox` sin cerrar el formulario (`:78-83`).

### Cambio a Anexos (`lblAnexos_Click`, `:32-55`)

1. `On Error GoTo errores` y `m_Error = ""` (`:33-34`).
2. Si el central **ya es** `FormUsuarioHistorico02Anexos`, sale sin tocar nada: es un guard contra re-entrada (`:35-37`).
3. `SeleccionarEnlace "ANEXO", m_Error` aplica la rama «Anexos» que fija el tag de `lblAnexos` a `Elegibles;Seleccionado` y borra el «;» de las demás etiquetas con «;» en el tag (`:38`; `:108-122`). Si falla, eleva 1000 (`:40-42`).
4. Duplica la asignación del tag de `lblAnexos` a `Elegibles;Seleccionado` (`:39`) — está en el helper, pero también se hace aquí.
5. `EstablecerControlesformlateral Me, m_Error` repinta los colores (`:43`; `FG.bas:963-1004`). Si falla, eleva 1000 (`:44-46`).
6. Fija `Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuarioHistorico02Anexos"` (`:47`).
7. En error: mensaje, correo y `MsgBox` sin cerrar el formulario (`:49-54`).

### Helper privado de selección (`SeleccionarEnlace`, `:88-130`)

No tiene manejador público; lo llaman los tres procedimientos de arriba. Dos ramas operativas y un handler de error sin cuerpo de notificación:

- **`p_TipoHPS = ""`** (`:94-107`): si `m_ObjUsuarioHistoricoActivo Is Nothing`, sale en silencio sin tocar nada (`:95-97`). Si está asignado, recorre `Me.Controls`, filtra los `Label` con `InStr(1, ctl.Tag, ";") <> 0` y les pone `Tag = "Elegibles;NoSeleccionado"` (`:98-105`).
- **`p_TipoHPS = "ANEXO"`** (`:108-122`): guarda el nombre de `lblAnexos`, fuerza `lblAnexos.Tag = "Elegibles;Seleccionado"` y, para cada otro `Label` con «;» en `Tag` (normalizado con `Nz(ctl.Tag, ";")`), lo pone a `Elegibles;NoSeleccionado` (`:109-121`).
- **Manejador de error** (`:126-129`): asigna `p_Error` con descripción si el error no es interno 1000; **no** envía correo ni muestra `MsgBox`. La notificación corresponde a los llamadores.

### Errores

Los tres procedimientos públicos (`Form_Load`, `lblAnexos_Click`, `lblDatosGenerales_Click`) siguen el patrón del repositorio: para errores no internos 1000, componen el mensaje, llaman a `CorreoAlAdministrador` y muestran un `MsgBox` crítico guardando el resultado en `pregunta` (`:22-28, :49-54, :78-83`; `FG.bas:67`). Solo `Form_Load` cierra el formulario tras el error; los `Click` no lo cierran para no perder la pestaña abierta (`:28, :54, :83`).

## Reglas, validaciones y cálculos

- **El lateral es un conmutador, no un formulario de negocio.** Su única acción útil es asignar `Me.Parent.SubFormCentral.SourceObject` a uno de los dos formularios central (`:47, 63`); la regla de qué contiene cada central vive en su formulario hijo.
- **Guard contra re-entrada en Anexos.** El primer paso de `lblAnexos_Click` sale sin tocar nada si la pestaña ya está activa (`:35-37`): no se reasigna `SourceObject` ni se repinta al pulsar dos veces.
- **Carga de entorno como side-effect.** `Form_Load` reordena los tags solo si `m_ObjUsuarioHistoricoActivo` está asignado (`:95-97`); si es `Nothing`, el lateral queda como esté y la siguiente línea —`lblDatosGenerales_Click`— igual monta la pestaña General porque la rama «Anexos» no se ejecuta. Sin usuario activo, el lateral cargará pero el subformulario central no abrirá datos (a verificar contra el binario, donde `m_ObjUsuarioHistoricoActivo` se rellena desde `FormUsuarioHistorico08Historicos` antes de abrir el `00`).
- **Estilo leído del entorno.** Los colores del fondo, bordes de etiqueta y selección se leen de `m_ObjEntorno.{CabeceraBackColor,EtiquetaBorderColor,EtiquetaLinkForeColor,EtiquetaLinkBackColorSeleccionado,EtiquetaForeColorSeleccionado}` en `EstablecerControlesformlateral` (`FG.bas:967, 981, 989-997`); vienen de `colColoresMaster` o `colColoresSlave` (`Entorno.cls:825-886`). Cambiar la paleta implica tocar `Entorno`, no el `.form.txt`.
- **Identidad del control por nombre.** `SeleccionarEnlace "ANEXO"` compara contra `Me.lblAnexos.Name` (`:109, 114`); cualquier futuro control que no sea `Label`, o cuyo `Tag` no contenga «;», queda fuera de la limpieza. Las cinco etiquetas declaradas en el `.form.txt` lo cumplen (`.form.txt:677-796`).
- **`Tag` con dos tramos separados por «;».** El primer tramo es la categoría (`Elegibles`) y el segundo el estado (`Seleccionado`/`NoSeleccionado`) (`:71, 102, 116`; `.form.txt:680, 708`). Cambiar el separador rompería `EstablecerControlesformlateral`, que parte por `;` (`FG.bas:982`).

## Lecturas y escrituras

Ninguna contra base de datos.

| Operación | Origen | Tipo | Evidencia |
|---|---|---|---|
| Estado del usuario histórico | `m_ObjUsuarioHistoricoActivo` | global público `As UsuarioHistorico` | `VariablesEntorno.bas:227` |
| Versión, dónde están los datos y modo escritura | `m_ObjEntorno.{VersionAplicacion,DatosEn,ModoEscritura}` | propiedades del `Entorno` | `Entorno.cls:1526-1561, 2219-2241, 2242-2253` |
| Cambio de pestaña central | `Me.Parent.SubFormCentral.SourceObject` | propiedad del control `SubForm` del padre | `:47, 63` |
| Estilo del lateral | `EstablecerControlesformlateral` | helper de `FG.bas` que pinta la paleta | `:43, 72`; `FG.bas:963-1004` |

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). La regla de visibilidad de la pestaña es lineal: primero `lblDatosGenerales_Click` fija `FormUsuarioHistorico02Edicion` (`:63`), y luego `EstablecerControlesformlateral` la pinta de seleccionado (`FG.bas:990-996`); si la fuente se cae entre los dos pasos (el `Entorno` lanza al pintar, p. ej.), la asignación queda hecha pero el estilo no. El `For Each` sobre `Me.Controls` (`:98-121`) es `TypeOf … Is Label`: si en algún momento se añade un control que no sea `Label` con `Tag` legible (un rectángulo-botón, p. ej.), no se actualizará al cambiar de pestaña, sin error. La ausencia de `OnUnload` no es un vacío: el estado global del usuario se mantiene deliberadamente entre pestañas (a verificar — a qué formularios les conviene que `m_ObjUsuarioHistoricoActivo` siga vivo al cerrar este lateral).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **`lblAnexos.Tag` se asigna dos veces** en `lblAnexos_Click`: primero dentro de `SeleccionarEnlace "ANEXO"` (`:109-110`) y luego otra vez en el cuerpo del procedimiento (`:39`). Es redundante; el segundo es el que aplica el estilo final. Si un día se quita uno, el visible es el del helper (a verificar contra el binario).
- **`Form_Load` llama a `lblDatosGenerales_Click`** (`:20`) sin restaurar antes `m_Error` ni `On Error`. Si el helper eleva 1000, el `On Error GoTo errores` del `Form_Load` (`:9`) lo atrapa y muestra el `MsgBox` crítico (`:27`), pero `Me.AllowEdits` (no usado aquí) ni el estilo del lateral no llegan a aplicarse. La consecuencia observable es que el central queda como heredado del padre y el lateral queda con los tags anteriores al helper.
- **`SeleccionarEnlace ""` aborta en silencio si no hay usuario activo** (`:95-97`); `Form_Load` lo llama antes de la llamada a `lblDatosGenerales_Click` (`:16, 20`) y entonces monta la pestaña General sobre un lateral «sin inicializar». Visualmente las etiquetas quedan con los tags del `.form.txt` (`Elegibles;Seleccionado` para General, `Elegibles;NoSeleccionado` para Anexos) hasta que `EstablecerControlesformlateral` las pinta (`:72`; `FG.bas:989-997`).
- **`SeleccionarEnlace` no notifica al administrador** (`:126-129`): se limita a poblar `p_Error`. Los llamadores sí lo hacen, así que cualquier excepción del propio helper llega al correo por su cuenta; pero un error interno 1000 elevado aquí no se propaga porque el `If Err.Number <> 1000` lo filtra (`FG.bas:1001, 1058`).
- **Sin `Form_Unload`.** El estado global sobrevive al cierre del lateral; no hay forma de invalidar `m_ObjUsuarioHistoricoActivo` desde aquí. Si el binario nunca lo limpia, el siguiente `00` lo encuentra poblado (a verificar con `VariableGlobal.bas` y `FG.bas`).
- **El `BackColor` del `Detalle` es `8210719`** (`.form.txt:230`); la paleta del `Entorno` puede sobrescribirlo con `m_ObjEntorno.CabeceraBackColor` (`FG.bas:971`). Si la paleta está sin inicializar, el binario pintado en el `.form.txt` es lo que se ve hasta la primera carga del `Entorno`.
- **Restos:** tres líneas en blanco entre `Form_Load` y `Public Sub lblAnexos_Click` (`:30-31` del `.cls`, huecos sin uso); misma línea de `m_ObjColParaFiltros` aquí no existe (era de `FormInicial03`); `pregunta` se guarda pero no se lee (VariablesEntorno compartido: `:27, 54, 83`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: cambiar el lateral a un componente router declarativo (`<router-outlet>` con dos rutas `general` y `anexos`) en lugar del `SourceObject` por código; derivar la pestaña activa del estado de la URL y no de un `Tag` de la etiqueta; mover la paleta a CSS variables y no a `On Error`-protegido en cada `Load`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al entrar con un usuario histórico activo, las etiquetas «Versión», «Datos en» y «Modo Escritura» muestran los valores leídos del `Entorno`.
2. La pestaña «Gral.» aparece marcada por defecto y el `SubFormCentral` del padre muestra `FormUsuarioHistorico02Edicion`.
3. Pulsar «Anexos» cambia el `SubFormCentral` a `FormUsuarioHistorico02Anexos` y lo repinta en estilo «seleccionado», dejando «Gral.» en «no seleccionado».
4. Pulsar la pestaña activa dos veces seguidas no produce reasignación ni repintado adicionales (guard de re-entrada).
5. Si no hay `m_ObjUsuarioHistoricoActivo`, el lateral carga sin notificar al usuario, y el `SubFormCentral` queda en el estado que tuviera el padre (a verificar).
6. Un fallo de `EstablecerControlesformlateral` se muestra como `MsgBox` crítico y no cierra la pestaña abierta.

## Pendientes

- Contenido de `m_ObjEntorno.ColEstadosVisiblesHPS` y el resto de paletas (`Entorno.cls:824-887`).
- Qué formularios llaman a `lblDatosGenerales_Click` o `lblAnexos_Click` desde fuera: solo el `Form_Load` propio (`:20`).
- Procedencia del `m_ObjUsuarioHistoricoActivo` que el padre (`FormUsuarioHistorico00`) carga; en `src/` no hay `DoCmd.OpenForm` para él.
- Si `EstablecerControlesformlateral` se aplica también al `FormInicial01Lateral` (mismo helper, `FG.bas:963-1004`; usado en `Form_FormInicial01Lateral.cls:333, 346, 381, 428, 438`).
