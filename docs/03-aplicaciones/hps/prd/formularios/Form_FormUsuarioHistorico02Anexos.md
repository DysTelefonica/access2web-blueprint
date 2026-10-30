# Form_FormUsuarioHistorico02Anexos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuarioHistorico02Anexos.cls` (156 líneas) y `.form.txt` (1018 líneas) · **PRD:** [README](../README.md)

Subformulario central de la pestaña «Anexos» del histórico: lista los anexos del usuario histórico activo y los abre con el visor del sistema al hacer doble clic. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos se leen cada vez a través de `m_ObjUsuarioHistoricoActivo.ColAnexos`, que a su vez pide a `Constructor.getAnexosUsuarioHistorico` un `Dictionary` de `AnexoUsuarioHistorico` por `IDUsuario`. Un único `ListBox` (`ListaAnexos`, dos columnas: identificador y nombre) soporta el navegador; la URL se construye con `AnexoUsuarioHistorico.URLAnexo`. Las citas son `Form_FormUsuarioHistorico02Anexos.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuarioHistorico01Lateral.lblAnexos_Click` | Fija `Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuarioHistorico02Anexos"` | `Form_FormUsuarioHistorico01Lateral.cls:47`, `Form_FormUsuarioHistorico00Principal.form.txt:890` |
| `Form_FormUsuarioHistorico00Principal` | El subformulario central declara `SourceObject = "Form.FormUsuarioHistorico02Edicion"` por defecto; este formulario es la alternativa | `Form_FormUsuarioHistorico00Principal.form.txt:891` |

Abre el visor asociado al `URLAnexo` mediante `Ejecutar Me.hwnd, "open", m_AnexoSeleccionado.URLAnexo, …` (`:144`; `VariablesEntorno.bas:5, 42`). No abre otros formularios de Access.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | Aplica estilo, aborta si no hay usuario activo, rellena la lista | `:47-77` |
| (formulario) | — | `OnUnload` implícito | Sin `Form_Unload` propio; `m_AnexoSeleccionado` queda como variable de módulo y sobrevive al cierre | — |
| `ListaAnexos` | listbox, `RowSourceType = Value List`, 2 columnas | `OnClick` → `ListaAnexos_Click` | Carga `m_AnexoSeleccionado` para el `IDAnexo` elegido | `:79-121`, `.form.txt:820-854` |
| `ListaAnexos` | listbox | `OnDblClick` → `ListaAnexos_DblClick` | Abre el archivo `URLAnexo` con el visor del sistema | `:123-156`, `.form.txt:837-838` |
| `Detalle` | sección | — | `Height = 7711`; sin lógica | `.form.txt:810` |
| (formulario) | — | `Caption = "Anexos"` | Rótulo de la pestaña en el `SourceObject` | `.form.txt:38` |

## Flujos

### Apertura (`Form_Open`, `:47-77`)

1. `On Error GoTo errores` (`:49`).
2. `Me.AllowEdits = False` durante el montaje (`:50`).
3. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores (`FG.bas:1005-1053`). Si falla, eleva 1000 (`:51-54`).
4. Si `m_ObjUsuarioHistoricoActivo Is Nothing`, fija `m_Error = "No se ha seleccionado ningún usuario"` y eleva 1000 (`:56-59`).
5. `EstablecerLista m_Error` rellena `ListaAnexos`. Si falla, eleva 1000 (`:60-63`).
6. Reactiva edición y habilita la lista: `Me.AllowEdits = True` y `Me.ListaAnexos.Enabled = True` (`:64-65`).
7. En error: compone mensaje, envía correo al administrador, **reactiva `AllowEdits = True`**, muestra `MsgBox` crítico guardando el resultado en `pregunta` y **cierra el formulario sin guardar** (`:68-75`). El orden con `AllowEdits = True` antes del `MsgBox` es deliberado: dejar el formulario editable evita el aspecto de «cargando» si el usuario cancela el mensaje.

### Cargar la lista (`EstablecerLista`, `:8-43`)

1. `On Error GoTo errores` (`:16`).
2. Toma la lista: `Set lst = Me.ListaAnexos`, vacía `RowSource` (`:18-19`).
3. Lee la colección del global: `Set m_Col = m_ObjUsuarioHistoricoActivo.ColAnexos` (`:20`; `UsuarioHistorico.cls:234-262`). Esa colección es perezosa: la crea `Constructor.getAnexosUsuarioHistorico` con `SELECT * FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=…` (`Constructor.bas:508-565`).
4. `p_Error = m_ObjUsuarioHistoricoActivo.Error`: recoge el último error del `UsuarioHistorico` y eleva 1000 si está relleno (`:21-24`).
5. Si la colección no es `Nothing`, recorre cada `IDAnexo`, resuelve el `AnexoUsuarioHistorico` del diccionario y añade `m_IDAnexo & ";" & m_Anexo.NombreAnexo` a la lista (`:25-31`).
6. Si la lista solo tiene un elemento, lo selecciona y llama a `ListaAnexos_Click` (`:32-35`): comportamiento de «autoselección cuando hay un único anexo».
7. Manejador de error: `p_Error = "El método EstablecerLista ha dado el error: " & vbNewLine & Err.Description` si no es 1000 (`:39-42`). No envía correo: el llamador (`Form_Open`) lo hace.

### Cambio de selección (`ListaAnexos_Click`, `:79-121`)

1. `On Error GoTo errores` (`:84`).
2. Lee `m_IDAnexoSeleccionado = Nz(Me.ListaAnexos.Column(0), "")` — columna del identificador, oculta (`ColumnWidths = "0;6462"`, `.form.txt:836`) (`:86`).
3. Si está vacío (o el usuario hizo clic en el encabezado), pone `m_AnexoSeleccionado = Nothing` y sale (`:87-90`): la pulsación queda en el aire, no se carga nada.
4. Si `m_AnexoSeleccionado` aún no está cargado, lo pide con `Constructor.getAnexoUsuarioHistorico(m_IDAnexoSeleccionado, m_Error)` (`:91-92`; `Constructor.bas:466-505`). Si falla, eleva 1000 (`:93-95`).
5. Si devolvió `Nothing`, el anexo no existe en `TbAnexosUsuariosHistoricos`: «El Anexo del Usuario Histórico seleccionado no parece estar registrado» y eleva 1000 (`:96-99`).
6. Si `m_AnexoSeleccionado` ya estaba cargado pero el `IDAnexo` cambió, **vuelve a buscarlo** (`:100-110`). Esta llamada a la línea `:102` está escrita sin el prefijo `Constructor.` (a verificar, ver defectos).
7. En error: mensaje, correo al administrador, apaga el reloj de arena, `MsgBox` crítico (`:114-120`).

### Doble clic para abrir (`ListaAnexos_DblClick`, `:123-156`)

1. `On Error GoTo errores` (`:126`).
2. `Application.Echo False` para evitar parpadeo durante la apertura externa (`:128`).
3. Si `m_AnexoSeleccionado` no está cargado (por ejemplo, pulsar doble clic antes que clic), lo carga igual que en `ListaAnexos_Click` (`:129-138`).
4. `If Not fso.FileExists(m_AnexoSeleccionado.URLAnexo)` (`:140`): `URLAnexo` se construye concatenando `m_ObjEntorno.URLCarpetaAnexosHistoricos`, el `DNI` del usuario y el `NombreAnexo` (`AnexoUsuarioHistorico.cls:56-89`; `UsuarioHistorico.cls:267-298`). Si no existe, eleva 1000 con un mensaje que incluye la URL (`:140-143`).
5. `Ejecutar Me.hWnd, "open", m_AnexoSeleccionado.URLAnexo, "", "", 1` abre el archivo con la aplicación registrada para su extensión — el sexto argumento (`nShowCmd = 1`) equivale a `SW_SHOWNORMAL` (`:144`; `VariablesEntorno.bas:5, 42`).
6. `Application.Echo True` (`:145`).
7. En error: mensaje, correo, reactiva `Echo`, apaga el reloj, `MsgBox` (`:148-155`).

### Errores

Los tres procedimientos públicos siguen el mismo patrón: si `Err.Number <> 1000`, componen el mensaje con `vbNewLine`, llaman a `CorreoAlAdministrador` (`FG.bas:67-149`) y muestran un `MsgBox` crítico guardando el resultado en `pregunta` (`:71-74, :114-120, :148-155`). `Form_Open` cierra el formulario tras el error; `ListaAnexos_Click` y `ListaAnexos_DblClick` no lo cierran: el doble clic puede haber sido legítimo y la siguiente selección debe seguir funcionando. Adicionalmente, `ListaAnexos_Click` y `ListaAnexos_DblClick` desactivan el reloj de arena (`DoCmd.Hourglass False`) —el reloj se enciende en algún llamador a `PasarAActual` del formulario padre, no aquí (a verificar contra el binario, donde podría estar vinculado a la operación padre)—.

## Reglas, validaciones y cálculos

- **Identidad por global, no por argumento.** El usuario y su colección de anexos vienen de `m_ObjUsuarioHistoricoActivo`, no de un `OpenArgs` (`:20, 56`; `VariablesEntorno.bas:227`). El control asume que el padre ya validó la existencia.
- **Una sola fila: autoselección.** Si la lista termina con `ListCount = 1`, se selecciona y se ejecuta `ListaAnexos_Click` para precargar `m_AnexoSeleccionado` (`:32-35`). Reduce un clic en el caso frecuente de usuario con un único anexo.
- **Cancelación silenciosa.** Pulsar sobre el encabezado (sin `IDAnexo` válido) pone `m_AnexoSeleccionado = Nothing` y sale sin error (`:87-90`): la pulsación no produce carga, doble clic ni `Err.Raise`.
- **Recarga solo si cambia el `IDAnexo`.** Si el usuario ya tenía un anexo cargado y vuelve a hacer clic en la misma fila, no se vuelve a `getAnexoUsuarioHistorico` (`:100-101`) — se evita una llamada a base por clic redundante. La comparación es por `IDAnexo` (`AnexoUsuarioHistorico.cls:14`).
- **No abrir si el archivo no existe.** `fso.FileExists(m_AnexoSeleccionado.URLAnexo)` falla antes de invocar `Ejecutar` (`:140-143`): `Ejecutar` (alias `ShellExecuteA`) devolvería un `ERROR_FILE_NOT_FOUND` opaco que aquí se evita con un mensaje que incluye la URL.
- **`URLAnexo` es de solo lectura y se memoiza.** `AnexoUsuarioHistorico` guarda el resultado la primera vez (`AnexoUsuarioHistorico.cls:22, 81-82`); pulsaciones repetidas no reconstruyen la ruta.
- **Capa de presentación silenciosa.** El formulario no escribe en la base: ni `UPDATE`, ni `INSERT`, ni `DELETE`. Es un visor.
- **Lista con dos columnas, una oculta.** `ColumnCount = 2` y `ColumnWidths = "0;6462"` (`.form.txt:826, 836`): `IDAnexo` queda en ancho 0 (técnicamente visible para VBA), `NombreAnexo` es el texto que ve el usuario.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de anexos del histórico | `TbAnexosUsuariosHistoricos` (`SELECT * … WHERE IDUsuario=…`) | `Recordset` abierto por `Constructor.getAnexosUsuarioHistorico` | `Constructor.bas:508-565` |
| Anexo individual por `IDAnexo` | `TbAnexosUsuariosHistoricos` (`SELECT * … WHERE IDAnexo=…`) | `Recordset` abierto por `Constructor.getAnexoUsuarioHistorico` | `:92, 130`; `Constructor.bas:466-505` |
| URL local del anexo | `m_ObjEntorno.URLCarpetaAnexosHistoricos & DNI & "\" & NombreAnexo` | propiedad calculada | `AnexoUsuarioHistorico.cls:56-89`; `UsuarioHistorico.cls:267-298` |
| Apertura con el visor del sistema | `ShellExecuteA(hwnd, "open", url, "", "", SW_SHOWNORMAL)` | API declarada en `VariablesEntorno.bas` | `:144`; `VariablesEntorno.bas:5, 42` |
| Estilo de controles | `EstablecerControlesDatos Me, m_Error` | helper de `FG.bas` | `:51`; `FG.bas:1005-1053` |

Sin escrituras, sin transacciones.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo concreto: `Application.Echo` y `DoCmd.Hourglass` son APIs de Access sin equivalente directo; la suspensión del repintado mientras el visor arranca (`:128, 145`) es un patrón legados que web ignora. `fso` (`FileSystemObject`) sigue siendo dependencia de VBA (`VariablesEntorno.bas:218`). El `PredeclaredId = True` y el `VB_Name` coinciden entre `.cls` y `.form.txt` (`:1`; `.form.txt:860`), pero el binario real del formulario podría diferir del snapshot (`vba-binary-drift` aplicaría si se recompila; el `5e64cf9` fija la lectura). La doble columna con ancho cero para el `IDAnexo` (`.form.txt:826, 836`) es el patrón equivalente a un `key` no visible en una tabla web; en la API debe enviarse el dato de algún modo sin mostrarlo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Falta el prefijo `Constructor.` en la línea :102.** `Set m_AnexoSeleccionado = getAnexoUsuarioHistorico(m_IDAnexoSeleccionado, m_Error)` aparece sin `Constructor.` delante, al contrario que las líneas :92 y :130 que sí lo usan. `getAnexoUsuarioHistorico` no es un procedimiento global del módulo `Form_`; en las otras dos llamadas la búsqueda se delega explícitamente en `Constructor`. Sin ese prefijo, VBA resuelve al procedimiento del propio módulo (`Public Sub …` no encontrado aquí) y, al no existir, eleva error 424 «Objeto requerido» o 438 «El método no es compatible» en tiempo de ejecución. El path «segundo clic en la misma fila cambia de anexo» es, en la práctica, **el único que ejecuta la línea :102** y, por tanto, **el único que dispara el defecto** (a verificar; el `getAnexoUsuarioHistorico` también se invoca directamente desde la línea :130 con prefijo correcto, pero esa es la rama «doble clic sin selección previa», no la del `:102`).
- **`EstablecerLista` sobrescribe `p_Error` antes de propagarlo** (`:21`). Si la carga de `ColAnexos` falla, `UsuarioHistorico.Error` puede contener el mensaje real, pero `p_Error` se asigna encima: la traza que llega al `MsgBox` menciona el método `EstablecerLista`, no la causa original. (`:60-63` lo eleva como 1000 igualmente, pero el mensaje del `.Error` que está asignando `:21` es ya el compuesto por el propio `EstablecerLista`, porque el bloque `On Error GoTo errores` ya está activo cuando se lee `:21` — comprobar si antes del `:21` `m_ObjUsuarioHistoricoActivo.Error` se ha quedado como «propagado» desde `getAnexosUsuarioHistorico`; a verificar contra el binario.)
- **`p_Error = m_Error` no aparece nunca escrito**, pero `m_Error` se vacía al principio (`:17`). Si el `getAnexosUsuarioHistorico` había dejado algo en `m_ObjUsuarioHistoricoActivo.Error`, la asignación a `p_Error` es correcta. Pero **no se vacía `m_ObjUsuarioHistoricoActivo.Error` antes de la lectura**, así que un error anterior que el `UsuarioHistorico` no haya limpiado por completo puede aparecer como `p_Error` en una carga exitosa. Convención a confirmar.
- **`AllowEdits = False` se restaura a `True` dentro del manejador de error** (`:73`) antes del `MsgBox`. El efecto es que el cierre con «Aceptar» hereda `AllowEdits = True` durante el `DoCmd.Close`; normalmente `acSaveNo` ignora el flag, pero queda la duda: en algunos paths podría dispararse `Dirty` antes de cerrar. (A verificar.)
- **`DoCmd.Hourglass False` aparece sin un `DoCmd.Hourglass True` correspondiente en este formulario** (`:119, 154`). Alguno de los formularios padre debe encenderlo (por ejemplo, al pasar a Actual) y este lo apaga para limpiar el cursor. Si no hay un padre con el reloj encendido, la llamada es inocua pero deja la duda sobre quién lo activa: a verificar en `FormUsuarioHistorico00Principal` y `Indicadores.cls`.
- **`ListaAnexos.Column(0)` devuelve la columna 0, que es el `IDAnexo`**, no el nombre. La rama «clic vacío en el encabezado» pone `m_AnexoSeleccionado = Nothing` (`:88-89`): la siguiente pulsación con un anexo válido no entra en la rama «ya cargado» (porque ahora es `Nothing`) y recarga — bien. La rama «clic en la misma fila» (`:100-101`) compara contra `m_AnexoSeleccionado.IDAnexo`: correcto.
- **`fso` como dependencia global.** `VariablesEntorno.bas:218` declara `Public fso As New Scripting.FileSystemObject`; no se inicializa aquí, pero se usa en `:140`. Cualquier test en sandbox debe inyectar `fso` o crear el suyo (a verificar con `vba-sandbox`).
- **`Constructor` no declarado como tipo concreto.** VBA resuelve en tiempo de ejecución; si más adelante se renombra, las tres llamadas (`:92, 102, 130`) necesitan el cambio manual. Es una de las convenciones del repositorio, mantenida en este formulario.
- **Restos:** `pregunta` se guarda pero no se lee (`:74, 120, 155`); `m_SQL`/`flag`/`dato` declarados en `VariablesEntorno.bas:220` son globales de amplio espectro cuyo valor aquí se ignora. Sin `Form_Unload`, `m_AnexoSeleccionado` y `m_Error` sobreviven al cierre del subformulario; si el padre abre una ficha de histórico distinta sin recargar, arrastran referencia al anexo anterior.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: vista de tabla o grid en lugar de `ListBox`; descarga directa del archivo en lugar de delegar en la aplicación del SO; thumbnail o previsualización en línea del PDF/imagen cuando lo soporte el tipo MIME; endpoint API con el archivo en lugar de una URL de sistema de archivos; estado de «anexo no alcanzable» explícito en la UI en vez de `MsgBox` con la URL concatenada.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al entrar con un usuario histórico activo, la lista muestra todos los anexos de `TbAnexosUsuariosHistoricos` con `IDUsuario` igual al del histórico, por nombre.
2. Si el histórico no tiene anexos, la lista queda vacía sin error y ningún anexo queda preseleccionado.
3. Si el histórico tiene un único anexo, queda preseleccionado en la lista y `m_AnexoSeleccionado` queda listo para abrirse sin un clic previo.
4. Pulsar un anexo distinto carga el `AnexoUsuarioHistorico` correspondiente (a través de `Constructor.getAnexoUsuarioHistorico`), y un segundo clic sobre la misma fila no recarga.
5. Doble clic sobre un anexo alcanzable lo abre con el visor asociado a su extensión; doble clic sobre un anexo no alcanzable muestra «No es alcanzable en estos momentos el anexo Histórico seleccionado» con la URL.
6. Si no hay usuario histórico activo al abrir, se muestra «No se ha seleccionado ningún usuario» y se cierra sin guardar.

## Pendientes

- Procedencia exacta de `m_ObjUsuarioHistoricoActivo.ColAnexos`: `UsuarioHistorico.cls:234-262` lo carga, `Constructor.bas:508-565` lo hidrata; confirmar que `getAnexosUsuarioHistorico` filtra por `IDUsuario` y no por `IDHistorico` u otro.
- Forma de `AnexoUsuarioHistorico.URLAnexo` cuando el usuario está en oficina vs. remoto: el helper concatena `m_ObjEntorno.URLCarpetaAnexosHistoricos`, que cambia por entorno (`Entorno.cls`); a verificar si en el historial real el directorio es estable o se reconstruye.
- Quién enciende `DoCmd.Hourglass`: ningún `Hourglass True` en este formulario; el apagado (`:119, 154`) presupone un encendido externo.
- `paso a Actual` desde el padre mueve anexos (`UsuarioLifecycleTransactionCoordinator`, ver ficha de `FormUsuarioHistorico00`): tras un movimiento correcto, la lista debería quedar vacía o señalar el nuevo estado del histórico; la UI actual no se refresca (a verificar en el central editor).
- Si la columna 0 (`IDAnexo`) oculta cumple alguna función interna que no se ve: la regla que compara con `m_AnexoSeleccionado.IDAnexo` (`:101`) la necesita sí o sí.
