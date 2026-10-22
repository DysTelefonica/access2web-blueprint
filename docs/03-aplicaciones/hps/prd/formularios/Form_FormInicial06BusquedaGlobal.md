# Form_FormInicial06BusquedaGlobal

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial06BusquedaGlobal.cls` (504 líneas) y `.form.txt` (1 234 líneas) · **PRD:** [README](../README.md)

Búsqueda transversal sobre nueve tablas del frontend local. Un cuadro de texto (`PalabraClave`) alimenta `RellenarListaConBusquedaGlobal`, que delega en `Constructor.getColResultadoBusqueda` (`Constructor.bas:1774-2077`); cada coincidencia añade una fila a `ListaResultados` (`ID;Apartado;Nombre Campo;Valor`); doble-clic o `Ver Dato Seleccionado` enruta al resultado abriendo el usuario, histórico, observación o anexo correspondiente en el central del `FormInicial00Principal`. Es un formulario **no enlazado**: la búsqueda recorre los recordsets y compara con `InStr(1, m_Valor, p_Palabra) <> 0` en VBA, **sin concatenar el texto del usuario en un `WHERE`** — no hay inyección de SQL desde el cuadro de búsqueda. Las citas son `Form_FormInicial06BusquedaGlobal.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`; `Constructor.bas` es `modules/Constructor.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial00Principal.ComandoBuscar_Click` | `DoCmd.OpenForm "FormInicial06BusquedaGlobal"` y, si `Palabra <> ""`, propaga `frm.Controls("PalabraClave") = Palabra` y dispara `PalabraClave_Exit False` | `Form_FormInicial00Principal.cls:55-63` |
| `FormInicial00Principal.SubFormCentral` | Las funciones `Filtrar*` reescriben `SourceObject` (vía `Form_FormInicial01Lateral`) y desplazan el `Bookmark` del subform de resultados | `:148, :243-245, :326-328` |
| `m_ObjUsuarioSeleccionado`, `m_ObjExpedienteSeleccionado` (declarados) | **Privadas y nunca asignadas en este `.cls`** (`:6-7`); fueron previstas y quedaron en desuso | `:6-7` |
| `m_ObjUsuarioActivo`, `m_ObjUsuarioHistoricoActivo`, `m_ObjAnexoUsuarioHPSActivo`, `m_ObjAnexoUsuarioSICAActivo` | Escritos por las funciones `Filtrar*` como efecto secundario | `:157, :258, :273, :295`, `VariablesEntorno.bas:225-234` |

`CmdSalir_Click` cierra el formulario y limpia `Forms("FormInicial00Principal").PalabraClave = Null` (`:12-15`).

## Controles y eventos

`OnLoad ="[Event Procedure]"` (`.form.txt:82`); `DefaultView =0` (formulario simple, `.form.txt:15`); sin `RecordSource` en el formulario.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Pinta estilo, inicializa lista y botones, abre y cierra el reloj de arena | `:76-113`, `.form.txt:82` |
| `lblTitulo` | etiqueta («BÚSQUEDA GLOBAL», `Tag = "Elegibles;NoSeleccionado"`) | — | Título del encabezado; repintado por `EstablecerControlesDatos` | `.form.txt:313-315` |
| `Etiqueta14` | etiqueta («Palabra Clave») | — | Rótulo del cuadro de texto | `.form.txt:354-355` |
| `PalabraClave` | caja de texto (`Tag = "DATOS"`) | `OnExit` → `PalabraClave_Exit`; `OnKeyDown` → `PalabraClave_KeyDown` | Ejecuta la búsqueda al perder el foco o al pulsar `Enter` | `:356-395, :396-446`, `.form.txt:381-383` |
| `ListaResultados` | listbox de 4 columnas (`ColumnCount =4`, `ColumnWidths ="1134;2268;2268;5393"`, `RowSourceType ="Value List"`, `AllowValueListEdits =0`) | `OnClick` → `ListaResultados_Click`; `OnDblClick` → `ListaResultados_DblClick` | Habilita «Ver Dato Seleccionado» y, en doble-clic, lanza la navegación | `:115-136, :347-354`, `.form.txt:416-425` |
| `ComandoLimpiarEmpresaUsuario` | botón (icono `Cepillo.ico`, `ControlTipText` «Quita el contenido de todos los campos de búsqueda») | `OnClick` → `ComandoLimpiarEmpresaUsuario_Click` | Vacía `PalabraClave` y reinicia el `RowSource` | `:17-22`, `.form.txt:447-455` |
| `ComandoVerDatoSeleccionado` | botón («Ver Dato Seleccionado», `Enabled = NotDefault`) | `OnClick` → `ComandoVerDatoSeleccionado_Click` | Despacha el resultado al formulario de detalle que corresponda | `:24-74`, `.form.txt:629-633` |
| `cmdSalir` | botón («Salir») | `OnClick` → `CmdSalir_Click` | Limpia la palabra del padre y cierra | `:12-15`, `.form.txt:686-688` |

`Detalle` (`.form.txt:331`) contiene el cuerpo; `EncabezadoDelFormulario` (`.form.txt:296`) y `PieDelFormulario` (`.form.txt:672`) reciben el color de la paleta en `EstablecerControlesDatos` (`FG.bas:1051-1053`).

## Flujos

### Apertura (`Form_Load`, `:76-113`)

1. `VBA.DoEvents` y `DoCmd.Hourglass True` (`:79-81`).
2. `m_Error = ""` y `Me.AllowEdits = False` (`:82-83`).
3. `EstablecerControlesDatos Me, m_Error` aplica la paleta a etiquetas, campos y listas (`:84`; `FG.bas:1005-1060`); eleva 1000 si falla (`:85-87`).
4. `CopiarEstiloBoton` sobre `ComandoVerDatoSeleccionado` (`:88`) y `cmdSalir` (`:93`); eleva 1000 si cualquiera falla (`:89-91, :94-96`).
5. `Me.ListaResultados.RowSource = "ID;Apartado;Nombre Campo;Valor"` y `Me.ListaResultados.Requery` (`:98-99`).
6. `Me.AllowEdits = True` y `DoCmd.Hourglass False` (`:100-103`).
7. En error: `m_Error` poblado, `CorreoAlAdministrador` si no es 1000, `Me.AllowEdits = True`, `DoCmd.Hourglass False`, `MsgBox` crítico y `DoCmd.Close acForm, Me.Name, acSaveNo` (`:104-112`).

### Salir (`CmdSalir_Click`, `:12-15`)

`Forms("FormInicial00Principal").PalabraClave = Null` y `DoCmd.Close acForm, Me.Name, acSaveNo` (sin manejador de error).

### Limpiar (`ComandoLimpiarEmpresaUsuario_Click`, `:17-22`)

`On Error Resume Next` (`:18`); limpia `Me.PalabraClave`, restaura `RowSource` y `Requery` (`:19-21`). Es la única sentencia del `.cls` que omite la propagación de errores.

### Selección en la lista (`ListaResultados_Click`, `:115-136`)

1. `Me.ComandoVerDatoSeleccionado.Enabled = False` (`:121`).
2. `IDSeleccionado = Nz(Me.ListaResultados.Column(0), "")` y `ApartadoSeleccionado = Nz(Me.ListaResultados.Column(1), "")` (`:124-125`).
3. Si `IDSeleccionado = ""`, sale sin habilitar el botón (`:126-130`).
4. Si hay `ID`, habilita `Me.ComandoVerDatoSeleccionado.Enabled = True` (`:131`).
5. En error: `m_Error = "Al abrir el formulario se ha producido el error: …"`, `CorreoAlAdministrador` si no es 1000, `MsgBox` (`:132-135`).

### Doble-clic en la lista (`ListaResultados_DblClick`, `:347-354`)

`On Error Resume Next` (`:348`); `Application.Echo False` (`:349`); si `ComandoVerDatoSeleccionado.Enabled`, llama a `ComandoVerDatoSeleccionado_Click` (`:350-352`); `Application.Echo True` (`:353`). Equivale al botón, pero sin comprobación de error.

### Ver dato seleccionado (`ComandoVerDatoSeleccionado_Click`, `:24-74`)

1. Lee las dos primeras columnas (`IDSeleccionado`, `ApartadoSeleccionado`, `:28-29`).
2. Despacha por el valor de `ApartadoSeleccionado` a uno de los siete `Filtrar*`: «UsuarioS SICA» → `FiltrarUsuarioSICA` (`:31-35`), «Usuarios Históricos» → `FiltrarUsuarioHistorico` (`:36-40`), «Anexos» → `FiltrarAnexos` (`:41-45`), «Anexos SICA» → `FiltrarAnexosSICA` (`:46-50`), «Observaciones» → `FiltrarObservaciones` (`:51-55`), «Observaciones Históricas» → `FiltrarObservacionesHistoricas` (`:56-60`); el resto cae al `Else` y dispara `FiltrarUsuario` (`:61-66`). Cada rama eleva 1000 si el `Filtrar*` devuelve `m_Error` (`:33, :38, :43, :48, :53, :58, :63`).
3. En error: `m_Error = "Al intentar ver el dato se ha producido el error: …"`, `CorreoAlAdministrador` si no es 1000, `MsgBox` crítico (`:68-73`).

### Cambio del cuadro de búsqueda (`PalabraClave_Exit`, `:356-395`)

1. `Palabra = Nz(Me.PalabraClave, "")` y `Me.AllowEdits = False` (`:360-361`).
2. Si `Palabra = ""`, vacía la lista y sale (`:362-368`).
3. Si `Palabra <> m_PalabraClaveAlInicio` **o** la lista tiene una sola fila, abre el reloj de arena, deshabilita `ComandoVerDatoSeleccionado` y rellena la lista (`:369-383`).
4. `m_PalabraClaveAlInicio = Palabra` (`:384`) y `Me.AllowEdits = True` (`:385`).
5. En error: `m_Error = "Al realizar la búsqueda se ha producido el error: …"`, `CorreoAlAdministrador` si no es 1000, restaura `AllowEdits` y el reloj, `MsgBox` (`:387-394`).

`PalabraClave_KeyDown` (`:396-446`) replica ese flujo cuando la tecla es `Enter` (`vbKeyReturn`, `:401`): habilita y enfoca `ComandoLimpiarEmpresaUsuario`, devuelve el foco a `PalabraClave` (`:406-408`), busca con la misma condición de idempotencia (`:409-429`), guarda `m_PalabraClaveAlInicio` y anula el pitido con `KeyCode = 0` (`:432, :443`).

### Rellenar lista (`RellenarListaConBusquedaGlobal`, `:448-504`)

1. Restablece `lst.RowSource = "ID;Apartado;Nombre Campo;Valor"` (`:465`) y, si `p_Palabra = ""`, sale (`:466-468`).
2. `Set ColResultado = Constructor.getColResultadoBusqueda(p_Palabra, p_Error)` (`:475`); eleva 1000 si el constructor devolvió error (`:476-479`).
3. Si `ColResultado Is Nothing`, sale (`:480-482`).
4. Itera `ColResultado.Keys`, parte cada clave por `|`, exige `UBound(dato) = 3` (`:485`) y añade `id;apartado;campo;valor` con `Replace(…, ";", ":")` en los cuatro tramos para no chocar con el separador de columnas (`:486-491`).
5. Si la lista tiene exactamente dos filas —cabecera más una coincidencia—, selecciona la fila 1 y dispara `ListaResultados_Click` para habilitar el botón (`:494-497`).
6. En error: rellena `p_Error` con descripción; **no** envía correo ni muestra `MsgBox` (`:500-503`).

### Despacho por `Filtrar*` (siete funciones públicas)

- **`FiltrarUsuario(p_ID)`, `:138-169`**: si `FormInicial00Principal` no está abierto, lo abre (`:144-146`); `Form_FormInicial01Lateral.lblUsuariosHPS_Click` fuerza el central a `FormInicial02ConsultasPrincipal` (`:148`); navega `frm → SubFormCentral.Form → SubFormResultados.Form` (`:150-151`); `rst.FindFirst "[ID] =" & p_ID` (`:152-153`); si hay coincidencia, `Bookmark`, `Constructor.getUsuarioHPS` en `m_ObjUsuarioActivo` (`:154-160`) y `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` (`:161`).
- **`FiltrarObservaciones(p_IDObservacion)`, `:170-202`**: exige `p_IDObservacion <> ""` (`:177-179`); carga `Constructor.getObservacion` (`:180-183`); si existe y su `.Usuario` no es `Nothing`, delega en `FiltrarUsuario(m_ID, p_Error)` (`:187-194`).
- **`FiltrarObservacionesHistoricas(p_IDObservacion)`, `:203-230`**: análogo, con `Constructor.getObservacionHistorica` (`:211-214`) y `FiltrarUsuarioHistorico` (`:218-225`).
- **`FiltrarUsuarioHistorico(p_ID)`, `:231-265`**: abre el principal si hace falta (`:236-238`); si `SubFormCentral.SourceObject <> "FormInicial08UsuariosHistoricos"`, llama a `Form_FormInicial01Lateral.lblUsuariosHistoricos_Click` (`:243-245`); navega hasta `SubFormResultados.Form` (`:246-248`); `FindFirst "[ID] =" & p_ID` (`:249-251`); `Bookmark`, `frm.IDExpediente.SetFocus` (`:252-254`), `Constructor.getUsuarioHistorico` en `m_ObjUsuarioHistoricoActivo` (`:255-260`) y `Form_FormInicial08UsuariosHistoricos.ComandoVerUsuarioHistorico_Click` (`:261`).
- **`FiltrarAnexos(p_ID)`, `:266-289`**: `m_ObjAnexoUsuarioHPSActivo = getAnexoUsuarioHPS(p_ID, p_Error)` (`:272-275`); si el anexo no existe en `URLAnexo`, compone el mensaje y eleva 1000 (`:279-281`); `Ejecutar Me.hWnd, "open", m_ObjAnexoUsuarioHPSActivo.URLAnexo, "", "", 1` (`:283`) lanza el visor por `ShellExecuteA` (`VariablesEntorno.bas:5-7`).
- **`FiltrarAnexosSICA(p_ID)`, `:290-313`**: análogo, sobre `m_ObjAnexoUsuarioSICAActivo` (`:295`), `getAnexoUsuarioSICA` y su mensaje «anexo SICA» (`:302-306`).
- **`FiltrarUsuarioSICA(p_ID)`, `:314-344`**: abre el principal si hace falta (`:321-323`); si central ≠ `FormInicial07UsuariosSICA`, `Form_FormInicial01Lateral.lblUsuariosSICA_Click` (`:326-328`); `rst.FindFirst "[ID] ='" & p_ID & "'"` (`:331-332`, con comillas simples porque la clave SICA es texto); `Bookmark`, `frm.ID.SetFocus` (`:333-335`) y `Form_FormInicial07UsuariosSICA.ComandoEditarUsuarioSICA_Click` (`:336`).

### Errores

Los handlers siguen el patrón del repositorio: para errores no internos 1000, `m_Error = "Al … se ha producido el error: " & Err.Description`, `CorreoAlAdministrador m_Error` y `MsgBox` crítico (bloques de error en `:68-73, :104-112, :130-135, :387-394, :438-445`; `FG.bas:67`). Las siete funciones `Filtrar*` y `RellenarListaConBusquedaGlobal` solo rellenan `p_Error` sin notificar (`:165-168, :198-201, :226-229, :261-264, :285-288, :309-312, :340-343, :500-503`).

## Reglas, validaciones y cálculos

- **Búsqueda en memoria, no SQL.** `Constructor.getColResultadoBusqueda` recorre cada tabla con `DAO.Recordset` (`.MoveFirst`, `Do While Not .EOF`, `For Each fld In .Fields`) y compara cada campo con `InStr(1, m_Valor, p_Palabra) <> 0` (`Constructor.bas:1803, 1837, 1868, 1900, 1931, 1961, 1991, 2022, 2052`). El texto del usuario **no se concatena en un `WHERE`** y no se escapa porque no llega al motor de datos. Es el patrón más seguro de esta ficha.
- **Cobertura de nueve tablas.** Cinco locales vía `CurrentDb()`: `TbDatosLocal` (`Constructor.bas:1793`), `TbUsuariosSICALocal` (`Constructor.bas:1825`), `TbUsuariosHistoricosLocal` (`Constructor.bas:1857`), `TbObservacionesLocal` (`Constructor.bas:1890`), `TbObservacionesHistoricasLocal` (`Constructor.bas:1921`); cuatro remotas vía `getdb()`: `TbUsuarioAnexos` (`Constructor.bas:1951`), `TbAnexosUsuariosSICA` (`Constructor.bas:1981`), `TbObservaciones` (`Constructor.bas:2012`), `TbObservacionesHistoricas` (`Constructor.bas:2042`). El constructor recorre todas y devuelve un `Scripting.Dictionary` con `CompareMode = TextCompare` (`Constructor.bas:1804-1810`).
- **`ListaResultados` con `Value List`.** Las filas se añaden en VBA con `lst.AddItem` y separador `;`; los `;` del contenido se sustituyen por `:` (`:490-491`) porque `;` es el separador de columna del `RowSource` (`Constructor.bas:1808` compone la clave con `|`; el formulario reescribe a `;`).
- **Filtro de resultados con `UBound(dato) = 3`.** Si la clave devolvida tiene menos o más de cuatro tramos, la fila se descarta (`:485`); evita que un valor con `|` rompa el layout.
- **Idempotencia por palabra.** `PalabraClave_Exit` y `PalabraClave_KeyDown` solo releen la lista si la palabra cambió o si la lista quedó en una sola coincidencia (`:369, :414`), usando el campo público `m_PalabraClaveAlInicio` (`:10, :384, :432`).
- **`AllowEdits` como cerrojo.** `Form_Load`, `PalabraClave_Exit` y `PalabraClave_KeyDown` ponen `Me.AllowEdits = False` durante la búsqueda y lo restauran al terminar o en el `errores:` (`:83, :100, :361, :385, :392, :405, :431, :444`). `PalabraClave_KeyDown` además consulta `Me.AllowEdits` para reentrada (`:402-404`).
- **Sin captura de errores propia en las funciones.** Las siete `Filtrar*` y `RellenarListaConBusquedaGlobal` devuelven el error en `p_Error`; los handlers son responsables de elevar 1000 (`:33, :38, :43, :48, :53, :58, :63, :85, :89, :94, :365, :377, :421, :433, :477`).
- **Anclas por nombre de columna.** El constructor compone el primer tramo con el campo clave de cada tabla: `ID`, `IDObservacion` o `IDAnexo` (`Constructor.bas:1808, 1907, 1929, 1947, 1971, 2003, 2024, 2055`). Cambiar el nombre de esa columna obliga a tocar el constructor, no este formulario.
- **`OpenForm` defensivo solo en los `Filtrar*` de usuario.** `FiltrarUsuario`, `FiltrarUsuarioHistorico` y `FiltrarUsuarioSICA` abren `FormInicial00Principal` si no está (`:144-146, :236-238, :321-323`); `FiltrarAnexos` y `FiltrarAnexosSICA` no lo hacen (`:266-289, :290-313`).
- **Sin valores de cliente.** Ni el `RowSource` ni los rótulos contienen nombres de personas o empresas: solo cabeceras de columna y el literal «BÚSQUEDA GLOBAL» (`.form.txt:39, :314`).

## Lecturas y escrituras

| Operación | Origen | Tipo | Evidencia |
|---|---|---|---|
| Búsqueda global | nueve tablas (`TbDatosLocal`, `TbUsuariosSICALocal`, `TbUsuariosHistoricosLocal`, `TbObservacionesLocal`, `TbObservacionesHistoricasLocal`, `TbUsuarioAnexos`, `TbAnexosUsuariosSICA`, `TbObservaciones`, `TbObservacionesHistoricas`) | `OpenRecordset` + `InStr` en VBA | `Constructor.bas:1774-2077` |
| Apertura de anexo | URL del anexo (`URLAnexo`) | `ShellExecuteA` (`Ejecutar`) | `:283, :306`, `VariablesEntorno.bas:5-7` |
| Navegación al resultado | `frm.Controls("SubFormCentral").Form.Controls("SubFormResultados").Form.Bookmark` | propiedades del subformulario | `:150-155, :246-254, :329-335` |
| Activar pestaña central | `Form_FormInicial01Lateral.lblUsuarios{HPS,Historicos,SICA}_Click` | procedimiento público del lateral | `:148, :244, :327` |
| Estado activo | `m_ObjUsuarioActivo`, `m_ObjUsuarioHistoricoActivo`, `m_ObjAnexoUsuarioHPSActivo`, `m_ObjAnexoUsuarioSICAActivo` | globales en `VariablesEntorno` | `:157, :258, :273, :295`, `VariablesEntorno.bas:225-234` |

Sin escrituras propias: ningún `INSERT`, `UPDATE` ni `DELETE` desde este formulario.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: el formulario se comunica con cuatro formularios hermanos **solo por nombre**, así que un cambio de nombre rompe la llamada sin que el compilador avise — `FiltrarUsuario` depende de `Form_FormInicial01Lateral.lblUsuariosHPS_Click` (`:148`) y de `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` (`:161`); `FiltrarUsuarioSICA` depende de `lblUsuariosSICA_Click` (`:327`); `FiltrarUsuarioHistorico` depende de `lblUsuariosHistoricos_Click` (`:244`). El `SourceObject` se alcanza por dos rutas distintas (`m_Frm.SubFormCentral.SourceObject` en el lateral, `frm.Controls("SubFormCentral").SourceObject` aquí) y los nombres `SubFormCentral` y `SubFormResultados` se repiten en las tres funciones de usuario y en el lateral, sin registro compartido. La doble gestión de errores (helper silencioso + handler que eleva 1000) es coherente con el repositorio, pero significa que un fallo en `getColResultadoBusqueda` se notifica desde `PalabraClave_Exit` y no desde el helper. Al salir, el padre queda con `PalabraClave = Null` (`:13`), de modo que la próxima búsqueda parte de cero.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Typo en el despacho por apartado:** el `If` compara `ApartadoSeleccionado = "UsuarioS SICA"` (`:31`) con la «S» mayúscula intermedia, mientras que el constructor escribe `Apartado = "Usuarios SICA"` (`Constructor.bas:1824`). La condición nunca es cierta, así que los resultados de SICA caen al `Else` y ejecutan `FiltrarUsuario` (`:61-66`), que busca con `FindFirst "[ID] =" & p_ID` **sin comillas** (`:153`) sobre un campo `ID` de texto en SICA. A verificar contra el binario, pero la fuente escrita basta para concluir que la rama nunca aplica.
- **`m_ObjUsuarioSeleccionado` y `m_ObjExpedienteSeleccionado`** (`:6-7`) se declaran y no se asignan en este `.cls`: quedan como variables muertas.
- **`ComandoLimpiarEmpresaUsuario_Click` corre con `On Error Resume Next`** (`:18`) y no pone `ComandoVerDatoSeleccionado.Enabled = False`. Si la última búsqueda dejó una selección y el usuario pulsa limpiar, el botón puede quedar habilitado con la lista vacía (a verificar en runtime).
- **`PalabraClave_KeyDown` salta el foco dos veces** (`:406-408`: `ComandoLimpiarEmpresaUsuario.SetFocus` y de vuelta a `PalabraClave`) y luego reproduce a mano lo que ya hace `PalabraClave_Exit` (`:409-429` frente a `:362-383`). Son dos copias del mismo algoritmo; cualquier corrección hay que aplicarla dos veces.
- **`CmdSalir_Click` no tiene manejador de error** (`:12-15`): si el padre se cerró antes, `Forms("FormInicial00Principal").PalabraClave = Null` falla con el error 2455 sin notificación (a verificar).
- **`FiltrarAnexos*` no abren `FormInicial00Principal`** (`:266-289, :290-313`) mientras que las tres funciones de usuario sí lo hacen (`:144-146, :236-238, :321-323`); la asimetría sugiere que el anexo debería poder abrirse sin la pantalla principal, pero no está justificado en el código (a verificar).
- **`getColResultadoBusqueda` carga las nueve tablas completas en memoria** y compara campo a campo en VBA (`Constructor.bas:1795-1818, 1827-1850, …`): con `TbDatosLocal` y `TbObservaciones` pobladas el coste es lineal en filas × campos (a verificar con el volumen real).
- **`FindFirst "[ID] ='" & p_ID & "'"`** (`:332`) concatena comillas simples sin escapar comillas internas. `p_ID` procede de la columna 0 de la lista, no del cuadro de búsqueda, así que el vector es indirecto, pero un `ID` con apóstrofe rompería la SQL (a verificar).
- **`AllowEdits` como cerrojo es frágil.** `PalabraClave_KeyDown` consulta `Me.AllowEdits` para descartar la reentrada (`:402-404`) pero lo restaura en tres sitios distintos (`:431, :440, :444`); si la búsqueda se relanza desde otro camino mientras el cerrojo está bajo, la lista puede quedar a medio rellenar (a verificar).
- **`FiltrarObservaciones*` ignoran su valor de retorno** (`:191, :222`): asignan `FiltrarObservaciones = FiltrarUsuario(...)` pero el llamador `ComandoVerDatoSeleccionado_Click` solo comprueba `m_Error`; una navegación fallida que no produzca error queda sin señal (a verificar).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: mover la búsqueda a un endpoint con parámetros vinculados (`ILIKE`) en vez de `InStr` sobre recordsets en memoria; corregir el literal `UsuarioS SICA` o sustituir el despacho por un identificador de apartado estable en vez de un texto traducido; sustituir `ShellExecuteA` por un endpoint que sirva los anexos con autorización, con el `URLAnexo` validado en servidor; unificar `PalabraClave_Exit` y `PalabraClave_KeyDown` en un único manejador reutilizado; y un registro declarativo de los nombres de formulario y de procedimiento para eliminar la dependencia por cadena.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al entrar, `ListaResultados.RowSource = "ID;Apartado;Nombre Campo;Valor"` y `ComandoVerDatoSeleccionado` queda deshabilitado.
2. Salir (`cmdSalir`) cierra el formulario y pone `Forms("FormInicial00Principal").PalabraClave` a `Null`.
3. Teclear una palabra y pulsar `Enter` (o perder el foco) rellena `ListaResultados` con los `ID;Apartado;Nombre Campo;Valor` que coincidan en alguna de las nueve tablas, sin incluir texto de clientes en la definición del recurso.
4. Si hay exactamente una coincidencia, la lista la deja seleccionada y «Ver Dato Seleccionado» queda habilitado sin clic adicional.
5. Seleccionar una fila habilita el botón; el doble-clic produce el mismo efecto que el botón.
6. «Ver Dato Seleccionado» con un resultado de «Usuarios» abre el central de HPS con el registro seleccionado; con «Anexos» abre el visor; con «Usuarios Históricos» abre el central de históricos y muestra el registro; con «Observaciones» navega al usuario dueño de la observación.
7. Limpiar vacía la lista y la palabra y deja «Ver Dato Seleccionado» deshabilitado.
8. Los errores no internos generan aviso al administrador y `MsgBox` crítico; los de las funciones auxiliares se atribuyen al evento que las invocó.

## Pendientes

- Confirmar en el binario si el literal correcto del apartado SICA es `"Usuarios SICA"` y por qué `ComandoVerDatoSeleccionado_Click` escribe `"UsuarioS SICA"`: si el constructor cambió el nombre del apartado sin actualizar este `Select`, la rama lleva tiempo muerta.
- Tamaño real de cada una de las nueve tablas para acotar el tiempo de búsqueda.
- Si existe algún otro consumidor de `RellenarListaConBusquedaGlobal` además de `PalabraClave_Exit` y `PalabraClave_KeyDown` (`src/` no muestra más llamadas).
- `Forms("FormInicial00Principal").PalabraClave` es un control del padre que este formulario escribe desde `CmdSalir_Click` (`:13`): documentar esa dependencia inversa en `../coexistencia.md`.