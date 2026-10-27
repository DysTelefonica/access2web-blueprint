# Form_FormUsuario02Anexos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario02Anexos.cls` (590 líneas) y `.form.txt` (1714 líneas) · **PRD:** [README](../README.md)

Subformulario central de la pestaña «Anexos» del usuario HPS en curso: gestiona los ficheros anexos al usuario activo repartidos en dos listas —`ACTUALES` e `HISTÓRICOS`—, permite anexar ficheros desde disco, eliminarlos, mover en lote los seleccionados de un lado al otro y abrirlos con la aplicación del sistema. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos llegan por el global `m_ObjUsuarioActivo`, declarado como `UsuarioHPS` en `VariablesEntorno.bas:225`, y sus colecciones `colAnexosActuales` / `colAnexosHistoricos` se releen de `TbAnexosUsuariosHPS` cada vez que el formulario invalida la caché. Las citas son `Form_FormUsuario02Anexos.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `Constructor.bas` `modules/Constructor.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuario01Lateral.lblAnexos_Click` | `SeleccionarEnlace "ANEXO"` y `Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuario02Anexos"` | `Form_FormUsuario01Lateral.cls:38, 47` |
| mismo lateral | Sale sin recargar si ese formulario ya es el `SourceObject` activo | `Form_FormUsuario01Lateral.cls:35` |
| `Form_FormUsuario01Lateral` (otros enlaces) | Los otros dos destinos del `SubFormCentral` son `FormUsuario02Edicion` y `FormUsuario02HPS` | `Form_FormUsuario01Lateral.cls:63, 102` |

Sale hacia fuera del sistema con `Ejecutar Me.hWnd, "open", <url>, "", "", 1` (`:474, 577`; `VariablesEntorno.bas:5`) y, si la apertura del formulario falla, se cierra sin guardar (`:374`). No abre otros formularios de Access.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | Ajusta tamaño, aplica la paleta, bloquea o habilita los cinco comandos según solo lectura, exige usuario activo y rellena las dos listas | `:332-376`; `.form.txt:39` |
| (formulario) | — | `OnUnload` implícito | Sin `Form_Unload`; `m_Error` y `m_AnexoSeleccionado` (`:4-5`) son de módulo | — |
| `ListaAnexosActuales` | listbox `Value List`, 2 columnas, `MultiSelect =2` | `OnClick` → `ListaAnexosActuales_Click`; `OnDblClick` → `ListaAnexosActuales_DblClick` | Selecciona y abre anexos del lado actual | `:378-487`; `.form.txt:835`, bloque `.form.txt:822-839` |
| `ListaAnexosHistoricos` | listbox `Value List`, 2 columnas, `MultiSelect =2` | `OnClick` → `ListaAnexosHistoricos_Click`; `OnDblClick` → `ListaAnexoshistoricos_DblClick` | Igual, en el lado histórico | `:489-590`; `.form.txt:902`, bloque `.form.txt:886-906` |
| `Etiqueta1`, `Etiqueta2` | etiqueta | — | Rótulos «ACTUALES» y «HISTÓRICOS» | `.form.txt:871, 938` |
| `ComandoAnexar` | botón «Anexar» | `OnClick` → `ComandoAnexar_Click` | Selección múltiple de ficheros; alta de un anexo por fichero | `:104-154`; `.form.txt:962` |
| `ComandoPasarAHistorico`, `ComandoPasarAActual` | botones «A Histórico», «A Actual» | `OnClick` → `ComandoPasarAHistorico_Click`, `ComandoPasarAActual_Click` | Movimiento en lote de las filas seleccionadas de la lista contraria | `:227-330`; `.form.txt:1058, 1105` |
| `ComandoEliminarAnexoActual`, `ComandoEliminarAnexoHistorico` | botones «Eliminar» | `OnClick` → `ComandoEliminarAnexoActual_Click`, `ComandoEliminarAnexoHistorico_Click` | Borrado del anexo seleccionado tras confirmación | `:156-225`; `.form.txt:1009, 1152` |

Ambos listbox llevan `ColumnCount =2` y `ColumnWidths ="0;6462"`: la columna 0 (el `IDAnexo`) queda con ancho cero y solo se ve el `NombreAnexo` (`.form.txt:827, 837` y `.form.txt:893, 904`).

## Flujos

### Apertura (`Form_Open`, `:332-376`)

1. Deja el formulario en `AllowEdits = False`, ajusta el tamaño y aplica la paleta de controles (`:336-338`; `FG.bas:901, 1005`).
2. Si `Application.TempVars("SoloLectura") = "Sí"`, desactiva los cinco comandos; si no, habilita al menos `ComandoAnexar` (`:342-350`).
3. Si `m_ObjUsuarioActivo Is Nothing`, compone «No se ha seleccionado ningún usuario» y eleva 1000 (`:351-354`).
4. Rellena ambas listas (`:355-358`), reactiva `AllowEdits` y las dos listas (`:359-361`) y aplica el estilo de los tres botones de acción (`:362-364`).
5. En error: compone el mensaje, avisa al administrador, reactiva `AllowEdits`, muestra `MsgBox` crítico y **cierra el formulario sin guardar** (`:367-374`).

### Rellenar las listas (`EstablecerListas`, `:8-61`)

1. Invalida la caché y fuerza la relectura: `Set m_ObjUsuarioActivo.colAnexosActuales = Nothing` y, acto seguido, lectura de la propiedad (`:19-20`), que a su vez llama a `Constructor.getAnexosUsuarioHPS(Me.ID, EnumSiNo.No, …)` (`UsuarioHPS.cls:384`).
2. Propaga el error pendiente del objeto y eleva 1000 si viene relleno (`:21-23`).
3. Añade una fila por anexo con `IDAnexo & ";" & NombreAnexo` (`:28`).
4. Si la lista queda con un único elemento, lo selecciona y dispara su `Click` (`:31-33`): autoselección cuando hay un solo anexo.
5. Repite el mismo bloque para el lado histórico, contra `colAnexosHistoricos` (`:36-53`; `UsuarioHPS.cls:413`).
6. El manejador solo añade la descripción cuando `Err.Number <> 1000`, para no pisar el mensaje ya compuesto por el llamador (`:57-60`).

### Mover en lote (`ComandoPasarAActual` / `ComandoPasarAHistorico`, `:227-330`)

1. Captura las filas seleccionadas **una sola vez** con `CapturarIDsSeleccionados`, que recorre `ItemsSelected`, se queda con la columna 0 y descarta en silencio todo valor no numérico o vacío (`:245-246`, `:79-86`).
2. Sin selección, compone «No hay ningún anexo histórico seleccionado para pasar a Actual.» (o su simétrico) y eleva 1000 (`:247-250`).
3. Delega todo el lote en una sola llamada al coordinador: `m_Coordinador.MoverLote(ids, False, Nothing, "", m_Error)` —`False` significa histórico→actual; el gemelo `True` significa actual→histórico (`:253, 302`; `AnexoSelectionTransactionCoordinator.cls:53`).
4. Si el lote falla y el mensaje contiene «Rollback», «compensación» o «compensacion», **bloquea los cuatro comandos de mutación** (`:254-262`, `:303-311`).
5. En éxito, invalida las dos cachés, reconstruye las listas, limpia la selección y notifica a los indicadores con la dirección `h2c` o `c2h` (`:266-273`, `:314-321`; `modIndicadores.bas:11-13`).

El coordinador envuelve el lote en una transacción `DAO.Workspace` (`ws.BeginTrans`, `ws.CommitTrans`, `ws.Rollback` en `AnexoSelectionTransactionCoordinator.cls:265, 323, 336`) y, por cada anexo, copia el fichero a su carpeta destino, hace `UPDATE TbAnexosUsuariosHPS SET EsHistorico='Sí'|'No' WHERE IDAnexo=` y borra el origen (`:280, 294-296, 310`). Como DAO no deshace operaciones de fichero, deja instantáneas en `%TEMP%\HPS\AnexoSelectionSnapshots\` (`:42, 250`) y compensa a mano si algo falla (`:354-386`).

### Anexar ficheros (`ComandoAnexar_Click`, `:104-154`)

1. Abre el diálogo con selección múltiple y concatena las rutas con `;` (`:114`; `FG.bas:837, 846`). Si el usuario cancela, `m_URL = ""` y el procedimiento sale sin error (`:117-119`).
2. Divide la cadena y, por cada fichero, crea un `AnexoUsuarioHPS` con `URLLocal`, `idUsuario = m_ObjUsuarioActivo.ID` y `EsHistorico = "No"`, y lo registra (`:122-132`; `AnexoUsuarioHPS.cls:221`).
3. Añade la fila a la lista actual y, al terminar, reconstruye **las dos** listas (`:135, 140`).
4. En error compone el mensaje, avisa al administrador y muestra `MsgBox` crítico (`:148-153`).

### Eliminar (`ComandoEliminarAnexoActual` / `ComandoEliminarAnexoHistorico`, `:156-225`)

1. Exige anexo seleccionado y que su `EsHistorico` corresponda al lado de la lista (`No` en `:160-167`, `Sí` en `:195-201`); el mensaje del lado histórico es «Seleccione una nexo de la tabla» (`:196`).
2. Pide confirmación con «¿Desea eliminar definitivamente del Sistema el anexo seleccionado?», con el botón No por defecto (`:168-171`, `:203-206`).
3. `Borrar` elimina el fichero y el registro (`:174`; `AnexoUsuarioHPS.cls:388-393`) y después reconstruye las listas (`:177`).

### Selección y doble clic (`:378-590`)

1. Al hacer clic se lee la columna 0; si está vacía, se descarta la selección y el handler sale (`:394-397`, `:500-503`).
2. Si el anexo cacheado no encaja con la lista de la que saltó el evento —por `IDAnexo` distinto o por `EsHistorico` que no corresponde—, se recarga con `getAnexoUsuarioHPS` (`:400-417`, `:506-523`): la reconciliación es explícita para que un clic posterior a un movimiento no opere sobre una instancia obsoleta.
3. Con `m_EstadoConexion <> "Solo lectura"` se habilitan el borrado y el movimiento hacia el lado contrario (`:421-424`, `:527-530`).
4. El doble clic **no confía en la selección cacheada**: relee el `IDAnexo` de la lista que disparó el evento, recarga, comprueba `EsHistorico` y, si el fichero no existe, eleva 1000 con la URL (`:449-474`, `:552-577`).

## Reglas, validaciones y cálculos

- **Identidad por global, no por argumento.** El usuario y sus anexos vienen de `m_ObjUsuarioActivo` (`VariablesEntorno.bas:225`); el control asume que el padre ya validó que existe (`:351-354`).
- **Separación por `EsHistorico`.** «Actual» e «Histórico» son la misma tabla filtrada por `EsHistorico='No'|'Sí'`, no dos entidades (`Constructor.bas:355-366`). La URL se compone con la carpeta que corresponde al lado: `Usuario.URLCarpetaAnexosActuales` o `Usuario.URLCarpetaAnexosHistoricos`, más el `NombreAnexo`, y se memoiza (`AnexoUsuarioHPS.cls:91-97`).
- **El movimiento de filas es atómico y en lote.** N filas se mueven con una sola transacción y una sola compensación de ficheros; el diseño anterior, que recorría `ItemsSelected` dos veces y movía fila a fila, queda descrito en los comentarios de los procedimientos (`:229-238`, `.form.txt:1395-1401` para la versión antigua).
- **Solo lectura se decide en dos sitios distintos.** `Form_Open` lee `Application.TempVars("SoloLectura")` (`:342`); los handlers de clic leen el global `m_EstadoConexion` (`VariablesEntorno.bas:259`; `:421, 527`).
- **No abrir si el fichero no existe.** `fso.FileExists` se comprueba antes de invocar `Ejecutar` (`:470, 573`), para no propagar un error opaco de `ShellExecuteA`.
- **`Err.Raise 1000` sin argumento.** Los mensajes viajan por el `m_Error` de módulo, y los manejadores solo los sobrescriben si `Err.Number <> 1000`.
- **Selección múltiple simple.** `MultiSelect =2` (`.form.txt:825, 891`): las filas se numeran por `ItemSelected`, no por índice de fila.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de anexos del usuario | `SELECT * FROM TbAnexosUsuariosHPS WHERE IDUsuario=… AND EsHistorico='No'|'Sí' ORDER BY NombreAnexo` | `Dictionary` de `AnexoUsuarioHPS` | `:20, 39`; `Constructor.bas:355-366`; `UsuarioHPS.cls:384, 413` |
| Anexo suelto por `IDAnexo` | `SELECT * FROM TbAnexosUsuariosHPS WHERE IDAnexo=…;` | `AnexoUsuarioHPS` | `:400, 455`; `Constructor.bas:292-294` |
| Alta de anexo | `AddNew` sobre `TbAnexosUsuariosHPS` + copia del fichero | escritura | `:130`; `AnexoUsuarioHPS.cls:259` |
| Movimiento en lote | `UPDATE TbAnexosUsuariosHPS SET EsHistorico='Sí'|'No' WHERE IDAnexo=` + `CopyFile` / `DeleteFile` | escritura con compensación por instantáneas | `:253, 302`; `AnexoSelectionTransactionCoordinator.cls:294-296, 310, 354-386` |
| Borrado de anexo | `DELETE * FROM TbAnexosUsuariosHPS …` + `DeleteFile` | escritura | `:174, 209`; `AnexoUsuarioHPS.cls:388-393` |
| Notificación de movimiento | evento `RaiseLifecycleMoveCompleted` | señal en memoria | `:273, 321`; `modIndicadores.bas:13` |
| Apertura con el visor del sistema | `ShellExecuteA(hwnd, "open", url, "", "", 1)` | API declarada en `VariablesEntorno.bas` | `:474, 577`; `VariablesEntorno.bas:5` |

Sin transacciones propias: la del lote la abre y cierra el coordinador.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo concreto: `Application.Echo`, `DoCmd.Hourglass`, `DoCmd.Close` y `Application.TempVars` son APIs de Access sin equivalente directo; el apagado del reloj en los cuatro manejadores de error (`:431, 485, 537, 588`) presupone un encendido externo (a verificar). `fso` (`Scripting.FileSystemObject`, `VariablesEntorno.bas:218`) y la compensación de ficheros bajo `%TEMP%` (`AnexoSelectionTransactionCoordinator.cls:42`) no tienen traducción directa: en web el anexo debe viajar como recurso del servidor, con el hash que ya existe en la tabla (`AnexoSelectionTransactionCoordinator.cls:28`). Las dos columnas del listbox con la primera a ancho cero (`.form.txt:837, 904`) equivalen a un `id` no visible que la API debe devolver igualmente.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **El `.cls` y el módulo embebido en el `.form.txt` son versiones distintas.** El código embebido (`.form.txt:1195` en adelante) no contiene `CapturarIDsSeleccionados`, `LimpiarSeleccionTrasCommit` ni `AnexoSelectionTransactionCoordinator`, y su `ComandoPasarAActual_Click` (`.form.txt:1385-1401`) mueve fila a fila sobre `ItemsSelected`. Cuál de las dos versiones ejecuta el binario no se puede decidir desde el snapshot.
- **El doble clic del lado histórico no limpia `m_Error` al entrar** (`:541-552`), mientras que los otros tres handlers sí lo hacen (`:113, 392, 446, 498`). Como `getAnexoUsuarioHPS` no vacía su `p_Error` de entrada (`Constructor.bas:277-290`), un `m_Error` que haya quedado relleno tras un error previo provoca un 1000 espurio en el primer doble clic sobre un anexo histórico. Impacto: falso error, sin efecto sobre los datos.
- **`ListaAnexosHistoricos_Click` desactiva el botón equivocado cuando el clic es vacío**: desactiva `ComandoPasarAHistorico` (`:502`), que actúa sobre la lista de actuales, mientras que en la misma rama vacía de la lista de actuales sí desactiva el correcto (`:396`). Quedan «A Actual» y «Eliminar» (histórico) habilitados sin selección (`:527-530`), y ambos fallan después con «No hay ningún anexo histórico seleccionado…» (`:248`) o «Seleccione una nexo de la tabla» (`:196`). Defecto de estado, no de datos.
- **Dos fuentes para la regla de solo lectura** (`Application.TempVars` en `:342`, `m_EstadoConexion` en `:421, 527`): si divergen, los handlers pueden habilitar comandos en una sesión de solo lectura. A verificar contra el binario.
- **`EstablecerListas` declara retorno `String` y nunca asigna valor** (`:8`): el error viaja solo por `p_Error`. Igual en `CapturarIDsSeleccionados` con `Collection` (`:69`), que sí usa `Set`.
- **`LimpiarSeleccionTrasCommit` no toca `ComandoAnexar`** (`:96-100`), mientras que el bloqueo por solo lectura sí lo incluye (`:345`). Inofensivo —un movimiento no añade ficheros—, pero asimétrico.
- **`CapturarIDsSeleccionados` descarta identificadores no numéricos sin avisar** (`:82-86`). Hoy solo se añaden filas con `IDAnexo` numérico (`:28`), así que el riesgo es teórico.
- **Restos:** `pregunta` es un global (`VariablesEntorno.bas:220`) reutilizado para confirmaciones y errores, y nunca se vuelve a leer. No hay `Form_Unload`: en un subformulario VBA reinicia el estado del módulo al cerrar, pero conviene confirmarlo contra el binario.
- **Las llamadas a `getAnexoUsuarioHPS` omiten el calificador `Constructor.`** (`:400, 411, 455, 506, 517, 558`), mientras que la copia antigua embebida en el `.form.txt` sí lo usa (`.form.txt:1609`). Al ser `Constructor.bas` un módulo estándar, la llamada resuelve globalmente; el cambio es ruido de estilo, no un fallo (a verificar contra el binario).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tabla o grid en lugar de dos listbox, con columna de estado en lugar del reparto en dos listas; arrastre y soltado para mover anexos; endpoint de descarga o visor en línea en lugar de delegar en la aplicación del sistema; estado explícito de «no alcanzable» en la fila en lugar de `MsgBox` con la URL; confirmación de borrado con el nombre del fichero y su tamaño.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir con un usuario activo, ambas listas muestran sus anexos por nombre, separadas por `EsHistorico`, ordenadas por `NombreAnexo`.
2. Si no hay usuario activo al abrir, se muestra «No se ha seleccionado ningún usuario» y el formulario se cierra sin guardar.
3. Con un único anexo en un lado, queda preseleccionado y listo para abrirse sin un clic previo.
4. Seleccionar una fila habilita únicamente el borrado y el movimiento que corresponden a ese lado; una fila vacía no deja comandos de mutación activos sobre una lista.
5. Mover varias filas de una vez es todo o nada: si algo falla, todas las filas y todos los ficheros permanecen en su lado de origen.
6. Con selección múltiple, mover N filas notifica una sola vez a los indicadores de la dirección correspondiente.
7. Doble clic sobre un anexo alcanzable lo abre con el visor asociado a su extensión; sobre uno no alcanzable, muestra el mensaje con su ruta.
8. En sesión de solo lectura, los cinco comandos permanecen deshabilitados.

## Pendientes

- Cuál de las dos versiones del código —`.cls` o la embebida en `.form.txt`— es la que ejecuta el binario, y si el `[anexos-batch-hotfix]` está desplegado.
- Si `UsuarioHPS.URLCarpetaAnexosActuales` y `URLCarpetaAnexosHistoricos` son rutas estables o dependen del entorno (`Entorno.cls`); de ello depende dónde acaba el fichero tras un movimiento.
- Qué mantiene coherente el campo `Hash` de `TbAnexosUsuariosHPS` (`AnexoSelectionTransactionCoordinator.cls:28`) y si el alta de `ComandoAnexar` lo calcula (`AnexoUsuarioHPS.cls:221-259`): en la lectura del snapshot no aparece.
- Quién enciende `DoCmd.Hourglass` antes de los cuatro apagados de los manejadores de error.
- Si el «A Actual» de la lista histórica debería recargarse o desatarse al pulsar sobre una fila vacía, y si el estado de `ComandoEliminarAnexoHistorico` debería seguir a la selección.