# Form_FormInicial07UsuariosSICA

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial07UsuariosSICA.cls` (236 líneas) y `.form.txt` (1708 líneas) · **PRD:** [README](../README.md)

Formulario central de la pestaña de usuarios SICA: filtra por nombre, exporta el listado filtrado a Excel, y da de alta, consulta y **borra en cascada** el usuario SICA seleccionado más sus anexos. Es un formulario **enlazado**: su subformulario `SubFormResultados` se alimenta del `RecordSource` SQL que compone `EstablecerFiltroUsuarios` sobre la caché local `TbUsuariosSICALocal`. A diferencia de los formularios de consulta de usuarios HPS, aquí la función de filtrado es **local** (`Public Function` en el propio módulo, `:104`), no la de `Filtro.bas`. Las citas son `Form_FormInicial07UsuariosSICA.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial01Lateral` (`lblUsuariosSICA`) | Cambia el `SourceObject` de `SubFormCentral` a `FormInicial07UsuariosSICA` | `Form_FormInicial01Lateral.cls:444-446` |
| `Form_FormInicial07UsuariosSICADatos` | El subform de detalle declara ese `SourceObject`; es el que fija `m_ObjUsuarioSICAActivo` | `.form.txt:897` |

Abre tres formularios: `FormUsuario07SICAEdicion` (`:21-23`), `FormUsuario07SICAAlta` (`:169-170`) y `FormUsuario02DatosObservacionHistorico` no aplica aquí. Cierra `FormUsuario07SICAEdicion` y `FormUsuario07SICAAlta` antes de reabrirlos si ya estaban abiertos (`:20-22, 167-170`). `ComandoListaEnExcel` invoca el visor del sistema sobre el `.xlsx` generado (`:149`; `VariablesEntorno.bas:5, 42`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Estilo de botones, primera carga del listado y habilitación según `SoloLectura` | `:183-217`, `.form.txt:375` |
| `SubFormResultados` | subformulario, `SourceObject = FormInicial07UsuariosSICADatos` | — | Recibe el `RecordSource` SQL; es el soporte de los tres botones de fila | `.form.txt:889-897`; `:118, 129` |
| `Nombre` | cuadro de texto | `OnExit` → `Nombre_Exit` | Filtra por nombre al salir del campo | `:221-236`, `.form.txt:1217, 1229` |
| `ComandoLimpiarNombre` | botón (`Caption = "Comando36"`) | `OnClick` → `ComandoLimpiarNombre_Click` | Vacía `Nombre` y refiltra | `:78-82`, `.form.txt:1254-1256` |
| `ComandoNuevoUsuarioSICA` | botón («Nuevo») | `OnClick` → `ComandoNuevoUsuarioSICA_Click` | Abre el alta en modo limpio | `:163-181`, `.form.txt:965-967` |
| `ComandoEditarUsuarioSICA` | botón («Ver/Editar») | `OnClick` → `ComandoEditarUsuarioSICA_Click` | Abre la edición del usuario activo | `:10-33`, `.form.txt:1010-1012` |
| `ComandoEliminarUsuarioSICA` | botón («Eliminar») | `OnClick` → `ComandoEliminarUsuarioSICA_Click` | Confirma, borra en cascada y refresca la caché local | `:35-76`, `.form.txt:1056-1058` |
| `ComandoListaEnExcel` | botón (`Caption = "Comando123"`) | `OnClick` → `ComandoListaEnExcel_Click` | Exporta el subform a Excel y lo abre | `:138-161`, `.form.txt:922-924` |
| `Comando5`, `Comando6` | 2 botones («Anexos», «Eliminar») | **sin** `OnClick` en el `.form.txt` | Sin handler en el módulo: código muerto | `.form.txt:1103-1104, 1149-1150` |
| `CursoEnVigor` | cuadro de texto con `ControlSource` | — | Rótulo de datos dentro del subform de detalle; sin lógica propia | `.form.txt:1435-1436` |
| `lblNombre`, `Etiqueta21` | etiquetas («Nombre», «CursoEnVigor») | — | Rótulos | `.form.txt:1201-1202, 1453-1454` |
| `Detalle` | sección | — | `Height = 7711`; sin lógica | `.form.txt:880` |

## Flujos

### Apertura (`Form_Load`, `:183-217`)

1. `On Error GoTo errores` (`:184`).
2. `CopiarEstiloBoton` sobre `ComandoNuevoUsuarioSICA` y `ComandoEditarUsuarioSICA`, encadenando `Err.Raise 1000` en cada fallo (`:188-195`).
3. `EstablecerFiltroUsuarios m_Error` hace la **primera carga** del listado con el filtro vacío (`:197-200`).
4. Si `Application.TempVars("SoloLectura") = "Sí"`, deshabilita Nuevo y Eliminar; en caso contrario los habilita (`:201-207`). `ComandoEditarUsuarioSICA` queda siempre habilitado (`:208`).
5. En error: mensaje + `CorreoAlAdministrador` + `DoCmd.Hourglass False` + `MsgBox` (`:210-216`). No cierra el formulario.

### Filtrado (`Filtrar` + `EstablecerFiltroUsuarios`, `:84-136`)

1. `Filtrar` (`:84-102`): **crea siempre** un `Scripting.Dictionary` nuevo con `CompareMode = TextCompare` (`:87-88`) — no reutiliza el global como los formularios de consulta HPS, y por tanto descarta cualquier otro filtro activo. Escribe la clave `"Nombre"` con `Nz` (`:89`).
2. Llama a su propia `EstablecerFiltroUsuarios m_Error` y propaga con `Err.Raise 1000` (`:91-94`).
3. `EstablecerFiltroUsuarios` (`:104-136`) toma `Set frm = Me.Controls("SubFormResultados").Form` (`:118`) — alcance **local**, no el del formulario principal, a diferencia de `Filtro.bas:24-26`.
4. Sin nombre, `m_Where = ";"` (`:121`); con nombre, un `LIKE '*<valor>*'` sobre la concatenación de nombre y dos apellidos, **con el valor interpolado sin parametrizar** (`:123`).
5. Compone `SELECT TbUsuariosSICALocal.* FROM TbUsuariosSICALocal <where>` y asigna `frm.RecordSource` (`:125-129`).
6. El valor de retorno de la función nunca se usa: los dos únicos llamantes pasan solo `p_Error` (`:90, 197`).

### Abrir la edición (`:10-33`)

1. Exige `m_ObjUsuarioSICAActivo` no nulo; si no, «Se ha de seleccionar el usuario SICA primero» y 1000 (`:16-20`).
2. Si `FormUsuario07SICAEdicion` está abierto, lo cierra sin guardar (`:21-22`) y lo abre (`:23`).
3. `DoCmd.OpenForm "FormUsuario07SICAEdicion"` (`:23`), sin `OpenArgs`: el estado viaja por el global `m_ObjUsuarioSICAActivo` (`VariablesEntorno.bas:226`).

### Borrar en cascada (`ComandoEliminarUsuarioSICA_Click`, `:35-76`)

1. Exige usuario activo y guarda `m_ID = m_ObjUsuarioSICAActivo.ID` **antes** de borrar (`:42-46`).
2. `MsgBox` de confirmación con `vbYesNo + vbDefaultButton2` (botón por defecto «No»); si no es `vbYes`, `Exit Sub` sin borrado (`:47-52`).
3. `VBA.DoEvents` / `DoCmd.Hourglass True` alrededor de la llamada (`:53-55`).
4. `m_ObjUsuarioSICAActivo.Borrar m_Error` (`UsuarioSICA.cls:857`) borra el usuario y sus anexos (`:56`; `UsuarioSICA.cls:876`). Si falla, **eleva 100** en vez de 1000 (`:58`).
5. `ActualizarUsuarioSICA m_ID, EnumSiNo.Sí, m_Error` (`FG.bas:3389`) refresca la caché local; si falla, eleva 1000 (`:60-63`).
6. Apaga el reloj de arena y sale (`:64-66`).
7. En error: mensaje + correo + apagado de reloj + `MsgBox` (`:69-75`).

### Exportar a Excel (`:138-161`)

`VBA.DoEvents` + `Hourglass True`, `ConsultaAExcel(Me.SubFormResultados.Form, , , EnumSiNo.Sí, m_Error)` (`FG.bas:382`), `Ejecutar Me.hWnd, "open", m_URLExcelGenerado, "", "", 1` y apagado del reloj (`:141-152`).

### Errores

Todos los handlers repiten el patrón `If Err.Number <> 1000` → mensaje → `CorreoAlAdministrador` → `DoCmd.Hourglass False` → `MsgBox` guardando el retorno en `pregunta` (`:28-32, 69-75, 96-101, 154-160, 173-180, 210-216, 230-235`). Ninguno cierra el formulario: el usuario pierde la sesión de edición pero conserva la pantalla de consulta.

## Reglas, validaciones y cálculos

- **El filtro es exclusivo y local.** `Filtrar` crea el diccionario desde cero (`:87-88`), así que el criterio de nombre se evalúa siempre contra un estado limpio; no convive con los filtros HPS de `Filtro.bas` porque aquel lee un global distinto. El global `m_ObjColParaFiltros` se escribe pero aquí no se lee (`VariablesEntorno.bas:256`).
- **Concatenación de nombre y apellidos.** El `LIKE` cubre la concatenación `[Nombre] & ' ' & [Apellido_1] & ' ' & [Apellido_2]` (`:123`), de modo que un texto que cruza el límite de dos apellidos no encuentra nada.
- **SQL por concatenación de valor de usuario.** El texto de `Nombre` se interpola sin comillas de escape ni parametrización (`:123`). No hay sanitización en ninguna parte del módulo.
- **`m_Where = ";"` como «sin filtro».** La rama vacía deja un punto y coma suelto en el SQL final (`:121, 125-127`): Access lo tolera.
- **Modo solo lectura por `TempVars`.** Es una variable de aplicación de Access, no un permiso del backend: se fija fuera de este formulario y solo controla la visibilidad de los botones (`:201-207`). `ComandoEditarUsuarioSICA` no se deshabilita nunca (`:208`).
- **Borrado irreversible tras confirmación.** `UsuarioSICA.Borrar` borra usuario y anexos (`UsuarioSICA.cls:857, 876`) sin papelera; la única red es el `MsgBox` con «No» por defecto (`:48`).
- **`vba.DoEvents` como transacción visible.** Cinco llamadas ceden el hilo para que la interfaz no se congele (`:53-55, 64-66, 141-143, 150-152`): Access es monohilo, la espera es cooperada.
- **Devolución de `EstablecerFiltroUsuarios` ignorada.** La función se declara `As String` y devuelve `p_Error` por `ByRef`; el valor de retorno no lo consulta nadie (`:104-106, 90, 197`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado de usuarios SICA | `TbUsuariosSICALocal` (`SELECT … WHERE nombre+apellidos Like`) | `RecordSource` del subform de resultados | `:118, 123, 125-129` |
| Caché local de SICA | `TbUsuariosSICALocal` | lectura indirecta del subform de detalle | `.form.txt:897` |
| Refresco de caché tras borrar | `TbUsuariosSICALocal` (reescritura del registro) | escritura vía `ActualizarUsuarioSICA` | `:60`; `FG.bas:3389` |
| Borrado del usuario y sus anexos | tablas de usuario SICA y anexos | escritura vía `UsuarioSICA.Borrar` | `:56`; `UsuarioSICA.cls:857-876` |
| Exportación a Excel | `ConsultaAExcel` sobre el `Form` del subform | lectura → fichero `.xlsx` en disco | `:144`; `FG.bas:382` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo concreto: `Me.SubFormResultados.Form` es un patrón de Access sin equivalente — en web el listado es un endpoint y el «formulario» hijo desaparece. `Application.TempVars("SoloLectura")` es estado de sesión de la aplicación de escritorio, no un permiso: en web la habilitación de botones debe derivar de la autorización del usuario en cada petición. `VBA.DoEvents`, `DoCmd.Hourglass` y `MsgBox` son tres APIs sin correspondencia directa; `ShellExecuteA` abre el `.xlsx` con la aplicación registrada en el sistema del cliente (`VariablesEntorno.bas:5, 42`), que en web es una descarga. La caché local `TbUsuariosSICALocal` es una réplica en el `.accdb` del cliente; el borrado obliga a refrescarla explícitamente (`:60`), un acoplamiento que en web desaparece si el listado se lee del servidor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **`Err.Raise 100` en `ComandoEliminarUsuarioSICA_Click`** (`:58`): el resto del módulo usa 1000 como error interno. Con 100, el manejador `If Err.Number <> 1000` (`:70`) **sí** entra, así que un fallo de `Borrar` compone mensaje y envía correo, pero con la etiqueta equivocada; más grave, `m_Error` ya viene relleno y el mensaje se prepende sin limpiarlo, de modo que el `MsgBox` puede mostrar dos textos concatenados (a verificar con un fallo real de base).
- **`ActualizarUsuarioSICA` solo se ejecuta si `Borrar` no falló**: tras `Err.Raise 100` (`:58`) el flujo salta al manejador, así que la secuencia es correcta. Pero si `Borrar` devolviera `m_Error` vacío sin haber borrado nada (a verificar en `UsuarioSICA.cls:857`), la caché se refrescaría para un usuario inexistente. El mensaje `:71` además carece del espacio entre el nombre del handler y «se».
- **`Comando5` y `Comando6` no tienen handler.** Existen en el diseño con los rótulos «Anexos» y «Eliminar» (`.form.txt:1103-1104, 1149-1150`) pero el módulo no define ningún `Comando5_Click` ni `Comando6_Click`: son inertes y sugieren acciones que no existen (a verificar en el binario; el snapshot `.cls` es la lectura fijada). Junto con `Comando123` y `Comando36` (`.form.txt:923, 1255`), los cuatro botones carecen de texto legible: rompen accesibilidad y traducción.
- **`EstablecerFiltroUsuarios` devuelve un valor que nadie lee** (`:104-106`). Firma confusa que invita a pensar que hay dos canales de error.
- **SQL sin parametrizar.** `m_Nombre` se interpola en el `LIKE` (`:123`). Con un apóstrofo en el nombre el SQL se rompe; Access puede mostrar un `MsgBox` en vez de un error claro.
- **`DoCmd.Hourglass False` en el manejador de error sin un `True` correspondiente en todos los caminos** (`:74`): inocuo, pero sugiere que quien lo enciende es otro formulario.
- **`Me.SubFormResultados.Form` en `EstablecerFiltroUsuarios`** (`:118`) requiere que el subformulario esté ya cargado. `Form_Load` lo invoca (`:197`) y funciona, pero el orden depende de que Access haya instanciado el hijo; en otro motor de renderizado esa llamada sería prematura (a verificar).
- **`m_ObjUsuarioSICAActivo` sobrevive al cierre**: se anula solo en `ComandoNuevoUsuarioSICA_Click` (`:166`), de modo que cerrar el formulario sin dar de alta deja el global apuntando al usuario anterior (a verificar frente a `VariablesEntorno.bas:484, 541`). `pregunta` es un `Long` global compartido (`VariablesEntorno.bas:220`) usado a la vez como retorno del `MsgBox` de confirmación (`:47`) y del de error.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: borrado en dos pasos con estado «pendiente de confirmar» en lugar de un `MsgBox` modal; pager y ordenación en servidor en lugar de `RecordSource` compuesto; exportación asíncrona con enlace temporal en lugar de `ShellExecuteA`; y permiso de escritura derivado de la sesión en lugar de `Application.TempVars("SoloLectura")`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, el listado muestra todos los registros de `TbUsuariosSICALocal`.
2. Escribir un nombre y salir del campo restringe el listado a los usuarios cuyo nombre o apellidos contienen ese texto.
3. «Limpiar» vacía el filtro y devuelve el listado completo.
4. Un nombre con apóstrofos no rompe el filtrado: el sistema comunica el error en lugar de devolver una lista arbitraria.
5. «Ver/Editar» sin usuario seleccionado muestra «Se ha de seleccionar el usuario SICA primero» y no abre nada.
6. «Eliminar» con usuario seleccionado pide confirmación con «No» por defecto; cancelar no borra nada; confirmar borra el usuario y sus anexos y refresca la caché local.
7. En modo solo lectura, «Nuevo» y «Eliminar» aparecen deshabilitados y «Ver/Editar» sigue disponible.
8. «Lista en Excel» exporta exactamente las filas que se están viendo y abre el fichero generado.

## Pendientes

- Origen del estado «seleccionar el usuario SICA»: `m_ObjUsuarioSICAActivo` (`VariablesEntorno.bas:226`) lo fija `FormInicial07UsuariosSICADatos`; falta la ficha que documente ese subformulario.
- Contenido y vida de la caché `TbUsuariosSICALocal`: se escribe en `Constructor.bas:1825` y se audita en `CacheConsistencyAudit.bas:141-143`; queda por documentar su sincronización.
- Si `Comando5` / `Comando6` tienen handler en el binario real y no en el snapshot (a verificar con `vba-binary-drift`).
- Firma real de `ConsultaAExcel` (`FG.bas:382`): los tres argumentos intermedios se pasan vacíos y no se documenta aquí qué hacen.