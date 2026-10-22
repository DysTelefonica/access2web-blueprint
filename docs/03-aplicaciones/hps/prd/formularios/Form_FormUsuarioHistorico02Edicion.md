# Form_FormUsuarioHistorico02Edicion

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuarioHistorico02Edicion.cls` (235 líneas) y `.form.txt` (2441 líneas) · **PRD:** [README](../README.md)

Subformulario central por defecto de la ficha de histórico: muestra los datos del usuario histórico activo como **solo lectura** y lista sus observaciones, en dos listas separadas (generales y de curso). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): los datos llegan del global `m_ObjUsuarioHistoricoActivo` y se vuelcan campo a campo con `DatosDelObjetoAlFormulario`. Hermano de `Form_FormUsuarioHistorico02Anexos`, que comparte contenedor y patrón de apertura por doble clic. Las citas son `Form_FormUsuarioHistorico02Edicion.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuarioHistorico00Principal` | `SubFormCentral` declara `SourceObject = "Form.FormUsuarioHistorico02Edicion"` por defecto | `Form_FormUsuarioHistorico00Principal.form.txt:891` |
| `Form_FormUsuarioHistorico01Lateral` | Vuelve a este formulario si el central está mostrando los anexos | `Form_FormUsuarioHistorico01Lateral.cls:62-63` |

Abre `FormUsuario02DatosObservacionHistorico` con `OpenArgs` = identificador de observación (`:200, 227`), cerrándolo antes sin guardar si ya estaba abierto (`:197-199, 225-227`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Valida el contexto, vuelca datos y **bloquea la escritura** control a control | `:6-57`, `.form.txt:375` |
| `ListaObservaciones` | listbox, 3 columnas (`Value List`) | `OnDblClick` → `ListaObservaciones_DblClick` | Abre la observación general seleccionada | `:184-209`, `.form.txt:1476, 1481-1496` |
| `ListaObservacionesCurso` | listbox, 3 columnas (`Value List`) | `OnDblClick` → `ListaObservacionesCurso_DblClick` | Abre la observación de curso seleccionada | `:211-235`, `.form.txt:1514, 1519-1532` |
| `ID`, `Nombre`, `Apellido_1`, `Apellido_2`, `DNI`, `Correo_e`, `F_Nacimiento`, `F_Curso`, `F_Baja`, `Telefono`, `Motivo_HPS`, `EmpresaUsuario`, `EmpresaTramitadora`, `FechaPrimeraConvocatoria`, `FechaSegundaConvocatoria`, `FechaCorreoNoCurso`, `JuridicaContrato`, `Titulo`, `LugarNacimiento`, `SuministroContrato` | ~20 cuadros de texto | — | Campos de solo lectura querellenados con el volcado; el bucle de `Form_Load` los marca `Locked = True` | `.form.txt:1007, 1037, 1429, 1459, 1097, 1157, 1247, 1308, 1369, 1653, 1743, 1684, 1713, 1844, 1906, 1968, 2061, 2121, 2182` |
| `Etiqueta116` … `Etiqueta68` | ~20 etiquetas («ID», «Empresa_Usuario», «DNI», «Correo_e», «F_Nacimiento», «F_Curso», «Telefono», «Motivo_HPS», «F_Baja», «F_Convocat.1/2», «F_Correo JS», «En Vigor», «Jurídica Contratación», «Expediente», «L_Nacim.», «OBSERVACIONES», «OBS CURSO») | — | Rótulos; sin lógica | `.form.txt:916-917` y el bloque de rótulos `.form.txt:1538-2186` |
| `EnVigor` | combo (`Value List`: `Sí;No`) | — | Estado del curso | `.form.txt:1761, 1782-1783` |
| `Detalle` | sección | — | `Height = 7711`; sin lógica | `.form.txt:893` |

## Flujos

### Apertura (`Form_Load`, `:6-57`)

1. `On Error GoTo errores` (`:9`).
2. `Me.AllowEdits = True` deliberadamente (`:16`), para que Access permita el clic en las listas; la protección se aplica después, campo a campo (comentario en `:14-15`).
3. Si `m_ObjUsuarioHistoricoActivo Is Nothing`, «No se ha abierto el formulario con los datos necesarios» y 1000 (`:18-21`).
4. `EstablecerControlesDatos Me, m_Error` aplica la paleta; si falla, 1000 (`:23-24`).
5. `EstablecerDatos m_Error` vuelca los campos y carga las listas; si falla, 1000 (`:26-27`).
6. **Bucle de bloqueo** (`For Each ctl In Me.Controls`, `:31-42`): marca `Locked = True` todo `TextBox`, `ComboBox` y `CheckBox`, y deja `Locked = False` + `Enabled = True` en todo `ListBox`.
7. En error: mensaje + `CorreoAlAdministrador` + `MsgBox` + **cierre sin guardar** (`:49-56`).

### Volcado de datos (`EstablecerDatos`, `:59-115`)

1. Vacía los `RowSource` de las dos listas (`:70-71`).
2. Si no hay usuario activo, sale sin error (`:73-75`) — guarda ya validado en `Form_Load`.
3. `DatosDelObjetoAlFormulario m_ObjUsuarioHistoricoActivo, Me, p_Error` vuelca las propiedades del objeto a los controles por nombre (`FG.bas`); si falla, 1000 (`:77-78`).
4. Segundo bucle, heredado y ya labeled redundante en el propio código (comentario `:80-82`): deshabilita (`Enabled = False`) los `TextBox`/`ComboBox` cuyo `Tag` contenga `DATOS;INTOCABLE` (`:83-91`).
5. Rellena los tres campos calculados con valor por defecto «sin dato»: `EmpresaUsuario` con `SuministradorContratoObj.NombreConCif`, `EmpresaTramitadora` con `SuministradorTramitadoraObj.NombreConCif` y `Titulo` con `CodExp & " " & Titulo` del expediente (`:93-102`), cada uno condicionado a que el objeto asociado exista.
6. `Me.Titulo.ForeColor = Me.Correo_e.ForeColor` (`:104`): el título de expediente adopta el color de un campo de correo, no el de su etiqueta.
7. `EstablecerListasObservaciones , p_Error` carga ambas listas; si falla, 1000 (`:107-108`).

### Cargar las listas (`EstablecerListasObservaciones`, `:116-182`)

1. `p_Error = ""`, `m_Error = ""` (`:126-127`).
2. `If p_Tipo = "Curso" Then GoTo Curso` (`:131-132`): el parámetro es una bandera de salto, y **el único llamante lo pasa vacío** (`:107`), así que la ejecución cae siempre por las dos ramas.
3. Rama «General» (`:134-155`): vacía `ListaObservaciones`, fuerza `Set m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricas = Nothing` para invalidar la caché (`:136`), lee `p_Error = m_ObjUsuarioHistoricoActivo.Error` y eleva 1000 si hay error (`:137-142`), y recorre las claves del diccionario añadiendo `IDObservacion & ";" & Fecha & ";" & Observacion` (`:143-152`).
4. Rama «Curso» (`:153-172`): idéntica sobre `ColObservacionesHistoricasCurso` y `ListaObservacionesCurso`.
5. Devuelve `"OK"` (`:174`); el valor no lo usa el llamante (`:107`).

### Abrir una observación (`ListaObservaciones_DblClick`, `:184-209`, y `:211-235`)

1. Lee `Nz(Me.ListaObservaciones.Column(0), "")` (`:190`) — columna del identificador, oculta con `ColumnWidths = "0;1443;10094"` (`.form.txt:1495`).
2. Si está vacía, `Exit Sub` sin error (`:191-195`).
3. Cierra `FormUsuario02DatosObservacionHistorico` sin guardar si está abierto y lo abre con `OpenArgs:=m_IDObservacionSeleccionada` (`:197-200`). El segundo handler es copia literal con `ListaObservacionesCurso` (`:216-227`).

### Errores

`Form_Load` compone el mensaje sin `vbNewLine` (`:51`) y cierra el formulario; los dos handlers de doble clic usan `CorreoAlAdministrador` + `MsgBox` sin cerrar (`:202-208, 230-235`). El retorno del `MsgBox` va al global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **Solo lectura por recorrido reflectante de controles, no por permiso.** `Form_Load` marca `Locked = True` en todo `TextBox`/`ComboBox`/`CheckBox` (`:33-35`) y deja las listas vivas (`:38-41`). No hay `DataAllowEdit`, no hay contraseña de protección de formulario: es una convención de presentación.
- **`AllowEdits = True` es deliberado.** Sin él Access no dispara `DblClick` en listas de un formulario no enlazado; el comentario en `:14-15` documenta el cambio como «CAMBIO CRÍTICO».
- **Dos listas, un solo `p_Tipo`.** `EstablecerListasObservaciones` acepta un discriminador `p_Tipo` que solo toma el valor `"Curso"` para saltar la primera rama (`:131-132`), y el llamante no lo usa (`:107`). La firma sugiere dos funciones cuando hay una.
- **Invalidación explícita de la caché.** Ambas ramas hacen `Set …Col… = Nothing` antes de leer (`:136, 155`), lo que fuerza a `UsuarioHistorico` a recargar por el `Property Set` (`UsuarioHistorico.cls:348-350`).
- **El texto de las listas se compone en VBA.** `ID;Fecha;Observacion` con `;` como separador (`:146, 166`) y anchos `0;…;…` para ocultar el identificador (`.form.txt:1495, 1531`): un texto de observación que contenga `;` se partirá en columnas de más (a verificar).
- **`RowSource` horneado en el diseño.** `ListaObservaciones` trae filas de ejemplo escritas en el `.form.txt` (`.form.txt:1493`); `EstablecerDatos` las vacía (`:70`) en cada apertura, pero son datos de personas reales incrustados en el binario del formulario.
- **Identidad por global, no por argumento.** Todo el contexto llega de `m_ObjUsuarioHistoricoActivo` (`VariablesEntorno.bas:227`); el mismo patrón que `Form_FormUsuarioHistorico02Anexos`.
- **Campos calculados con guarda de nulidad.** Los tres objetos auxiliares (`SuministradorContratoObj`, `SuministradorTramitadoraObj`, `Expediente`) se comprueban antes de dereferenciarlos (`:93, 96, 100`), de modo que un histórico sin expediente deja `Titulo` vacío en vez de fallar.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Datos del usuario histórico | `Constructor.get…` vía `UsuarioHistorico` | diccionario en memoria, volcado por reflexión | `:77`; `UsuarioHistorico.cls:321-350` |
| Observaciones generales del histórico | `Constructor.getObservacionesHistoricas(id, "General")` | `Dictionary` de `ObservacionHistorica` | `:136-152`; `UsuarioHistorico.cls:335` |
| Observaciones de curso | `Constructor.getObservacionesHistoricas(id, …)` | `Dictionary` de `ObservacionHistorica` | `:155-171`; `ObservacionHistorica.cls:13` |
| Apertura de la observación | — | `OpenArgs` a `FormUsuario02DatosObservacionHistorico` | `:200, 227` |
| Estilo de controles | `EstablecerControlesDatos Me, m_Error` | helper de `FG.bas` | `:23`; `FG.bas:1005-1053` |

Sin escrituras propias: el formulario es un visor. La escritura ocurre en el formulario que abre el doble clic.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo concreto: el volcado por `DatosDelObjetoAlFormulario` es reflexión de nombres de control a propiedades del dominio; en web desaparece y pasa a ser un contrato de serialización explícito, donde la lista de campos deja de depender de que el control y la propiedad se llamen igual. El `For Each ctl In Me.Controls` con `TypeOf` es la manera de Access de declarar «esto es solo lectura» y no tiene equivalente: en web es una ruta sin `POST` o un campo `disabled` por definición de permisos. `DoCmd.OpenForm … OpenArgs:=` es un paso de estado por cadena, sin equivalente directo; el patrón moderno es parámetro de ruta. `Lock` y `Enabled` son atributos de presentación, no de autorización: un cliente web puede ignorar ambos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Ejecución de las dos ramas por el `GoTo`.** El llamante pasa `p_Tipo` vacío (`:107`), así que el flujo entra en la rama «General», ejecuta su bucle y **cae** en la etiqueta `Curso` (`:154`) sin `GoTo`: las dos listas se rellenan en una sola llamada. Funciona por accidente; si alguien ocupa `GoTo Curso` explícitamente, la lista general no se rellena (a verificar el propósito original de la bandera).
- **`Set …Col… = Nothing` puede dejar la lista vacía.** La propiedad `Property Get` de `UsuarioHistorico` recarga de la base solo si la privada es `Nothing` (`UsuarioHistorico.cls:327-339`). Asignar `Nothing` al `Property Set` (`:136`) borra la caché; la lectura posterior la repone. Si la recarga falla, el `If Not … Is Nothing` (`:137`) es falso y la lista queda vacía sin error visible, porque el `p_Error` solo se lee dentro de la rama verdadera.
- **`p_Error` se lee pero se sobrescribe.** `p_Error = m_ObjUsuarioHistoricoActivo.Error` (`:138, 157`) sin limpiar antes: un error residual de una lectura anterior puede aparecer como error de carga de una operación exitosa.
- **Separador `;` en el texto de las observaciones** (`:146, 166`): un texto con punto y coma genera una columna extra y desalinea la fila. `ColumnCount = 3` (`.form.txt:1482, 1520`) descarta el sobrante, no falla.
- **`Me.Titulo.ForeColor = Me.Correo_e.ForeColor`** (`:104`) es una regla sin justificación documentada: el color de un campo de correo define el del título de expediente.
- **El bucle `INTOCABLE` es redundante** (`:83-91`), tal como admite el propio comentario (`:80-82`): `Form_Load` ya marca `Locked = True` en todos los `TextBox`/`ComboBox` (`:33-35`). Se conserva y conviene decidir si se retira.
- **Mensajes y comentarios escritos en segunda persona y con tono de cuaderno.** «Ponemos AllowEdits = True…», «Ahora funcionarán los clics…», «Si tienes botones…» (`:14-15, 38-41, 106`): no son documentación, son notas de desarrollo incrustadas en el código de producción.
- **`Form_Load` es `Public`** (`:6`) y sus helpers `Private` (`:59, 116, 184, 211`): la asimetría sugiere que se llamó desde VBA desde fuera del formulario en algún momento (a verificar con `vba-handler-backtrace`).
- **Sin `Form_Unload`.** `m_ObjObservacionSeleccionada` (`Form_FormUsuarioHistorico02Edicion.cls:5`) se declara pero no se usa en ningún punto del módulo: resto de una versión anterior.
- **Datos de personas reales horneados en el diseño** (`.form.txt:1493`): el `RowSource` de ejemplo se vacía en cada apertura (`:70`), pero el texto permanece en el binario distribuible.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: contrato explícito de campos en lugar del volcado por reflexión; tabla en lugar de dos `ListBox` con separador `;`; badges de solo lectura por rol en lugar del bucle de bloqueo; y ruta de navegación `/historico/:id/observacion/:idObs` en lugar de `OpenArgs`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un usuario histórico activo, el formulario muestra todos sus campos en modo solo lectura y las listas de observaciones quedan cargadas.
2. Sin usuario histórico activo, se muestra «No se ha abierto el formulario con los datos necesarios» y el formulario se cierra.
3. Ningún campo de datos es editable; las listas sí responden al clic.
4. Las observaciones generales y las de curso aparecen en listas separadas, con fecha y texto visibles y sin el identificador.
5. Doble clic sobre una observación abre su ficha con el identificador correcto; doble clic sobre el encabezado no hace nada.
6. Un histórico sin expediente, sin suministrador de contrato o sin suministrador tramitadora se muestra sin error y con esos tres campos calculados vacíos.
7. Un texto de observación con punto y coma no rompe la fila ni oculta el resto de columnas.

## Pendientes

- Contrato de `FormUsuario02DatosObservacionHistorico`: qué espera recibir por `OpenArgs` y si escribe (el formulario parece ser el de alta/edición de observación, no un visor).
- Por qué `Form_Load` es `Public` y qué otros módulos lo invocan (a verificar con `vba-handler-backtrace`).
- Si `Me.ComandoObservacionEditar` existió alguna vez: el comentario de `:38-41` lo menciona en futuro («si los tienes en este formulario»), pero no hay ningún control con ese nombre en el `.form.txt`.
- Regla de pertenencia de una observación a la categoría «Curso»: hoy se decide por la llamada a `Constructor.getObservacionesHistoricas` con el discriminante desde `UsuarioHistorico`, no desde este formulario.
- Procedencia de los datos de ejemplo horneados en `.form.txt:1493` y si deben retirarse del binario distribuible.