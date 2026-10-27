# Form_FormUsuario02Edicion

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario02Edicion.cls` (976 líneas) y `.form.txt` (3844 líneas) · **PRD:** [README](../README.md)

Subpantalla central de la edición de un usuario **actual**: carga el usuario activo, muestra sus datos, sus observaciones generales y sus observaciones de curso, y los graba. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): los valores viajan por reflexión con `RellenarObjetoConDatos` y `HaHabidoCambios` (`FG.bas:4124`, `FG.bas:4147`), no por enlace de controles. Las citas son `Form_FormUsuario02Edicion.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuario00Principal` | Se abre con el usuario ya establecido en el global `m_ObjUsuarioActivo`; el formulario no recibe argumentos | `:590-593` |
| `Form_FormUsuario02Alta` | Tras dar de alta, `ComandoGrabar_Click` cierra el principal sin guardar (`acSaveNo`), lo que descarta el formulario de alta | `:288-289` |
| `Form_FormExpedientesBusqueda` | Recibe el evento `m_FormExpedientes_Seleccionado(m_Exp As Expediente)` | `:941` |
| `Form_FormUsuario07SICAEdicion` | `ComandoDetalleUsuarioSICA_Click` graba los cambios pendientes y abre la edición SICA | `:200-232` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Detecta el modo lectura/escritura, carga el usuario, habilita botones y bloquea los campos | `:570-673` |
| (formulario) | — | declara `Public Event EdicionRegistro()` | Evento que el contenedor usa para recargar; se emite solo al grabar, no al observar | `:11`, `:277` |
| `DNI` | texto | `Exit` → `DNI_Exit` | Valida longitud, formato y duplicado contra actual e histórico | `:508-547` |
| `Correo_e` | texto | `Exit` → `Correo_e_Exit` | Valida formato de correo si no está vacío | `:548-569` |
| `PalabraClave` | texto | `KeyDown` → `PalabraClave_KeyDown` | Enter lanza la búsqueda de expediente | `:141-160` |
| `ListaObservaciones` | lista | `Click`, `DblClick` | Selecciona y abre una observación general | `:819-878` |
| `ListaObservacionesCurso` | lista | `Click`, `DblClick` | Selecciona y abre una observación de curso | `:879-940` |
| `ID`, `Nombre`, `LBLDNI`, `F_Nacimiento`, `F_Curso`, `Telefono`, `Apellido_1`, `Apellido_2`, `IDEmpresaUsuario` | texto y etiquetas | — | Datos del usuario; los carga y lee `RellenarObjetoConDatos` por reflexión | `.form.txt:698-1681`, `FG.bas:4124` |
| `ComandoGrabar` | botón | `Click` | Valida que haya cambios y registra el usuario | `:240-293` |
| `ComandoBuscarExpediente`, `ComandoIrABusquedaExpedientes` | botones | `Click` | Búsqueda de expediente por palabra clave | `:12-104`, `:105-140` |
| `ComandoObservacionAlta`, `ComandoObservacionCursoAlta` | botones | `Click` | Alta de observación | `:295-339` |
| `ComandoObservacionEditar`, `ComandoObservacionCursoEditar` | botones | `Click` | Edición de la observación seleccionada | `:340-404`, `:405-432` |
| `ComandoObservacionEliminar`, `ComandoObservacionCursoEliminar` | botones | `Click` | Baja de la observación seleccionada | `:367-404`, `:433-470` |
| `ComandoPasarAHistorico` | botón | `Click` | Pasa el usuario a histórico; se deja **invisible** al abrir | `:471-507`, `:608-609` |
| `ComandoDetalleUsuarioSICA` | botón | `Click` | Graba lo pendiente y abre la edición SICA | `:200-232` |
| `Detalle`, `Etiqueta*`, `lblExpedienteActividad` | marco y etiquetas | — | Estructura visual; sin lógica | `.form.txt:585-1074` |

## Flujos

### Apertura (`Form_Load`, `:570-673`)

1. Lee `Application.TempVars("SoloLectura")` y guarda el modo en `bSoloLectura` (`:580-584`).
2. **Activa la edición a nivel de formulario siempre** (`:587-589`): el propio comentario explica que si no, las listas se congelan. La protección se hace control a control.
3. Exige `m_ObjUsuarioActivo` y falla si falta (`:591-594`).
4. Carga una **copia de inicio** del usuario con `getUsuarioHPS` (`:597-598`; `Constructor.bas:743`), que sirve de referencia para detectar cambios.
5. Aplica la paleta de colores (`:599-601`).
6. Deja apagados los cuatro botones de edición/eliminación de observación y **oculta** `ComandoPasarAHistorico` (`:605-609`).
7. Activa o desactiva nueve controles según el modo (`:612-635`).
8. Copia el estilo de seis botones (`:637-642`).
9. `EstablecerDatos` rellena los campos y las dos listas (`:644-646`; `:675-762`, `:763-818`).
10. Si es solo lectura, recorre los controles y pone `Locked = True` a textos, combos y casillas, dejando las listas vivas (`:649-667`).
11. En error: correo al administrador, `MsgBox` crítico y cierre (`:666-673`).

### Grabar (`ComandoGrabar_Click`, `:240-293`)

1. Pone el reloj de arena (`:248-249`).
2. Exige `m_ObjUsuarioActivo` y `m_ObjUsuarioAlInicio` (`:251-257`).
3. `HaHabidoCambios` compara la copia de inicio con lo que hay en pantalla; si no hay cambios, aborta con «No hay cambios que grabar» (`:260-267`; `FG.bas:4147`).
4. `RellenarObjetoConDatos` vuelca los controles al objeto (`:268-271`; `FG.bas:4124`).
5. Marca `ParaAlta = No` y llama a `Registrar` pasando la copia de inicio (`:272-276`).
6. Emite `EdicionRegistro` (`:277`), apaga el reloj y avisa «Edición correcta» (`:279-283`).
7. **Cierra `FormUsuario00Principal` sin guardar** (`:284`).
8. En error: apaga el reloj, correo al administrador y `MsgBox` (`:286-293`).

### Validación del DNI (`DNI_Exit`, `:508-547`)

1. Solo valida si el valor **cambia** respecto a la copia de inicio (`:515-517`).
2. Exige longitud 9 (`:517-520`) y delega el formato en `FormatoDatoValido(…, "DNI")` (`:522-525`).
3. `DNIEnActivo` busca duplicado en usuarios actuales (`:526-530`) y `DNIEnHistorico` en históricos (`:532-536`).
4. En ambos casos **vacía el campo** (`Me.DNI = Null`) y relanza el error indicando el usuario en conflicto (`:527-529`, `:533-535`).

### Editar SICA (`ComandoDetalleUsuarioSICA_Click`, `:200-232`)

1. Sale si no hay usuario SICA activo (`:205-207`).
2. Si hay cambios pendientes, **dispara `ComandoGrabar_Click` sin esperar** (`:214-215`): al no capturarse `m_Error`, si el guardado falla la navegación continúa igualmente.
3. Cierra `FormUsuario07SICAEdicion` si estuviera abierta (`:221-223`).
4. Marca `ParaAlta = No` y abre `FormUsuario07SICAEdicion` (`:224-225`).

### Errores

Todos los manejadores repiten el patrón: `Err.Number <> 1000` ⇒ anexar descripción, `CorreoAlAdministrador` y `MsgBox` crítico. El retorno del `MsgBox` va al global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **Todo el estado de sesión es global.** El usuario en curso (`m_ObjUsuarioActivo`), la copia de inicio (`m_ObjUsuarioAlInicio`), el usuario SICA (`m_ObjUsuarioSICAActivo`) y el retorno de los `MsgBox` (`pregunta`) son variables públicas de módulo, no parámetros. El formulario no es reentrante.
- **El modo lectura/escritura se decide dos veces.** Una con `TempVars` para los botones (`:612-635`) y otra recorriendo controles para `Locked` (`:649-667`). Las dos listas se traverses en un segundo bucle solo para forzar `Locked = False`.
- **Comparación por reflexión.** `HaHabidoCambios` y `RellenarObjetoConDatos` recorren los controles y los contrastan contra las propiedades del objeto (`FG.bas:4147`, `FG.bas:4124`). Añadir un control al formulario cambia el comportamiento sin tocar este `.cls`.
- **El evento `EdicionRegistro` solo se emite al grabar** (`:277`). Dar de alta, editar o eliminar una observación no lo emite, aunque cambia datos visibles en esta misma pantalla.
- **Validación de correo solo de formato** (`:548-569`); la de DNI es la única que consulta duplicados, y contra dos tablas.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga del usuario | vía `Constructor.getUsuarioHPS` | lectura por `OpenRecordset` | `:597-598`, `Constructor.bas:743` |
| Registro de la edición | `Registrar` sobre `UsuarioHPS` | escritura con la copia de inicio como referencia | `:277-279` |
| Duplicado de DNI | usuarios actuales e históricos | consulta auxiliar | `:526-537` |
| Observaciones | `EstablecerListasObservaciones` sobre el objeto de usuario | colección en memoria | `:763-818` |
| Copia de inicio y volcado de pantalla | `HaHabidoCambios`, `RellenarObjetoConDatos` | reflexión sobre controles | `:259-272`, `FG.bas:4124`, `FG.bas:4147` |

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). El patrón dominante aquí no es el SQL sino el **estado global**: la pantalla no recibe identificadores, los lee de variables públicas de módulo, y emite un evento propio que el contenedor debe conocer. En web eso es un parámetro de ruta y un callback explícito. El reloj de arena (`:247-249`, `:279`, `:291`) y el cierre con `acSaveNo` del principal (`:284`) no tienen equivalente directo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **`ComandoDetalleUsuarioSICA_Click` no comprueba el resultado del guardado.** Llama a `ComandoGrabar_Click` y sigue aunque falle (`:212-214`): si había cambios pendientes y el registro falla, el usuario abre la edición SICA con la edición sin grabar y solo un `MsgBox` que puede pasar desapercibido.
- `DNIEnActivo` y `DNIEnHistorico` relanzan el error con `Err.Raise 1000` (`:529`, `:535`), así que el `CorreoAlAdministrador` del manejador (`:544`) no se dispara por duplicado: el aviso al administrador depende del tipo de error, no del problema.
- Los dos `MsgBox` de formato incorrecto (`:518`, `:522`) usan `Exit Sub` sin limpiar `m_Error`, que quedó vacío en `:513`; los dos de duplicado vacían el campo pero no `Cancel`, de modo que la validación depende del efecto colateral de `Me.DNI = Null`.
- `Form_Load` activa `AllowEdits` siempre (`:587-589`) y compensa después con un bucle que solo recorre `Me.Controls`, no los controles de subformulario: un control anidado quedaría editable en modo solo lectura.
- `ComandoPasarAHistorico` se oculta al abrir (`:609`) y `ComandoObservacionEditar`/`Eliminar` arrancan apagados (`:605-608`); el estado depende de que los `Click` de las listas los repongan, lo que no está en este `.cls`.
- `m_Error` es una variable de módulo del formulario, y los helpers `RellenarObjetoConDatos`, `HaHabidoCambios` y `DNIEnActivo` la sobrescriben desde fuera al pasarla `ByRef`.
- `ComandoGrabar_Click` cierra `FormUsuario00Principal` con `acSaveNo` (`:284`): el usuario no puede volver atrás sin perder lo no grabado en otras subpantallas.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pasar el identificador del usuario por ruta en vez de por global; comparator explícito de campos modificados en vez de reflexión sobre controles; y un único guard de permisos que devuelva el modo en vez de decidirlo dos veces con `TempVars` y con `Locked`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin usuario activo se comunica al usuario y no deja el formulario a medias.
2. En modo solo lectura los campos no admiten escritura pero las dos listas se pueden recorrer y abrir observaciones.
3. Grabar sin cambios no escribe en la base y explica por qué.
4. Un DNI con formato válido pero ya usado por un usuario actual o histórico se rechaza, vacía el campo y nombra el conflicto.
5. Un correo con formato inválido se rechaza al salir del control.
6. Editar el detalle SICA con cambios pendientes guarda antes de abrir la pantalla SICA, y si el guardado falla no navega.
7. Tras una edición correcta, el contenedor se recarga con los datos nuevos.

## Pendientes

- Quién escucha `EdicionRegistro` y si recarga o solo refresca: no hay `WithEvents` ni manejador en `src`.
- Por qué existe `ComandoPasarAHistorico` si se oculta al abrir (`:609`); el resto de su lógica está en `:471-507`.
- Implementación de `DNIEnActivo` y `DNIEnHistorico`, en `Constructor.bas`.
- Si `m_FormExpedientes_Seleccionado` (`:941`) tiene algún punto de enganche declarado en `src`.
- Qué componente cierra o conserva `m_ObjUsuarioAlInicio` tras grabar: el formulario solo lo lee.