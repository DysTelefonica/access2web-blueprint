# Form_frmBuscarJustificacion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_frmBuscarJustificacion.cls` (333 líneas) y `.form.txt` (26 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`; las clases de `src/classes/` y los repositorios de `src/modules/` se citan por el nombre del fichero. Las citas son `Form_frmBuscarJustificacion.cls:línea` salvo indicación.

Buscador y gestor de justificaciones, con dos modos en un mismo formulario. En **modo gestión** (el que usa el menú de entidades) ofrece alta, edición y baja; en **modo selección** —activado con `OpenArgs = "MODO_SELECCION"` (`:34`)— sólo permite elegir una justificación activa y devuelve el identificador al formulario que la pidió. Es **no enlazado**: la lista se puebla a mano desde la caché global de `Entorno`, sin `RowSource` de consulta.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú de entidades de `Form0BDOpciones` | `Entorno.EntidadesFormularios` asocia `EnumEntidades.JustificacionHPS` a este formulario; se abre sin `OpenArgs`, luego en modo gestión | `classes/Entorno.cls:228` |
| `Form_FormSolicitudesAltaTraspasoDatos` | `DoCmd.OpenForm "frmBuscarJustificacion", OpenArgs:="MODO_SELECCION"`, y se enlaza con `WithEvents` para recibir `JustificacionElegida` | `Form_FormSolicitudesAltaTraspasoDatos.cls:934-938` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesAltaTraspasoDatos` | `RaiseEvent JustificacionElegida(id, Titulo)` al elegir, manualmente o tras un alta, y cierre inmediato | `:216-217, 247-250` |
| `Form_frmJustificacion` | `DoCmd.OpenForm "frmJustificacion"` para el alta y `…, idSel` para la edición; el buscador se suscribe con `WithEvents` | `:148-151, 165-168` |
| `JustificacionServicio` | `VerificarUso` en cada refresco de botones y `EliminarJustificacion` en la baja | `:184, 315` |
| `Entorno` | `Justificaciones` para pintar la lista y `RecargarJustificaciones` al recibir el evento del hijo | `:98, 204` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | `Ajustar`, fija el modo por `OpenArgs`, pone «Activas» en el filtro y carga la lista | `:30-44` |
| `txtBuscar` | texto | `OnAfterUpdate` → `txtBuscar_AfterUpdate` | Filtra por título o descripción y, si queda un único resultado, actúa sobre él | `:45-68`, `.form.txt:703` |
| `txtBuscar` | texto | `OnChange` → `txtBuscar_Change` | Procedimiento vacío: sólo comentario | `:74-79` |
| `cboFiltroEstado` | combo | `OnAfterUpdate` → `cboFiltroEstado_AfterUpdate` | Recarga conservando el texto ya escrito | `:81-87`, `.form.txt:839` |
| `lstResultados` | lista de 5 columnas | `OnClick` / `OnAfterUpdate` / `OnDblClick` | `Click` y `AfterUpdate` recalculan botones; el doble clic elige o edita según el modo | `:253-260, 325-332`, `.form.txt:535, 540-542` |
| `btnAlta` | botón | `OnClick` → `btnAlta_Click` | Abre `frmJustificacion` en modo alta | `:144-152`, `.form.txt:573-575` |
| `btnEditar` | botón | `OnClick` → `btnEditar_Click` | Abre `frmJustificacion` con el id de la fila seleccionada | `:154-169`, `.form.txt:631-633` |
| `btnEliminar` | botón | `OnClick` → `btnEliminar_Click` | Pide confirmación y desactiva la justificación | `:171-192`, `.form.txt:732-734` |
| `cmdElegir` | botón | `OnClick` → `cmdElegir_Click` | Lanza `JustificacionElegida` y cierra; sólo visible y habilitado en modo selección y si la fila está activa | `:36-39, 237-251, 301-309`, `.form.txt:781-784` |
| `f_Detalle` (varias) | `WithEvents` | — | Referencia al editor abierto, para recibir `JustificacionGuardada` | `:15, 151, 168, 201-232` |
| `lblTitulo`, `Etiqueta80`, `Etiqueta264`, `Etiqueta237` | etiquetas | — | Rótulos «JUSTIFICACIONES», «Justificación», «Activas» y una sin rótulo | `.form.txt:499-500, 678-679, 867-868, 612` |
| `lblEstado` | etiqueta | — | Rótulo de diseño `Caption ="URLDirectorioAplicaciones"`; el `.cls` no la usa | `.form.txt:901-903` |

## Flujos

### Normal — carga y filtrado

1. `Form_Load` dimensiona el formulario y lee el modo (`:31-40`): con `OpenArgs = "MODO_SELECCION"` muestra `cmdElegir`; en caso contrario lo oculta.
2. Pone «Activas» en el filtro de estado y llama a `ConfigurarLista` y `CargarLista` (`:41-43`). `ConfigurarLista` fija cinco columnas y `ColumnWidths = "0cm;1cm;4cm;14cm;0cm"`, ocultando el identificador y la quinta columna (`:71-72`).
3. `CargarLista` (`:88-139`) vacía la lista, mete una cabecera fija `"ID;Est.;Justificación;Descripción;RawEstado"` (`:102`) y recorre `m_ObjEntorno.Justificaciones` (`:98`) filtrando en memoria: primero por el valor del combo (`:108-112`), después por coincidencia de texto en título o descripción con `InStr` sin distinguir mayúsculas (`:116-118`).
4. Cada fila se añade como una única cadena de cinco campos separados por `;`, con un símbolo ✔/✘ construido con `ChrW(10004)`/`ChrW(10006)` y la quinta columna con `CStr(just.activa)` (`:121-133`).
5. `ActualizarEstadoBotones` (`:280-324`) decide el estado de los tres botones a partir de la fila seleccionada.

### Normal — alta y edición

1. `btnAlta_Click` abre el editor sin argumentos y se enlaza con `WithEvents` (`:148-151`); el comentario de `:146-147` recuerda que no se use `acDialog` justamente para poder capturar el evento.
2. Al guardar el editor, el buscador recibe `f_Detalle_JustificacionGuardada` (`:201`), invalida la caché (`:204`) y recarga la lista manteniendo el filtro de texto (`:209`).
3. En modo selección relanza el evento con el id recibido y se cierra (`:212-217`); en modo gestión resalta la fila editada con `ResaltarFila` y recalcula los botones (`:218-228`). En ambos casos suelta la referencia al editor (`:231`).
4. `btnEditar_Click` exige selección y abre el editor con `idSel` (`:154-168`).

### Alternativos

- **Auto-selección:** `txtBuscar_AfterUpdate` dispara el flujo de la fila cuando `ListCount = 2`, es decir una cabecera más un resultado (`:51-62`): en modo selección elige y cierra, en modo gestión abre el editor. Con más de un resultado sólo pasa el foco a la lista (`:64-66`).
- **Filtro de estado:** «Activas», «Inactivas» y «Todas» (`:109-111`); el valor inicial es «Activas» (`:41`) y el de diseño del combo no decide.
- **Baja:** `btnEliminar_Click` pide confirmación con `vbYesNo` y llama a `EliminarJustificacion` (`:177-184`), que hace un `UPDATE … SET activa = False` (`JustificacionRepositorio.bas:62-67`) y recarga. Es una baja lógica, no un borrado.
- **Restricciones por fila:** elegir exige que la justificación esté activa (`:301-308`, con el texto emergente explicando el motivo); editar siempre está permitido (`:312`); eliminar exige que no esté en uso (`:315-323`). Las tres reglas se expresan con `ControlTipText`.

### Errores

- `btnEliminar_Click` captura el fallo con `On Error GoTo ErrorEliminar` y muestra «No se pudo eliminar: » con `Err.Description` (`:183, 190-191`).
- `Form_Load`, `CargarLista`, `ResaltarFila`, `ActualizarEstadoBotones` y los handlers de la lista **no tienen manejador de error**.

## Reglas, validaciones y cálculos

- **Modo por `OpenArgs`:** la única diferencia entre los dos modos es la visibilidad de `cmdElegir` (`:36, 39`) y tres ramificaciones en el guardado, el doble clic y los botones. Todo lo demás es común.
- **Sólo se pueden elegir justificaciones activas:** se lee la quinta columna, oculta, de la fila (`:299`), no el símbolo visible.
- **No se puede eliminar lo que está en uso:** `VerificarUso` cuenta las filas de `TbSolicitudes` con ese `idjustificacion` (`JustificacionRepositorio.bas:75`).
- **La lista no viene de una consulta:** `Justificaciones` es la caché de `Entorno`, que se rellena una sola vez con `ObtenerTodas` (`Entorno.cls:125-128`; `JustificacionRepositorio.bas:9-17`, sin `WHERE`). El buscador filtra ese `Collection` con `InStr`.
- **La búsqueda sólo mira título y descripción:** el identificador y el estado no son texto buscable (`:117-118`).
- **`Activas` es sólo el valor inicial:** `CargarLista` vuelve a leer el combo en cada llamada y usa `Nz(..., "Activas")` si está vacío (`:96`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listar | `SELECT idjustificacion, titulo, descripcion, activa FROM TbJustificaciones ORDER BY titulo` | DAO sobre el back-end, cacheada en `Entorno` | `JustificacionRepositorio.bas:12-16`, `Entorno.cls:125-128` |
| Comprobar uso | `SELECT Count(*) as Total FROM TbSolicitudes WHERE idjustificacion = ` & `id` | DAO, SQL concatenado | `JustificacionRepositorio.bas:75-77` |
| Desactivar (baja) | `UPDATE TbJustificaciones SET activa = False WHERE idjustificacion = ` & `id` | DAO, SQL concatenado | `JustificacionRepositorio.bas:65-66` |
| Invalidar caché | (ninguna) | Derivado | `:204`, `Entorno.cls:132-134` |

Las dos consultas del buscador concatenan un `Long`, así que no hay superficie de inyección en ellas; el punto sensible está en el filtro de texto, que no llega al SQL porque se evalúa en VBA.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que este formulario aporta al plan: es el único consumidor de la caché de justificaciones y el que decide las tres reglas de gestión (activa para elegir, en uso bloquea el borrado, baja lógica). En web las tres se resuelven en el servidor —una consulta con `WHERE activa` y paginación, y una clave foránea real que impida el borrado—, pero eso exige decidir si la baja lógica sigue siendo un `activa = False` o pasa a un estado explícito (`baja`), porque el filtro actual sólo conoce dos valores.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-38.** Los tres handlers comprueban `IsNull(Me.lstResultados)` en lugar de mirar la selección (`:156` en `btnEditar`, `:172` en `btnEliminar`, `:238` en `cmdElegir`): la propiedad `IsNull` de un cuadro de lista no enlazado es siempre `False`, de modo que la guarda nunca detiene nada. La comprobación correcta es la de `ActualizarEstadoBotones`, `Nz(Me.lstResultados.Column(0)) = 0` (`:285`). Sin selección, `Column(0)` devuelve `Null` y `idSel` recibe `0`: `btnEliminar` llega a mostrar «¿Estás seguro de eliminar…?» y sólo falla después, con «ID de justificación no válido» (`JustificacionServicio.cls:64-66`).
- **H-HPS-S-39.** La lista se construye concatenando los cinco campos con `;` en una sola cadena (`:102, 129-133`), de modo que un punto y coma en el título o en la descripción desplaza todas las columnas siguientes. La quinta columna es la que decide si se puede elegir (`:299`), y `CBool` sobre un fragmento que ya no es un booleo lanza error de conversión en `ActualizarEstadoBotones`, que no tiene manejador. El riesgo es real porque `descripcion` es `MEMO` en el ERD de origen (`schema.sql:75`) y puede contener cualquier carácter.
- **H-HPS-S-40.** El filtrado y la búsqueda se resuelven recorriendo en memoria la caché completa de justificaciones, que `ObtenerTodas` carga sin `WHERE` (`:98, 106-136`; `JustificacionRepositorio.bas:9-17`). No hay `LIKE` en la base, ni paginación, ni orden estable más allá del `ORDER BY titulo` de la consulta de origen; el coste crece con el número de filas y la tabla completa reside en memoria durante toda la sesión.
- **H-HPS-S-41.** `txtBuscar_AfterUpdate` ejecuta la acción sobre el único resultado sin pedir confirmación (`:51-62`): en modo gestión, una búsqueda que afine a una sola fila abre el editor de esa fila; en modo selección, lanza `JustificacionElegida` y cierra el buscador. Basta con una pulsación de tecla o un `Tab` para confirmar una elección que el usuario no ha hecho.
- `VerificarUso` lanza una consulta a `TbSolicitudes` en cada `Click`, en cada `AfterUpdate` de la lista y en cada recarga (`:315`, invocada desde `:138, 227, 326, 331`): cuatro rutas de consulta por interacción, sobre un contador que ya está disponible en la fila.
- `txtBuscar_Change` está vacío y sólo contiene comentarios (`:74-79`); el filtrado real ocurre en `AfterUpdate`, así que la lista no se mueve mientras se escribe, sólo al salir del campo.
- Rótulos heredados de plantilla: el botón de selección dice «Elejir» (`.form.txt:782`) y la etiqueta de estado lleva el texto de desarrollador `URLDirectorioAplicaciones` (`.form.txt:902`), que el `.cls` nunca usa.
- `ResaltarFila` (`:267-276`) convierte con `CLng` la primera columna de cada fila sin proteger el caso `Null`; la lista sólo contiene filas reales más la cabecera, que es texto no numérico, pero el recorrido empieza en `1` precisamente para esquivarla.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: consulta de búsqueda en servidor con paginación y orden explícito en lugar de recorrido de la caché; la lista como tabla con columnas reales en lugar de cadenas con separador; una ruta de alta y edición que devuelva el resultado al listado sin evento entre ventanas; y confirmación explícita antes de aceptar el resultado único de una búsqueda.

## Criterios de aceptación

Propuestos, no aprobados:

1. En modo gestión se ven todas las justificaciones con su estado, y en modo selección sólo las activas.
2. El filtro de estado «Inactivas» muestra exactamente las no activas y «Todas» las muestra todas.
3. Una búsqueda coincide por título o por descripción, sin distinguir mayúsculas, y la lista se actualiza al confirmar el texto.
4. Sin fila seleccionada, «Editar», «Eliminar» y «Elegir» no ejecutan ninguna acción ni muestran una confirmación.
5. Un título o una descripción con `;` se muestran íntegros y no alteran el estado de los botones.
6. Una justificación con solicitudes asociadas no se puede eliminar; al intentarlo, el botón explica por qué.
7. Tras un alta o una edición, la lista refleja el cambio sin reiniciar la aplicación y respeta el filtro activo.
8. En modo selección, elegir una justificación inactiva no es posible y el motivo se explica.
9. La baja deja la justificación inactiva, no la borra, y sigue consultable con el filtro «Inactivas».

## Pendientes

- Tamaño real de `TbJustificaciones`: el ERD de origen declara «-- Rows: 7» (`schema.sql:77`); es el volumen que asume el filtrado en memoria.
- Si la etiqueta `Etiqueta237` (`.form.txt:612`) tiene rótulo o sólo ocupa espacio: el `.form.txt` no declara `Caption` y el `.cls` no la menciona.
- Origen de `EnumEntidades.JustificacionHPS` y cómo lo recorre el menú de `Form0BDOpciones`: la asociación está en `Entorno.cls:228`, pero el consumo del diccionario queda en el formulario del menú, fuera de esta ficha.
- Los dos modos comparten el mismo formulario; no consta en `src` si algún tercer lanzador además de `Form_FormSolicitudesAltaTraspasoDatos` los usa con `MODO_SELECCION`.