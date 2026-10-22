# Form_FormRiesgosGestion

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgosGestion.cls` (1784 líneas) y `src/forms/FormRiesgosGestion.form.txt` (5044 líneas; el código empieza en `.form.txt:3276`) · **PRD:** [README](../README.md)

Listado de los riesgos de una edición. Es el contenedor del dominio: un árbol de cuatro niveles (edición → riesgo → mitigación o contingencia → acción) a la izquierda, un subformulario de detalle a la derecha, y una barra de separación arrastrable entre ambos. Es además el **reloj de la máquina de estados del riesgo**: fija los objetos globales que el resto de formularios leen. Las citas son `Form_FormRiesgosGestion.cls:línea` salvo indicación.

La estructura del árbol se reconstruye siempre, y su esquema de claves es la columna vertebral de todo el código: `EDICION|IDE`, `RIESGO|IDR`, `PM|IDR|IDPM`, `PC|IDR|IDPC`, `PMACCION|IDPM|IDACC` y `PCACCION|IDPC|IDAC` (`:1149-1157`, `:1106-1115`). La decodificación vive en `Funciones Generales.bas:6157`.

## Entrada y salida

Es el único punto de entrada al árbol y se abre desde nueve sitios distintos, con y sin argumento:

| Quién lo abre | Con `OpenArgs` | Evidencia |
|---|---|---|
| `Form_FormGestionRiesgos` (edición de proyecto) | no | `Form_FormGestionRiesgos.cls:224` |
| `Form_FormProyectosGestion` | no | `Form_FormProyectosGestion.cls:453` |
| `Form_Form0BDOpciones` y `Form_Form0BDOpcionesTecnico` | no | `Form_Form0BDOpciones.cls:345`, `Form_Form0BDOpcionesTecnico.cls:138` |
| `Form_FormCalidadTareasDetalleEdicion`, `Form_FormTecnicoTareasDetalleEdicion` | no | `Form_FormCalidadTareasDetalleEdicion.cls:193`, `Form_FormTecnicoTareasDetalleEdicion.cls:101` |
| `Form_FormExpedientesBusqueda`, `Form_FormGestionRiesgosRiesgosOferta` | no | `Form_FormExpedientesBusqueda.cls:177`, `Form_FormGestionRiesgosRiesgosOferta.cls:343` |
| `Form_FormRiesgo` (volver desde la ficha) | no | `Form_FormRiesgo.cls:122` |
| `Form_FormCalidadRiesgoMaterializaciones` | **sí**, con el identificador de la edición | `Form_FormCalidadRiesgoMaterializaciones.cls:229` |

El `OpenArgs` solo se acepta si `IsNumeric` lo confirma (`:985`); en ese caso fija la edición de arranque (`:986`). Si no es numérico se ignora en silencio y se usa la última edición del proyecto activo (`:163`).

Salidas: `FormPublicacionTecnico` o `FormPublicacionCalidad` según rol (`:60-70`), `FormRiesgosEstablecerPrioridades` (`:99`), `FormGestionRiesgos` (`:804`), `FormAnexos` con el argumento literal «Edición» o «Riesgo» (`:672-686`), y la ayuda (`:707`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Resuelve edición, carga datos y selecciona el nodo raíz | `:976-1019` |
| `Arbol` | control ActiveX de árbol | `Arbol_NodeClick` | **Enruta el contexto**: decide el subformulario y fija los globales | `:242-516`, `.form.txt:944` |
| `FormDetalle` | subformulario | — | Cambia de `SourceObject` según el tipo del nodo | `:2390`, `:280, 296, 338, 380, 426, 467` |
| `ListaImagenes` | control de lista de imágenes | — | Catálogo de iconos del árbol, OCX embebido | `:183`, `.form.txt:2412` |
| `ComboEdicion` | combo | `ComboEdicion_BeforeUpdate` | Cambia de edición y recarga el árbol | `:821-900`, `.form.txt:601` |
| `ComboVerRetirados` | combo | `ComboVerRetirados_Change` | Filtra o no los riesgos retirados | `:943-974`, `.form.txt:560` |
| `ComboVerDescripcion` | combo | `comboVerDescripcion_BeforeUpdate` | Alterna entre título descriptivo y título de estado del nodo | `:906-941`, `.form.txt:872` |
| `Sep` | etiqueta arrastrable | `Sep_MouseDown` / `MouseMove` / `MouseUp` | Reparte el ancho entre árbol y detalle | `:1021-1070` |
| `lblEstadoRiesgo` | etiqueta | — | Muestra el estado del riesgo seleccionado | `:300`, `.form.txt:634` |
| `lblUltimoCambio` | etiqueta | — | Quién y cuándo tocó la edición por última vez | `:228-240`, `.form.txt:3169` |
| `ComandoActualizarParteArbol` | botón | `ComandoActualizarContador_Click` | Vacía el detalle, recarga la edición y redibuja el árbol | `:522-558`, `.form.txt:3201` |

## Flujos

### Apertura (`Form_Load`, `:976-1019`)

1. Si el `OpenArgs` es numérico, resuelve esa edición (`:985-990`).
2. `EstablecerDatos` (`:117-226`): elige la edición activa (`:137-155`), rellena el combo con todas las ediciones del proyecto (`:160-164`), fija los dos filtros desde las preferencias del entorno (`:167-176`), carga los iconos (`:184`), registra el último proyecto (`:188`) y dibuja el árbol (`:192`).
3. Calcula el permiso de edición (`:202-215`): solo la edición activa concede permisos, y exige que el usuario esté autorizado en la edición.
4. Coloca el separador y selecciona el nodo raíz (`:997-1002`).

### Normal: selección de nodo (`Arbol_NodeClick`, `:242-516`)

Decodifica la clave (`:254`) y **ramifica por tipo**, seis ramas con la misma forma:

| Nodo | Subformulario | Globales que fija | Línea |
|---|---|---|---|
| Edición | `FormRiesgosGestionEdicion` | `m_ObjEdicionActiva`, permiso | `:259-288` |
| Riesgo | `FormRiesgosGestionRiesgo` | `m_RiesgoSeleccionado`, `m_EstadoRiesgoActivo` | `:290-331` |
| Plan de mitigación | `FormRiesgosGestionPlanPrincipal` | `m_ObjPMActivo`, `m_EsMitigacion = Sí` | `:332-380` |
| Plan de contingencia | `FormRiesgosGestionPlanPrincipal` | `m_ObjPCActivo`, `m_EsMitigacion = No` | `:381-424` |
| Acción de mitigación | `FormRiesgosGestionPlanAcciones` | `m_ObjPMAccionActiva` | `:425-460` |
| Acción de contingencia | `FormRiesgosGestionPlanAcciones` | `m_ObjPCAccionActiva` | `:461-503` |

Todas repiten el mismo bloque final: pintar la etiqueta de estado (solo si la edición está activa), releer el riesgo para detectar si cambió y cerrar el indicador de progreso. El cierre lo hace `AvanceCerrar` (`Funciones Generales.bas:4920`).

### Alternativo: filtrar y recargar

`ComboVerRetirados_Change` y `comboVerDescripcion_BeforeUpdate` no filtran en memoria: **vuelven a dibujar el árbol entero** (`:960`, `:921`). El filtro se traduce en qué colección se lee: `colRiesgos` o `colRiesgosActivos` (`:775-781`), y el título del nodo se recalcula con `setTitulosNodosRiesgos` (`:921`). El botón de refresco parcial (`:522-558`) invalida además los títulos e iconos cacheados en el propio riesgo (`:1383-1391`).

### De error

Todas las routines terminan en el mismo patrón `errores:`: si el número no es 1000, componen el mensaje, **envían correo al administrador** y muestran un `MsgBox` crítico; si es 1000, muestran solo una advertencia (`:507-515`). La excepción es `Form_Load`, que además **cierra el formulario** (`:1018`).

## Reglas, validaciones y cálculos

- **La máquina de estados vive fuera.** Este formulario no la implementa: la enumeración `EnumRiesgoEstado` tiene catorce estados (`Constantes.bas:153-167`), y el árbol solo la refleja a través del icono (`getNombreIconoRiesgoArbolDeRiesgos`, `:1392`) y del texto de la etiqueta (`RellenaEtiquetaRiesgo` → `Funciones Generales.bas:6832`, `:6740`).
- **El permiso se calcula dos veces con dos criterios distintos**, según la rutina (`:202-210` frente a `:267-276`).
- **Los booleanos de negocio viajan como texto**: «Sí» y «No» en los combos, en `setModoVerArbolRiesgos` y en `p_EsMitigacion` (`:121-134`, `:921`).
- **Dibujar el árbol escribe en base de datos**: `GrabarNombreNodoDesc` (`:1398`) y `GrabarNombreNodoEstado` (`:1408`) persisten el título cacheado de cada riesgo.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Ediciones, riesgos y planes | `Constructor`, objetos de dominio | `:138-163`, `:254`, `:861` |
| Colección de riesgos a listar | `Edicion.colRiesgos` o `colRiesgosActivos` | `:775-781`, `:1213` |
| Iconos y títulos de nodo | properties del riesgo, con escritura | `:1383-1412` |
| Último proyecto del usuario | `RegistrarUltimoProyecto` | `:188` |
| Estado global de la sesión | `Variables Globales.bas:75-80` | `:266, 293, 336, 383, 428, 464` |

## Coexistencia

El árbol es un control OCX de Windows con lista de imágenes y una base de datos_ACCESS abierta en exclusiva; ninguna de las dos cosas tiene equivalente directo en un navegador. La lista de nodos, el orden y los textos cacheados en el riesgo son datos de presentación guardados junto al dato de negocio, de modo que una web que los calcule por su cuenta necesita una regla de equivalencia explícita.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-158.** El permiso `blnPermitidoEditar` se calcula con **las dos ramas asignando lo mismo** al seleccionar un nodo de edición: si el usuario está autorizado sale `True` y si no también (`:270-275`). En la apertura el mismo permiso sí distingue (`:202-210`). El efecto real es que seleccionar un nodo de edición deja el permiso en `True` siempre que la edición esté activa.
- **H-GR-159.** Al cambiar de edición se rehabilitan anexos, publicación y edición de proyecto sin comprobar nada (`:869-871`), mientras que en la apertura esos mismos comandos sí dependen del permiso (`:211-215`). Conviene verificar si es intencionado.
- **H-GR-160.** El título del formulario y la etiqueta de último cambio incluyen el nombre de la persona conectada y de quien hizo el último cambio (`:164`, `:228-238`; `.form.txt:40` y `:3170` guardan además un nombre real como valor de diseño). Es información personal en el diseño, no solo en ejecución.
- **H-GR-161.** Cada clic relee el riesgo desde `Constructor.getRiesgo` para detectar si el identificador cambió (`:311`, `:319`, `:355`, `:363`, `:399`, `:407`, `:441`, `:449`, `:479`, `:487`). Con un árbol de cientos de nodos, navegar es un número alto de lecturas repetidas del mismo objeto.
- **H-GR-162.** El `OpenArgs` se acepta solo si es numérico (`:985`), pero no se comprueba que esa edición pertenezca al proyecto del usuario conectado: la edición se carga y se muestra sin más (`:986-990`).
- **H-GR-163.** `Err.Raise 1000` se lanza en decenas de puntos sin descripción (`:256`, `:302`, `:315`, …). El manejador de error muestra `m_Error` (`:514`), que en esos casos está vacío: el usuario ve una advertencia sin texto.
- **H-GR-164.** `EstablecerDatos` reactiva `AllowEdits` **también en su rama de error** (`:224-226`), dejando el formulario editable pese a haber fallado la carga.
- **H-GR-165.** `Form_Load` cierra el formulario tras avisar (`:1018`), con lo que el usuario pierde el árbol y cualquier filtro que hubiera elegido sin llegar a ver el motivo del fallo más de una vez.
- **H-GR-166.** Al construir el árbol se **escribe en la base de datos** (`GrabarNombreNodoDesc`, `:1398`; `GrabarNombreNodoEstado`, `:1408`). Un simple listado modifica los riesgos que muestra.
- **H-GR-167.** Al refrescar se vacían los títulos e iconos cacheados en el riesgo (`:1383-1391`) y se vuelven a grabar. Si el refresco falla a mitad, el riesgo queda con el título vacío y el árbol lo muestra sin etiqueta.
- **H-GR-168.** El separador depende de tres manejadores con `On Error Resume Next` anidados (`:1022`, `:1036-1037`) y un mínimo de 60 twips (`:1063`). Ningún error de redimensionado se comunica y el ancho no se persiste entre sesiones.
- **H-GR-169.** El resultado de los `MsgBox` se guarda en la global `pregunta` (`Variables Globales.bas:69`), que es un `Long` compartido por toda la aplicación: la variable sobrevive al mensaje y se lee desde otras rutinas.
- **H-GR-170.** El botón cuyo rótulo es «Contador» (`.form.txt:3201-3202`) ejecuta `ComandoActualizarContador_Click`, que recarga el árbol completo (`:522-545`). El nombre, el rótulo y el efecto no coinciden, y no hay contador visible en el formulario.
- **H-GR-171.** El estado de contexto vive en variables `Public` del propio formulario (`:30-49`), entre ellas `m_NodoSeleccionado`, `m_ColRiesgosAplicados` y `m_EsMitigacion`. Otros formularios los leen y escriben directamente, así que el orden de recarga del árbol es parte del contrato informal entre ellos.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el árbol por una tabla con agrupación y filtros por estado, tipo de plan y edición; derivar el permiso de edición de una decisión del servidor en lugar de una bandera compartida; calcular los títulos de nodo en el servidor sin persistirlos en el riesgo; devolver el listado de riesgos como recurso con filtros en la URL, sin el estado global que hoy comparte el árbol con sus subformularios.

## Criterios de aceptación

Propuestos, no aprobados:

1. El listado muestra los riesgos de la edición activa con su estado y su plan asociado.
2. Cambiar de edición recarga el listado y respeta el permiso del usuario en esa edición.
3. Filtrar por riesgos retirados o por modo de título no deja el árbol a medias ni pierde la selección razonable.
4. Un riesgo en estado alto o muy alto se distingue visualmente en el listado.
5. Ante un fallo de carga, el usuario recibe un mensaje con la causa y no un formulario que se cierra solo.

## Pendientes

- Quién escribe y lee las variables públicas de contexto, y con qué orden, más allá de los subformularios ya fichados.
- Qué garantiza que el listado de una edición y el estado guardado en el árbol no se desincronicen si dos usuarios trabajan a la vez.
- Si el separador arrastrable y el ancho del árbol forman parte de la experiencia esperada o son un residuo del formulario de escritorio.
- Cómo se traduce `EnumRiesgoEstado` en la web: catorce estados (`Constantes.bas:153-167`) en una columna, o una agrupación distinta.