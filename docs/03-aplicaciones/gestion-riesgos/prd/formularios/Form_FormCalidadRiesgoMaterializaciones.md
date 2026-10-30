# Form_FormCalidadRiesgoMaterializaciones

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadRiesgoMaterializaciones.cls` (718 líneas) y `src/forms/FormCalidadRiesgoMaterializaciones.form.txt` · **PRD:** [README](../README.md)

Histórico de materializaciones de un riesgo y decisión sobre su no conformidad. Muestra una lista de materializaciones y, sobre la seleccionada, permite **vincularla a una NC existente**, **marcar que no requiere NC** o **revocar una decisión ya tomada**. No acepta parámetros: opera sobre la global `m_ObjRiesgoActivo`. Las citas son `Form_FormCalidadRiesgoMaterializaciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormRiesgoMaterializado` | Cierra la instancia previa si existe y abre esta sin argumentos | `Form_FormRiesgoMaterializado.cls:157-160` |

No acepta `OpenArgs`. El riesgo es la global `m_ObjRiesgoActivo` (`Variables Globales.bas`), y el estado se toma de ella en `:385`. La materialización seleccionada llega por la columna 0 del cuadro de lista (`:59`, `:328`, `:467`, `:568`).

**Salidas:** una sola, y no es explícita: el botón «Ver PC» **cierra** la ventana de gestión de riesgos si está abierta y abre otra apuntando al plan de contingencia de la materialización seleccionada (`:225-229`, `:240`). El resto de acciones no cierra la ventana; solo recargan la lista y recalculan las tareas de calidad (`:92-107`, `:170-185`, `:514-529`, `:614-629`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Rótulo, tamaño, combo de NC, botónera y lista | `:258-287` |
| `ListasMaterializaciones` | cuadro de lista | `ListasMaterializaciones_Click` | Resuelve la selección y repinta la grupo de decisión | `:549-601`, `.form.txt:694-699` |
| `ComboNC` | desplegable | — | Lista de no conformidades candidatas | `:662`, `.form.txt:738-740` |
| `DescripcionDecisionMaterializado` | cuadro de texto | — | Explica en texto el estado de la decisión | `:350`, `:359`, `:366`, `.form.txt:1596` |
| `ComandoVincularNC` | botón | `ComandoVincularNC_Click` | Vincula la materialización a la NC elegida | `:460-547`, `.form.txt:1652-1656` |
| `ComandoNoParaNC` | botón | `ComandoNoParaNC_Click` | Registra que no genera NC | `:53-125`, `.form.txt:1688-1692` |
| `ComandoRevocarDecision` | botón | `ComandoRevocarDecision_Click` | Devuelve la materialización a «por decidir» | `:127-205`, `.form.txt:1725-1729` |
| `ComandoVerPlanesContingencia` | botón | `ComandoVerPlanesContingencia_Click` | Abre la gestión de riesgos en el plan de contingencia | `:211-256`, `.form.txt:1849-1851` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` | Abre la ayuda | `:22-51`, `.form.txt:613-616` |
| `ComandoSalir` | botón | `ComandoSalir_Click` | Cierra la ventana | `:207-209`, `.form.txt:1771-1774` |
| `m_FormNC` | formulario de NC | `m_FormNC_NCRegistrada` | Refrescar la lista cuando se registra una NC | `:14`, `:603-642` |

## Flujos

### Apertura (`Form_Load`, `:258-287`)

1. Reloj de arena, `Me.Caption = m_ObjEntorno.TituloUsuarioConectado` y `AjustarTamaño Me` (`:261-267`).
2. `EstablecerDatos` (`:268`), que a su vez: oculta el grupo de NC (`:294`), fija el `RowSource` de la lista (`:388`), rellena el desplegable (`:390`), prepara la botónera (`:394`) y carga la lista (`:398`).
3. En error, el manejador avisa, **vuelve a avisar** y cierra (`:280-286`).

### Normal: decisión sobre la materialización seleccionada

Los tres manejadores de decisión comparten la misma forma:

1. Resolver la selección desde la columna 0 de la lista, o bien recuperar el objeto si la selección cambió (`:57-70`, `:132-145`, `:465-478`).
2. Si no hay nada seleccionado, avisar y salir (`:72-74`, `:147-149`, `:480-482`).
3. Preguntar con confirmación y botón predeterminado «No» (`:75-79`, `:152-156`, `:496-500`).
4. Escribir: `RegistrarParaNONC` (`:84`), `RegistrarPorDecidir` (`:162`) o `VincularNC` (`:506`).
5. Invalidar la caché de materializaciones del riesgo, recargar la lista, reseleccionar la fila y recalcular las tareas de calidad (`:92-107`, `:170-185`, `:514-529`).

### Alternativo: ver el plan de contingencia (`:211-256`)

Si no hay selección, sale sin hacer nada (`:216-218`). Si la hay, cierra la gestión de riesgos si está abierta (`:225-227`), la abre apuntando a la edición de la materialización (`:229`), recupera el riesgo (`:230`) y selecciona el nodo correspondiente (`:240`).

### De error

Todos los manejadores repiten el patrón de la aplicación: número de error distinto de 1000 → mensaje con detalle, correo al administrador y diálogo crítico; 1000 exacto → solo advertencia. En `Form_Load` se **muestra además un segundo diálogo** (`:285`). El manejador de la lista (`:594-600`) y el de la ayuda (`:41-47`) no cierran la ventana.

## Reglas, validaciones y cálculos

- **La decisión depende del estado calculado de la materialización**, no de los permisos de quien la toma (`:346-370`): si no está calculada, no se ofrece nada; si no tiene fecha de decisión, se ofrece «decidir»; si la decisión fue «no NC», se ofrece revocar; si tiene NC vinculada, se muestra su código y se ofrece revocar.
- **La columna «pte.decidir NC» de la lista está invertida**: `Decidido = EnumSiNo.Sí` produce el texto «No» y `Decidido = EnumSiNo.No` produce «Sí» (`:437-441`). Traduce el estado a lenguaje de usuario, no al revés.
- **Las decisiones son tres y todas irreversibles desde aquí**: el camino de vuelta es revocar, y el texto emergente del botón lo dice («el sistema la desvincula, no borra la NC», `.form.txt:1729`).
- **La lista se arma concatenando cadenas** con `;` como separador (`:449-450`), el mismo separador que usan el `RowSource` y el desplegable de NC (`:388`, `:424`, `:662`). Un `;` en cualquiera de los tres campos rompería el número de columnas.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo y materializaciones | Global `m_ObjRiesgoActivo` y su colección, invalidada con `Set … = Nothing` | `:92`, `:170`, `:426`, `:514`, `:614` |
| Materialización seleccionada | `Constructor.getRiesgoMaterializado` (`Constructor.bas:6666`) | `:59`, `:328`, `:467`, `:568` |
| No conformidades candidatas | `Constructor.getNCs` (`Constructor.bas:5506`), `Constructor.getNC` (`Constructor.bas:3643`) | `:663`, `:488` |
| Decisiones | `RegistrarParaNONC`, `RegistrarPorDecidir`, `VincularNC` | `:84`, `:162`, `:506` |
| Tareas de calidad | `EstablecerTareasCalidad` (`Funciones Generales.bas:2054`) | `:105`, `:183`, `:527`, `:627` |
| Gestión de riesgos | `Form_FormRiesgosGestion.SeleccionarNodo` | `:240` |

## Coexistencia

Pendiente de `../../forms.md`. La lista se pinta con un `RowSource` de cadenas separadas por `;`, y el desplegable de NC se alimenta de un constructor que devuelve una colección. Ninguna de las dos cosas tiene equivalente directo: la web necesita una consulta con columnas reales, lo que obliga a decidir qué pasa con el separador embebido en los datos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-94.** `m_FormNC` se declara `WithEvents` (`:14`) y **nunca se instancia**: no hay ningún `Set m_FormNC` en el fichero. El manejador `m_FormNC_NCRegistrada` (`:603-642`) no puede dispararse, aunque el evento existe en el formulario de NC (`Form_FormRiesgoNC.cls:15, 62`). Es código muerto que aparenta una integración que no está.
- **H-GR-95.** `Form_Load` muestra **dos diálogos de error**: el del manejador (`:283`) y uno incondicional detrás (`:285`), ambos con el mismo texto, y después cierra (`:286`). El usuario que falle al abrir tiene que cerrar dos ventanas para volver.
- **H-GR-96.** `ComandoSalir_Click` cierra **sin** `acSaveNo` (`:208`), mientras que todos los demás cierres del fichero lo usan. Es el único punto del formulario que pide confirmación de guardado.
- **H-GR-97.** El diseño declara seis columnas y sus anchos (`.form.txt:687, 697`), pero guarda un `RowSource` de cinco (`.form.txt:696`). En ejecución, `EstablecerDatos` fija cinco columnas (`:388`) que `EstablecerLista` sobrescribe de inmediato con seis (`:424`): la primera asignación es código muerto. Como la lista **no muestra cabeceras** (`ColumnHeads = NotDefault`, `.form.txt:682`), las columnas se distinguen solo por su contenido.
- **H-GR-98.** `EstablecerComboNC` sale sin rellenar el desplegable si `m_ObjRiesgoAlInicio` está vacío (`:654-656`), una variable distinta del `m_ObjRiesgoActivo` que usa todo el resto del formulario. Si esa condición se cumple, «Vincular NC» queda sin candidatas y el error solo aparece al pulsar el botón.
- **H-GR-99.** `VisibilizarDesvisibilizarGrupoNC` no toca `ComandoRevocarDecision` en su rama de activar (`:293-299`) y lo pone visible en la de desactivar (`:307`). La visibilidad de ese botón queda heredada de la llamada anterior (`:322`, `:355`, `:360`, `:367`), no de la decisión actual.
- **H-GR-100.** El desplegable de NC **guarda filas de ejemplo en el diseño** (`.form.txt:740`, y el mismo origen repetido en `:1155`). No se reproduce aquí. Son datos reales de una no conformidad guardados en el formulario: si el `RowSource` llegara a ejecutarse sin ser sustituido por `:662`, la lista mostraría esos datos.
- **H-GR-101.** `ComandoAyuda` tiene rótulo «Salir» (`.form.txt:614`), texto emergente «Ayuda» (`.form.txt:616`) e imagen de ayuda (`.form.txt:617`), y convive con `ComandoSalir`, también rotulado «Salir» (`.form.txt:1772`). Dos botones con el mismo texto y efectos opuestos.
- **H-GR-102.** `EstablecerBotoneraGrupoNC` sale sin dejar texto ni habilitar acciones cuando la materialización todavía no está calculada (`:346-349`): el usuario ve un grupo de decisión vacío, sin explicación y sin salida visible, cuando lo que corresponde es un mensaje.
- **H-GR-103.** El bloque «resolver selección → preguntar → escribir → invalidar caché → recargar lista → reseleccionar → recalcular tareas» está **copiado cuatro veces** (`:53-118`, `:127-192`, `:460-537`, `:603-629`), con la única diferencia del método de escritura y del texto de la pregunta. Cualquier corrección hay que hacerla por quadruplicado.
- **H-GR-104.** `blnCalculadoNC` (`:18`) y `m_EstadoRiesgo` (`:20`, asignado en `:385`) se declaran y se asignan, y no se leen en ningún punto del fichero.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: decidir por tabla y no por separador, con columnas reales en la respuesta; dejar de invalidar caché y recargar la lista tras cada decisión, y devolver el estado ya calculado; substituir el grupo de botones por una acción única con el verbo explícito («vincular», «no requiere NC», «revocar») decidida por parámetro; revisar si la integración con el formulario de NC sigue siendo necesaria o si la web resuelve la vinculación desde la propia pantalla.

## Criterios de aceptación

Propuestos, no aprobados:

1. La ventana lista las materializaciones del riesgo activo con su estado de decisión.
2. Al seleccionar una, el grupo de decisión refleja si esa materialización está pendiente, decidida sin NC o vinculada a una NC concreta.
3. Vincular, marcar sin NC y revocar producen un cambio visible en la lista y actualizan las tareas de calidad.
4. Ninguna de esas tres acciones es posible sin una materialización seleccionada, y el usuario recibe un aviso si lo intenta.

## Pendientes

- Si la colección de materializaciones del riesgo se puede dejar vacía a mitad de una operación (`:426`, `:514`, `:614` invalidan la caché antes de recargar) y qué muestra la lista en ese caso.
- Qué criterio usa `RiesgoMaterializacion.EsMaterializacionCalcuado` para distinguir «pendiente» de «todavía no evaluada» (`:346`).
- Si el formulario de NC puede alcanzarse desde otro camino, dado que aquí la instancia `m_FormNC` no existe (H-GR-94).
- Qué hace `AjustarTamaño` con una ventana de lista de altura variable (`Funciones Generales.bas:1043`).