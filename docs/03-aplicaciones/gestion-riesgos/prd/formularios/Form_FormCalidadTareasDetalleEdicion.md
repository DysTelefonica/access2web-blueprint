# Form_FormCalidadTareasDetalleEdicion

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadTareasDetalleEdicion.cls` (361 líneas) y `src/forms/FormCalidadTareasDetalleEdicion.form.txt` · **PRD:** [README](../README.md)

Ficha de una edición dentro del árbol de tareas de calidad. No nace de forma independiente: se incrusta en el subformulario del árbol y **lee la edición seleccionada de la instancia global del formulario padre**. Muestra los datos de cabecera del proyecto y de la edición, marca en rojo los plazos vencidos, y ofrece cinco salidas: recalcular la fecha de publicación, generar el informe, generar el informe de publicabilidad, ir al flujo de publicación y saltar a la gestión de riesgos. Las citas son `Form_FormCalidadTareasDetalleEdicion.cls:línea` salvo indicación.

## Entrada y salida

**No tiene entrada propia.** No lee `OpenArgs`: en `Form_Open` comprueba que el padre tenga algo seleccionado y aborta si no (`:252-255`).

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalidadTareas` | `Me.FormDetalle.SourceObject = "FormCalidadTareasDetalleEdicion"` | `Form_FormCalidadTareas.cls:50` |

La lectura del contexto es una **referencia directa al miembro público** `Form_FormCalidadTareas.m_EdicionSeleccionada` (`:252`, `:256`), declarado en `Form_FormCalidadTareas.cls:37`. El valor se copia a la variable de instancia `m_Edicion` (`:13`) y de ahí sale todo lo demás.

Salidas:

| Botón | Efecto | Evidencia |
|---|---|---|
| Ir a publicar | Fija proyecto y edición activos y abre `FormPublicacionCalidad` | `:155-160` |
| Ver detalle | Cierra `FormRiesgosGestion` si está abierto y lo reabre | `:186-193` |
| Ver proyecto | Cierra `FormGestionRiesgos` si está abierto y lo reabre | `:218-225` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Valida que haya edición activa y llama a `EstablecerDatos` | `:243-276` |
| `Proyecto` / `NombreProyecto` | texto | — | Nombre del proyecto | `:309`, `:310` |
| `Supervisor` | texto | — | Usuario de calidad asignado | `.form.txt:736` |
| `EdicionActual` | texto | — | Etiqueta de la edición | `:315`, `.form.txt:1452` |
| `EdicionAnterior` / `FechaPublicacionEdicionAnterior` | texto / fecha | — | Edición previa y su fecha de publicación | `:303-308`, `.form.txt:580`, `.form.txt:639` |
| `FechaProgPublicar` / `DiasParaPublicar` | fecha / numérico | — | Fecha máxima de publicación y días restantes, en rojo si vence | `:316-324`, `.form.txt:1482`, `.form.txt:1605` |
| `FechaCierre` / `FechaFinContrato` | fecha | — | Cierre del proyecto y fin de contrato, en rojo si el contrato venció | `:329-345`, `.form.txt:670`, `.form.txt:1291` |
| `IDExpediente` / `IDProyecto` / `IDEdicion` (+ etiquetas) | texto | — | Identificadores técnicos, visibles solo para administradores | `.form.txt:951`, `.form.txt:980`, `.form.txt:1077` |
| `ComandoVerDetalle` | botón | `Click` | Salta al detalle de riesgos | `:179-209`, `.form.txt:1213` |
| `ComandoVerProyecto` | botón | `Click` | Salta a la gestión de proyectos | `:211-241`, `.form.txt:1251` |
| `ComandoIrAPublicar` | botón | `Click` | Abre el flujo de publicación | `:148-177`, `.form.txt:1176` |
| `ComandoGenerarInforme` | botón | `Click` | Genera el informe de la edición | `:63-110`, `.form.txt:1352` |
| `ComandoInformePublicabilidad` | botón | `Click` | Genera el informe de publicabilidad | `:112-146`, `.form.txt:1137` |
| `ComandoActualizarFechaMaxPublicacion` | botón | `Click` | Recalcula la fecha máxima de publicación | `:15-61`, `.form.txt:1634` |

## Flujos

### Apertura (`Form_Open`, `:243-276`)

1. Si `Form_FormCalidadTareas.m_EdicionSeleccionada` es `Nothing`, compone el mensaje «No se conoce la edición activa» y lanza el error interno (`:252-255`).
2. Copia la selección a `m_Edicion` y llama a `EstablecerDatos` (`:256-260`).

El manejador de error **cierra el formulario** (`:275`) sin asignar `Cancel`, y el subformulario queda incrustado en el árbol hasta que otra selección lo sustituya.

### Carga de datos (`EstablecerDatos`, `:279-357`)

1. Deshabilita `ComandoVerDetalle` (`:286`) y decide la visibilidad de los identificadores según `EsAdministrador` o `EsAdministradorConectadoInicialmente` (`:287-301`).
2. Si la edición tiene edición anterior, copia su etiqueta y su fecha de publicación si es válida (`:303-308`).
3. Vuelca los datos del proyecto: nombre, supervisor y los tres identificadores (`:309-314`).
4. Calcula los días hasta la publicación programada y pinta `DiasParaPublicar` en rojo si el plazo se ha vencido o es hoy (`:316-324`).
5. Copia las fechas de preparación, cierre y fin de contrato (`:326-345`).
6. Rehabilita `ComandoVerDetalle` (`:351`).

### Alternativo: `ComandoActualizarFechaMaxPublicacion_Click` (`:15-61`)

Recarga el proyecto, llama a `FechaMaxProximaPublicacionGrabar` y **vuelve a leer la edición**, como si la fecha se recalculara en base de datos (`:22-33`). Después repite el cálculo de días y, aquí sí, **limpia el campo cuando la fecha no es válida**: `Me.FechaProgPublicar = Null` (`:35-36`).

### Alternativo: generación de informes (`:63-110` y `:112-146`)

- **Informe de la edición.** Comprueba que todos los riesgos estén priorizados y, si no, avisa y sale (`:69-72`). Si pasa el filtro, llama a `GenerarInformePublicacion` con el tipo HTML y guarda la ruta en una variable local `url` (`:82`). Esa ruta **no se usa**: la línea que la abriría está comentada, con la explicación de que la propia función ya se encarga (`:88-89`).
- **Informe de publicabilidad.** Pasa el identificador de la ventana del formulario a `GenerarInformePublicabilidadEdicionInteractivoHTML` (`:123`), que se ocupa de abrirlo, y deja la ruta en la **global `m_URLHTMLActivo`** (`:127`), que es lo que lee el visor HTML (`Form_FormWeb`) para decidir qué muestra.

### Navegación (`ComandoIrAPublicar`, `ComandoVerDetalle`, `ComandoVerProyecto`)

Los tres escriben las globales de contexto (`m_ObjProyectoActivo`, `m_ObjEdicionActivo`, y en el tercero `m_ObjProyectoAlInicio` y `m_EsAlta` en `:220-221`) y luego abren el formulario destino. Los dos últimos **cierran antes el destino si ya estaba abierto** (`:189-190`, `:222-223`) para reutilizar la instancia, siguiendo el mismo patrón de instancia única que el visor HTML.

### Error (`:52-59`, `:97-105`, `:134-144`, `:168-176`, `:200-208`, `:232-240`, `:268-276`)

Los siete manejadores son idénticos: `AvanceCerrar` solo en la rama de error, reloj de arena apagado, y según `Err.Number` compone un mensaje con número y detalle, **envía correo al administrador** y muestra un `MsgBox` crítico; si el error es `1000`, muestra solo un `MsgBox` de advertencia con `m_Error`.

## Reglas, validaciones y cálculos

- **Priorización completa es condición de generación.** Es la única regla de negocio del formulario, y vive solo en el botón de informe (`:69`), no en la carga ni en la generación del informe de publicabilidad.
- **Los avisos son de color, no de estado.** Rojo (`RGB(255,0,0)`) para días no positivos hasta la publicación (`:323`, `:41`) y para contrato vencido (`:341`). El resto hereda el color del campo de fecha (`:322`, `:338`, `:343`). En web el color no es un canal accesible: necesita un estado textual además del tono.
- **Los identificadores técnicos se ocultan por rol**, con una condición que admite dos caminos distintos de administración (`:287`), y se evalúa **una sola vez**, al cargar.
- **El cálculo de días se duplica** literal entre `EstablecerDatos` (`:318-324`) y el botón de recálculo (`:35-42`), con la única diferencia del `Else` quepone a `Null` el campo.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Edición activa | miembro público del formulario padre | `:252`, `:256` |
| Proyecto | `m_Edicion.Proyecto` | `:302` |
| Cálculo de fecha máxima | `Constructor.getProyecto` y `FechaMaxProximaPublicacionGrabar` | `:22`, `:26` |
| Recarga de la edición | `Constructor.getEdicion` | `:29` |
| Priorización de riesgos | `m_Edicion.TodosLosRiesgosPriorizados` | `:69`; `Edicion.cls:1817` |
| Informe de la edición | `GenerarInformePublicacion` | `Funciones Generales.bas:1409` |
| Informe de publicabilidad | `GenerarInformePublicabilidadEdicionInteractivoHTML` | `PublicabilidadEdicion.bas:781` |
| Ruta del informe activo | global `m_URLHTMLActivo` | `:127` |

## Coexistencia

Pendiente de `../coexistencia.md`. Los dos informes de este formulario se generan **como ficheros** y se abren con un visor local (`Form_FormWeb`) que depende de una variable global compartida. Es la misma cadena que atraviesa ese visor: generar, fijar la ruta en una global y confiar en que el visor este abierto. En una web, el informe tendría que servirse desde una ruta autenticada y el identificador de la edición viajaría en la URL, no en el estado de un formulario abierto.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-149.** El contexto entra por una **referencia directa a la instancia global del formulario padre** (`:252`), no por parámetro ni por propiedad pública propia. El formulario no es reutilizable: no se puede abrir desde ningún otro sitio sin replicar antes el árbol, y cualquier macro que intente abrirlo directamente obtiene «No se conoce la edición activa».
- **H-GR-150.** `Form_Open` no asigna `Cancel` y el manejador de error cierra el formulario por nombre con `DoCmd.Close acForm, Me.Name` (`:275`). Al estar incrustado como subformulario, cerrarse a sí mismo desde su propio evento de apertura deja el contenedor del árbol en un estado intermedio que solo se recompone al seleccionar otro nodo.
- **H-GR-151.** `EstablecerDatos` deshabilita `ComandoVerDetalle` al empezar (`:286`) y lo rehabilita al final (`:351`), sin protección: cualquier error intermedio deja el botón deshabilitado para siempre en esa instancia. El mismo patrón se repite en los cuatro manejadores de error sin reactivar el botón.
- **H-GR-152.** `ComandoGenerarInforme_Click` asigna la ruta a una variable local `url` (`:64`, `:82`) que **nunca se lee**, con la apertura comentada (`:88-89`). El formulario depende de un efecto secundario dentro de `GenerarInformePublicacion` (`Funciones Generales.bas:1409`) que no es visible desde aquí: si esa función deja de abrir, este camino deja de hacer algo sin ningún error.
- **H-GR-153.** El filtro de generación solo contempla `EnumSiNo.No` (`:69`). Cualquier otro valor del enumerado pasa el filtro, y el aviso distingue entre «Error» y «Advertencia» por el título del `MsgBox`, no por un código.
- **H-GR-154.** El aviso de fin de contrato compara `Me.FechaCierre <> ""` (`:337`) contra una cadena vacía, cuando el resto del procedimiento usa `IsDate` para lo mismo (`:329`). Con un valor nulo la comparación no es la esperada, así que el color del fin de contrato depende de un criterio distinto al de los demás campos.
- **H-GR-155.** `EstablecerDatos` **no limpia ningún campo** cuando la fecha no es válida: las cuatro comprobaciones con `IsDate` (`:304`, `:316`, `:326`, `:329`, `:333`) no tienen rama `Else`. El botón de recálculo sí lo hace (`:35-36`). Como el subformulario se reutiliza sin recargar (`Form_FormCalidadTareas.cls:46`, `:348`), al encadenar dos ediciones en la misma sesión las fechas no asignables de la segunda **heredan el valor de la primera**.
- **H-GR-156.** Los dos nombres de destino son casi idénticos y significan cosas distintas: `FormRiesgosGestion` (`:190`, `:193`) y `FormGestionRiesgos` (`:223`, `:225`). La diferencia de orden de las palabras es la única frontera entre el detalle de riesgos y la gestión de proyectos, y no hay ninguna señal en el código que la refuerce.
- **H-GR-157.** Los tres botones de navegación cambian **globales de contexto** y abren formularios no modales sin cerrar los anteriores ni dejar estado de retorno (`:155-160`, `:186-193`, `:218-225`). Si el usuario cierra la ventana de destino sin terminar, las globales quedan apuntando al último proyecto abierto.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pasar la edición por identificador en la ruta y carregar los datos en servidor, en lugar de leer el estado de otro formulario; devolver los informes desde una ruta autenticada en vez de generar ficheros y abrirlos con un visor local; reemplazar los avisos por color con estados textuales accesibles («plazo vencido», «contrato finalizado»); una única acción de navegación que cover proyecto, edición y riesgos en vez de tres formularios abiertos a la vez.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al elegir una edición en el árbol, el detalle muestra su proyecto, su edición y las fechas asociadas.
2. Solo un administrador ve los identificadores técnicos.
3. Un plazo vencido o un contrato finalizado se identifican sin depender solo del color.
4. No se puede generar el informe de una edición con riesgos sin priorizar.
5. Los tres botones de navegación llevan al destino correcto.

## Pendientes

- Qué escribe exactamente `FechaMaxProximaPublicacionGrabar` (`Proyecto.cls`) y por qué el cálculo obliga a releer la edición después (`:26-29`).
- Si `GenerarInformePublicacion` abre el fichero generado y con qué mecanismo, dado que aquí la ruta se descarta.
- Qué diferencia funcional hay entre `FormRiesgosGestion` y `FormGestionRiesgos`, y cuál de los dos es la pantalla de destino real.
- Si `EsAdministradorConectadoInicialmente` (`:287`) es un valor que puede diferir del rol con el que se está operando durante la sesión.

Árbol padre: `Form_FormCalidadTareas`, que publica la edición seleccionada e incrusta este formulario en su subformulario de detalle. Visor de informes: `Form_FormWeb`, que consume la global `m_URLHTMLActivo` escrita en `:127`. Catálogo completo en el [índice de formularios](../../forms.md).