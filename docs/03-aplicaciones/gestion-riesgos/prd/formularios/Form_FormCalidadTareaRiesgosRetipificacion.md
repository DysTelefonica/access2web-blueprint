# Form_FormCalidadTareaRiesgosRetipificacion

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadTareaRiesgosRetipificacion.cls` (199 líneas) y `src/forms/FormCalidadTareaRiesgosRetipificacion.form.txt` · **PRD:** [README](../README.md)

Hoja de detalle de la worklist de calidad que reúne los **riesgos pendientes de retipificar**, es decir, aquellos cuya clasificación debe ajustarse al catálogo de riesgos. Se muestra como subformulario embebido, no se abre de forma independiente. Es la única de las tres que **no abre ningún formulario de trabajo**: muestra el estado de la tarea y deja el criterio de retipificación al catálogo, al que llega por un evento que este formulario declara pero nunca conecta. Las citas son `Form_FormCalidadTareaRiesgosRetipificacion.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`. Toma el riesgo de la **variable pública de instancia** del formulario padre.

| Quién lo carga | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalidadTareas` | `Me.FormDetalle.SourceObject = "FormCalidadTareaRiesgosRetipificacion"` | `Form_FormCalidadTareas.cls:498` |
| `Form_FormCalidadTareas` (nodos) | `SeleccionarRiesgoParaRetipificar`, invocado desde `Arbol_NodeClick` | `Form_FormCalidadTareas.cls:125`, `:488-504` |

`SeleccionarRiesgoParaRetipificar` vacía antes el subformulario, copia `m_ObjSeleccionado` a `m_RiesgoSeleccionado` y guarda su estado en `m_EstadoRiesgoSeleccionado` (`Form_FormCalidadTareas.cls:492-497`). Los nodos que disparan la carga son `RIESGOSPENDIENTESRETIPIFICAR` y `RIESGOPENDIENTERETIPIFICAR` (`Form_FormCalidadTareas.cls:123`; `ArbolTareasCalidad.cls:436-448`).

Única salida: `FormRiesgo` (`:34`), cerrando antes la instancia previa (`:29-31`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Carga los datos y cierra el formulario si falla | `:134-162`, `.form.txt:1` |
| (formulario) | — | `m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido` | Recoge el riesgo de biblioteca elegido | `:164-197` |
| `ComandoVerRiesgo` | botón | `ComandoVerRiesgo_Click` | Abre `FormRiesgo` con el riesgo activo | `:20-52`, `.form.txt:1640` (rótulo «Ver riesgo») |
| `CodigoRiesgo`, `CodRiesgoBiblioteca` | cuadros de texto | — | Código propio del riesgo y código del catálogo | `:108-109`, `.form.txt:1080`, `:1018` |
| `TextoPendienteRetipificacion` | cuadro de texto | — | Aviso «Pendiente de Calidad», visible solo cuando hace falta | `:110-115`, `.form.txt:1049` |
| `Descripcion`, `CausaRaiz` | cuadros de texto | — | Descripción y causa raíz del riesgo | `:117-118`, `.form.txt:889`, `:986` |
| `FechaRiesgoParaRetipificar`, `DiasSinRetipificar` | cuadros de texto | — | Fecha en que entró en la lista y días sin respuesta | `:119-120`, `.form.txt:1111`, `:1145` |
| `Proyecto`, `Supervisor` | cuadros de texto | — | Nombre del proyecto y usuario de calidad | `:105-106`, `.form.txt:922`, `:955` |
| `IDExpediente`, `IDProyecto`, `IDEdicion`, `IDRiesgo` y sus etiquetas | varios | — | Identificadores, visibles solo para administrador | `:67-84`, `.form.txt:1389-1453` |

## Flujos

### Apertura (`Form_Load`, `:134-162`)

Usa `Form_Load`, no `Form_Open` como las otras dos hojas. Pone el reloj de arena (`:136-138`), llama a `EstablecerDatos` (`:143-146`) y, en error, avisa y **se cierra con `acSaveNo`** (`:161`). No hay ninguna salida por botón de cierre dentro de la hoja: el cierre es el de la ventana padre.

### Carga de datos (`EstablecerDatos`, `:54-132`)

1. Deshabilita el botón de ver riesgo (`:59`).
2. **Fija `AllowEdits` según el rol** (`:61-65`): técnico no edita, el resto sí. Es la única de las tres hojas que deja el formulario editable.
3. Muestra u oculta los cuatro identificadores según `EsAdministrador` o `EsAdministradorConectadoInicialmente` (`:67-84`).
4. Toma el riesgo del padre (`:86`) y **lo vuelve a leer de la base de datos** con `Constructor.getRiesgo` (`:91`), guardándolo en `m_ObjRiesgoAlInicio`. A partir de ahí, todos los campos se proyectan desde esa copia, no desde `m_Riesgo`.
5. Muestra el aviso «Pendiente de Calidad» si `RequiereRiesgoDeBibliotecaCalculado` es afirmativo; si no, lo oculta (`:110-115`).
6. Proyecta el resto de datos (`:117-120`) y rehabilita el botón (`:123`).

### Normal — ver la ficha del riesgo (`ComandoVerRiesgo_Click`, `:20-52`)

Marca `m_EsAlta = EnumSiNo.No` (`:32`), publica el riesgo en la global `m_ObjRiesgoActivo` (`:33`) y abre `FormRiesgo` (`:34`).

### Alternativo — elegir riesgo de catálogo (`m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido`, `:164-197`)

Ante un riesgo de biblioteca elegido: copia su código (`:174`), lo asigna también al objeto `m_Riesgo` (`:175`) y, según `RiesgoParaRetipificar`, o bien copia su descripción y **desbloquea y habilita el campo de descripción** (`:177-179`), o bien lo vacía y lo deja deshabilitado (`:180-183`). Es el único punto del formulario en el que se altera la edición de un campo de negocio.

### De error

Los tres manejadores repiten el bloque habitual: quitar el reloj y, **según `Err.Number`**, componer el mensaje, avisar por correo al administrador y mostrar un `MsgBox` crítico, o bien mostrar solo un `MsgBox` de advertencia (`:41-50`, `:152-160`, `:189-197`). El cierre del formulario solo ocurre en `Form_Load` (`:161`).

## Reglas, validaciones y cálculos

- **El formulario se valida a sí mismo con dos objetos.** `m_Riesgo` viene del padre (`:86`) y `m_ObjRiesgoAlInicio` se relee de la base de datos (`:91`). Los datos que se muestran proceden del segundo; si divergen, la hoja muestra el estado persistido, no el que el árbol tenía en memoria.
- **La edición depende del rol, no del dato.** El único control que se habilita es la descripción, y solo dentro del evento de catálogo (`:178`, `:182`). Para el usuario técnico, `AllowEdits = False` lo impide (`:62`).
- **El aviso «Pendiente de Calidad» es condicional y temporal.** Se escribe siempre con el mismo texto (`:112`) y no se limpia cuando deja de aplicar: solo se oculta (`:114`). Al reutilizarse el subformulario entre nodos, conviene verificar que ocultar y mostrar no dejan el valor anterior.
- **Los días sin retipificar siempre son los calculados**, nunca el valor almacenado (`:120`). En las otras dos hojas de riesgo hay una elección entre valor almacenado y calculado; aquí no la hay.
- **`EstablecerDatos` se declara `Private`** y devuelve `String` sin devolver nada (`:54`, `:126-127`): el resultado viaja en `ByRef p_Error`, igual que en el resto de la aplicación.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo seleccionado | `Form_FormCalidadTareas.m_RiesgoSeleccionado` | `:86` |
| Copia persistida del riesgo | `Constructor.getRiesgo(m_Riesgo.IDRiesgo, , , p_Error)` | `:91` |
| Datos mostrados | propiedades de `m_ObjRiesgoAlInicio` | `:101-121` |
| Aviso de retipificación pendiente | `RequiereRiesgoDeBibliotecaCalculado` | `:110` |
| Clasificación y descripción resultantes | propiedades de `RiesgoBiblioteca` | `:174-182` |
| Riesgo activo publicado | global `m_ObjRiesgoActivo` | `:33`; `Variables Globales.bas:80` |

La única escritura es en memoria: `m_Riesgo.CodRiesgoBiblioteca` (`:175`). La persistencia ocurre en otro punto de la aplicación.

## Coexistencia

No aporta coexistencia propia. Su dependencia característica es un evento recibido de otro formulario mediante `WithEvents` (`:11-13`): la hoja no pide nada, espera a que el catálogo le avise. En web esto se resuelve con un retorno de la acción, y es una de las razones por las que el flujo de retipificación necesita rediseño, no traducción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-130.** La instancia `m_FormRiesgosBibliotecaGestion` se declara con eventos (`:11-13`) y se escucha su evento (`:164-197`), pero **nunca se crea ni se asigna** en todo el módulo. Sin `Set`, el manejador no puede dispararse desde aquí, y el formulario no contiene ningún botón ni evento que abra el catálogo. Es decir, la rama que decide si la descripción se habilita o se vacía parece inalcanzable desde esta hoja; hay que confirmar si la retipificación se hace en otro formulario.
- **H-GR-131.** El único botón del formulario abre la ficha del riesgo (`:34`). Si la tarea se resuelve desde otra parte de la aplicación, esta hoja es solo informative y su posición dentro de la worklist no ofrece ninguna acción. Conviene verificar qué se espera del usuario al abrir una tarea de retipificación.
- **H-GR-132.** `EstablecerDatos` relee el riesgo desde la base de datos (`:91`) y muestra esa copia, pero el botón de ver riesgo publica **`m_Riesgo`**, la versión heredada del padre (`:33`), no `m_ObjRiesgoAlInicio`. Si ha habido un cambio entre la selección y la carga, la hoja muestra un estado y abre otro. Es el mismo nombre de variable con dos orígenes distintos.
- **H-GR-133.** `m_ObjRiesgoAlInicio` solo se usa como fuente de lectura (`:101-121`). El sufijo «AlInicio» sugiere una comparación con el estado posterior que no se realiza en ninguna parte: no hay validación de cambios ni de concurrencia.
- **H-GR-134.** `TextoPendienteRetipificacion` recibe el texto «Pendiente de Calidad» (`:112`) y, cuando el riesgo ya no lo requiere, solo se oculta (`:114`) sin vaciar. Como el subformulario se reutiliza para cada nodo seleccionado (`Form_FormCalidadTareas.cls:494`), conviene comprobar que un nodo posterior que sí lo requiere muestra el texto y no el valor heredado de otro caso.
- **H-GR-135.** `AllowEdits` se decide solo por `EsTecnico` (`:61-65`): no se tiene en cuenta que el usuario sea el supervisor del proyecto ni que la edición esté bloqueada por estado. Un usuario de calidad sin relación con el riesgo puede habilitar la descripción si llega al evento de `:177-179`.
- **H-GR-136.** `CodRiesgoBiblioteca` se muestra siempre (`:109`), incluso cuando el riesgo no requiere biblioteca y el valor debería estar vacío. La visibilidad del aviso sí es condicional (`:110-115`), pero el código de biblioteca no, de modo que la hoja puede mostrar una clasificación obsoleta sin ninguna señal de que no es aplicable.
- **H-GR-137.** El evento de `:164-197` escribe en `m_Riesgo`, no en `m_ObjRiesgoAlInicio` (`:175`), lo que rompe la simetría de las dos variables descritas en **H-GR-132** y deja sin origen conocido cuál de las dos es la buena.
- **H-GR-138.** `Form_Load` no acepta `Cancel` ni comprueba nada al abrir, y su manejador de error cierra con `acSaveNo` (`:161`). Si el formulario se reutiliza y hay cambios sin guardar, se descartan sin aviso.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: resolver la retipificación dentro de la propia hoja, con el catálogo como servicio y un retorno explícito en lugar de un evento de formulario; mantener la hoja como informativo y sacar la acción a la ficha del riesgo; comparar el estado de la copia con el del árbol para avisar de cambios concurrentes; condicionar también el campo de código de biblioteca, no solo su aviso.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al elegir un riesgo pendiente de retipificar, la hoja muestra código propio, código de catálogo, causa raíz, fecha y días sin retipificar.
2. Un usuario técnico ve la hoja sin posibilidad de editar.
3. Cuando el riesgo requiere clasificación de catálogo, se indica de forma visible que la decisión está pendiente de calidad.
4. El botón de ver riesgo abre la ficha del riesgo que la hoja está mostrando.
5. Un usuario no administrador no ve los identificadores.
6. Si no se puede determinar el riesgo, se avisa y la hoja se cierra sin dejar el formulario a medias.

## Pendientes

- Dónde se abre realmente el catálogo de riesgos durante una tarea de retipificación y si el evento de `:164` llega a dispararse alguna vez.
- Si existe un guard de guardado en `FormRiesgo` o en el catálogo que recargue la worklist al cambiar una clasificación.
- Qué calcula `RequiereRiesgoDeBibliotecaCalculado` y con qué frecuencia.
- Si `DiasSinRespuestaCalidadRetipificacionCalculado` incluye días de calendario o días laborables, y desde qué fecha cuenta.