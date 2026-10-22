# Form_FormPublicacionCalidadPublicarEjecutar

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionCalidadPublicarEjecutar.cls` (324 líneas) y `src/forms/FormPublicacionCalidadPublicarEjecutar.form.txt` · **PRD:** [README](../README.md)

Diálogo modal que ejecuta la publicación de una edición. Es la pantalla donde el usuario de calidad confirma el envío al RAC, decide si la edición es la última, fija la fecha de la próxima publicación y pulsa el botón que genera el informe y lanza el aviso por correo. Al terminar emite el evento `Publicado`, que el formulario padre escucha para cerrarse. Las citas son `Form_FormPublicacionCalidadPublicarEjecutar.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPublicacionCalidadPublicar` | `DoCmd.OpenForm`, tras cerrar cualquier instancia anterior; el padre se engancha al evento `Publicado` con `WithEvents` | `Form_FormPublicacionCalidadPublicar.cls:18, 191-196` |

No acepta parámetros ni declara `TempVars`: lee el contexto de los objetos públicos de sesión (`m_ObjEdicionActiva`, `m_ObjEntorno`, `m_ObjProyectoActivo`), igual que el resto de la cadena. Salidas: cierra `FormRiesgosGestion` si está abierto, cierra `FormPublicacionCalidad` y se cierra a sí misma (`:188, 191, 193`). Opcionalmente abre el informe generado con el visor del sistema mediante `Ejecutar Me.hWnd, "open", …` (`:169`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Tamaño, título, fechas del contrato, correo del RAC y fecha máxima de publicación | `:265-320`, `.form.txt:1495` |
| (formulario) | — | `ComandoPublicar_Click` | Publica, refresca el proyecto, recalcula tareas, avisa por correo y cierra la cadena | `:95-204`, `.form.txt:859` |
| (formulario) | — | `ComandoGenerarInforme_Click` | Valida priorización y abre el informe previo | `:45-92`, `.form.txt:797` |
| (formulario) | — | `ComandoSalir_Click` | Cierra sin publicar | `:207-221`, `.form.txt:1121` |
| (formulario) | — | `ComandoAyuda_Click` | Abre la ayuda del formulario | `:17-42` |
| (formulario) | — | `EstablecerFechaMax` (procedimiento privado) | Recalcula la fecha máxima de la próxima publicación | `:224-258` |
| `ComboUltimaEdicion` | combo | `ComboUltimaEdicion_AfterUpdate` | Llama a `EstablecerFechaMax` | `:261-263`, `.form.txt:735` (`RowSource "Sí;No"`) |
| `CorreoRAC` | texto | — | Destinatario opcional del RAC; se persiste si difiere del proyecto | `:113-121`, `.form.txt:768` |
| `FechaMaxProximaPublicacion` | texto | — | Fecha calculada por `CalcularFechaMaxProximaPublicacion` | `:236-242`, `.form.txt:1018` |
| `FechaInicioContrato`, `FechaFinContrato` | texto | — | Rango del expediente, solo informativo | `:293-298`, `.form.txt:927, 957` |
| `cboTipoDocumento` | combo | — | Selector `HTML;PDF` del informe previo | `.form.txt:1083` |

## Flujos

### Apertura (`Form_Load`, `:265-320`)

1. `AjustarTamaño Me` y `Me.Caption = m_ObjEntorno.TituloUsuarioConectado` (`:272-274`).
2. Si `m_ObjEdicionActiva Is Nothing`, escribe el mensaje «El formulario se ha abierto con parámetros insuficientes» y lanza el error (`:276-279`).
3. Lee la fecha prevista de cierre; si faltan menos de quince días marca «Sí» en el combo (`:283-289`) y, a continuación, **sobrescribe ese valor con «No»** (`:292`).
4. Copia las fechas del contrato si son fechas válidas (`:293-298`).
5. Rellena `CorreoRAC` con el valor calculado del proyecto si no viene vacío (`:299-301`).
6. Llama a `EstablecerFechaMax` (`:302`).

### Publicación (`ComandoPublicar_Click`, `:95-204`)

1. Pregunta si el correo debe enviarse también al RAC, con «No» como botón por defecto (`:107`). Si se acepta y el campo está vacío, avisa y **aborta sin publicar** (`:109-112`).
2. Si el valor del campo difiere del que tiene el proyecto, lo registra con `Proyecto.RegistrarCorreoRAC` (`:113-121`; `src/classes/Proyecto.cls:4097`).
3. Si el combo marca «Sí», usa la fecha del sistema como fecha de cierre (`:127-129`).
4. `Publicar` genera el documento y devuelve la ruta del informe en la variable global `m_URLInforme` (`:134-139`; `src/modules/Funciones Generales.bas:5002`).
5. Recarga el proyecto y recalcula las tareas de calidad (`:143-147`).
6. Si el proyecto tiene activated avisos, envía el correo de nueva publicación (`:148-156`; `Funciones Generales.bas:2024`).
7. Si el fichero existe, ofrece ver el informe y, en ese caso, guarda la ruta en la variable global `m_URLHTMLActivo` y la abre con el visor (`:157-171`).
8. Lanza `RaiseEvent Publicado` (`:181`), tras lo cual el padre cierra su propio formulario, y el código cierra los formularios abiertos y se cierra a sí mismo (`:188-193`).

### Informe previo (`ComandoGenerarInforme_Click`, `:45-92`)

Rechaza la operación si la edición no tiene todos sus riesgos priorizados (`:51-54`). Calcula `bGenerarPDF` a partir del combo (`:58-62`), llama a `GenerarInformePublicacion` (`:68`; `Funciones Generales.bas:1409`) y abre la ruta devuelta con el visor (`:75`).

### Alternativo (`ComandoSalir_Click`, `:207-221`)

Cierra sin publicar y sin escribir nada. Es la vía de escape del diálogo.

### Error

Los seis manejadores repiten el mismo patrón: si `Err.Number` no es `1000`, se compone un mensaje con número y descripción, se envía por `CorreoAlAdministrador` y se muestra crítico; si es `1000`, solo se muestra el texto que el propio formulario ya dejó en `m_Error` (`:36-41, 88-91, 199-203, 216-220, 252-256, 314-318`). El de `Form_Load` cierra además el formulario (`:320`).

## Reglas, validaciones y cálculos

- **Priorización obligatoria para el informe previo.** `TodosLosRiesgosPriorizados = EnumSiNo.No` bloquea la generación (`:51`), pero **no** bloquea la publicación: `ComandoPublicar_Click` no repite la comprobación.
- **El envío al RAC es opt-in y exige destinatario.** Sin valor en el campo, la publicación se cancela (`:109-112`).
- **La fecha máxima se recalcula, no se fija.** `CalcularFechaMaxProximaPublicacion` recibe la edición, la fecha del sistema y el combo de última edición, y devuelve una fecha o cadena vacía (`:232-242`; `Funciones Generales.bas:8650`).
- **El estado de salida es global.** `m_URLInforme` y `m_URLHTMLActivo` son variables de módulo públicas, no propiedades del formulario (`src/modules/Variables Globales.bas:123-124`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Edición activa y proyecto | Objetos de sesión | `:271, 108` |
| Fechas del expediente y correo del RAC | `Proyecto.Expediente`, `Proyecto.CorreoRACCalculado` | `:281, 293-301` |
| Documento, fecha de cierre y registro de publicación | `Publicar` | `:134-139` |
| Correo del RAC | `Proyecto.RegistrarCorreoRAC` | `:113-121` |
| Tareas de calidad del usuario | `EstablecerTareasCalidad` | `:144` |
| Aviso de nueva publicación | `EnvioCorreoNuevaPublicacion` | `:148-156` |
| Fichero del informe | Ruta devuelta por `Publicar`, comprobada con `fso.FileExists` | `:157` |

## Coexistencia

Pendiente de `../coexistencia.md`. Dos dependencias que no tienen equivalente directo: el resultado se devuelve por **variables globales** en lugar de por retorno o por el formulario, y la cadena se comunica mediante el evento `Publicado`, que el padre debe haber enganchado **antes** de abrir este formulario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **El cálculo de «última edición» es código muerto.** `Form_Load` decide «Sí» si faltan quince días o menos para el cierre previsto (`:283-289`) y en la línea siguiente fuerza «No» (`:292`). La rama calculada nunca llega a verse, y la etiqueta «4) Es última publicación» solo refleja el valor forzado.
- **`m_FechaRef` se declara y se pasa, pero nunca se asigna.** `:97` la declara, `:139` la entrega a `Publicar` y queda siempre vacía; justo encima hay una asignación comentada con una fecha fija de 2025 (`:129`), señal de un valor en depuración que quedó en el código.
- **El selector de formato del informe no hace nada.** El combo se lee en `:58-62` para calcular `bGenerarPDF`, pero la llamada de `:68` pasa siempre `EnumTipoInformePublicacion.html`; la variable calculada no se usa en ninguna otra línea.
- **El comentario sobre la doble apertura está sin resolver.** `:73-74` avisa de que el documento puede abrirse dos veces y propone «comentar o borrar la siguiente línea»; el código ejecuta igualmente la apertura (`:75`, `:169`), así que el duplicado sigue presente salvo que la función interna lo evite.
- **`Err.Raise 1000` se usa como control de flujo propio.** Todas las validaciones internas lanzan ese número y los seis manejadores distinguen `Err.Number <> 1000` para decidir si hay que notificar al administrador. Si Access reasigna el número, un fallo de negocio acaba notificado como error técnico.
- **La publicación no comprueba la priorización de riesgos.** `ComandoGenerarInforme_Click` la exige (`:51`) y `ComandoPublicar_Click` no; un usuario que publique sin pasar por el informe previo puede generar un documento con riesgos sin priorizar.
- **El fallo de apertura deja al usuario sin salida.** El manejador de `Form_Load` avisa y cierra el formulario (`:314-320`); si la causa es que falta el contexto de sesión, no queda ninguna acción posible dentro de la propia aplicación.
- **El cierre encadena varios formularios sin coordinación.** Tras `RaiseEvent Publicado` el código cierra `FormRiesgosGestion`, `FormPublicacionCalidad` y sí mismo (`:186-193`), mientras el padre también se cierra al recibir el evento (`Form_FormPublicacionCalidadPublicar.cls:449-454`). Hay dos rutas de cierre para la misma ventana.
- **El resultado se propaga por estado global.** `m_URLInforme` y `m_URLHTMLActivo` (`src/modules/Variables Globales.bas:123-124`) se sobrescriben en cada publicación; dos sesiones abiertas en el mismo proceso de Access se pisarían entre sí.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: devolver el identificador del documento generado y su ruta como resultado explícito de la operación en lugar de por variables globales; resolver la fecha de referencia como parámetro de la publicación en vez de una cadena vacía; sustituir el salto de `DateDiff` y el posterior borrado por una regla única y visible; validar la priorización en el servidor como precondición única, tanto para el informe previo como para la publicación; y reemplazar el doble `MsgBox` de confirmación por un diálogo con estado explícito.

## Criterios de aceptación

Propuestos, no aprobados:

1. Publicar exige confirmación y avisa del resultado, incluidos los avisos por correo.
2. Sin destinatario del RAC, la publicación se cancela con un mensaje que lo explique.
3. La fecha de la próxima publicación se recalcula al cambiar la marca de última edición.
4. El informe solo se genera si todos los riesgos están priorizados.
5. Al terminar, la cadena de formularios queda cerrada y la edición queda publicada.

## Pendientes

- Qué hace exactamente `Publicar` con `p_FechaRef` cuando llega vacía (`Funciones Generales.bas:5002`) y si esa es la razón de la asignación comentada.
- Si el selector `HTML;PDF` llegó a tener efecto en alguna versión anterior, o si es un resto de la interfaz.
- Si `AvanceCerrar` (`:155`) y la barra de progreso que abre `Form_Load` son del mismo mecanismo de avance.
- Qué ocurre con las tareas de calidad recalculadas en `:144` cuando la publicación falla después.