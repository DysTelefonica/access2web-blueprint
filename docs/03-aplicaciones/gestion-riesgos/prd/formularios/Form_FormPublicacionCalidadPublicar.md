# Form_FormPublicacionCalidadPublicar

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionCalidadPublicar.cls` (472 líneas) y `src/forms/FormPublicacionCalidadPublicar.form.txt` (1519 líneas) · **PRD:** [README](../README.md)

Diálogo de publicación. Resume el estado de la propuesta, muestra si las dos evidencias exigidas están complementadas y concentra las cuatro acciones posibles: avisar al técnico, publicar, rechazar y revisar historiales. Es la pantalla que decide. Las citas son `Form_FormPublicacionCalidadPublicar.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormPublicacionCalidad` | Se le fija como destino del panel lateral y como fuente del subformulario | `Form_FormPublicacionCalidad.cls:139-140` |

La edición sobre la que opera se toma de `m_ObjEdicionActiva` (`Variables Globales.bas:79`). Si no hay ninguna, `Form_Open` avisa y no continúa (`:274-276`).

Salidas: abre `FormPublicacionCalidadPublicarEjecutar` (`:187-188`) y `FormPublicacionCalidadRechazoObservaciones` (`:163-165, 236-238`); cierra `FormPublicacionCalidad` tras un rechazo (`:422-424`) y se cierra a sí misma cuando el diálogo de publicación confirma que se publicó (`:457`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Valida que hay edición activa y compagina toda la pantalla | `:263-299`, `.form.txt:40` |
| `lbl1`, `lbl2` | etiquetas | — | Rótulos de las dos evidencias exigidas: acuerdo con UTE y solicitud de riesgos a suministradores | `.form.txt:628-653` |
| `lbl1Requerido`, `lbl1Complementado` | etiquetas | — | Estado «Requerido / No Requerido» y «Complementado / No complementado / N/A» de la primera evidencia | `:311-330`, `.form.txt:676-701` |
| `lbl2Requerido`, `lbl2Complementado` | etiquetas | — | Lo mismo para la segunda evidencia | `:331-345`, `.form.txt:724-749` |
| `lblPublicacion` | etiqueta | — | Estado global de la propuesta, con color | `:344-380`, `.form.txt:959-960` |
| `ComandoEnviarCorreoTecnicoPropuestaPublicacion` | botón | `ComandoEnviarCorreoTecnicoPropuestaPublicacion_Click` | Avisa a los responsables para que revisen la edición | `:23-65`, `.form.txt:984` |
| `ComandoGenerarInformePublicabilidad` | botón | `ComandoGenerarInformePublicabilidad_Click` | Genera el informe de publicabilidad y recuerda su URL | `:67-95`, `.form.txt:810` |
| `ComandoHistorialPublicaciones` | botón | `ComandoHistorialPublicaciones_Click` | Abre el historial de publicaciones en una ventana HTML | `:97-124`, `.form.txt:772` |
| `ComandoVerCorreosEnviados` | botón | `ComandoVerCorreosEnviados_Click` | Abre el historial de correos a responsables | `:126-150` |
| `ComandoVerRechazo` | botón | `ComandoVerRechazo_Click` | Abre el formulario de motivos de rechazo | `:152-180`, `.form.txt:924` |
| `ComandoPublicar` | botón | `ComandoPublicar_Click` | Abre el diálogo de ejecución de la publicación | `:182-216`, `.form.txt:886` |
| `ComandoRechazar` | botón | `ComandoRechazar_Click` | Pide motivos y abre el formulario de rechazo | `:218-261`, `.form.txt:848` |

El diseño conserva controles heredados sin código en esta clase (etiqueta, rectángulo, imagen, grupo de opciones, lista, desplegable, subformulario y marco de objeto sin enlazar: `.form.txt:419-605`).

## Flujos

### Apertura (`Form_Open`, `:263-299`)

1. Si no hay edición activa, mensaje «No hay ninguna edición activa» y `Err.Raise 1000` (`:274-276`).
2. `EstablecerDatos` (`:301-394`) y, si devuelve error, `Err.Raise 1000` (`:283-286`).
3. Rótulo de la ventana con el usuario conectado (`:287`).

### Estado de las evidencias (`EstablecerDatos`, `:311-345`)

- **Evidencia de acuerdo UTE** (`:312-329`): si el proyecto la exige, la etiqueta de requerida dice «Requerido» y la de complementado dice «Complementado» (verde) o «No complementado» (rojo) según `TieneAnexoEvidenciaUTE`. Si no se exige: «No Requerido» y «N/A», **siempre en verde** (`:325-326`).
- **Evidencias de suministradores** (`:330-345`): misma estructura con `EvidenciasSuministradoresRequeridas` y `EvidenciasSuministradoresCompletadas`.

### Estado de la propuesta (`:339-380`)

- **Edición activa** (`:339-360`): se habilitan rechazar, publicar, generar informe y enviar correo; el botón de publicar rotula «Publicar».
  - Sin fecha de propuesta: «SIN PROPUESTA TÉCNICA» en rojo (`:346-347`).
  - Con propuesta y sin fecha de rechazo: «PROPUESTA (fecha) ESPERANDO A CALIDAD», y el botón de ver motivos se oculta (`:350-354`).
  - Con propuesta y con fecha de rechazo: «PROPUESTA (fecha) RECHAZADA (fecha)», y el botón de ver motivos aparece (`:355-358`).
- **Edición no activa** (`:362-380`): se deshabilitan todas las acciones, el botón rotula «Publicado» y la etiqueta global muestra la publicación si hay fecha (`:369-378`).

### Normal: enviar correo al técnico (`:23-65`)

Confirmación con `MsgBox` de sí/no (`:29`). Si se acepta, crea una `EdicionCorreoRevision`, la registra (`:37-40`), **recarga la edición activa** desde el constructor (`:44`) y cierra el avance con `AvanceCerrar` (`Funciones Generales.bas:4920`, `:49`). Termina con un mensaje de confirmación (`:53`).

### Normal: generar informe (`:67-95`)

Llama a `GenerarInformePublicabilidadEdicionInteractivoHTML` (`PublicabilidadEdicion.bas:781`) pasando el hwnd del formulario y guarda la URL devuelta en la global `m_URLHTMLActivo` (`Variables Globales.bas:124`, `:76-80`). No abre nada ni avisa: el resultado se ve por el mecanismo de avance.

### Normal: ver historiales (`:97-124, 126-150`)

Ambas acciones entregan a `HTMLENTXT` (`Funciones Generales.bas:3438`) el HTML que construye la edición (`Edicion.cls:529` y `Edicion.cls:579`) y lo abre fuera de Access. El segundo argumento se omite; el tercero es `EnumSiNo.No`.

### Alternativo: rechazar (`:218-261`)

1. Si la propuesta ya está rechazada (`PublicacionRechazada`, `:224`), pregunta si quiere ver las observaciones ya registradas (`:225-230`).
2. Si no, pide confirmación antes de abrir el formulario de motivos (`:231-235`).
3. Cierra el formulario de motivos si ya estaba abierto y lo abre (`:236-238`).

### Alternativo: publicar (`:182-216`)

Cierra el diálogo de ejecución si estaba abierto, lo abre de nuevo y, si sigue abierto, guarda la referencia en la variable `WithEvents` (`:186-192`).

### De error: rechazo consumado (`:396-447`)

`m_FormMotivos_Motivado` recibe el motivo y, si no está vacío (`:405-409`), lo escribe en `PropuestaRechazadaPorCalidadMotivo` y llama a `RechazoPropuestaParaPublicacion` (`Edicion.cls:3296`, `:410-415`). Avisa de que el correo llegó a los responsables (`:420`), cierra `FormPublicacionCalidad` si está abierta (`:422-424`) y, si el formulario de gestión de riesgos tiene cargada la edición como subformulario, refresca su etiqueta de rechazo con `EstablecerlblRechazado` (`Funciones Generales.bas:8058`, `:425-431`).

### De error: publicación consumada (`:449-471`)

`m_FormPublicar_Publicado` se limita a cerrar este formulario (`:457`). La recarga del estado de la pantalla que eso implica no está en este código.

## Reglas, validaciones y cálculos

- **El permiso no se comprueba aquí.** El código usa `EsActivo` como única condición para habilitar acciones (`:339, 362`); el permiso de escritura global lo calcula el contenedor (`Form_FormPublicacionCalidad.cls:83-91`).
- **«N/A» se pinta en verde** en ambos pares de etiquetas (`:326, 329, 343`), igual que «Complementado»: el color comunica estado, no disponibilidad.
- **`PublicacionRechazada` no es un campo.** Es una propiedad calculada de `Edicion` (`Edicion.cls:1636`) que depende de que la fecha de rechazo esté vacía.
- **El color es la única señal de estado** para el usuario: los rótulos «Requerido» y «No requerido» no cambian de estilo.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Estado completo de la propuesta | `m_ObjEdicionActiva` (`Variables Globales.bas:79`) | `:311-380` |
| Envío de correo a responsables | `EdicionCorreoRevision.Registrar` | `:37-40` |
| Recarga de la edición | `Constructor.getEdicion` | `:44` |
| Rechazo de la propuesta | `Edicion.RechazoPropuestaParaPublicacion` (`Edicion.cls:3296`) | `:410-415` |
| Informe de publicabilidad | `PublicabilidadEdicion.bas:781`; URL en `m_URLHTMLActivo` | `:76-80` |
| Historiales en HTML | `Edicion.cls:529, 579` y `HTMLENTXT` (`Funciones Generales.bas:3438`) | `:106, 134` |
| Colores de estado | `m_ColorComplementado`, `m_ColorNoComplementado` (globales de módulo) | `:316, 319, 352, 356` |

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí ya se ve: la pantalla es un coordinador de diálogos y de ventanas HTML externas, con estado repartido entre la sesión global, la clase de edición y tres formularios hermanos; en web sería una vista con acciones asíncronas y un recurso de estado de publicación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-31.** El diálogo de publicación se abre sin confirmación ni comprobación de que las evidencias estén complementadas (`:182-192`): `EstablecerDatos` calcula y pinta el estado de las dos evidencias (`:311-345`), pero ningún camino impide publicar con «No complementado». La validación, si existe, tiene que estar en `FormPublicacionCalidadPublicarEjecutar`.
- **H-GR-32.** Tras `m_FormPublicar_Publicado` solo se cierra el formulario (`:457`): ni se recarga la edición ni se repinta el estado. Si el usuario cierra después el contenedor, los rótulos pueden seguir mostrando «PROPUESTA ESPERANDO A CALIDAD» en lugar de la publicación ya realizada.
- **H-GR-33.** El mensaje de consulta al ver un rechazo ya registrado dice «¿Desea ver las observaciones registradaSí» (`:226`): le falta el signo de interrogación de cierre y la palabra está mal construida. Es un texto visible para el usuario final.
- **H-GR-34.** Tras enviar el correo a los responsables se recarga la edición activa (`:44`), pero `EstablecerDatos` **no se vuelve a llamar**: la pantalla no refleja el nuevo estado hasta que se reabra. La recarga solo sirve para la sesión.
- **H-GR-35.** `ComandoGenerarInformePublicabilidad` guarda la URL en la global `m_URLHTMLActivo` (`:80`) y no muestra ningún mensaje al usuario: el resultado solo es visible a través del mecanismo de avance, lo que deja la acción sin confirmación visible.
- **H-GR-36.** Los dos handlers de historial entregan HTML a `HTMLENTXT` (`:106, 134`) sin comprobar el valor devuelto ni el error: si la generación falla, el usuario no recibe ningún aviso desde esta pantalla.
- **H-GR-37.** `Private WithEvents m_FormAnexos As Form_FormAnexos` (`:16`) se declara y nunca se asigna: la referencia queda siempre vacía. Conviene confirmar que no falta un cableado de eventos.
- **H-GR-38.** `Form_Open` (`:263-299`) y `ComandoPublicar_Click` (`:182-216`) no ponen el reloj de arena en la rama normal, a diferencia del resto de handlers: la apertura y el cierre del diálogo no muestran espera.
- **H-GR-39.** El rechazo ofrece dos intenciones indistintas: si ya está rechazada, el botón muestra los motivos (`:224-230`); si no, pide rellenarlos (`:231-235`). Ambos caminos abren el mismo formulario y el rótulo del botón siempre dice «Rechazar» (`.form.txt:848`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único endpoint de estado de publicación que exponga evidencias, propuesta, fechas y motivo; acciones asíncronas con realimentación inmediata en lugar de diálogos HTML externos; historiales como recursos consultables y paginados; estados explícitos en lugar de rótulos con color.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla muestra, para cada evidencia exigida, si está requerida y si está complementada.
2. Con la edición activa se pueden proponer, publicar y rechazar; con la edición no activa, ninguna de esas acciones está disponible.
3. El estado global de la propuesta y sus fechas son visibles y su color corresponde al estado.
4. Enviar el correo a los responsables pide confirmación y avisa del resultado.
5. Rechazar exige motivos y deja constancia de la fecha y del motivo.
6. Publicar deja constancia en el historial de publicaciones.

## Pendientes

- Si `FormPublicacionCalidadPublicarEjecutar` bloquea la publicación cuando las evidencias no están complementadas (**H-GR-31**).
- Si la recarga de la edición tras el correo tiene efecto observable en otra pantalla (**H-GR-34**).
- Para qué se declara `m_FormAnexos` sin usar (**H-GR-37**).
- Dónde se recarga la edición tras publicar, si no es en este formulario (**H-GR-32**).