# Form_FormPublicacionCalidad

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionCalidad.cls` (148 líneas) y `src/forms/FormPublicacionCalidad.form.txt` (1000 líneas) · **PRD:** [README](../README.md)

Contenedor principal de la cadena de publicación de calidad. No pinta datos: prepara el contexto de la edición activa, decide qué páginas del panel lateral quedan disponibles y hospeda un subformulario. Todas las citas son `Form_FormPublicacionCalidad.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| El formulario de gestión de riesgos | Se abre como formulario con subformulario de una edición concreta; no lee argumentos | `Form_FormPublicacionCalidad.cls:140` |

La edición sobre la que trabaja no llega como parámetro: se toma de la variable global `m_ObjEdicionActiva` (`Variables Globales.bas:79`), que el formulario de gestión ha dejado establecida. Si esa variable está vacía, el formulario se cierra (`:69-72`).

Salidas: **no abre ningún formulario**. El único botón de salida, `ComandoSalir`, cierra el propio (`:47-49`). La navegación hacia otras páginas se hace con `DoCmd.OpenForm` desde los botones de navegación guardados en el diseño (`.form.txt:428, 481, 534, 588, 641`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Todo el posicionamiento, el control de escritura y la habilitación de páginas | `:51-127`, `.form.txt` sin evento de carga declarado |
| (formulario) | — | — | Fija el destino del panel lateral al subformulario de publicación | `:139-140` |
| `lblTitulo` | etiqueta | — | Proyecto, número de edición y marca de edición activa o no | `:81`, `.form.txt:306` |
| `NavCtrl` | control de navegación | — | Panel lateral con seis páginas: Publicación, Datos generales, Acuerdo UTE, Suministradores, Notas Calidad y Agregar nuevo | `.form.txt:384`, `:412-678` |
| `NavUTE`, `NavAcuerdoSuministradores` | botón de navegación | — | Se habilitan o deshabilitan según lo que exige la edición | `:96-105` |
| `NavNotasCalidad` | botón de navegación | — | Abre `FormPublicacionCalidadNotas`; su habilitación no se toca en código | `.form.txt:624-641` |
| `FormDetalle` | subformulario | — | Aloja la página activa; en la apertura se le fija `FormPublicacionCalidadPublicar` | `:140`, `.form.txt:721-728` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` | Abre la ayuda | `:19-45`, `.form.txt:331` |
| `ComandoSalir` | botón | `ComandoSalir_Click` | Cierra el formulario | `:47-49`, `.form.txt:764` |
| `lblEstado` | etiqueta | — | Texto de estado; en el diseño aparece con el contenido de la variable de URL de ayuda | `.form.txt:835-836` |

## Flujos

### Apertura (`Form_Load`, `:51-127`)

1. Poner el reloj de arena y limpiar `m_Error` (`:57-61`).
2. **Tampona dos globales** antes de comprobar nada: `m_ObjAnexoEvicenciaUTE = Nothing` y `blnPermitidoEscribir = False` (`:67-68`).
3. Fija el tamaño del formulario en twips (`:64-65`).
4. Si no hay edición activa, mensaje y cierre (`:69-72`).
5. Con la edición activa (`:74-106`):
   - `EsActivo` marca la edición como activa o no y compone la partícula del título (`:75-80`).
   - Compone el rótulo con proyecto y número de edición (`:81`).
   - `blnPermitidoEscribir` depende de dos condiciones: edición activa **y** usuario conectado autorizado (`:83-91`).
   - Habilita la página de acuerdo UTE y carga el anexo de evidencia UTE cuando el proyecto lo pide; si no, la deja deshabilitada (`:95-100`).
   - Habilita la página de suministradores cuando la edición exige evidencias de suministradores (`:101-105`).
6. `EstablecerDatos` y, si falla, mensaje y cierre (`:108-111`).
7. Rótulo de la ventana con el usuario conectado (`:112`).

El manejador de error cierra siempre el formulario (`:117-127`).

### Normal (`ComandoAyuda_Click`, `:19-45`)

Poner el reloj, llamar a `AbrirAyuda` (`Funciones Generales.bas:5327`) y, si devuelve error, `Err.Raise 1000`. En caso contrario, quitar el reloj y salir. El error inesperado se notifica al administrador antes de mostrarse (`:38-40`).

### Alternativo: `EstablecerDatos` (`:129-144`)

Solo dos líneas efectivas: fija el destino del panel lateral y la fuente del subformulario a `FormPublicacionCalidadPublicar` (`:139-140`). El resto del cuerpo está vacío.

### De error

Los tres manejadores comparten el mismo patrón: si el número de error no es 1000, se compone el texto con número y descripción, se envía al administrador y se muestra crítico; si es 1000, se muestra como advertencia sin notificar (`:36-44, 117-126, 148-155` en el orden del fichero). El cierre del formulario solo ocurre en `Form_Load`.

## Reglas, validaciones y cálculos

- **El permiso de escritura es una global, no una propiedad del formulario.** `blnPermitidoEscribir` (`Variables Globales.bas:121`) se calcula en `:83-91` y se pone a `False` al entrar (`:68`), incluso antes de saber si hay edición activa.
- **La edición activa se comprueba dos veces**, una para el título y otra para el permiso: si la edición no está activa, el usuario autorizado no puede escribir (`:83-91`).
- **El proyecto decide dos páginas.** Un proyecto con acuerdo UTE y una edición que exija evidencias de suministradores habilitan páginas distintas; el resto de páginas no se condiciona.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Edición activa | `m_ObjEdicionActiva` (`Variables Globales.bas:79`) | `:69-106` |
| Anexo de evidencia UTE | `m_ObjAnexoEvicenciaUTE`, propiedad `AnexoEvidenciaUTE` | `:67, 98` |
| Permiso de escritura | `blnPermitidoEscribir` (`Variables Globales.bas:121`) | `:67-91` |
| URL de informe en curso | `m_URLInforme` (`Variables Globales.bas:123`), puesta a cadena vacía | `:69` |
| Usuario conectado | `m_ObjEntorno.TituloUsuarioConectado` | `:112` |

No escribe en base de datos: es puro andamiaje de navegación.

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí ya se ve: el formulario depende de `m_ObjEdicionActiva`, una sesión global compartida, y de `FormularioAbierto` (`Funciones Generales.bas:832`), que en la web no tiene equivalente. El panel lateral con subformulario anidado es un patrón de Access sin traducción directa: en web serían rutas o pestañas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-19.** El permiso de escritura se decide con una comprobación anidada que se puede leer al revés: la rama `Else` de `:83-84` vuelve a poner `blnPermitidoEscribir = False` y la rama de edición activa vuelve a comprobar `UsuarioConectadoAutorizado` (`:86-91`). La lógica es correcta, pero el valor depende de una global modificada también por otras pantallas; conviene verificar que ningún otro formulario la pise durante la vida de esta pantalla.
- **H-GR-20.** `Form_Load` fija `Me.InsideHeight` e `Me.InsideWidth` en twips (`:64-65`) sin relación con ninguna propiedad del diseño. Es una restricción de tamaño embebida en el código y no en el formulario.
- **H-GR-21.** El destino del panel lateral se fija siempre a la página de publicación (`:139-140`), de modo que **el formulario no recuerda en qué página estaba el usuario**; cada reapertura vuelve a la primera.
- **H-GR-22.** `lblEstado` se declara en el diseño con el contenido de la variable global de URL de ayuda (`.form.txt:835-836`), pero ningún código de este formulario la escribe. Queda como rótulo muerto o como dependencia no trazada desde otra pantalla.
- **H-GR-23.** La etiqueta `Tag = "OBLIGATORIO"` que aparece en la página de notas (`.form.txt:967` de la ficha de notas) no se comprueba aquí: el contenedor no valida que las páginas obligatorias estén cumplidas antes de dejar publicar. Conviene verificar si la validación vive en la ejecución de la publicación.
- **H-GR-24.** `EstablecerDatos` (`:129-144`) declara `Optional ByRef p_Error As String` y devuelve `As String`, pero no devuelve nada: el contrato de error se cumple solo por el segundo argumento. Es una firma copiada de otros formularios.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir la sesión global por un identificador de edición en la URL y resolver el contexto en el servidor; convertir el panel lateral en pestañas o en rutas con estado en la URL; derivar la habilitación de páginas de las capacidades reales del usuario en vez de una bandera global; validar obligatoriedad antes de habilitar la publicación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla de publicación de una edición muestra su título y el estado activo o no activo.
2. Las páginas de acuerdo UTE y de suministradores solo aparecen cuando la edición las exige.
3. Sin edición activa, la pantalla no se abre y se avisa.
4. Salir devuelve al llamante sin guardar nada.
5. El botón de ayuda abre la ayuda de la pantalla.

## Pendientes

- Quién abre `FormPublicacionCalidad` y con qué edición: la búsqueda en `src` no devuelve ningún `DoCmd.OpenForm` que lo nombre, así que el punto de entrada está fuera del código fuente.
- Si la publicación comprueba la obligatoriedad de las páginas antes de publicar (afecta a **H-GR-23**).
- Qué escribe `lblEstado` (**H-GR-22**).