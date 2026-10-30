# Form_FormWeb

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormWeb.cls` (68 líneas) y `src/forms/FormWeb.form.txt` · **PRD:** [README](../README.md)

Visor embebido de una página HTML ya generada en disco. Es el punto de llegada de varios flujos: quien fija la global `m_URLHTMLActivo` decide qué se ve aquí. No escribe nada. Las citas son `Form_FormWeb.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`: lee la **global** `m_URLHTMLActivo`.

| Quién fija `m_URLHTMLActivo` | Evidencia |
|---|---|
| `Form_FormCalidadTareasDetalleEdicion` | `:127` |
| `Form_FormPublicacionCalidadPublicar` | `:80` |
| `Form_FormPublicacionCalidadPublicarEjecutar` | `:163` (desde `m_URLInforme`) |
| `Form_FormPublicacionTecnicoPropuesta` | `FormPublicacionTecnicoPropuesta.form.txt:1042` |
| `Form_FormTecnicoTareasDetalleEdicion` | `FormTecnicoTareasDetalleEdicion.form.txt:1572` |

El único punto que abre el formulario es `Funciones Generales.bas:3476`, y está precedido de un `DoCmd.Close` si ya estaba abierto (`:3473-3474`): **el visor es de instancia única** y se reutiliza. `Form_FormPublicacionCalidadPublicarEjecutar.cls:164-167` tiene la misma secuencia **comentada**, es decir, ese camino está desactivado a propósito.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Comprueba el fichero, ajusta el tamaño y navega | `:17-56` |
| (formulario) | — | `Form_Resize` | Reajusta el control `Navegador` al tamaño del formulario | `:62-68` |
| `Navegador` | control de navegación web | — | Muestra el HTML | `:38`, `.form.txt:363` |

## Flujos

### Apertura (`Form_Load`, `:17-56`)

1. `DoEvents`, reloj de arena y `AllowEdits = False` (`:22-27`).
2. `AjustarTamaño Me` (`:28`).
3. **Si el fichero no existe, cierra el formulario y sale sin avisar** (`:32-35`): ni `MsgBox` ni correo. El usuario ve abrirse y cerrarse una ventana sin explicación.
4. Si existe, `Me.Navegador.Navigate m_URLHTMLActivo` (`:38`), quita el reloj de arena y reactiva la edición (`:40-43`).

### Redimensionado (`Form_Resize`, `:62-68`)

`On Error Resume Next`. Ajusta `Navegador.Width` y `Navegador.Height` al interior del formulario. Es lo que hace que el visor ocupe toda la ventana.

### Errores (`:46-56`)

Quita el reloj de arena y, **según `Err.Number`**:

- **Distinto de 1000:** compone el mensaje con número y detalle, **envía correo al administrador** (`CorreoAlAdministrador`) y muestra un `MsgBox` crítico (`:48-50`).
- **1000 exacto:** muestra un `MsgBox` de advertencia con `m_Error`, sin correo (`:51-52`).

En ambos casos reactiva la edición y **cierra el formulario** (`:55-56`).

## Reglas, validaciones y cálculos

- **La visibilidad depende de una global.** No hay forma de abrir este formulario con un contenido concreto: hay que escribir la global antes. Es el patrón opuesto a `OpenArgs`, que es lo que usa el resto de la aplicación.
- **El código distinguido de «avisado» y «error».** El error interno 1000 se trata como advertencia y **no notifica por correo**; cualquier otro sí. Es la misma convención que usa el resto de la aplicación para los errores internos (`Form_frmSplash.cls:62`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Comprobación de existencia | `fso.FileExists` sobre la ruta de la global | `:32` |
| Lectura del HTML | fichero local, servido por el control `Navegador` | `:38` |

Sin escrituras en base de datos ni en disco: solo lee.

## Coexistencia

Pendiente de `../coexistencia.md`. Es el punto donde la dependencia de ficheros se hace explícita: la web genera un HTML en disco y el legacy lo abre con un control de navegador. Mientras conviven, ese fichero tiene que existir en un sitio que ambos vean, lo que ata el despliegue de la web al sistema de ficheros del usuario. **Es el punto más frágil de la coexistencia de esta aplicación**, junto con las rutas de anexos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-9.** Si el fichero no existe, el formulario se cierra **en silencio** (`:32-35`): sin `MsgBox`, sin correo y sin rastro. El usuario que pidió un informe y no lo ve no tiene forma de saber si el informe no existe, si se generó en otra ruta o si falló la generación.
- **H-GR-10.** El camino de commented-out en `Form_FormPublicacionCalidadPublicarEjecutar.cls:164-167` sugiere que la apertura desde ahí se desactivó porque duplicaba la de `Funciones Generales.bas:3476`. Conviene confirmar que no queda otro camino sin cerrar el visor previo, porque el control es de instancia única y una apertura sin cierre previo dejaría el contenido anterior.
- **H-GR-11.** Hay una URL de ejemplo **completamente escrita** en un comentario (`:33`), con ruta de usuario y de perfil de Windows. No se reproduce aquí. Es una ruta local de una máquina concreta que quedó en el código; conviene señalarla al propietario, que decide si se limpia.
- **H-GR-12.** `Form_Resize` usa `On Error Resume Next` (`:63`) y no distingue un fallo real de una excepción de control: si `Navegador` no existe o falla el ajuste, el visor se queda con el tamaño anterior sin avisar.
- **H-GR-13.** La ruta se lee de una global que cualquier otro formulario puede cambiar entre que se asigna y que este lee. No hay ninguna garantía de que el valor siga siendo el que el emisor puso.
- **H-GR-14.** `Me.AllowEdits` se pone a `True` en el final del manejador de error (`:55`) justo antes de cerrar: es inocuo aquí, pero si alguien reutiliza este patrón en un formulario que no cierre, dejaría el formulario editable tras un fallo.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: servir el informe desde una ruta autenticada de la web en lugar de un fichero local; pasar el contenido por parámetro en la ruta en lugar de una global; devolver un error explícito cuando el recurso no exista en vez de cerrar en silencio.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un HTML disponible, el visor lo muestra ocupando todo el formulario.
2. Redimensionar la ventana reajusta el visor.
3. Si el recurso no existe, se avisa al usuario en vez de cerrar sin más.
4. Abrir dos veces seguidas muestra el contenido de la segunda, no el de la primera.

## Pendientes

- Cómo se resuelve la ruta que hay que compartir entre el proceso que genera el HTML y el visor, y si esa ruta es alcanzable desde la red.
- Qué hace `AjustarTamaño` con este formulario en particular (`Funciones Generales.bas:1043`).
- Si el control `Navegador` es un `Microsoft Web Browser` activo o embebido, y qué versión del motor de render usa.