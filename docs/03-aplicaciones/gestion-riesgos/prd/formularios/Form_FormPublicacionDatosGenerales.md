# Form_FormPublicacionDatosGenerales

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionDatosGenerales.cls` (88 líneas) y `src/forms/FormPublicacionDatosGenerales.form.txt` · **PRD:** [README](../README.md)

Ficha de identification de la edición en curso: código de documento, número de edición, fecha de edición, firmas de elaborado, revisado y aprobado, y el correo del RAC. Es de solo lectura y de solo presentación. No tiene botones, ni manejadores de clic, ni escritura en base de datos. Se abre como subformulario desde el control de navegación «Datos generales» de `FormPublicacionCalidad`. Las citas son `Form_FormPublicacionDatosGenerales.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormPublicacionCalidad` | Subformulario del control de navegación «Datos generales», con `NavigationTargetType =32768` | `src/forms/FormPublicacionCalidad.form.txt:460-481` |

No tiene salida: no se abre con `DoCmd.OpenForm`, no emite eventos y no cierra nada. **No acepta parámetros**: lee la edición del objeto de sesión `m_ObjEdicionActiva`, igual que el resto de la cadena de publicación. Es, por tanto, un consumidor silencioso del contexto global, y su única dependencia es que ese contexto esté poblado antes de que se abra la página.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Ajusta el tamaño y llama a `EstablecerDatos` | `:16-45`, `.form.txt:41` (`OnOpen`) |
| (formulario) | — | `EstablecerDatos` (procedimiento privado) | Rellena los ocho campos visibles | `:48-80`, `.form.txt:1457` |
| `CodigoDocumento` | texto | — | Código del documento del proyecto | `:59`, `.form.txt:852` |
| `Edicion` | texto | — | Número de edición | `:60`, `.form.txt:942` |
| `FechaEdicion` | texto | — | Fecha de la edición, en día/mes/año | `:58`, `.form.txt:972` |
| `Elaborado`, `Revisado`, `Aprobado` | texto | — | Firmas; su origen depende de si la edición está publicada | `:61-71`, `.form.txt:882, 912, 1002` |
| `CorreoRAC` | texto | — | Destinatario del RAC, si el proyecto lo tiene calculado | `:73-77`, `.form.txt:1032` |

Los ocho controles están marcados `Enabled = NotDefault` en el diseño (`.form.txt:619, 650, 682, 714, 746, 778, 810`), es decir, sin habilitado propio: su estado es puramente el de la vista, no una decisión del código.

## Flujos

### Apertura (`Form_Open`, `:16-45`)

1. Pone el reloj de sistema y llama a `AjustarTítulo`/`AjustarTamaño` sobre el propio formulario (`:20-24`; el tamaño se ajusta con `AjustarTamaño Me`).
2. Llama a `EstablecerDatos m_Error` (`:26`).
3. Si la función ha devuelto texto de error, lanza `Err.Raise 1000` (`:27-29`).
4. Quita el reloj (`:33-35`).

### Relleno (`EstablecerDatos`, `:48-80`)

1. Copia la fecha de edición con formato día/mes/año, el código de documento y el número de edición (`:58-60`).
2. **Ramifica según esté publicada o no la edición** (`:61-72`):
   - **Publicada** (`.FechaPublicacion <> ""`): las tres firmas se toman de la propia edición (`:63-65`).
   - **No publicada:** las tres firmas se toman de los valores por defecto del proyecto (`:69-71`).
3. Si el proyecto tiene un correo del RAC calculado y no está vacío, lo muestra (`:73-77`).

### Alternativo

No hay flujo alternativo con decisiones de negocio: cualquier camino que no sea el anterior acaba en el manejador de error.

### Error (`:36-44`)

Si `Err.Number` no es `1000`, compone un mensaje con número y descripción, lo envía con `CorreoAlAdministrador` y muestra un `MsgBox` crítico; si es `1000`, muestra solo el texto acumulado en `m_Error`. A diferencia de los formularios hermanos, **este no cierra la ventana**: el subformulario se queda visible con los campos vacíos o a medio rellenar.

## Reglas, validaciones y cálculos

- **Las firmas dependen del estado de publicación, no de quién mira.** La misma edición muestra firmas distintas a un técnico y a calidad según tenga fecha de publicación o no (`:61-71`); no hay ninguna validación que lo impida.
- **El correo del RAC solo se muestra si existe.** No hay campo para introducirlo ni para corregirlo: es estrictamente informativo, y el valor editable de esa dirección vive en el formulario de publicación.
- **No hay validación de entrada.** `EstablecerDatos` no comprueba nada: si `.FechaEdicion` no es una fecha, `Format` falla o devuelve el valor sin formatear, y el flujo continúa.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Fecha de edición, número de edición, fecha de publicación y firmas propias | `m_ObjEdicionActiva` | `:58, 60-65` |
| Código de documento y firmas por defecto | `m_ObjEdicionActiva.Proyecto` | `:59, 69-71` |
| Destinatario del RAC | `Proyecto.CorreoRACCalculado` | `:73-77`; `src/classes/Proyecto.cls:498` |

Sin escrituras: el formulario no guarda nada.

## Coexistencia

Pendiente de `../coexistencia.md`. Es el caso más simple de los tres: datos derivados, presentación pura, sin estado propio. Lo que no tiene equivalente directo es la dependencia implícita del contexto de sesión, que en web se sustituiría por un parámetro explícito de la vista.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **`EstablecerDatos` se declara con `On Error Resume Next`** (`:53`), de modo que un fallo al leer una propiedad —por ejemplo, si el proyecto no está cargado— no se propaga: la ejecución continúa y el campo se queda en blanco sin aviso, con la pantalla aparentemente correcta pero incompleta.
- **La fecha de edición se formatea sin comprobar nada.** `:58` pasa la propiedad directamente a `Format` con el formato día/mes/año; si el valor llega vacío o no es una fecha, el resultado es una cadena vacía o el propio error, sin mensaje para el usuario.
- **El fallo de apertura deja un formulario a medio rellenar.** El manejador de `:36-44` avisa pero **no cierra ni cancela**: a diferencia de los formularios hermanos, que cierran al fallar, este deja la ficha visible con una parte de los campos informed y otra no.
- **El valor por defecto del correo del RAC sobrevive entre aperturas.** `Me.CorreoRAC` solo se escribe cuando el valor no está vacío (`:74-76`); si el subformulario se reutiliza para otra edición cuyo proyecto no tiene correo, el valor anterior permanece en pantalla.
- **El subformulario hereda la identidad global sin comprobarla.** A diferencia de los otros dos formularios de la cadena, aquí no se verifica que `m_ObjEdicionActiva` exista antes de usarlo (`:56-58`): si se abre sin contexto, `EstablecerDatos` falla en la primera línea y el resultado es una ficha completamente vacía, sin mensaje.
- **El título del formulario lleva datos de identidad incrustados en el diseño.** `.form.txt:40` fija un texto de ejemplo con usuario y rol; el formulario no lo actualiza al abrirse, a diferencia de los otros dos, que sí reciben el título del usuario conectado en `Form_Load`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: convertir esta ficha en una sección más del detalle de la edición, en lugar de un formulario independiente, ya que no interactúa; pasar la edición como parámetro explícito de la vista en vez de leerla del contexto de sesión; mostrar las firmas como un bloque único con su origen distinguido (edición o proyecto) para que el lector entienda por qué cambian; y limpiar los campos al inicializar para que un valor heredado de una edición anterior nunca sobreviva.

## Criterios de aceptación

Propuestos, no aprobados:

1. La ficha muestra código, edición, fecha de edición, las tres firmas y el correo del RAC cuando exista.
2. Las firmas provienen de la edición si está publicada y del proyecto si no lo está.
3. Si falta el contexto de la edición, se avisa en lugar de mostrar una ficha vacía.
4. La ficha no permite editar ningún campo ni escribe nada.

## Pendientes

- Qué valores exactos toma `Proyecto.Elaborado`, `Proyecto.Revisado` y `Proyecto.Aprobado` cuando la edición aún no está publicada: si son nombres de persona, roles o rótulos de plantilla.
- Si `FechaPublicacion` es la única señal de «edición publicada» o si hay otro indicador más fiable.
- Si esta ficha se reutiliza para ediciones distintas sin recargar el subformulario, caso en el que el valor heredado del correo del RAC sería un defecto observable.
- Si la navegación de `FormPublicacionCalidad` recarga el subformulario en cada cambio de página o conserva la instancia anterior.