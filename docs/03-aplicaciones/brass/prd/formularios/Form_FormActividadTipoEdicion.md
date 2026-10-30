# Form_FormActividadTipoEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormActividadTipoEdicion.cls` (323 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición del texto de un tipo de actividad del catálogo `TbCodActividad` (campo `ACTIVIDADNOPROGRAMADA`) identificado por su código. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el código en `OpenArgs`, lo relee con `Dame` y persiste con un `Recordset` DAO propio. `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormActividadTipoEdicion.cls:línea` salvo indicación. Capacidad de todos los controles: CAP-4 Actividades y tipos de actividad.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormActividadTipoGestion` | `ComandoEditar_Click`: cierra el formulario si estaba abierto y lo abre pasando el código seleccionado en `OpenArgs` | `Form_FormActividadTipoGestion.cls:10-29` |

Sin código en `OpenArgs`, `Form_Open` muestra «No se han cargado todos los datos» y se cierra (`:175-179, 204-210`). Tras guardar, refresca `ListaTiposDeActividad` de `FormActividadTipoGestion` si está abierto (`:111-113`) y se cierra (`:114`). `cmdSalir` pregunta si guardar cuando hay cambios (`:34-67`).

Controles: 14 (22 `Begin` en el `.form.txt`: 1 `Form`, 6 estilos por defecto, 1 `Section` y 14 controles).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `CodTipo` | texto no enlazado | — | Código del tipo; se rellena con `OpenArgs` y no se vuelve a escribir al guardar | `:180`; `.form.txt:496-521` |
| `TIPOACTIVIDAD` | texto no enlazado | — | Texto editable del tipo; es lo único que se guarda | `:197`, `:146`; `.form.txt:523-549` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Valida cambios, guarda, refresca la lista y cierra | `:76-125`; `.form.txt:337-339` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra; si hay cambios ofrece guardarlos | `:34-75`; `.form.txt:380-382` |
| `ComandoAyudaFormulario` | botón | `OnClick = [Event Procedure]` sin handler en el `.cls` | Ninguno visible | `.form.txt:423-425` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «Edición de Tipo de Actividad»; `lblTitulo1` se copia a `Me.Caption` | `:174`; `.form.txt:560-561, 580-581` |
| `Etiqueta92`, `Etiqueta118`, `Etiqueta1`, `Etiqueta2` | etiquetas | — | Rótulos «Tipo», «Cod» y dos «*» | `.form.txt:265-266, 481, 624, 646` |
| `CuadroSuperior`, `Cuadro2`, `Marco122` | rectángulos y marco | — | Decoración | `.form.txt:290, 313, 602` |

## Flujos

### Apertura (`Form_Open`, `:167-211`)

1. `AjustarTamaño Me` (`FG.bas:261`) y título de ventana (`:173-174`).
2. `strCodActividadAlInicio = Nz(OpenArgs, "")`; vacío ⇒ error (`:175-179`). Rellena `CodTipo` (`:180`).
3. `Dame("TbCodActividad", "ACTIVIDADNOPROGRAMADA", "CODIGO", código)` (`:186`); un resultado con `|` es error (`:187-191`); vacío ⇒ «No se han cargado todos los datos» (`:193-196`).
4. Guarda el valor en `strTipoActividadAlInicio` y lo muestra en `TIPOACTIVIDAD` (`:192, 197`).
5. Cualquier error: `MsgBox` crítico y cierre del formulario (`:204-210`).

### Guardado (`ComandoRegistrar_Click`, `:76-125`)

1. `HaHabidoCambios` compara `TIPOACTIVIDAD` y `CodTipo` con los valores de apertura (`:21-33`). Sin cambios: «No ha habido cambios que guardar» (`:88-94`). Ante error interno devuelve `True` (`:31-32`).
2. `Registrar` (`:126-166`) exige código de apertura, lee `TIPOACTIVIDAD` y llama a `editar` (`:142-152`); éxito: «Edición Correcta» (`:158`).
3. Refresca la lista de `FormActividadTipoGestion` si está abierto (`:111-113`), cierra con `acSaveNo` (`:114`).
4. Error: `MsgBox` crítico y el formulario queda abierto (`:119-124`).

### Salir con cambios (`cmdSalir_Click`, `:34-75`)

Pregunta «¿Desea guardar los cambios producidos?» (`:39`). Sí ⇒ `Registrar` y refresco (`:40-61`). No ⇒ descarta. En ambos casos cierra (`:67`). Si `Registrar` falla, muestra el error y **no** cierra (`:69-74`).

## Reglas, validaciones y cálculos

`editar` (`:212-312`) aplica en orden:

- Código de apertura y tipo no vacíos: «No se ha indicado el código…» / «…el tipo de actividad» (`:233-240`).
- El código existe en `TbCodActividad` (`:246-255`).
- **Unicidad del texto:** si ya hay un código con ese `ACTIVIDADNOPROGRAMADA`, «Ya existe un tipo de actividad como esa» (`:261-270`).
- **Sin uso previo:** si `Dame("TbActividades", "Actividad", "Actividad", código)` devuelve algo, «No se puede cambiar ya que forma parte de al menos una actividad ya registrada» (`:276-285`).
- Solo se cambia el texto del tipo, nunca el código (`:215-223`, `:296`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del tipo al abrir | `TbCodActividad` | `Dame`: SQL dinámico concatenado; reintenta con comillas si falla | no | `:186`; `FG.bas:300-310` |
| Comprobación de existencia, duplicado y uso | `TbCodActividad`, `TbActividades` | `Dame` (idem) | no | `:246, 261, 276` |
| Actualización | `TbCodActividad` | SQL dinámico concatenado + DAO `Edit`/`Update` | no | `:286-298` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `Dame` construye SQL por concatenación y depende de un fallo de Jet (valor no numérico) para pasar a comillas (`FG.bas:303-310`); en PostgreSQL con tipos estrictos el primer intento podría no fallar o fallar de otro modo. La comparación del campo de código usa el valor concatenado sin tipo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La comprobación de unicidad (`:261-270`) hace `Dame` con el texto sin escapar y sin comillas la primera vez (`FG.bas:300-304`); un texto con apóstrofo o con números/operadores podría romper o falsear la consulta.
- `Dame` devuelve vacío tanto si no hay fila como si el campo es nulo (`FG.bas:316-319`): «no existe» y «sin valor» no se distinguen (`:193-196`).
- La comprobación de uso (`:276`) busca el **código** en la columna `Actividad` de `TbActividades`, no el texto; que ese campo guarde el código es una suposición a verificar.
- `Option Compare Database` (`:2`) podría hacer insensible a mayúsculas `HaHabidoCambios` (`:25`); cambiar solo la caja daría «No ha habido cambios».
- Mensaje de error de `ComandoRegistrar_Click` habla de «registrar el gasto» (`:121`), texto heredado.
- `ComandoAyudaFormulario` tiene `OnClick` sin handler (`.form.txt:425`): al pulsarlo no hay efecto o error de Access.
- Comprobar y escribir no están en transacción (`:246-298`): otra sesión podría crear el mismo texto entre ambas.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: restricción única sobre el texto del tipo y clave foránea desde actividades, para no depender de comprobaciones previas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin código muestra «No se han cargado todos los datos» y no abre el editor (`Form_Open`).
2. Abrir con un código existente muestra el código y su tipo actuales (`CodTipo`, `TIPOACTIVIDAD`).
3. Guardar sin modificar el texto muestra «No ha habido cambios que guardar».
4. Guardar un texto ya usado por otro código muestra «Ya existe un tipo de actividad como esa» y no escribe.
5. Guardar un tipo cuyo código figura en alguna actividad muestra «No se puede cambiar ya que forma parte…» y no escribe.
6. Guardar un cambio válido actualiza solo el texto, refresca la lista de gestión y cierra.
7. Salir con cambios pregunta si guardar; «No» cierra sin escribir (`cmdSalir`).

## Pendientes

- Contenido de `Dame`/`FormularioAbierto` frente a `TbActividades.Actividad`: qué guarda realmente esa columna (ERD de origen).
- `ComandoAyudaFormulario` sin handler (`.form.txt:425`): si abre alguna ayuda por macro.
- Efecto de `DoCmd.Restore` dentro de `AjustarTamaño` (`FG.bas:261`), no analizado.
