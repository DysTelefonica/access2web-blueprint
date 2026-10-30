# Form_FormUbicacionGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormUbicacionGestion.cls` (219 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Mantenimiento del catálogo de ubicaciones (`TbUbicacion`, un único campo `UBICACION`): lista, alta por `InputBox` y baja con confirmación. Es un formulario sin `RecordSource`; la lista se alimenta con un `RowSource` propio (`Form_FormUbicacionGestion.form.txt:328`). Las citas son `:línea` del `.cls` del formulario salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | Clic en la etiqueta `lblGestionUbicaciones`: `DoCmd.OpenForm "FormUbicacionGestion"` | `Form_Form0BDOpcionesOtras.cls:129` |

No lanza eventos ni devuelve valores. `cmdSalir` cierra sin guardar (`:96-99`). Las funciones `Alta` y `Eliminar` son públicas (`:100, 149`) pero ninguna otra pieza las invoca según el índice (a verificar; véase Pendientes).

## Controles y eventos

Controles: 18 (18 `Begin` en el `.form.txt`): 1 formulario, 5 plantillas de estilo por defecto (Label, Rectangle, CommandButton, OptionGroup, ListBox; `.form.txt:144-210`), 1 sección, 3 rectángulos, 1 lista, 4 botones, 1 grupo de opciones y 2 etiquetas.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| Formulario | formulario | `OnOpen` → `Form_Open` | Ajusta tamaño, copia `lblTitulo1` al título de ventana | CAP-12 | `:83-95`; `Form_FormUbicacionGestion.form.txt:42` |
| `ListaUbicacion` | lista | `OnClick` → `ListaUbicacion_Click` | Guarda la ubicación seleccionada (columna 0) en `strUbicacionSeleccionada` y habilita o deshabilita `ComandoEliminar` | CAP-12 | `:212-220`; `Form_FormUbicacionGestion.form.txt:326-330` |
| `ComandoAlta` («Alta») | botón | `OnClick` → `ComandoAlta_Click` | Pide el nombre con `InputBox`, llama a `Alta` y refresca la lista | CAP-12 | `:5-37`; `Form_FormUbicacionGestion.form.txt:356-358` |
| `ComandoEliminar` («Eliminar») | botón | `OnClick` → `ComandoEliminar_Click` | Confirma, llama a `Eliminar`, limpia la selección y refresca | CAP-12 | `:39-82`; `Form_FormUbicacionGestion.form.txt:433-435` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra con `acSaveNo` | CAP-12 | `:96-99`; `Form_FormUbicacionGestion.form.txt:382-384` |
| `ComandoAyudaFormulario` | botón | `OnClick = [Event Procedure]` sin handler en el `.cls` | Ninguno visible | sin asignar | `Form_FormUbicacionGestion.form.txt:462-464` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «UBICACIONES» | — | Título | CAP-12 | `Form_FormUbicacionGestion.form.txt:502-523` |
| `Detalle`, `CuadroFormulario`, `CuadroSuperior`, `Cuadro2`, `Marco120` y 5 plantillas | sección, rectángulos, grupo de opciones, estilos | — | Solo diseño | sin asignar | `Form_FormUbicacionGestion.form.txt:233-409` |

## Flujos

### Apertura (`:83-95`)
Con `On Error Resume Next` (`:84`): cursor de espera, `AjustarTamaño` (`FG.bas:261-287`), título de ventana y repintado. Un fallo no se notifica.

### Alta (`:5-37`)
1. `InputBox("Nombre de la Ubicación")`; vacío o cancelar ⇒ sale sin hacer nada (`:8-11`).
2. `Alta(nombre)` (`:100-148`): vacío ⇒ error (`:118-121`); busca la ubicación con SQL concatenado (`:122-125`); si existe, «Ya existe una ubicación registrada con el nombre…» (`:127-130`); si no, `AddNew`/`Update` (`:131-134`).
3. Los errores vuelven como texto `#ERR|detalle` (`:147`); el llamador detecta `|` (`:21-25`) y muestra `MsgBox` crítico (`:31-36`).
4. Éxito: `ListaUbicacion.Requery` y fin del cursor de espera (`:26-28`). El texto de éxito de `Alta` (`:137`) no se muestra.

### Baja (`:39-82`)
1. Sin selección ⇒ «Se ha de seleccionar la ubicación que desea eliminar» (`:45-48`).
2. Confirmación Sí/No con «No» por defecto (`:49-53`).
3. `Eliminar` (`:149-210`): comprueba con `Dame("TbActividades","Ubicacion","Ubicacion",valor)` (`:176`; `FG.bas:290-326`) si la ubicación se usa en alguna actividad (`:182-185`); busca el registro (`:186-189`), si no existe error (`:191-194`) y lo borra con `.Delete` (`:195`).
4. Éxito: vacía la selección, `Requery`, foco en `ComandoAlta` y deshabilita `ComandoEliminar` (`:68-71`).

## Reglas, validaciones y cálculos

- **Obligatorio:** nombre no vacío (`:9, 118`).
- **Unicidad:** consulta de existencia por texto exacto en `TbUbicacion` (`:122-130`); la comparación la resuelve el motor.
- **Integridad al borrar:** solo se comprueba `TbActividades.Ubicacion` (`:158, 176-185`); no se consultan otras tablas.
- **Confirmación** de la baja con botón por defecto «No» (`:49-50`).
- **Sin transacción**: cada operación es una única escritura.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Rellenar la lista | `TbUbicacion` | SQL guardada en `RowSource`, ordenada por `UBICACION` | — | `Form_FormUbicacionGestion.form.txt:328` |
| Existencia en alta | `TbUbicacion` | SQL dinámico concatenado con el nombre entre comillas | no | `:122-125` |
| Alta | `TbUbicacion` | DAO `AddNew`/`Update` | no | `:131-134` |
| Uso en actividades | `TbActividades` | SQL dinámico concatenado, primero sin comillas y, si falla, entre comillas | no | `FG.bas:290-326` |
| Existencia en baja | `TbUbicacion` | SQL dinámico concatenado | no | `:186-189` |
| Baja | `TbUbicacion` | DAO `Delete` | no | `:195` |

## Coexistencia

Hay que adaptar: SQL concatenado con valor entre comillas simples (`:124, 188`); `Dame` genérico que prueba el valor sin comillas y recurre a comillas tras un error (`FG.bas:290-326`), lo que en PostgreSQL daría error de columna inexistente en el primer intento y no en un parámetro; comparación de texto sin distinguir mayúsculas heredada del motor Jet/ACE (a verificar en el vinculado); `CurrentDb()` frente a la conexión vinculada; borrado físico sin restricción de clave foránea.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un nombre con apóstrofo rompe la consulta de existencia (`:124`) y el alta falla con el mensaje genérico de error.
- Al responder «No» a la confirmación se sale con `Exit Sub` (`:51-53`) sin apagar el cursor de espera activado en `:43`.
- El texto de éxito de `Alta` habla de «originador» (`:137`), aunque no se muestra.
- Tras un error de `Dame` bajo `On Error Resume Next`, el uso de `rcdDatos` podría fallar (`FG.bas:290-326`).
- `ComandoAyudaFormulario` declara `OnClick` sin handler (`Form_FormUbicacionGestion.form.txt:464`) y conserva el rótulo de diseño «Comando46» (`Form_FormUbicacionGestion.form.txt:463`).
- Solo se protege el borrado frente a `TbActividades`; si otras tablas guardan la ubicación quedarían huérfanas.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: alta en línea sin `InputBox`; restricción única e integridad referencial en base de datos; bloqueo de la baja si hay referencias en cualquier tabla.

## Criterios de aceptación

Propuestos, no aprobados:

1. Dar de alta una ubicación nueva la deja en `ListaUbicacion` ordenada alfabéticamente (`:26`).
2. Dar de alta un nombre ya existente muestra «Ya existe una ubicación registrada con el nombre…» y no escribe (`:127-130`).
3. Cancelar o dejar vacío el `InputBox` no hace nada (`:9-11`).
4. `ComandoEliminar` solo está habilitado con una ubicación seleccionada (`:215-219`).
5. Eliminar una ubicación usada en alguna actividad muestra «forma parte de alguna actividad» y no borra (`:182-185`).
6. Eliminar una ubicación libre pide confirmación y, al aceptar, la quita de la lista (`:49-71`).
7. Cualquier error se muestra en un cuadro crítico y no deja el cursor de espera activo (`:35-36`).

## Pendientes

- Si otras tablas distintas de `TbActividades` referencian `UBICACION` (ERD de origen).
- Quién llama a `Alta` y `Eliminar` desde fuera del formulario (referencia dinámica no resuelta).
- Cuál era la acción prevista de `ComandoAyudaFormulario` (`Form_FormUbicacionGestion.form.txt:462-464`).
- Longitud y colación de `UBICACION` en el backend.
