# Form_FormBUIGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormBUIGestion.cls` (229 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Mantenimiento mínimo del catálogo de BUI: una lista con todos los BUI, un botón para dar de alta uno nuevo y otro para eliminar el seleccionado. La lista tiene `RowSource` propio sobre `TbBUI` (`Form_FormBUIGestion.form.txt:329`), pero el formulario no tiene `RecordSource` propio. Las altas y bajas se hacen con DAO desde las funciones públicas `Alta` y `Eliminar` del propio formulario. Las citas son `Form_FormBUIGestion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | `lblGestionBUI_Click` abre `FormBUIGestion` sin argumentos | `Form_Form0BDOpcionesOtras.cls:69-72` |
| `Form_FormNodoBUIGestion` | `ComandoIrABUI_Click` abre `FormBUIGestion` sin argumentos | `Form_FormNodoBUIGestion.cls:46-49` |

No recibe `OpenArgs` ni devuelve datos. Tras eliminar, si `FormNodoBUIGestion` está abierto, refresca su lista `ListaBuisEnNodo` (`:72-74`). Sale con `cmdSalir` (`:95-98`). Variable de módulo: `strBUISeleccionado` (`:4`).

## Controles y eventos

Controles: 10 (17 `Begin` en el `.form.txt`: 1 `Form`, 5 plantillas de estilo por defecto en cabecera, 1 `Section` y 10 controles de la sección Detalle).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaBUI` | lista | `OnClick` → `ListaBUI_Click` | Guarda `Column(0)` en `strBUISeleccionado` y habilita `ComandoEliminar` si no está vacío | CAP-12 | `:99-107`, `.form.txt:327-331` |
| `ComandoAlta` («Alta») | botón | `OnClick` | Pide el nombre con `InputBox` y llama a `Alta` | CAP-12 | `:5-36`, `.form.txt:357-359` |
| `ComandoEliminar` («Eliminar») | botón | `OnClick` | Confirma y llama a `Eliminar` | CAP-12 | `:37-85`, `.form.txt:433-435` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-12 | `:95-98`, `.form.txt:383-385` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «GESTIÓN DE LOS BUIS»; `lblTitulo1` pasa a `Caption` del formulario | CAP-12 | `:89`, `.form.txt:463-464, 483-484` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:252, 276, 300` |
| `Marco120` | grupo de opciones | — | Marco decorativo sin eventos en el código | sin asignar | `.form.txt:410` |

## Flujos

### Apertura (`Form_Open`, `:86-94`)

Con `On Error Resume Next`: `AjustarTamaño Me` (`FG.bas:261-288`), copia el título y deshabilita `ComandoEliminar` (`:88-90`).

### Alta (`ComandoAlta_Click`, `:5-36`)

1. `InputBox("Nombre del BUI")`; vacío o cancelado ⇒ sale sin mensaje (`:8-11`).
2. `Alta(strBUI)` (`:19`, definición `:179-228`): vacío ⇒ error; busca `TbBUI` por `BUI` y, si existe, «El BUI ya está dado de alta previamente» (`:198-210`); si no, `AddNew` con el campo `BUI` (`:211-213`).
3. Devuelve el texto de éxito o `#ERR|texto`; el formulario detecta el `|` y lanza el error 1000 (`:20-24`). Tras el éxito, `ListaBUI.Requery` (`:25`).
4. Error: `MsgBox` crítico con el texto (`:30-35`).

### Baja (`ComandoEliminar_Click`, `:37-85`)

1. Sin selección ⇒ «Se ha de seleccionar el Nodo que desea eliminar» (`:40-43`).
2. Confirmación Sí/No con «No» por defecto (`:44-48`).
3. `Eliminar` (`:108-178`): comprueba con `Dame("TbEventos","IDEvento","BUI",…)` que ningún evento use el BUI (`:137-146`); abre `TbBUI` por `BUI` y borra el registro (`:147-157`); después borra de `TbNodoBUI` las filas de ese BUI (`:160-165`).
4. Éxito: limpia la selección, `Requery`, deshabilita el botón y refresca `ListaBuisEnNodo` si está abierto (`:63-74`).

## Reglas, validaciones y cálculos

- Nombre de BUI no vacío (`:8-11, 198-201`); unicidad exacta por texto en el alta (`:202-210`).
- Un BUI usado en `TbEventos` no se elimina (`:137-146`, mensaje «forma parte de algún evento»).
- Un BUI inexistente en `TbBUI` no se elimina (`:152-155`).
- La baja arrastra las filas de `TbNodoBUI` con ese BUI (`:160-164`).
- El mensaje de error de selección habla de «Nodo», no de BUI (`:41`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de BUI | `TbBUI` | `RowSource` guardado en el control, ordenado por `BUI` | — | `.form.txt:329` |
| Uso en eventos | `TbEventos` | `Dame` (SQL dinámico concatenado, con reintento entre comillas) | no | `:137`; `FG.bas:290-330` |
| Existencia en alta | `TbBUI` | SQL dinámico concatenado, DAO | no | `:202-205` |
| Alta | `TbBUI` | DAO `AddNew`/`Update` | no | `:211-213` |
| Existencia en baja | `TbBUI` | SQL dinámico concatenado, DAO | no | `:147-150` |
| Baja | `TbBUI` | DAO `Delete` | no | `:156` |
| Baja de relaciones | `TbNodoBUI` | SQL dinámico concatenado con `DoCmd.RunSQL` y avisos desactivados | no | `:160-165` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL con el BUI concatenado entre comillas simples; `DELETE tabla.campo FROM` es sintaxis Jet/ACE (`:160-161`); `DoCmd.RunSQL` y `SetWarnings` no tienen equivalente directo contra PostgreSQL vinculado; las dos bajas (`TbBUI` y `TbNodoBUI`) van en pasos separados sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Baja en dos pasos sin transacción (`:156` y `:164`): un fallo entre ambos dejaría el BUI borrado de `TbBUI` y las relaciones en `TbNodoBUI`.
- Si `RunSQL` falla, `Eliminar` ya ha borrado de `TbBUI` y devuelve error (`:156-177`).
- Los nombres con apóstrofo rompen el SQL concatenado (`:149, 162, 204`).
- `Dame` está pensada para un valor numérico y reintenta con comillas si falla (`FG.bas:300-312`); para un BUI de texto el primer intento puede fallar o dar un resultado inesperado: a verificar.
- `Form_Open` y `ListaBUI_Click` usan `On Error Resume Next` (`:87, 100`), por lo que los fallos pasan inadvertidos.
- Un `InputBox` vacío o cancelado se trata igual (`:8-11`), sin mensaje.
- El texto de error «Se ha de seleccionar el Nodo» no corresponde al objeto (`:41`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: borrado transaccional y restricción de integridad en lugar de comprobaciones en VBA.

## Criterios de aceptación

Propuestos, no aprobados:

1. «Alta» con un nombre nuevo lo añade a `ListaBUI` (`:5-25`).
2. «Alta» con un nombre ya existente muestra «El BUI ya está dado de alta previamente» y no escribe (`:207-210`).
3. «Eliminar» está deshabilitado al abrir y sin selección (`:90, 102-103`).
4. «Eliminar» de un BUI usado en algún evento muestra el aviso y no borra (`:143-146`).
5. «Eliminar» de un BUI libre pide confirmación y, si se acepta, lo quita de la lista y de `TbNodoBUI` (`:44-64, 147-165`).
6. Tras eliminar con `FormNodoBUIGestion` abierto, su lista de BUI se refresca (`:72-74`).

## Pendientes

- Contenido de la tabla `TbBUI` (clave, índices) y si `TbNodoBUI` tiene integridad referencial: ERD de origen.
- Si hay algún otro formulario o consulta que dependa de `TbBUI` además de los dos citados: no resuelto con el código leído.
