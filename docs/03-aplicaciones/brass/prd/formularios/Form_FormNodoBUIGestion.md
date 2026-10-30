# Form_FormNodoBUIGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormNodoBUIGestion.cls` (320 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de la relación nodo–BUI (`TbNodoBUI`): se elige un nodo, se ve la lista de sus BUI y se puede asignar un BUI libre o retirar uno ya asignado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); lee y escribe `TbNodoBUI` por DAO desde dos funciones públicas del propio módulo, `AsignarBUIANodo` (`:180-239`) y `DesAsignarBUIANodo` (`:240-312`). Abreviatura: `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormNodoBUIGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` (`lblRelacionNodoBUI`) | `DoCmd.OpenForm "FormNodoBUIGestion"` sin argumentos | `Form_Form0BDOpcionesOtras.cls:134-137` |

No recibe `OpenArgs` ni devuelve valores. Navega a `FormBUIGestion` (`:46-49`) y `FormNodoGestion` (`:50-53`), ambos con `On Error Resume Next`. Otros formularios le piden refrescar controles por nombre mientras está abierto: `Form_FormBUIGestion.cls:72-73` sobre `ListaBuisEnNodo` y `Form_FormNodoGestion.cls:77-78` sobre `ListaNodos` (este último control no existe aquí, véase defectos). Sale con `cmdSalir` (`:176-179`).

## Controles y eventos

Controles: 25 (25 `Begin` en el `.form.txt`): el propio formulario (1), la sección Detalle (1), 7 plantillas de estilo por tipo de control en la cabecera y 16 controles de la sección.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboNodos` | combo, `LimitToList` | `OnEnter` → `ComboNodos_Enter`; `OnChange` → `ComboNodos_Change`; `OnExit` → `ComboNodos_Exit` | `Enter` rehace la lista de nodos; `Change` y `Exit` ejecutan el mismo código: recargan `ListaBuisEnNodo` con los BUI del nodo (con «-1» si está vacío) y deshabilitan Quitar | CAP-12 | `:101-162`, `.form.txt:585-610` |
| `ListaBuisEnNodo` | lista de 2 columnas (BUI, nodo) con encabezados | `OnClick` → `ListaBuisEnNodo_Click` | Habilita Quitar si la columna 0 no está vacía | CAP-12 | `:313-321`, `.form.txt:463-483` |
| `ComboBuiParaNodo` | combo, `LimitToList` | `OnEnter` → `ComboBuiParaNodo_Enter` | Rehace la lista de BUI candidatos: los de `TbBUI` sin ninguna fila en `TbNodoBUI` | CAP-12 | `:97-100`, `.form.txt:715-739` |
| `ComandoAñadirBUI` («Añadir BUI») | botón | `OnClick` → `ComandoAñadirBUI_Click` | Valida nodo y BUI, llama a `AsignarBUIANodo`, refresca ambas listas | CAP-12 | `:5-45`, `.form.txt:501-513` |
| `ComandoQuitarBUI` («Quitar BUI») | botón | `OnClick` → `ComandoQuitarBUI_Click` | Toma BUI y nodo de la fila seleccionada y llama a `DesAsignarBUIANodo` | CAP-12 | `:54-95`, `.form.txt:844-857` |
| `ComandoIrABUI`, `ComandoIrANodos` | botones (rótulo de diseño «ComandoLimpiarEquipo») | `OnClick` | Abren `FormBUIGestion` y `FormNodoGestion` | CAP-12 | `:46-53`, `.form.txt:652-663, 781-792` |
| `ComandoAyudaFormulario` (rótulo «Comando46») | botón | `OnClick` declarado, sin handler en el `.cls` | Sin código asociado | sin asignar | `.form.txt:408-419` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-12 | `:176-179`, `.form.txt:344-355` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «BUIS EN NODOS» | — | `lblTitulo1` pasa a ser el `Caption` del formulario | CAP-12 | `:169`, `.form.txt:544-574` |
| `Etiqueta142` («NODOS»), `Etiqueta1` («BUI») | etiquetas | — | Rótulos de los dos combos | CAP-12 | `.form.txt:625-637, 754-766` |
| `Marco120` | grupo de opciones | — | Sin código asociado | sin asignar | `.form.txt:386` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | CAP-12 | `.form.txt:271-320` |

## Flujos

### Apertura (`Form_Open`, `:163-175`)

`AjustarTamaño Me` (`FG.bas:261`) y título desde `lblTitulo1`, todo bajo `On Error Resume Next`. No carga datos: la lista nace con el `RowSource` de diseño, filtrado por un nodo inexistente (`.form.txt:480`).

### Elegir nodo (`ComboNodos_Change` / `_Exit`, `:101-162`)

Recompone el `RowSource` de `ListaBuisEnNodo` con `SELECT BUI, NODO FROM TbNodoBUI WHERE NODO='<nodo>'` ordenado por BUI, hace `Requery` y deshabilita `ComandoQuitarBUI` (`:107-117, 141-151`).

### Añadir BUI (`ComandoAñadirBUI_Click`, `:5-45`)

1. Sin nodo: «No se puede determinar para qué NODO se quiere añadir el BUI»; sin BUI: «Se ha de seleccionar del desplegable un BUI a añadir» (`:13-20`).
2. `AsignarBUIANodo`; si devuelve «\|» se trata como error `#ERR\|texto` (`:26-31`).
3. Éxito: `Requery` de la lista, vacía `ComboBuiParaNodo` y lo rehace (`:32-34`). Error: `MsgBox` crítico; el formulario sigue abierto (`:39-44`).

### Quitar BUI (`ComandoQuitarBUI_Click`, `:54-95`)

Lee BUI y nodo de las columnas 0 y 1 de la fila (`:60-61`), valida (`:62-69`), llama a `DesAsignarBUIANodo` (`:75`), refresca ambas listas, mueve el foco a Añadir y deshabilita Quitar (`:81-84`).

### `AsignarBUIANodo` (`:180-239`) y `DesAsignarBUIANodo` (`:240-312`)

- Asignar: exige BUI y nodo (`:200-207`); si ya existe la pareja, «Ya está registrado para el Nodo … el BUI …» (`:211-220`); si no, `AddNew` con `BUI` y `NODO` (`:221-224`).
- Desasignar: exige BUI y nodo (`:261-268`); busca el BUI en `TbEventos` con `Dame` (`:274`, `FG.bas:290-330`) y, si hay algún evento, rechaza «La pareja Nodo … BUI … se ha registrado en al menos un evento» (`:280-283`); localiza la pareja y la borra con `.Delete` (`:287-297`).

## Reglas, validaciones y cálculos

- **Nodo y BUI obligatorios** en ambas operaciones, comprobados en el botón y otra vez en la función (`:13-20, 62-69, 200-207, 261-268`).
- **Pareja única** nodo + BUI al asignar (`:211-220`).
- **BUI sin nodo:** `ComboBuiParaNodo` solo ofrece BUI de `TbBUI` sin fila alguna en `TbNodoBUI` (`.form.txt:736`); por la interfaz, un BUI pertenece a un único nodo.
- **Retirada bloqueada** si el BUI figura en cualquier evento de `TbEventos` (`:274-283`); la comprobación es por BUI, no por la pareja.
- **Nodo vacío** se sustituye por «-1» para que la lista quede vacía (`:108-110, 142-144`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de nodos | `TbNodos` | `RowSource` guardado | — | `.form.txt:606` |
| BUI libres | `TbBUI`, `TbNodoBUI` | `RowSource` guardado, `LEFT JOIN … Is Null` | — | `.form.txt:736` |
| BUI del nodo | `TbNodoBUI` | SQL dinámico con el nodo entre comillas | — | `:111-115, 145-149` |
| Pareja existente | `TbNodoBUI` | SQL dinámico concatenado, `OpenRecordset` | no | `:211-215, 287-291` |
| Uso en eventos | `TbEventos` | `Dame`: primero SQL con el valor sin comillas y, si falla, con comillas | no | `FG.bas:300-310` |
| Alta de pareja | `TbNodoBUI` | DAO `AddNew`/`Update` | no | `:221-224` |
| Baja de pareja | `TbNodoBUI` | DAO `Delete` | no | `:297` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `LEFT JOIN … Is Null` y `ORDER BY` en `RowSource` (dialecto Jet, `.form.txt:736`); SQL concatenado con texto entre comillas; `.Delete` sobre un `Recordset` de selección (`:291-297`), que en tablas vinculadas a PostgreSQL requiere clave primaria o índice único en `TbNodoBUI`; el intento de `Dame` con valor sin comillas depende de que Jet falle con texto (`FG.bas:302-310`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Form_FormNodoGestion.cls:78` refresca `Controls("ListaNodos")` de este formulario, que no tiene ese control (el existente es `ComboNodos`, `.form.txt:585`); con el formulario abierto daría error de ejecución en `FormNodoGestion`.
- La comprobación de eventos es por BUI, no por pareja (`:274`): un BUI usado en cualquier evento no puede retirarse aunque la pareja con ese nodo nunca se usara; el mensaje habla de la pareja (`:281`).
- En `DesAsignarBUIANodo`, si la pareja no existe el mensaje es «Ya está registrado…» (`:293-295`), copiado de la función de alta.
- `Dame` concatena el BUI sin comillas en el primer intento (`FG.bas:302`); un BUI con texto no numérico depende del fallo y la segunda consulta (`FG.bas:304-310`), con `On Error Resume Next` (`FG.bas:303`). Un BUI vacío en `TbEventos` devuelve cadena vacía y no bloquea (`FG.bas:314`).
- `ComboNodos_Change` y `_Exit` repiten el código (`:101-162`); `Change` se dispara en cada pulsación y hace `Requery` cada vez.
- `strNodoSeleccionado` recibe la columna 0 (el BUI) y no se usa (`:4, 315`).
- `ComandoAyudaFormulario` declara `OnClick` sin handler (`.form.txt:419`); `ComandoIrABUI` y `ComandoIrANodos` conservan el rótulo `ComandoLimpiarEquipo` (`.form.txt:662, 791`).
- Texto con apóstrofo en nodo o BUI rompería el SQL concatenado (`:113, 213-214, 289-290`).
- El estado inicial de `ComandoQuitarBUI` depende de `Enabled` en diseño (`.form.txt:845`); solo se deshabilita por código tras elegir nodo o quitar (`:84, 117, 151`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: clave única nodo + BUI y comprobación de uso por pareja; un único manejador para la carga de la lista.

## Criterios de aceptación

Propuestos, no aprobados:

1. Elegir un nodo en `ComboNodos` muestra en `ListaBuisEnNodo` sus BUI y deja Quitar deshabilitado.
2. Añadir sin nodo o sin BUI muestra el mensaje correspondiente y no escribe.
3. Añadir un BUI libre crea la pareja, lo muestra en la lista y lo retira de `ComboBuiParaNodo`.
4. Añadir una pareja ya existente se rechaza con el mensaje de duplicado.
5. Seleccionar una fila habilita Quitar; quitar borra la pareja y deshabilita el botón.
6. Quitar un BUI presente en algún evento se rechaza con el mensaje de uso.
7. `ComandoIrABUI` y `ComandoIrANodos` abren las gestiones de BUI y de nodos.

## Pendientes

- Clave e índices de `TbNodoBUI` y de `TbEventos.BUI` (ERD de origen).
- Contenido de `AjustarTamaño` (`FG.bas:261`) y de `FormBUIGestion`/`FormNodoGestion`.
- Función de `Marco120` y estado inicial de `ComandoQuitarBUI` (`.form.txt:386, 845`).
