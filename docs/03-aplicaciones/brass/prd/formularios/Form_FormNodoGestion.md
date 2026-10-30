# Form_FormNodoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormNodoGestion.cls` (235 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Mantenimiento del catálogo de nodos (`TbNodos`): lista, alta por `InputBox` y baja con confirmación. Es un formulario **no enlazado** (sin `RecordSource`): la lista `ListaNodos` es un cuadro de lista con `RowSource` sobre `TbNodos` (`.form.txt:323-325`). La lógica de alta y baja vive en dos funciones públicas del propio formulario, `Alta` y `Eliminar` (`:184-232`, `:125-183`). `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormNodoGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormNodoBUIGestion` | `ComandoIrANodos_Click` abre el formulario sin argumentos | `Form_FormNodoBUIGestion.cls:50-53` |
| `Form_Form0BDOpcionesOtras` | `lblGestionNodos_Click` lo abre sin argumentos | `Form_Form0BDOpcionesOtras.cls:85-88` |

No recibe `OpenArgs`. Tras una baja correcta refresca `ListaNodos` de `FormNodoBUIGestion` si está abierto (`:77-79`); el alta **no** refresca ese formulario (`:25`). `cmdSalir` cierra sin guardar (`:106-109`).

## Controles y eventos

Controles: 17 `Begin` en el `.form.txt` = 1 `Form` + 1 `Section` (`Detalle`) + 5 plantillas por defecto (`.form.txt:144-210`) + 10 controles reales (tabla). Cuadre: 1 + 1 + 5 + 10 = 17.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `Form` | formulario | `OnOpen` → `Form_Open` | Ajusta tamaño y título desde `lblTitulo1`; errores ignorados (`On Error Resume Next`) | CAP-12 | `:93-105`, `.form.txt:42` |
| `ListaNodos` | cuadro de lista, columna 0 = nodo | `OnClick` → `ListaNodos_Click` | Guarda el nodo elegido y habilita `ComandoEliminar` si no está vacío | CAP-12 | `:110-124`, `.form.txt:323-327` |
| `ComandoAlta` («Alta») | botón | `OnClick` | Pide el nombre con `InputBox` y llama a `Alta` | CAP-12 | `:5-36`, `.form.txt:355-357` |
| `ComandoEliminar` («Eliminar») | botón, deshabilitado de diseño | `OnClick` | Pide confirmación y llama a `Eliminar` | CAP-12 | `:37-90`, `.form.txt:454, 465-467` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-12 | `:106-109`, `.form.txt:398-400` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2` | 3 rectángulos | — | Decoración | sin asignar | `.form.txt:251, 274, 297` |
| `Marco120` | grupo de opciones | — | Marco visual | sin asignar | `.form.txt:440` |
| `lblTitulo1`, `lblTitulo2` | 2 etiquetas «GESTIÓN DE LOS NODOS» | — | Título; `lblTitulo1` alimenta `Me.Caption` | sin asignar | `:99`, `.form.txt:508-509, 529-530` |

Desglose de los 10: 3 rectángulos, 1 cuadro de lista, 3 botones, 1 marco y 2 etiquetas.

## Flujos

### Alta (`ComandoAlta_Click`, `:5-36`)
1. `InputBox("Nombre del Nodo")`; vacío o cancelado: sale sin hacer nada (`:8-11`).
2. `Alta(strNodo)` (`:19`): rechaza nombre vacío (`:202-205`) y duplicado exacto, consulta por texto (`:206-214`); si no, `AddNew` del campo `NODO` (`:215-217`).
3. Un resultado con `\|` es error y se muestra (`:20-24`); éxito: `ListaNodos.Requery` (`:25`). El texto de éxito de `Alta` no se muestra (`:221`).

### Baja (`ComandoEliminar_Click`, `:37-90`)
1. Sin nodo seleccionado: «Se ha de seleccionar el Nodo que desea eliminar» (`:43-47`).
2. Confirmación `MsgBox` Sí/No con «No» por defecto (`:51-55`).
3. `Eliminar(strNodo)` (`:64`): rechaza si algún evento usa el nodo (`:153-162`); borra de `TbNodos` y de `TbNodoBUI` (`:163-174`).
4. Éxito: deshabilita `ComandoEliminar`, refresca la lista y la de `FormNodoBUIGestion` si está abierto (`:70-79`).
5. Error: `MsgBox` crítico y `Hourglass` apagado (`:84-89`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Nombre de nodo obligatorio (alta y baja) | funciones | `:144-147, 202-205` |
| Nodo único por coincidencia exacta de texto | `Alta` | `:206-214` |
| No se puede borrar un nodo usado en `TbEventos.NODO` (`Dame` devuelve valor no vacío) | `Eliminar` | `:153-162`; `FG.bas:290-320` |
| La baja elimina también las filas de `TbNodoBUI` con ese nodo | `Eliminar` | `:169-174` |
| Eliminar requiere selección y confirmación | formulario | `:43-55` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de nodos | `TbNodos` | SQL guardada en el `RowSource` | — | `.form.txt:325` |
| Comprobación de duplicado | `TbNodos` | DAO sobre SQL concatenado con comillas | no | `:206-209` |
| Alta | `TbNodos` | DAO `AddNew`/`Update` | no | `:215-217` |
| Comprobación de uso | `TbEventos` | función `Dame` (SQL dinámico concatenado) | — | `:153`; `FG.bas:290-320` |
| Baja del nodo | `TbNodos` | SQL dinámico concatenado con `DoCmd.RunSQL` y avisos apagados | no | `:163-168` |
| Baja de uniones | `TbNodoBUI` | SQL dinámico concatenado con `DoCmd.RunSQL` y avisos apagados | no | `:169-174` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE tabla.campo FROM` es sintaxis Jet/ACE (`:163, 169`); `DoCmd.RunSQL` con `SetWarnings` sobre tablas vinculadas; las dos bajas no están en una transacción; SQL concatenado con el nodo entre comillas simples; clave de `TbNodos` de tipo texto (campo único `NODO`, `:216`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Las dos bajas no son atómicas (`:163-174`): si falla la segunda, el nodo desaparece de `TbNodos` y quedan filas en `TbNodoBUI`. `SetWarnings True` se restaura en el manejador de errores (`:181`).
- Un apóstrofo en el nombre rompe la consulta de duplicado y las bajas (`:208, 165, 171`).
- Con `Option Compare Database` (`:2`) y SQL de Access el duplicado probablemente no distingue mayúsculas; no se normalizan espacios (`:8`).
- El alta no refresca `FormNodoBUIGestion` (`:25`), la baja sí (`:77-79`).
- La comprobación de uso solo mira `TbEventos.NODO` (`:153`); otras tablas que referencien el nodo, salvo `TbNodoBUI`, no se revisan en este código.
- El mensaje de error ante un nodo usado incluye el valor devuelto por `Dame` (`:160`), un identificador de evento mostrado al usuario.
- Los errores de `Form_Open` y `ListaNodos_Click` se ignoran (`:94, 111`).
- Parámetro `Canel` en `Form_Open` (`:93`): errata de nombre sin efecto.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: baja en una sola transacción con clave foránea en `TbNodoBUI`; validación de duplicados insensible a mayúsculas; refresco coherente de la gestión de nodos y BUI.

## Criterios de aceptación

Propuestos, no aprobados:

1. «Eliminar» está deshabilitado hasta elegir un nodo de la lista (`.form.txt:454`; `:110-120`).
2. Dar de alta un nodo nuevo lo muestra en la lista (`:215-217, 25`).
3. Dar de alta un nombre ya existente muestra «Ya existe un nodo registrado con el nombre» y no escribe (`:211-214`).
4. Cancelar o dejar vacío el `InputBox` no hace nada (`:8-11`).
5. Eliminar pide confirmación; «No» no borra (`:51-55`).
6. Eliminar un nodo usado en algún evento se rechaza y no borra nada (`:153-162`).
7. Eliminar un nodo libre lo quita de la lista y de las uniones con BUI, y refresca `FormNodoBUIGestion` si estaba abierto (`:163-179`, `:70-79`).

## Pendientes

- Implementación completa de `Dame` y su comportamiento con textos que no son numéricos (`FG.bas:290-320`).
- Contenido de `AjustarTamaño` (`:98`) y `FormularioAbierto` (`:77`, `FG.bas:171-186`).
- Claves e índices de `TbNodos` y `TbNodoBUI` (ERD de origen).
