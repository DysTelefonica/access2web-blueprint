# Form_FormEventoTipoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoTipoGestion.cls` (222 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Mantenimiento del catálogo de tipos de evento (`TbTipoEvento`, campo `TIPOEVENTO`): lista los tipos existentes, permite dar de alta uno nuevo y eliminar el seleccionado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la única fuente de datos es el `RowSource` de la lista `ListaTipoEventos` (`.form.txt:326`). Las citas son `Form_FormEventoTipoGestion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | La etiqueta `lblGestionTipoEventos` abre `FormEventoTipoGestion` sin `OpenArgs` | `Form_Form0BDOpcionesOtras.cls:119-122` |

No recibe argumentos ni devuelve valores ni lanza eventos. Se cierra con `cmdSalir` (`:96-99`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaTipoEventos` | lista de una columna | `OnClick` → `ListaTipoEventos_Click` | Guarda en `strTipoEventoSeleccionado` el valor de la columna 0 y habilita `ComandoEliminar` si no está vacío | CAP-3 | `:100-108`, `.form.txt:324-329` |
| `ComandoAlta` («Alta») | botón | `OnClick` → `ComandoAlta_Click` | Pide el nombre con `InputBox` y llama a `Alta` | CAP-3 | `:5-37`, `.form.txt:355-357` |
| `ComandoEliminar` («Eliminar») | botón, deshabilitado en diseño | `OnClick` → `ComandoEliminar_Click` | Confirma y llama a `Eliminar` | CAP-3 | `:39-82`, `.form.txt:422, 432-434` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra el formulario sin guardar | CAP-3 | `:96-99`, `.form.txt:381-383` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «Tipos de Eventos» | — | `lblTitulo1` pasa a `Caption` del formulario | CAP-3 | `:89`, `.form.txt:462, 482` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:250, 274, 298` |
| `Marco120` | grupo de opciones | — | Marco decorativo, sin handler | sin asignar | `.form.txt:408` |

Controles: 10 (17 Begin en el .form.txt: 1 `Form`, 5 plantillas de estilo por defecto —`Label`, `Rectangle`, `CommandButton`, `OptionGroup`, `ListBox`—, 1 `Section` y los 10 controles anteriores).

## Flujos

### Apertura (`Form_Open`, `:83-95`)

Con `On Error Resume Next`: reloj de arena, `AjustarTamaño Me` (`FG.bas:261-288`), título de ventana desde `lblTitulo1` y repintado (`:84-94`). `ComandoEliminar` arranca deshabilitado por diseño (`.form.txt:422`).

### Alta (`ComandoAlta_Click`, `:5-37`)

1. `InputBox` «Nombre el Tipo de Evento»; si queda vacío o se cancela, sale sin más (`:8-11`).
2. Llama a `Alta` (`:20`). Si el texto devuelto contiene «|», lo trata como error y lo muestra (`:21-25, 31-36`).
3. Éxito: `Requery` de la lista (`:26`). `Alta` rechaza el nombre vacío y el duplicado y, si no, inserta con `AddNew`/`Update` (`:128-143`).
4. Error: reloj de arena fuera y `MsgBox` crítico (`:35-36`).

### Selección y baja (`ListaTipoEventos_Click`, `ComandoEliminar_Click`, `:100-108, 39-82`)

1. Seleccionar una fila habilita `ComandoEliminar` (`:102-106`).
2. Eliminar: lee la columna 0; vacía ⇒ «Se ha de seleccionar el Tipo de Evento que desea eliminar» (`:45-49`). `MsgBox` Sí/No con «No» por defecto (`:50-54`).
3. `Eliminar` comprueba con `Dame("TbEventos", "TIPOEVENTO", ...)` que ningún evento use el tipo (`:187-196`) y que exista en `TbTipoEvento` (`:202-205`); borra con `.Delete` (`:206`).
4. Éxito: `Requery`, foco en `ComandoAlta` y `ComandoEliminar` deshabilitado (`:69-71`).

## Reglas, validaciones y cálculos

- **Nombre obligatorio:** vacío ⇒ «No se ha indicado el nombre del Tipo de Evento» (`:128-131`).
- **Unicidad:** si ya existe un tipo con ese texto exacto ⇒ «Ya existe un Tipo Evento registrado con el nombre…» (`:132-140`).
- **Baja bloqueada por uso:** si algún registro de `TbEventos` tiene ese `TIPOEVENTO` ⇒ «…forma parte de algún evento» (`:187-196`).
- **Protocolo de error propio:** las funciones devuelven «#ERR|texto» y el llamador detecta el carácter «|» (`:21-25, 64-68, 157, 220`). Un texto de éxito que contenga «|» (por ejemplo, un nombre con ese carácter) se interpretaría como error (`:147, 21`).
- No hay confirmación en el alta ni límite de longitud en el código.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Rellenar la lista | `TbTipoEvento` | `RowSource` guardado en el formulario | — | `.form.txt:326` |
| Comprobar duplicado | `TbTipoEvento` | SQL dinámico concatenado (nombre entre comillas simples) + `OpenRecordset` | no | `:132-135` |
| Alta | `TbTipoEvento` | DAO `AddNew`/`Update` | no | `:141-143` |
| Comprobar uso | `TbEventos` | SQL dinámico concatenado en `Dame`, primero sin comillas y, si falla, con comillas | no | `FG.bas:290-310`, `:187` |
| Comprobar existencia y baja | `TbTipoEvento` | SQL dinámico concatenado + DAO `Delete` | no | `:197-206` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con el nombre entre comillas simples (`:134, 199`); `Dame` depende de que Jet/ACE acepte el valor sin comillas y falle para reintentar con comillas (`FG.bas:299-308`), comportamiento que habrá que revisar con un tipo de dato textual en PostgreSQL; `OpenRecordset` y `Delete` sobre tabla vinculada sin transacción explícita.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un nombre con apóstrofo rompería el SQL concatenado de duplicado y de borrado (`:134, 199`); el mensaje de error sería el genérico.
- `Dame` prueba primero el valor sin comillas; con un texto sin comillas Jet/ACE podría tratarlo como parámetro y fallar o no, y de ello depende que la comprobación de uso funcione (`FG.bas:299-308`, `:187`).
- En `ComandoEliminar_Click`, si el usuario responde «No» sale con `Exit Sub` sin apagar el reloj de arena activado en `:43` (`:50-54`).
- La comparación de duplicados puede no distinguir mayúsculas (`Option Compare Database`, `:2`, y criterio de Jet), por lo que «Mantenimiento» y «mantenimiento» podrían considerarse el mismo tipo; a verificar.
- La baja se hace por texto y no por clave: si dos filas tuvieran el mismo texto, `.Delete` borraría solo la primera (`:197-206`).
- Un nombre que contenga «|» haría que el alta se muestre como error aunque se haya insertado (`:21, 141-147`).
- Los comentarios de `Eliminar` hablan de «Originador», lo que apunta a código copiado de otro formulario (`:165`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: restricción única sobre el nombre y borrado por clave; comprobación de uso mediante clave ajena en vez de consulta previa.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista muestra todos los tipos ordenados por nombre y «Eliminar» está deshabilitado (`.form.txt:326, 422`).
2. Seleccionar un tipo habilita «Eliminar» (`:100-108`).
3. Cancelar o dejar vacío el nombre en el alta no escribe nada (`:8-11`).
4. Dar de alta un nombre ya existente muestra «Ya existe un Tipo Evento registrado…» y no inserta (`:137-140`).
5. Dar de alta un nombre nuevo lo añade y aparece en la lista (`:141-143, 26`).
6. Eliminar un tipo usado por algún evento muestra «…forma parte de algún evento» y no borra (`:193-196`).
7. Responder «No» a la confirmación no borra (`:52-54`).
8. Eliminar un tipo sin uso lo quita de la lista y deja «Eliminar» deshabilitado (`:206, 69-71`).

## Pendientes

- Si `TbTipoEvento.TIPOEVENTO` tiene índice único y su longitud (ERD de origen).
- Qué otras tablas referencian el tipo además de `TbEventos` (`:187`).
- Contenido de `AjustarTamaño` en ejecución y efecto de `DoCmd.Restore` (`FG.bas:261-288`, ya leído; comportamiento visual a verificar).
