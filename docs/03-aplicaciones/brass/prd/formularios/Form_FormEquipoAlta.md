# Form_FormEquipoAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEquipoAlta.cls` (329 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un equipo en `TbEquipos` (identificador, nodo, BUI, subsistema, nombre y descripción opcional) a partir de tres combos en cascada Nodo → BUI → Subsistema y dos cuadros de texto. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la validación y la escritura viven en una función pública del propio módulo, `AltaIDEquipo` (`:197-322`). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormEquipoAlta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEquipoGestion` (`ComandoAltaEquipo`) | Cierra el formulario si estaba abierto, lo abre sin argumentos y, si la lista tiene filas, vuelca nodo, BUI y subsistema de la primera fila, llamando a mano a `ComboNodo_BeforeUpdate` y `ComboBui_BeforeUpdate` | `Form_FormEquipoGestion.cls:33-53` |
| `Form_FormEventoAlta` (`ComandoNuevoIDEquipo`) | Exige nodo, BUI y subsistema; abre con `OpenArgs` = `nodo\|BUI\|subsistema` | `Form_FormEventoAlta.cls:182-204` |

Salida: tras guardar, si `FormEquipoGestion` está abierto llama a su `ActualizarLista` (`:35-37`; `Form_FormEquipoGestion.cls:715`) y se cierra con `acSaveNo` (`:38`). No devuelve el identificador creado a quien lo abre. `cmdSalir` cierra sin guardar (`:111-114`).

## Controles y eventos

Controles: 30 (30 `Begin` en el `.form.txt`): el propio formulario (1), la sección Detalle (1), 7 plantillas de estilo por tipo de control en la cabecera (`.form.txt:738-841`) y 21 controles de la sección.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboNodo` | combo, `LimitToList` | `BeforeUpdate` → `ComboNodo_BeforeUpdate` (público) | Si el nodo cambia, vacía BUI y subsistema; recompone el `RowSource` de `ComboBUI` con los BUI del nodo (o todos si está vacío) | CAP-12 | `:80-109`, `.form.txt:1099-1121` |
| `ComboBUI` | combo, `LimitToList` | `BeforeUpdate` → `ComboBui_BeforeUpdate` (público) | Recompone el `RowSource` de `ComboSubsistema` con los subsistemas del BUI (o todos si está vacío) | CAP-12 | `:51-77`, `.form.txt:1136-1160` |
| `ComboSubsistema` | combo, `LimitToList` | — | Subsistema; obligatorio al guardar | CAP-12 | `.form.txt:936-960`, `:233-236` |
| `Equipo` | texto, `Tag` «OBLIGATORIO» | — | Nombre del equipo; obligatorio al guardar | CAP-5 | `.form.txt:1175-1190`, `:237-240` |
| `Descripcion` | texto multilínea | — | Descripción opcional; vacía se envía como «#» y no se guarda | CAP-5 | `.form.txt:873-885`, `:18, 304-306` |
| `Numero_Actividad` | texto con `ControlSource` «IDActividad» y `DefaultValue` con `DCount` | — | Residuo de diseño: `Enabled`/`Locked` activados; el código no lo lee ni lo escribe | sin asignar | `.form.txt:901-916` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Lee los cinco campos, llama a `AltaIDEquipo`, refresca la gestión y cierra | CAP-5 | `:6-49`, `.form.txt:1022-1032` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-5 | `:111-114`, `.form.txt:1060-1070` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «ALTA DE EQUIPO»; `lblTitulo1` pasa a ser el `Caption` del formulario | CAP-5 | `:125`, `.form.txt:1339-1371` |
| `Etiqueta142`, `Etiqueta196`, `Etiqueta197`, `Etiqueta198`, `Etiqueta199` | etiquetas de campo (NODO, Equipo, DESCR., BUI., SUBSIST.) | — | Rótulos | CAP-5 | `.form.txt:1204-1324` |
| `Etiqueta99`, `Etiqueta1`, `Etiqueta2`, `Etiqueta3` | etiquetas «*» | — | Marcas de obligatoriedad | CAP-5 | `.form.txt:1382-1457` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | CAP-5 | `.form.txt:974-1007` |

## Flujos

### Apertura (`Form_Load`, `:115-196`)

1. `AjustarTamaño Me` (`FG.bas:261`) y título desde `lblTitulo1` (`:123-125`).
2. Con `OpenArgs`, lo trocea por «\|»; si no resulta en tres partes, «Se ha abierto el formulario con parámetros insuficientes», `MsgBox` y cierre (`:126-132, 189-195`). Con argumentos, foco en `Equipo` (`:136`).
3. Por cada parte no vacía (nodo, BUI, subsistema) asigna el valor al combo, deja el combo `Enabled = False`, `Locked = True` y con `lngColorCampoNoEditable`; los combos sin valor quedan editables (`:138-180`). Para nodo y BUI reconstruye además el `RowSource` del combo siguiente con SQL concatenado (`:140-145, 156-161`).

### Alta (`ComandoRegistrar_Click`, `:6-49`)

1. Cursor de espera; lee los campos con `Nz`; `Descripcion` vacía vale «#» (`:11-18`).
2. `AltaIDEquipo` (`:24`). Si el retorno contiene «\|» lo trata como error (`#ERR\|texto`) y muestra «El método AltaIDEquipo ha devuelto un error» más el detalle (`:25-29`).
3. Sin error: refresca `FormEquipoGestion` si está abierto y cierra (`:35-40`).
4. Error: `MsgBox` crítico y el formulario queda abierto (`:43-48`).

### `AltaIDEquipo` (`:197-322`)

1. Valida obligatorios en el orden nodo, BUI, subsistema, equipo (`:225-240`).
2. Comprueba que la terna es compatible: `TbNodoBUI` unida a `TbSubsistemaBui` por BUI; si no hay fila, «El subsistema … no está registrado para el NODO … y el BUI …» (`:244-256`).
3. Calcula el identificador con `DameID("TbEquipos", "IDEquipo")` (`:265`), que recorre toda la tabla y devuelve máximo + 1 como texto (`FG.bas:409-428`); exige que sea numérico (`:272-275`).
4. Comprueba que no exista ya ese identificador (`:276-285`) ni otro equipo con el mismo nodo, BUI, subsistema y nombre (`:286-297`).
5. `AddNew` con `IDEquipo`, `NODO`, `BUI`, `SUBSISTEMA`, `Equipo` y, si procede, `Descripcion`; `Update` (`:298-308`). Devuelve el identificador (`:311`).

## Reglas, validaciones y cálculos

- **Obligatorios:** nodo, BUI, subsistema y nombre (`:225-240`). La descripción es opcional; vacía o «#» no se escribe (`:304-306`).
- **Compatibilidad de la terna** nodo/BUI/subsistema contra `TbNodoBUI` y `TbSubsistemaBui` (`:244-256`).
- **Unicidad:** nodo + BUI + subsistema + nombre de equipo (`:286-297`).
- **Identificador:** máximo existente + 1, sin reutilizar huecos (`FG.bas:409-428`).
- **Cascada de combos:** cambiar el nodo vacía BUI y subsistema (`:87-90`); cambiar el BUI solo cambia la lista de subsistemas (`:51-77`).
- **Errores:** los controlados se señalan con `Err.Raise 1000` y texto propio; el resto muestra número y descripción (`:314-316`, `:44-46`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de nodos | `TbNodos` | `RowSource` guardado en el diseño | — | `.form.txt:1119` |
| Lista de BUI por nodo | `TbNodoBUI` | SQL dinámico concatenado (`Like '*'` si no hay nodo) | — | `:92-103, 140-145` |
| Lista de subsistemas por BUI | `TbSubsistemaBui` | SQL dinámico concatenado (`Like '*'` si no hay BUI) | — | `:60-72, 156-161` |
| Compatibilidad de la terna | `TbNodoBUI`, `TbSubsistemaBui` | SQL dinámico concatenado, `OpenRecordset` | no | `:244-249` |
| Máximo identificador | `TbEquipos` | `SELECT` de la columna completa y bucle en VBA | no | `FG.bas:409-426` |
| Existencia del identificador | `TbEquipos` | SQL dinámico concatenado | no | `:276-279` |
| Duplicado por nombre | `TbEquipos` | SQL dinámico con cuatro textos entre comillas | no | `:286-292` |
| Alta | `TbEquipos` | DAO `AddNew`/`Update` | no | `:298-307` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: dialecto Jet en `Like '*'` y `SELECT DISTINCT` (`:62, 92`); identificador `Max + 1` en VBA sin transacción, que el PostgreSQL vinculado resolvería con secuencia; el campo `IDEquipo` se asigna como texto (`:299`, `FG.bas:428`) sobre una columna que podría ser numérica; SQL con texto entre comillas simples.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Max + 1` sin transacción ni bloqueo (`FG.bas:409-428`; `:265, 298-307`): dos altas simultáneas podrían calcular el mismo identificador. La comprobación de `:276-285` se hace antes del `AddNew`, no lo evita.
- El mensaje «Ya existe un IDEquipo que tiene el ID» muestra el nombre del equipo, no el identificador (`:282`).
- El `Recordset` de `:279` no se cierra antes de reasignar `rcdDatos` en `:292`; si hay fila existente, el error salta sin cerrar el primero, aunque el manejador cierra el último (`:317-320`).
- `ComboBui_BeforeUpdate` no vacía el subsistema al cambiar el BUI (`:51-77`) y la variable `strValorAnteriorBUI` solo se escribe (`:73`); un subsistema ajeno al nuevo BUI solo se detectaría al guardar (`:244-256`).
- `ComboNodo_BeforeUpdate` no reinicia el `RowSource` de `ComboSubsistema` al vaciar BUI y subsistema (`:80-109`).
- Ambos manejadores usan `On Error Resume Next` (`:53, 82`): un fallo en la recomposición de listas pasa inadvertido.
- Texto con apóstrofo en nodo, BUI, subsistema o equipo rompería los SQL concatenados (`:246-248, 288-291`).
- El comentario de cabecera dice que lo llama `ComandoAlta_Click` (`:217`); el botón real es `ComandoRegistrar` (`:6`).
- Residuos de diseño: `RowSource` guardado en `ComboBUI` y `ComboSubsistema` con valores concretos (`.form.txt:957, 1157`) y `Numero_Actividad` con origen `IDActividad` (`.form.txt:915-916`).
- «Nodo no vacío, BUI vacío» desde `OpenArgs` deja `ComboBUI` con el `RowSource` del nodo, pero `ComboSubsistema` con el del diseño (`:138-169`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos, restricción única sobre nodo + BUI + subsistema + nombre, y combos dependientes que reinicien el nivel siguiente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Guardar sin nodo, BUI, subsistema o nombre muestra el mensaje del campo faltante y no escribe (`ComboNodo`, `ComboBUI`, `ComboSubsistema`, `Equipo`).
2. Elegir un nodo limita `ComboBUI` a sus BUI y vacía BUI y subsistema.
3. Elegir un BUI limita `ComboSubsistema` a sus subsistemas.
4. Una terna no registrada en `TbNodoBUI`/`TbSubsistemaBui` se rechaza con el mensaje de compatibilidad.
5. Un equipo con igual nodo, BUI, subsistema y nombre se rechaza.
6. Un alta correcta crea la fila con identificador máximo + 1, cierra el formulario y refresca la gestión si está abierta.
7. Abierto con `OpenArgs`, los combos con valor aparecen bloqueados y el foco queda en `Equipo`; con argumentos incompletos se avisa y se cierra.
8. Una descripción vacía no se guarda.

## Pendientes

- Tipos y longitudes de `TbEquipos.IDEquipo` y restricciones únicas (ERD de origen).
- Contenido de `AjustarTamaño` (`FG.bas:261`) y valor de `lngColorCampoNoEditable` (global de `Variables Globales.bas`).
- Si `ControlSource` de `Numero_Actividad` provoca error con el formulario sin `RecordSource` (`.form.txt:915`).
