# Form_FormTecnicoAusencias

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnicoAusencias.cls` (215 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta por lotes y edición de una libranza (ausencia) de un técnico en `TbTecnicosAusencias`. Es un formulario **no enlazado** (sin `RecordSource`): el técnico y, en edición, la libranza llegan por las globales `m_ObjTecnicoActivo` y `m_ObjLibranzaActiva` (`VG.bas:245-246`), y el registro lo hace `Tecnico.AltaAusencias` (`Tecnico.cls:415-507`) apoyándose en `LIbranza.Registrar` (`LIbranza.cls:180-261`). `VG.bas` = `modules/Variables Globales.bas`. Las citas son `:línea` del `.cls` del formulario salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTecnicoAusenciasGestion` (alta) | `m_ObjLibranzaActiva = Nothing`, fija el técnico activo, cierra el formulario si estaba abierto y lo abre | `Form_FormTecnicoAusenciasGestion.cls:63, 78-81` |
| `Form_FormTecnicoAusenciasGestion` (edición) | `m_ObjLibranzaActiva = m_ObjLibranzaSeleccionada` y lo abre | `Form_FormTecnicoAusenciasGestion.cls:113, 115-118` |

El formulario declara los eventos `Alta` y `Editado` (`:7-8`); solo se lanza `Alta`, tras guardar (`:47`), también en edición. `Form_FormTecnicoAusenciasGestion` guarda una referencia `WithEvents` (`Form_FormTecnicoAusenciasGestion.cls:7`) pero no hay ningún manejador `m_Form_*` en ese fichero (búsqueda textual). Tras guardar se cierra solo (`:54`); `cmdSalir` cierra sin guardar (`:63-66`).

## Controles y eventos

Controles: 28 (28 `Begin` en el `.form.txt`): 1 formulario, 7 plantillas de estilo por defecto (Label, Rectangle, Image, CommandButton, OptionGroup, TextBox, ComboBox; `Form_FormTecnicoAusencias.form.txt:141-244`), 1 sección, 3 rectángulos, 3 botones, 4 etiquetas de campo, 3 cuadros de texto, 1 combo, 1 grupo de opciones, 2 títulos y 2 asteriscos.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| Formulario | formulario | `OnLoad` → `Form_Load` | Valida técnico activo, rellena combo y títulos, carga la libranza en edición | CAP-9 | `:67-109`; `Form_FormTecnicoAusencias.form.txt:82` |
| `FechaInicial` («FECHA 1») | texto, formato fecha corta | — | Primer día; obligatorio al guardar | CAP-9 | `:29`; `Form_FormTecnicoAusencias.form.txt:447-448`; `Tecnico.cls:437-441` |
| `FechaFinal` («FECHA 2») | texto, formato fecha corta | — | Último día; si no es fecha se toma la inicial | CAP-9 | `:30`; `Form_FormTecnicoAusencias.form.txt:503-504`; `Tecnico.cls:442-445` |
| `ComboMotivo` | combo lista de valores | — | Motivo; obligatorio; se rellena con los motivos ya usados | CAP-9 | `:39, 186-216`; `Form_FormTecnicoAusencias.form.txt:586-587`; `Tecnico.cls:450-453` |
| `Observaciones` | texto | — | Comentario libre opcional | CAP-9 | `:40`; `Form_FormTecnicoAusencias.form.txt:640` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Registra la libranza o el lote y cierra | CAP-9 | `:10-62`; `Form_FormTecnicoAusencias.form.txt:353-355` |
| `CmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra con `acSaveNo` | CAP-9 | `:63-66`; `Form_FormTecnicoAusencias.form.txt:664-666` |
| `ComandoAyudaFormulario` | botón | sin evento | Ninguno; rótulo de diseño «Comando46» | sin asignar | `Form_FormTecnicoAusencias.form.txt:379-380` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «Alta de Libranzas para <alias>» o «Edición de Libranza para <alias>» | CAP-9 | `:120-122, 129-131` |
| `Etiqueta5`, `Etiqueta10`, `Etiqueta15`, `Etiqueta18`, `Etiqueta2`, `Etiqueta3` | etiquetas | — | Rótulos «FECHA 1», «FECHA 2», «MOTIVO», «OBS.» y dos asteriscos | CAP-9 | `Form_FormTecnicoAusencias.form.txt:420, 475, 553, 612, 733, 755` |
| `Detalle`, `CuadroPrincipal`, `CuadroSuperior`, `Cuadro2`, `Marco13` y 7 plantillas | sección, rectángulos, grupo de opciones, estilos | — | Solo diseño | sin asignar | `Form_FormTecnicoAusencias.form.txt:267-529` |

## Flujos

### Apertura (`Form_Load`, `:67-109`)
1. `AjustarTamaño` (`FG.bas:261-287`).
2. Sin técnico activo se toma el de la libranza activa; si tampoco hay libranza, «No se conoce el técnico» (`:78-88`).
3. `EstablecerDatos` (`:111-149`): rellena `ComboMotivo` con `SELECT DISTINCT Motivo` (`:195`). Sin libranza activa pone título de alta y termina (`:119-124`); con libranza **relee** el registro con `Constructor.getLibranza` (`:125`; `Constructor.bas:4067-4116`) y copia fecha (en ambos campos de fecha), motivo y observaciones no vacíos (`:132-143`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:101-108`); el texto dice «Form_Open» aunque es `Form_Load` (`:103`).

### Registro (`ComandoRegistrar_Click`, `:10-62`)
1. Crea `m_ObjLibranzaActiva` si no existía; lee las fechas con `Nz` como texto (`:25-30`).
2. Asigna técnico y alias, motivo y observaciones (`:32-42`).
3. `AltaAusencias(libranza, libranza inicial, fechaInicial, fechaFinal)` (`:43`; `Tecnico.cls:415-507`) valida, recorre cada día del rango y registra uno a uno (`Tecnico.cls:459-483`).
4. Lanza `Alta`, muestra el resumen en `MsgBox` «Alta por lotes» y cierra (`:47-54`).
5. Error: se muestra en `MsgBox` crítico y el formulario queda abierto (`:56-61`).

## Reglas, validaciones y cálculos

- **Alias obligatorio** del técnico (`Tecnico.cls:434-437`). **Fecha inicial obligatoria** y válida (`Tecnico.cls:437-441`). **Fecha final** vacía o no válida ⇒ igual a la inicial (`Tecnico.cls:442-445`); no puede ser anterior a la inicial (`Tecnico.cls:446-449`). **Motivo obligatorio** (`Tecnico.cls:450-453`).
- **Días del rango:** diferencia absoluta de días entre ambas fechas más uno (`Tecnico.cls:454-458`).
- **Día registrable:** se omiten sábados y domingos (`Tecnico.cls:528-529`) y las fechas de `TbTecnicosFiestas` (`Tecnico.cls:552-553`); la llamada pasa `EnumSino.No`, de modo que no se consultan las ausencias existentes (`Tecnico.cls:470`).
- **Duplicado:** en alta, `MotivoNoOK` rechaza «La fecha ya ha sido registrada previamente» si el técnico ya tiene libranza ese día (`LIbranza.cls:139-160`); `Registrar` lo convierte en error (`LIbranza.cls:196-202`).
- **Horas:** `Registrar` fija siempre 7,5 horas (`LIbranza.cls:224`).
- **Identificador:** `IDAusencia` = máximo numérico existente + 1 (`LIbranza.cls:34`; `FG.bas:408-420`).
- **Resumen:** uno registrado ⇒ «Ausencia registrada correctamente»; ninguno ⇒ aviso; varios ⇒ «De los N días… se ha registrado un total de M…» (`Tecnico.cls:489-496`).
- **Edición:** `ParaAlta = No` si hay libranza inicial (`Tecnico.cls:461-465`); `Registrar` hace `Edit` sobre la fila localizada por alias y fecha (`LIbranza.cls:219-222, 239`).
- `HaHabidoCambios` (`:151-184`) está definida pero este formulario no la llama.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Motivos del combo | `TbTecnicosAusencias` | SQL `SELECT DISTINCT` fija, recorrida con DAO | — | `:195-208` |
| Releer libranza | `TbTecnicosAusencias` | SQL dinámico concatenado con el ID | — | `Constructor.bas:4067-4116` |
| Día laborable | `TbTecnicosFiestas` | SQL dinámico concatenado con fecha `#mm/dd/yyyy#` | no | `Tecnico.cls:552-554` |
| Duplicado | `TbTecnicosAusencias` | SQL dinámico concatenado con alias y fecha | no | `LIbranza.cls:147-150` |
| Siguiente ID | `TbTecnicosAusencias` | DAO: recorre todos los IDs | no | `FG.bas:408-420` |
| Alta | `TbTecnicosAusencias` | DAO `AddNew`/`Update` | no (una escritura por día) | `LIbranza.cls:228-238` |
| Edición | `TbTecnicosAusencias` | DAO `Edit`/`Update`, SQL concatenado por alias y fecha | no | `LIbranza.cls:219-222, 239-250` |

## Coexistencia

Hay que adaptar: literales de fecha `#mm/dd/yyyy#` y `Format` con locale (`Tecnico.cls:555`; `LIbranza.cls:150`); fechas manejadas como `String` con `Format "dd/mm/yyyy"` (`Tecnico.cls:442-446`); `Max + 1` calculado en VBA recorriendo la tabla, sin transacción ni bloqueo (`FG.bas:408-420`); `.AddNew`/`.Edit` DAO sobre tabla o consulta; alias del técnico como clave de texto en el SQL; horas fijas 7,5 (`LIbranza.cls:224`); `CurrentDb()` frente a `getdb()` (ambos se usan: `LIbranza.cls:225`, `Form_FormTecnicoAusencias.cls:196`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El lote no es atómico: un error en un día intermedio deja registrados los anteriores y aborta el resto (`Tecnico.cls:459-483`).
- Un día con libranza ya existente no se omite, como sugiere el resumen (`Tecnico.cls:492-496`): `EsDiaLaborable` se llama sin considerar libranzas (`Tecnico.cls:470`) y `MotivoNoOK` devuelve error (`LIbranza.cls:158`), que aborta el lote entero.
- En edición, `Registrar` localiza la fila por alias y fecha nuevos (`LIbranza.cls:219-222`): cambiar la fecha o editar con un rango podría hacer `Edit` sobre un conjunto vacío y fallar; además no se arrastra `IDAusencia` a la libranza que se guarda (`Tecnico.cls:460-469`).
- En edición el formulario lanza `Alta` (`:47`) y nunca `Editado` (`:8`); sin manejador en `Form_FormTecnicoAusenciasGestion`, la lista podría no refrescarse.
- `HaHabidoCambios` no se usa: editar sin cambios vuelve a escribir.
- El mensaje de error de carga dice «Form_Open» (`:103`).
- Comparación de fechas como texto en `MotivoNoOK` (`LIbranza.cls:142`), dependiente del formato regional.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: registro del lote en una transacción con informe previo de días omitidos; edición por identificador y no por alias y fecha; calendario de ausencias con selección de rango; horas configurables.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin motivo, el registro se rechaza con «El motivo es obligatorio» y no escribe (`Tecnico.cls:450-453`).
2. Con `FechaInicial` vacía se muestra «Se ha de indicar la fecha inicial» (`Tecnico.cls:437-441`).
3. Con `FechaFinal` vacía se registra solo el día inicial (`Tecnico.cls:442-445`).
4. Una `FechaFinal` anterior a la inicial se rechaza (`Tecnico.cls:446-449`).
5. Un rango de varias fechas crea una libranza por día laborable, sin sábados, domingos ni fiestas (`Tecnico.cls:459-483, 528-553`).
6. Cada libranza nueva se guarda con 7,5 horas y el alias del técnico (`LIbranza.cls:224-238`).
7. Un día ya registrado para el técnico se informa como error (`LIbranza.cls:158`).
8. En edición el formulario muestra fecha, motivo y observaciones de la libranza seleccionada y el título «Edición de Libranza para <alias>» (`:129-143`).
9. Al guardar se muestra el resumen y el formulario se cierra (`:48-54`).

## Pendientes

- Si `TbTecnicosAusencias` tiene índice único sobre alias y fecha (ERD de origen).
- Contenido del evento `Alta` en el llamador (`Form_FormTecnicoAusenciasGestion.cls:7`): no hay manejador localizado.
- Cómo se comporta `.Edit` cuando el filtro por alias y fecha devuelve cero filas (`LIbranza.cls:219-239`).
- Acción prevista de `ComandoAyudaFormulario` (`Form_FormTecnicoAusencias.form.txt:379-380`).
