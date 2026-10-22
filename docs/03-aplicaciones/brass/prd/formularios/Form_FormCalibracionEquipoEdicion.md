# Form_FormCalibracionEquipoEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormCalibracionEquipoEdicion.cls` (407 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición de un equipo calibrable (`TbEquiposCalibrables`): BUI, equipo, modelo, número de serie, observaciones y fecha de baja para calibración. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el identificador del equipo en `OpenArgs`, lee el registro con DAO y lo guarda con una función pública del propio formulario, `EdicionEquipoACalibrar`. Las citas son `Form_FormCalibracionEquipoEdicion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Capacidad de todo el formulario: CAP-5 Equipos, equipos de medida y calibraciones.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalibracionEquipoGestion` (botón «Editar Equipo») | Cierra el formulario si estaba abierto y lo abre pasando el ID del equipo seleccionado en `OpenArgs` | `Form_FormCalibracionEquipoGestion.cls:143-174` |

`OpenArgs` debe ser numérico; si no, error «No se han cargado todos los datos» y el formulario se cierra (`:224-228, 246-252`). Hacia fuera: tras guardar, si `FormCalibracionEquipoGestion` está abierto llama a su `ActualizarLista` (`:158-160, 199-201`) y se cierra (`:161, 207`). «Calibraciones» abre `FormCalibracionGestion` pasando solo el ID (`:113-116`).

## Controles y eventos

Controles: 33 (33 Begin en el .form.txt: 24 en la sección `Detalle`, más `Form`, `Section` y 7 plantillas de controles predeterminados).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Equipo` | texto no enlazado | — | Nombre del equipo; obligatorio al guardar | `Form_FormCalibracionEquipoEdicion.form.txt:463`; `:325-328` |
| `ComboBui` | combo `Table/Query` | — | BUI; origen `SELECT TbBUI.BUI FROM TbBUI ORDER BY TbBUI.BUI`; obligatorio | `Form_FormCalibracionEquipoEdicion.form.txt:633-635`; `:321-324` |
| `Modelo`, `NS`, `Observaciones` | textos no enlazados | — | Opcionales; vaciar el control los deja a `Null` | `Form_FormCalibracionEquipoEdicion.form.txt:413, 491, 661`; `:361-381` |
| `FechaBajaParaCalibracion` | texto con formato `Short Date` | — | Fecha de baja; si es válida se guarda, vacía se pone a `Null` | `Form_FormCalibracionEquipoEdicion.form.txt:840-841`; `:329-334, 382-388` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Valida cambios, guarda y cierra | `:128-173`; `Form_FormCalibracionEquipoEdicion.form.txt:356-358` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Pregunta si guardar cuando hay cambios y cierra | `:174-215`; `Form_FormCalibracionEquipoEdicion.form.txt:382-384` |
| `ComandoCalibraciones` («Calibraciones») | botón | `OnClick` | Abre la gestión de calibraciones del equipo | `:98-127`; `Form_FormCalibracionEquipoEdicion.form.txt:892-894` |
| `ComandoAyudaFormulario` | botón con imagen | — | Sin evento en el `.form.txt` | `Form_FormCalibracionEquipoEdicion.form.txt:753-754` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «EDICIÓN DE EQUIPO CALIBRABLE»; `lblTitulo1` pasa a `Caption` del formulario | `:223`; `Form_FormCalibracionEquipoEdicion.form.txt:713, 733` |
| `Etiqueta538`, `Etiqueta222`, `Etiqueta225`, `Etiqueta234`, `Etiqueta238`, `Etiqueta3` | etiquetas | — | Rótulos de campo (EQUIPO, MODELO, NÚM. SERIE, BUI, OBS., FECHA DE BAJA PARA CALIBRACIÓN) | `Form_FormCalibracionEquipoEdicion.form.txt:518-869` |
| `Etiqueta0`, `Etiqueta2` | etiquetas «*» | — | Marca de obligatorio | `Form_FormCalibracionEquipoEdicion.form.txt:793, 815` |
| `Cuadro69`, `CuadroSuperior`, `Cuadro2`, `Marco175` | rectángulos y grupo de opciones | — | Decoración | `Form_FormCalibracionEquipoEdicion.form.txt:288, 311, 334, 439` |
| `Form`, `Detalle` y 7 plantillas | formulario, sección, valores predeterminados | `OnLoad` → `Form_Load` | Ver Flujos | `:216-253`; `Form_FormCalibracionEquipoEdicion.form.txt:85` |

## Flujos

### Apertura (`Form_Load`, `:216-253`)

1. `AjustarTamaño Me` (`FG.bas:261`), `Caption` del formulario = `lblTitulo1` y `strIDEquipoCalibrable = Nz(OpenArgs, "")` (`:222-224`).
2. `EstablecerDatos` (`:7-62`) abre `SELECT TbEquiposCalibrables.* ... WHERE IDEquipoCalibrable=<id>` (`:15-18`). Sin fila: «No se han cargado todos los datos» (`:20-23`).
3. Guarda los valores iniciales en variables de módulo y rellena cada control solo si el valor no está vacío; la fecha, con formato `dd/mm/yyyy` (`:24-50`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:246-252`).

### Guardado (`ComandoRegistrar_Click`, `:128-173`)

1. `HaHabidoCambios` compara los seis controles con los valores iniciales; sin cambios, «No hay cambios que registrar» (`:139-142, 254-292`).
2. `Registrar` (`:63-97`) convierte cada control vacío en `"#"` (`:71-76`) y llama a `EdicionEquipoACalibrar` (`:82-83`).
3. Con éxito refresca la gestión si está abierta y cierra (`:158-161`). Con error muestra `MsgBox` y el formulario queda abierto (`:166-173`).

### Salir (`cmdSalir_Click`, `:174-215`)

Con cambios pregunta «¿Desea guardar los cambios producidos?»; «Sí» ejecuta `Registrar` y refresca la gestión; «No» descarta. Después cierra sin guardar el objeto (`:178-207`). Si `Registrar` falla, el error se muestra y el formulario no se cierra (`:190-194, 209-214`).

## Reglas, validaciones y cálculos

- **Obligatorios:** BUI y Equipo, vacíos o `"#"` ⇒ «Es obligatorio indicar el BUI» / «el Equipo» (`:321-328`).
- **Unicidad:** no puede existir otro equipo (distinto ID) con el mismo `BUI` y `Equipo`; si existe, «Para <BUI> el equipo <Equipo> ya ha sido registrado previamente» (`:335-345`).
- **Fecha de baja:** si no está vacía ni es `"#"`, debe cumplir `IsDate` («La fecha de baja de calibración no es válida») (`:329-334`).
- **Semántica de `"#"`:** en `Modelo`, `NS` y `Observaciones`, `"#"` escribe `Null`, un valor no vacío lo escribe y vacío no toca el campo (`:361-381`); en la fecha, `"#"` escribe `Null` (`:382-388`).
- **Cambios:** se comparan como texto; la fecha se normaliza a `dd/mm/yyyy` (`:279-284`). Un error al calcularlo devuelve «hay cambios» (`:290-291`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga del equipo | `TbEquiposCalibrables` | SQL dinámico con el ID concatenado (`OpenRecordset`) | — | `:15-18` |
| Comprobación de duplicado | `TbEquiposCalibrables` | SQL dinámico con `BUI` y `Equipo` concatenados entre comillas | — | `:335-340` |
| Relectura previa a editar | `TbEquiposCalibrables` | SQL dinámico con el ID concatenado | — | `:349-352` |
| Edición | `TbEquiposCalibrables` | DAO `Edit`/`Update` | no | `:358-389` |
| Lista de BUI | `TbBUI` | `RowSource` de diseño | — | `Form_FormCalibracionEquipoEdicion.form.txt:635` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: el SQL concatena cadenas con comilla simple (`:338-339`); la fecha se asigna como texto `dd/mm/yyyy` a un campo de fecha (`:383`), dependiente de la configuración regional; el `Edit`/`Update` DAO sobre tabla vinculada exige clave primaria visible; no hay transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un valor literal `"#"` tecleado en `Modelo`, `NS` u `Observaciones` se confunde con el marcador de «vaciar» y guarda `Null` (`:71-76, 361-381`).
- La comprobación de duplicado y la carga concatenan `BUI` y `Equipo` sin escapar (`:338-339`): un apóstrofo rompería la consulta.
- Comprobar duplicado y guardar son dos operaciones sin bloqueo (`:335-358`): dos usuarios podrían crear el mismo par BUI/equipo.
- `ComandoCalibraciones` pasa solo el ID a `FormCalibracionGestion` (`:116`), mientras que la gestión de equipos pasa `ID|texto` (`Form_FormCalibracionEquipoGestion.cls:117`); a verificar que el destino admita ambas formas.
- `On Error Resume Next` en el manejador de `ComandoRegistrar_Click` (`:169`) puede ocultar fallos posteriores al mensaje.
- El código del formulario no consulta permisos de usuario (`:1-407`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: restricción única `BUI`+`Equipo` en la base; valores nulos explícitos en lugar del marcador `"#"`; fecha enviada en formato ISO.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin identificador numérico muestra «No se han cargado todos los datos» y cierra (`:224-252`).
2. Los campos se cargan con los valores del equipo; los vacíos quedan en blanco (`:24-50`).
3. Guardar sin cambios muestra «No hay cambios que registrar» (`:139-142`).
4. Guardar sin BUI o sin Equipo muestra el mensaje de obligatoriedad y no escribe (`:321-328`).
5. Guardar con un par BUI/Equipo ya usado por otro equipo muestra el aviso de duplicado (`:335-345`).
6. Una fecha de baja no válida se rechaza; vaciarla la deja a `Null` (`:329-334, 385-386`).
7. Tras guardar, la lista de la gestión se refresca y el formulario se cierra (`:158-161`).
8. Salir con cambios ofrece guardar; «No» cierra sin escribir (`:178-207`).

## Pendientes

- Índice único, longitudes y tipo de `FechaBajaParaCalibracion` en `TbEquiposCalibrables` (ERD de origen).
- Contenido de `FormCalibracionGestion` al recibir `OpenArgs` sin `|` (`:116`).
- Contenido de `AjustarTamaño` más allá de `FG.bas:261-288`.
