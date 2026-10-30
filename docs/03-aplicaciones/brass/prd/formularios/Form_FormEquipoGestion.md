# Form_FormEquipoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEquipoGestion.cls` (838 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión del catálogo de equipos (`TbEquipos`): filtra por BUI, subsistema, equipo y «en mantenimiento», lista el resultado en `ListaFiltrados` y lanza alta, edición, baja, exportación a Excel y accesos a eventos y planificación. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`; la lista se rellena asignando `RowSource` por código, `Form_FormEquipoGestion.cls:746`). Las citas son `Form_FormEquipoGestion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Controles: 37 (37 Begin en el .form.txt): 27 controles, el `Form`, la sección `Detalle` y 8 plantillas de control por defecto (`Form_FormEquipoGestion.form.txt:738-862`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form0BDOpciones` | etiqueta `lblGestionEquipos`: `DoCmd.OpenForm` sin argumentos, con `On Error Resume Next` | `Form_Form0BDOpciones.cls:136-139` |
| `Form0BDOpcionesOtras` | etiqueta `lblGestionIDEquipos`: ídem | `Form_Form0BDOpcionesOtras.cls:77-80` |
| `FormEquipoAlta` | tras guardar, si este formulario está abierto, llama a `ActualizarLista` | `Form_FormEquipoAlta.cls:35-37` |

No lee `OpenArgs`. Hacia fuera abre `FormEquipoAlta` (`:46`), `FormEquipoEdicion` con el `IDEquipo` como argumento (`:87`), `FormPlanificacionAltaEquipo` con `IDEquipo|Equipo` (`:213`) y `FormEventoGestion` (`:266, :361`). No declara eventos propios. Cada apertura cierra antes la instancia previa si estaba abierta (`:43-45, :83-85, :210-212`). `cmdSalir` cierra sin guardar (`:601-604`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `Form` | formulario | `OnOpen` → `Form_Open` | Ajusta tamaño, fija título, vacía combos dependientes y carga la lista | CAP-5 | `:565-600`, `Form_FormEquipoGestion.form.txt:55` |
| `Detalle` | sección | — | Única sección con controles | CAP-5 | `Form_FormEquipoGestion.form.txt:883` |
| 8 plantillas por defecto (Label, Rectangle, Image, CommandButton, OptionGroup, TextBox, ListBox, ComboBox) | plantillas | — | Valores por defecto de diseño, sin comportamiento | sin asignar | `Form_FormEquipoGestion.form.txt:738-862` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `Form_FormEquipoGestion.form.txt:894, 916` |
| `Marco122` | grupo de opciones | — | Marco decorativo, sin eventos | sin asignar | `Form_FormEquipoGestion.form.txt:1019` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Rótulo «GESTIÓN DE LOS EQUIPOS»; `lblTitulo1` da el `Caption` de la ventana | CAP-5 | `:573`, `Form_FormEquipoGestion.form.txt:1106, 1127` |
| `Etiqueta142`, `Etiqueta267`, `Etiqueta268`, `Etiqueta273` | etiquetas | — | Rótulos «Mto.», «BUI», «SUBS.», «Equipo» | CAP-5 | `Form_FormEquipoGestion.form.txt:1410, 1437, 1464, 1530` |
| `ListaFiltrados` | lista, 6 columnas con cabecera (anchos `0;3402;3402;4536;0;0`) | `OnClick` → `ListaFiltrados_Click` | Columnas: 0 `IDEquipo`, 1 `BUI`, 2 `SUBSISTEMA`, 3 `Equipo`, 4 id de planificación, 5 `NODO`. Habilita o deshabilita los botones según haya fila elegida | CAP-5 | `:606-642`, `Form_FormEquipoGestion.form.txt:940-962`; `:665-668` |
| `ComboBUIBusqueda` | combo (`TbBUI`) | `OnChange`, `OnExit` | Si cambia el valor: reconstruye subsistemas y refresca la lista | CAP-12 | `:377-423`, `Form_FormEquipoGestion.form.txt:1378-1383` |
| `ComboSubsistemaBusqueda` | combo (`TbSubsistemaBui`) | `OnChange`, `OnExit` | Si cambia: reconstruye equipos y refresca | CAP-12 | `:517-563`, `Form_FormEquipoGestion.form.txt:1340-1346` |
| `ComboEquipoBusqueda` | combo (`TbEquipos.Equipo`) | solo `OnExit` | Si cambia, refresca. El handler `ComboIDEquipoBusqueda_Change` (`:471`) no corresponde a ningún control con ese nombre | CAP-5 | `:494-515`, `Form_FormEquipoGestion.form.txt:1498-1503` |
| `ComboEnMantenimientoBusqueda` | combo lista de valores `Sí;No` | `OnChange`, `OnExit` | Si cambia, refresca | CAP-5 | `:425-469`, `Form_FormEquipoGestion.form.txt:1300-1305` |
| `ComandoLimpiarBUI` / `ComandoLimpiarSubSistema` / `ComandoLimpiarIDEquipo` / `ComandoLimpiarEnMantenimiento` | botones | `OnClick` | Vacían su combo (si tiene valor), refrescan y reconstruyen el combo dependiente | CAP-5 | `:281-342`, `Form_FormEquipoGestion.form.txt:1734, 2094, 1914, 1554` |
| `ComandoActualizarTodo` | botón | `OnClick` | Ejecuta los cuatro «limpiar»; no recarga si todos los combos están vacíos | CAP-5 | `:12-32`, `Form_FormEquipoGestion.form.txt:1147-1149` |
| `ComandoAltaEquipo` | botón «Alta Equipo» | `OnClick` | Abre `FormEquipoAlta`; con más de una fila precarga Nodo, BUI y Subsistema de la primera fila de datos | CAP-5 | `:33-65`, `Form_FormEquipoGestion.form.txt:2274-2276` |
| `ComandoEditar` | botón | `OnClick` | Abre `FormEquipoEdicion` con el `IDEquipo`; error si no es numérico | CAP-5 | `:68-98`, `Form_FormEquipoGestion.form.txt:2314-2316` |
| `ComandoEliminar` | botón | `OnClick` | Comprueba referencias, confirma y borra el equipo | CAP-5 | `:99-193`, `Form_FormEquipoGestion.form.txt:2355-2357` |
| `ComandoEstablecerParaPlanificacion` | botón | `OnClick` | Abre `FormPlanificacionAltaEquipo` con `IDEquipo|Equipo` | CAP-6 | `:194-224`, `Form_FormEquipoGestion.form.txt:2478-2480` |
| `ComandoHistoricoDeEquipo` | botón | `OnClick` | Abre `FormEventoGestion` filtrado por el equipo | CAP-3 | `:248-280`, `Form_FormEquipoGestion.form.txt:2438-2440` |
| `ComandoMantenimientosRealizados` | botón | `OnClick` | Ídem, con tipo de evento `MANT. PROGRAMADO` | CAP-3 | `:343-376`, `Form_FormEquipoGestion.form.txt:2398-2400` |
| `ComandoExportarExcel` | botón | `OnClick` | Exporta a Excel el `RowSource` actual de la lista | CAP-5 | `:225-247`, `Form_FormEquipoGestion.form.txt:1052-1054` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra el formulario | CAP-5 | `:601-604`, `Form_FormEquipoGestion.form.txt:988-990` |

## Flujos

### Apertura (`Form_Open`, `:565-600`)

1. `AjustarTamaño Me` (`:571`; `FG.bas:261-288`) y título tomado de `lblTitulo1` (`:573`).
2. Deja `ComboEquipoBusqueda` y `ComboSubsistemaBusqueda` con un `RowSource` centinela (`'-1'`) que no devuelve filas (`:574-585`).
3. `ActualizarLista` (`:586`): con los cuatro filtros vacíos, `Filtrar` devuelve `IDEquipo=-1`, es decir, lista vacía (`:669-674`).
4. Error: `MsgBox` y cierre del formulario (`:593-599`).

### Filtrado (`ActualizarLista`, `:715-778`)

1. Deshabilita los seis botones de acción (`:725-731`), llama a `Filtrar` con los cuatro combos y, si devuelve `…|…`, muestra el error (`:737-744`).
2. Asigna el SQL al `RowSource` y hace `Requery` (`:745-747`). Con 0 filas desactiva la exportación; con 1 fila la selecciona y simula el clic (`:753-756`); con varias, recorre las seleccionadas y simula el clic en la primera (`:758-764`); con alguna fila activa la exportación (`:766`).
3. Cascada: cambiar BUI reconstruye subsistemas y vacía subsistema y equipo (`:384-387, :787-803`); cambiar subsistema reconstruye equipos (`:524-527, :804-825`). Si el combo reconstruido queda con una sola opción, la preselecciona (`:801-803, :822-824`).
4. Cada combo recuerda su último valor (`strBUIAnterior`, etc.; `:4-5`) para no refrescar si no ha cambiado; `LimpiarUltimos` los reinicia (`:6-11`).

### Baja (`ComandoEliminar_Click`, `:99-193`)

1. Exige fila elegida (`:105-109`).
2. Con `Dame` (`FG.bas:290-330`) busca el `IDEquipo` en `TbEventos` (`:115`), `TbPlanificacion` (`:130`) y `TbMaterial` (`:145`); si hay coincidencia aborta con «Al menos forma parte de un Evento», «…de una Planificación» o «Al menos hay registrado un material de este equipo» (`:121-124, :136-139, :151-154`).
3. Pregunta «¿Desea eliminar el equipo seleccionado?» con «No» por defecto (`:158-161`).
4. Abre `TbEquipos` por `IDEquipo`, borra con DAO y refresca la lista (`:165-178`). Si no existe, «No existe Equipo con IDEquipo» (`:170-173`).

### Accesos a eventos (`:248-280, :343-376`)

Si `FormEventoGestion` está abierto llama a su `ComandoActualizarTodo_Click` (`Form_FormEventoGestion.cls:27`), si no lo abre; fija `ComboEquipoBusqueda` y dispara `ComboEquipoBusqueda_AfterUpdate` (`:268-269`; `Form_FormEventoGestion.cls:525`). En «Mtos. Realizados» fija además `ComboTipoEventoBusqueda` y dispara `ComboIDEventoBusqueda_AfterUpdate` (`:363-365`; `Form_FormEventoGestion.cls:551`).

### Errores

Patrón común: error de negocio con `Err.Raise 1000` y mensaje previo (`:77, :108`); cualquier otro error compone «…error nº …» (`:93-95`); en ambos casos `MsgBox` crítico y fin del cursor de espera (`:96-97`). Los cuatro «limpiar» usan `On Error Resume Next` y no informan de fallos (`:282, :298, :313, :328`).

## Reglas, validaciones y cálculos

- Filtro por BUI y subsistema por igualdad; por equipo, `Like '*texto*'` (`:676-689`). Vacío equivale a `Like '*' Or Is Null` (`:676, :681, :686`).
- Filtro «en mantenimiento» (`:690-698`): con valor no vacío (`Sí` o `No`) aplica `Like '*' Or Is Null`, es decir, no filtra; con valor vacío aplica `TbPlanificacionEquipos.IDEquipo Is Null`. La rama `"Sí"` es inalcanzable (`:693`).
- Sin ningún filtro, la lista queda vacía (`:669-674`).
- Botones de acción activos solo con fila elegida (`:612-630`); `Eliminar` y `Editar` además validan al pulsar (`:75, :106`).
- Integridad en la baja: solo `TbEventos`, `TbPlanificacion` y `TbMaterial` (`:115, :130, :145`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de equipos | `TbEquipos` LEFT JOIN `TbPlanificacionEquipos` | SQL dinámico concatenado (`RowSource`) | — | `:665-704, :745-746` |
| Combo BUI | `TbBUI` | `RowSource` de diseño | — | `Form_FormEquipoGestion.form.txt:1380` |
| Combo subsistema | `TbSubsistemaBui` | SQL dinámico concatenado | — | `:580-584, :795-799` |
| Combo equipo | `TbEquipos` | SQL dinámico concatenado | — | `:574-578, :815-820` |
| Comprobación de referencias | `TbEventos`, `TbPlanificacion`, `TbMaterial` | DAO con SQL concatenado (`Dame`, con reintento entrecomillado si falla) | no | `FG.bas:300-311`; `:115, :130, :145` |
| Baja de equipo | `TbEquipos` | DAO `OpenRecordset` + `.Delete` | no | `:165-175` |
| Exportación | `RowSource` de la lista | DAO `OpenRecordset` + Excel; fichero `.xlsx` en el directorio local configurado | no | `FG.bas:1606-1625`; `:231-232` |

## Coexistencia

Adaptar a PostgreSQL vinculado: comodines `Like '*…*'` y `Like '*'` (`:676-696`) pasan a `%` y distinguen mayúsculas (se requeriría `ILIKE`); `Is Null` y `LEFT JOIN` son portables; `IDEquipo` se concatena sin comillas como número (`:167`) y `Dame` depende de que falle la primera consulta para reintentar con comillas (`FG.bas:303-311`); sentinelas `'-1'` en columnas de texto (`:576, :582`). `Delete` sobre dynaset de tabla vinculada requiere clave primaria. No inserta, así que no interviene el autonumérico. Pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Filtro «en mantenimiento» invertido o inoperante: con combo vacío solo salen equipos sin planificación y con `Sí`/`No` no se filtra (`:690-698`).
- `ComboEquipoBusqueda` no tiene `OnChange` porque el handler se llama `ComboIDEquipoBusqueda_Change` (`:471`; `Form_FormEquipoGestion.form.txt:1503`).
- `ComandoAltaEquipo` precarga Nodo, BUI y Subsistema de la primera fila de la lista y no de los filtros activos (`:47-52`).
- Mensajes con «gasto» en alta y edición, copiados de otro formulario (`:61, :76, :94`).
- La baja no comprueba `TbPlanificacionEquipos`, que la propia lista une (`:130, :667`); ni hay transacción entre comprobaciones y borrado (`:115-175`).
- `ComandoMantenimientosRealizados` dispara `ComboIDEventoBusqueda_AfterUpdate` y no el del equipo (`:365`).
- `ComandoLimpiarBUI` refresca la lista antes de reconstruir los combos dependientes (`:291-292`).
- `ListaFiltrados.RowSource` de diseño referencia `TbIDEquipos.EquipoID` y `EquipoID` (`Form_FormEquipoGestion.form.txt:957-960`), sustituido al abrir (`:586`).
- Filtros concatenados con comillas simples: un apóstrofo en el texto rompería la consulta (`:683, :688`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros parametrizados, exportación en servidor, comprobación de referencias por restricciones de la base.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin filtros, la lista está vacía (`Filtrar`, `:669-674`).
2. Filtrar por BUI limita subsistemas a los de ese BUI y vacía subsistema y equipo (`EstablecerCombo`).
3. Filtrar por texto de equipo devuelve los que lo contienen (`:688`).
4. Sin fila elegida están deshabilitados Editar, Eliminar, Histórico, Planificación y Mtos. Realizados (`ListaFiltrados_Click`).
5. Eliminar un equipo con eventos, planificación o material muestra el mensaje correspondiente y no borra (`:121-154`).
6. Eliminar tras confirmar borra el equipo y actualiza la lista (`:158-178`).
7. «Histórico» abre los eventos del equipo elegido (`:263-269`).
8. El resultado del filtro «en mantenimiento» se define con el propietario antes de migrarlo (`:690-698`).

## Pendientes

- Resto de `GenerarConsultasSQL` tras `FG.bas:1625` (generación y apertura del Excel).
- Contenido de `FormularioAbierto` (`FG.bas:171-185`, parcial).
- Si hay clave o restricciones en `TbEquipos` que impidan el borrado (ERD de origen).
- Cómo se comporta la selección de `ListaFiltrados` tras `Requery` (`:758-764`).
