# Form_FormEquipoEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEquipoEdicion.cls` (602 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición de un equipo del catálogo (`TbEquipos`): nodo, BUI, subsistema, nombre de equipo, descripción y fecha de obsolescencia. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el `IDEquipo` por `OpenArgs`, carga los datos con DAO y guarda con su propia función pública `EditarIDEquipo` (`Form_FormEquipoEdicion.cls:360-526`). Capacidad: CAP-5 Equipos, equipos de medida y calibraciones (el desplegable de nodo, BUI y subsistema pertenece a CAP-12). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormEquipoEdicion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEquipoGestion` (`ComandoEditar_Click`) | Toma la columna 0 de `ListaFiltrados`; si no es numérica, avisa; cierra `FormEquipoEdicion` si ya estaba abierto y lo abre pasando el ID en `OpenArgs` | `Form_FormEquipoGestion.cls:68-92` |

El formulario no devuelve valores ni lanza eventos. Tras guardar (o al salir guardando) refresca `ListaFiltrados` de `FormIDEquiposGestion` si está abierto (`:71-73`, `:114-116`) y se cierra sin guardar el diseño (`:79`, `:117`). `Form_Close` vuelve a cerrarlo con `On Error Resume Next` (`:235-238`).

## Controles y eventos

Controles: 32 (32 Begin en el .form.txt): `Form` 1, plantillas de formato (Label, Rectangle, Image, CommandButton, OptionGroup, TextBox, ComboBox) 7, sección `Detalle` 1 y 23 controles de la sección.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboNodo` (1) | combo | `OnExit` → `ComboNodo_Exit` | Origen `TbNodos`. Al salir, si el nodo cambia vacía `ComboBui` y `ComboSubsistema`, y reconstruye el origen de `ComboBui` con `TbNodoBUI` filtrado por nodo (o `Like '*'` si está vacío) | CAP-12 | `Form_FormEquipoEdicion.form.txt:516-552`; `:158-186` |
| `ComboBUI` (1) | combo | `OnExit` → `ComboBui_Exit` | Si el BUI cambia vacía `ComboSubsistema`; reconstruye su origen con `TbSubsistemaBui` filtrado por BUI. El origen de diseño lleva un nodo literal | CAP-12 | `Form_FormEquipoEdicion.form.txt:553-591`; `:129-157` |
| `ComboSubsistema` (1) | combo | — | Origen `TbSubsistemaBui` (lo reasignan los `Exit` anteriores) | CAP-12 | `Form_FormEquipoEdicion.form.txt:353-390` |
| `Equipo` (1) | texto no enlazado | — | Nombre del equipo; recibe el foco al abrir | CAP-5 | `Form_FormEquipoEdicion.form.txt:592-620`; `:332` |
| `Descripcion` (1) | texto no enlazado | — | Descripción opcional | CAP-5 | `Form_FormEquipoEdicion.form.txt:290-317` |
| `FechaObsoleto` (1) | texto, formato `Short Date` | — | Fecha de obsolescencia; vacía = no obsoleto | CAP-5 | `Form_FormEquipoEdicion.form.txt:887-916` |
| `Numero_Actividad` (1) | texto bloqueado y deshabilitado | — | `ControlSource` `IDActividad` y valor por defecto con `DCount` sobre otra tabla; el código no lo usa | sin asignar | `Form_FormEquipoEdicion.form.txt:318-352` |
| `ComandoRegistrar` (1, rótulo «Registrar») | botón | `OnClick` | Valida cambios, guarda con `Registrar`, refresca y cierra | CAP-5 | `Form_FormEquipoEdicion.form.txt:439-476`; `:88-128` |
| `cmdSalir` (1, rótulo «Salir») | botón | `OnClick` | Si hay cambios pregunta si guardar; cierra | CAP-5 | `Form_FormEquipoEdicion.form.txt:477-515`; `:46-87` |
| `CuadroSuperior`, `Cuadro2` (2) | rectángulos | — | Decoración | sin asignar | `Form_FormEquipoEdicion.form.txt:391-438` |
| `Etiqueta142`, `Etiqueta196`, `Etiqueta197`, `Etiqueta198`, `Etiqueta199`, `Etiqueta4` (6) | etiquetas | — | Rótulos «NODO», «EQUIPO», «DESCR.», «BUI.», «SUBSIST.», «F.Obsoleto» | sin asignar | `Form_FormEquipoEdicion.form.txt:621-755, 917-940` |
| `lblTitulo1`, `lblTitulo2` (2) | etiquetas | — | «EDICIÓN DE EQUIPO»; `lblTitulo1` da título a la ventana | sin asignar | `Form_FormEquipoEdicion.form.txt:756-798`; `:247` |
| `Etiqueta99`, `Etiqueta1`, `Etiqueta2`, `Etiqueta3` (4) | etiquetas «*» | — | Marca visual de obligatorio | sin asignar | `Form_FormEquipoEdicion.form.txt:799-886` |

## Flujos

### Apertura (`Form_Open`, `:239-277`)

1. `AjustarTamaño Me` (`FG.bas:261`), título de ventana desde `lblTitulo1` y `strIDEquipo = Nz(OpenArgs, "")`; vacío ⇒ «No se han cargado todos los datos» (`:245-252`).
2. `EstablecerDatos` (`:278-359`): exige ID numérico (`:283-285`); lee `TbEquipos` por ID con SQL concatenado (`:286-289`); sin fila, mismo error (`:290-294`). Guarda en variables de módulo los valores iniciales de `NODO`, `BUI`, `SUBSISTEMA`, `Equipo` y `DESCRIPCION` (`:295-299`), reconstruye los orígenes de `ComboBui` y `ComboSubsistema` para ese nodo y BUI y rellena los controles (`:301-331`). La fecha se obtiene aparte con `Dame("TbEquipos","FechaObsoleto","IDEquipo",...)` (`:338`; `FG.bas:290-330`) y solo se muestra si es fecha (`:344-347`).
3. Cualquier error: `MsgBox` crítico y cierre del formulario (`:267-277`).

### Guardado (`ComandoRegistrar_Click`, `:88-128`)

1. `HaHabidoCambios` (`:8-45`) compara los seis campos con los valores iniciales; si no hay diferencias: «Ningún cambio que grabar» (`:93-98`). Si la propia función falla, devuelve «hay cambios» (`:43-44`).
2. `Registrar` (`:187-234`) exige `strIDEquipo` (`:204-206`), lee los controles con `Nz` (descripción y fecha vacías pasan como `"#"`, `:212-213`) y llama a `EditarIDEquipo` (`:221`).
3. Éxito: refresca `FormIDEquiposGestion` si está abierto y cierra (`:114-117`). Error: `MsgBox` con el texto y el formulario queda abierto (`:122-127`).

### Salir (`cmdSalir_Click`, `:46-87`)

Sin cambios, cierra. Con cambios, pregunta «¿Desea guardar los cambios producidos?»: Sí ⇒ mismo `Registrar` y refresco; No ⇒ cierra sin guardar. Si `Registrar` falla, muestra el error y no cierra (`:50-67`, `:80-86`).

## Reglas, validaciones y cálculos

- **Obligatorios** (`EditarIDEquipo`): ID, nodo, BUI, subsistema y equipo no vacíos ni `"#"` (`:389-408`).
- **Fecha de obsolescencia:** si se recibe valor (`"#"` = vaciar), llama a `MarcarDesMarcarEquipoComoObsoleto`, que exige ID numérico y fecha válida y escribe la fecha o `Null` (`:409-423`, `:527-597`). Esto ocurre **antes** de las comprobaciones de ubicación.
- **Cambio de ubicación:** si cambia nodo, BUI o subsistema, `EstaFranqueado("EQ", id)` recorre `TbEventos` por `IDEquipo` y devuelve «Sí» si algún evento tiene `Franqueado` igual a «Verdadero» (`FG.bas:6131-6150`); en ese caso se rechaza: «No se puede cambiar el nodo, bui o subsistema de un equipo que pertenece a un evento franqueado» (`:424-446`).
- **Compatibilidad:** la terna debe existir en `TbNodoBUI` unida a `TbSubsistemaBui`; si no, «El subsistema … no está registrado para el NODO … y el BUI …» (`:448-466`).
- **Escritura:** `NODO`, `BUI`, `SUBSISTEMA` e `IDEquipo` solo se escriben si `strEventoFranqueado = "No"`, es decir, si hubo cambio de ubicación (`:478-484`). Descripción: vacía no cambia nada, `"#"` la deja `Null` (`:485-491`).
- **Alias:** la rama de `AliasParaExcel` exige `strAlias`, que este formulario no pasa (`:221`, `:495-511`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura inicial del equipo | `TbEquipos` | DAO sobre SQL dinámico (ID concatenado) | — | `:286-289` |
| Lectura de la fecha | `TbEquipos` | SQL dinámico de `Dame` (reintenta con el valor entre comillas si falla) | — | `FG.bas:299-312` |
| Orígenes de combos | `TbNodos`, `TbNodoBUI`, `TbSubsistemaBui` | SQL dinámico concatenado con nodo/BUI entre comillas | — | `:169-180`, `:139-151`, `:305-319` |
| ¿Evento franqueado? | `TbEventos` | SQL dinámico (ID concatenado) | — | `FG.bas:6131-6135` |
| Compatibilidad nodo/BUI/subsistema | `TbNodoBUI`, `TbSubsistemaBui` | SQL dinámico con tres literales concatenados | — | `:451-456` |
| Escritura de fecha obsoleta | `TbEquipos` | DAO `Edit`/`Update` | no | `:558-579` |
| Escritura de ubicación y descripción | `TbEquipos` | DAO `Edit`/`Update` | no | `:469-492` |
| Alias (no alcanzable desde aquí) | `TbMantenimientoPreventivoEquipos` | DAO `Edit`/`Update` | no | `:497-509` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL Jet con comillas simples y ID concatenado; el valor de `Franqueado` se compara como texto «Verdadero» (`FG.bas:6141`), dependiente de cómo Access convierte un booleano a cadena; escrituras DAO sobre `TbEquipos` en dos pasos sin transacción; el campo `IDEquipo` se reescribe con el mismo valor (`:483`), lo que sobre un autonumérico en PostgreSQL vinculado habría que revisar.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El nombre `Equipo` se valida y se detecta como cambio (`:30`, `:211`), pero `EditarIDEquipo` nunca lo escribe en `TbEquipos` (`:469-492`): editar solo el nombre mostraría «Edición correcta» sin persistirlo.
- `HaHabidoCambios` formatea la fecha con `"ddd/mm/yyyy"` (`:34`) y la compara con el texto devuelto por `Dame` sin formatear (`:344`): con fecha informada podría dar siempre «hay cambios».
- La fecha de obsolescencia se guarda antes de validar la ubicación (`:409-423` frente a `:424-466`): si la validación falla, la fecha queda ya escrita (sin transacción).
- Si cambia solo el nombre, descripción o fecha, `strEventoFranqueado` queda vacío y se omite la escritura de ubicación, coherente con `:479`; pero con un cambio de ubicación la comprobación depende de que `EstaFranqueado` devuelva exactamente «No» (`:438-446`).
- El refresco apunta a `FormIDEquiposGestion` (`:71`, `:114`), mientras que el formulario se abre desde `FormEquipoGestion` (`Form_FormEquipoGestion.cls:83-87`): su lista no se refrescaría tras guardar.
- `FG.bas:6141` compara con «Verdadero», dependiente de la configuración regional (la rama `EV` admite también «True», `FG.bas:6086`).
- SQL concatenado con nodo/BUI/subsistema entre comillas simples (`:451-455`, `:305-319`): un apóstrofo en el valor rompería la consulta.
- Mensaje de error «No está registrado un equipo con ID» concatena el nombre, no el ID (`:476`); los textos de error citan otros métodos (`:264`, `:222`).
- `Numero_Actividad` conserva un valor por defecto con `DCount` ajeno al formulario (`Form_FormEquipoEdicion.form.txt:333`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: guardar en una sola transacción; que la restricción por evento franqueado sea una regla de servidor; listas dependientes nodo → BUI → subsistema desde el catálogo en lugar de SQL en el cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir un equipo existente muestra nodo, BUI, subsistema, nombre, descripción y fecha (`:278-347`).
2. Abrir sin ID o con ID inexistente muestra «No se han cargado todos los datos» y cierra (`:248-252`, `:290-294`).
3. Cambiar el nodo vacía BUI y subsistema, y las listas se limitan al nodo y al BUI elegidos (`:158-186`, `:129-157`).
4. Guardar sin cambios muestra «Ningún cambio que grabar» (`:93-98`).
5. Dejar nodo, BUI, subsistema o equipo vacío impide guardar con el mensaje correspondiente (`:389-408`).
6. Cambiar la ubicación de un equipo de un evento franqueado se rechaza (`:439-441`).
7. Una terna nodo/BUI/subsistema inexistente se rechaza (`:458-461`).
8. Vaciar la fecha deja el equipo como no obsoleto; una fecha no válida se rechaza (`:552-554`, `:571-579`).
9. «Salir» con cambios ofrece guardar; tras un guardado correcto el formulario se cierra (`:46-87`).

## Pendientes

- `FormIDEquiposGestion`: no existe como formulario en `src/forms`; solo aparece en citas y en una apertura dinámica desde otro formulario (`:71`, `:114`).
- Contenido de `AjustarTamaño` más allá de su cabecera (`FG.bas:261`) y comportamiento de `FormularioAbierto` (`FG.bas:171`) en el caso de error.
- Si `Franqueado` es booleano en `TbEventos` y, por tanto, qué texto produce `Nz` (`FG.bas:6141`), a contrastar con el ERD.
- Quién escribe `Equipo` en `TbEquipos` si no es este formulario (`:469-492`).
