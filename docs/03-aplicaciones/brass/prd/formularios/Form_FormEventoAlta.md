# Form_FormEventoAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoAlta.cls` (880 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un evento en `TbEventos`: el usuario elige nodo/BUI, subsistema, equipo, tipo, criticidad, técnico, originador, fecha y descripción, más los datos SLA, y el formulario genera el `IDEvento` y lo graba por DAO. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) cuya lógica de negocio vive en el propio formulario (`Registrar`, `Alta`, `EstablecerDatos`). Capacidad: CAP-3 Eventos (los equipos de medida del evento, CAP-3/CAP-5). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`, `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son `Form_FormEventoAlta.cls:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` | `ComandoAltaEvento_Click` abre `FormEventoAlta` sin `OpenArgs` (alta vacía) | `Form_FormEventoGestion.cls:1131-1134` |
| Con `OpenArgs` = `IDEvento` | Precarga los datos de ese evento; rótulo «Alta de Evento de …». Ningún llamador del `src` pasa este argumento (ver Pendientes) | `:375-390` |

Salida: `ComandoAlta` devuelve el `IDEvento` generado en un `MsgBox` y cierra el formulario (`:131-142`). Si `FormEventoGestion` está abierto, lo refresca con `ComandoActualizarTodo_Click`, fija `ComboIDEventoBusqueda` y ejecuta su `AfterUpdate` (`:116-120`); si además su `ComandoAltaActividad` está habilitado, ofrece dar de alta una actividad (`:129-135`). `cmdSalir` cierra sin guardar (`:358-361`). Abre `FormEquipoAlta` con `Nodo|BUI|SUBSISTEMA` (`:203-204`) y `FormEventoEqMedida` (`:162-167`).

## Controles y eventos

Controles: 50 (62 `Begin` en el `.form.txt` = 1 `Form` + 8 plantillas de estilo + 3 secciones + 50 controles).

| Control | Tipo | Evento y handler | Qué hace / validación | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboBUI` | combo, `LimitToList`, `OBLIGATORIO` | `OnEnter`, `AfterUpdate` | `RowSource` `TbNodoBUI`; col. 0 = BUI, col. 1 = nodo. Al cambiar de BUI vacía subsistema y equipo, recarga subsistemas y contactos y habilita/deshabilita `+` y equipos de medida | CAP-12 | `.form.txt:750-788`; `:217-281, 284-287` |
| `ComboSubsistema` | combo, `LimitToList`, `OBLIGATORIO` | `OnEnter`, `OnChange` | `RowSource` en ejecución con `EstablecerComboSubSistema`; al cambiar vacía el equipo y recarga `ComboIDEquipo` | CAP-12 | `.form.txt:713-749`; `:299-353` |
| `ComboIDEquipo` | combo 3 columnas (ID;equipo;obsolescencia), `LimitToList`, `OBLIGATORIO` | `OnEnter` (requery) | `RowSource` en ejecución con `EstablecerComboEquipo` | CAP-5 | `.form.txt:965-1001`; `:290-293` |
| `ComandoNuevoIDEquipo` («+») | botón, deshabilitado en diseño | `OnClick` | Exige nodo, BUI y subsistema; cierra `FormEquipoAlta` si estaba abierto y lo abre con `Nodo\|BUI\|SUBSISTEMA` | CAP-5 | `.form.txt:1002-1041`; `:182-215` |
| `ComboTipoEvento` | combo, `LimitToList`, `OBLIGATORIO` | `OnEnter` | `RowSource` `TbTipoEvento` | CAP-3 | `.form.txt:818-856`; `:354-357` |
| `ComboCriticidad` | combo 2 columnas, `LimitToList`, `OBLIGATORIO` | — | `RowSource` `TbCriticidad`; valor 3 por defecto en alta vacía | CAP-3 | `.form.txt:857-895`; `:393` |
| `ComboOriginador` | combo, `LimitToList`, `OBLIGATORIO` | `OnEnter` | `RowSource` `TbOriginador` | CAP-12 | `.form.txt:896-934`; `:295-298` |
| `ComboTecnico` | combo 3 columnas (ALIAS;NOMBRE;TIPO), `LimitToList`, `OBLIGATORIO` | — | Se rellena en `Form_Load` con `RellenarComboTecnicos`; se guarda la col. 0 (alias) | CAP-9 | `.form.txt:1813-1850`; `:398`, `FG.bas:7224-7260` |
| `FechaAlta` | texto, `OBLIGATORIO` | — | Fecha de alta; por defecto la de hoy | CAP-3 | `.form.txt:935-964`; `:394` |
| `DescripcionEvento` | texto, `OBLIGATORIO` | — | Descripción del evento | CAP-3 | `.form.txt:649-677` |
| `PMPR` | texto | — | Dato PM/PR; solo se guarda si no está vacío | CAP-3 | `.form.txt:789-817`; `:708-710` |
| `Contacto` | combo editable | `OnEnter` (`Me.Requery`) | Contactos ya usados en eventos del mismo BUI; `LimitToList` solo sin BUI | CAP-3 | `.form.txt:1851-1888`; `:363-365, 803-829` |
| `FechaRecepcionNotificacion`, `FechaInicioContactoCliente` | texto | — | Datos SLA; se graban como fecha o `Null` | CAP-3 | `.form.txt:1564-1648`; `:728-738` |
| `cmbIncidenciaAveriaOReparacion`, `cmbEventoConServicioAfectado` | combos «Sí;No», `LimitToList` | — | Se graban como booleano o `Null` | CAP-3 | `.form.txt:1676-1770`; `:740-765` |
| `chkUrgente` | casilla | — | `SI`/`NO`/vacío a `Urgente` | CAP-3 | `.form.txt:1798-1812`; `:41-47` |
| `ComandoEquiposDeMedida` | botón | `OnClick` | Abre `FormEventoEqMedida`; su rótulo muestra «Alta equipos de Medida (n)» | CAP-3 | `.form.txt:1916-1953`; `:154-180, 831-880` |
| `ComandoAlta` (pie) | botón | `OnClick` | Valida, graba y cierra | CAP-3 | `.form.txt:1987-2022`; `:92-151` |
| `cmdSalir` (pie) | botón | `OnClick` | Cierra sin guardar | CAP-3 | `.form.txt:2023-2060`; `:358-361` |
| `ComandoAyudaFormulario` (cabecera) | botón | `OnClick` declarado, sin handler | No hace nada (ver defectos) | sin asignar | `.form.txt:587-610` |
| `Numero_Actividad` | texto residual, deshabilitado, ancho 6 | — | `ControlSource` `IDActividad`, `DefaultValue` con `DCount`; sin uso en el código | sin asignar | `.form.txt:678-712` |
| 28 etiquetas: `lblTitulo`, `lblSLA`, `Etiqueta13`, 11 rótulos de campo (`Etiqueta206-244`), 9 asteriscos y 5 rótulos SLA/booleanos (`lbl…`) | etiquetas | — | Rótulos y marca visual de obligatoriedad; solo `lblTitulo` se escribe en `Form_Load` | sin asignar | `.form.txt:576, 1042-1536, 1889-1975`; `:389, 392` |

Cuenta: 22 controles con fila propia + 28 etiquetas = 50. Los 8 `Begin` de estilo (etiqueta, rectángulo, imagen, botón, casilla, grupo de opciones, texto, combo) son plantillas por defecto y no controles (`.form.txt:430-536`).

## Flujos

### Apertura (`Form_Load`, `:367-418`)
1. `AjustarTamaño Me` (`:373`; `FG.bas:261`). Con `OpenArgs` llama a `EstablecerDatos` y enfoca `FechaAlta` (`:375-390`); sin él, criticidad 3, fecha de hoy y foco en `ComboBUI` (`:391-397`).
2. Rellena `ComboTecnico` con `RellenarComboTecnicos Me.ComboTecnico, No, Sí` y copia el título a `Me.Caption` (`:398-404`).
3. Error: `MsgBox` crítico y cierre del formulario (`:411-417`).

### Precarga (`EstablecerDatos`, `:420-548`)
Lee `TbEventos` por `IDEvento` con SQL concatenado (`:431-434`); si no existe, «No se ha indicado un IDEvento existente» (`:436-439`). Copia BUI, subsistema, PM/PR, tipo, criticidad, técnico, originador, descripción y fecha (`:457-484`). `Contacto` y `Notas` están retirados de la precarga (`:486-487`). Muestra el equipo por su nombre vía `Dame("TbEquipos","Equipo",…)` (`:488-501`) y recompone los `RowSource` de subsistema y equipo (`:503-532`).

### Alta (`ComandoAlta_Click` → `Registrar` → `Alta`, `:92-151, 7-91, 549-801`)
1. `Registrar` recoge los controles; si `ComboIDEquipo` no es numérico, busca el ID por nodo+BUI+subsistema+nombre (`:49-65`; `FG.bas:5279-5330`). Llama a `Alta` (`:72-75`).
2. `Alta` valida (reglas abajo), calcula el `IDEvento` con `DameIDEvento` (o usa `strIDEventoManual`, no informado desde este formulario), comprueba que no exista y hace `AddNew`/`Update` (`:655-770`).
3. Registra los equipos de medida del global `m_ColEquiposMedidaPorEventoActivos` con `Evento.RegistrarEquipoMedida` (`:777-788`; `Evento.cls:1859-1927`).
4. Errores: devuelven `#ERR|texto` y `ComandoAlta_Click` muestra `MsgBox` crítico; el formulario queda abierto (`:77-81, 105-109, 144-150`).

### Equipos de medida (`:154-180, 831-880`)
Abre `FormEventoEqMedida` y se suscribe a sus eventos `Alta` y `Eliminado` (`Form_FormEventoEqMedida.cls:9-10, 42, 87`) para actualizar el contador del botón.

## Reglas, validaciones y cálculos

- **Obligatorios en `Alta`:** nodo, BUI, subsistema, `IDEquipo`, tipo de evento, criticidad numérica, técnico, originador, descripción y fecha válida, cada uno con su mensaje (`:595-635`). Los asteriscos y la etiqueta `OBLIGATORIO` son solo visuales (`.form.txt:662`).
- **Equipo obsoleto:** si `TbEquipos.FechaObsoleto` es anterior a la fecha de alta, «El equipo seleccionado para el evento es obsoleto desde el …» (`:642-652`).
- **`IDEvento`:** iniciales de `TbBuiIDEvento` (o 3 primeras letras del BUI, mínimo 3) + `aamm` de la fecha + ordinal de 3 cifras = máximo existente + 1 (`FG.bas:6205-6242`).
- **Valores fijos:** `HORAINICIALEVENTO` = «7:30» y `TIEMPORESPUESTAEVENTO` = «00:00» (`:716-717`); `FechaRegistroAlta` y `FechaRegistroModificacion` = `Now()` (`:767-768`).
- **Contacto y notas:** `Notas` siempre vacía (`:37`); `Contacto` se envía sin validar (`:36, 720-722`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Subsistemas por BUI | `TbSubsistemaBui` | SQL dinámico concatenado | — | `FG.bas:1214-1218` |
| Equipos por BUI+subsistema | `TbEquipos` | SQL dinámico concatenado | — | `FG.bas:1198-1200` |
| Contactos por BUI | `TbEventos` | SQL dinámico concatenado | — | `:816-821` |
| ID de equipo por nombre | `TbEquipos` | SQL dinámico concatenado | — | `FG.bas:5316-5322` |
| Lectura genérica (obsolescencia, existencia de ID, nombre de equipo) | varias | `Dame`: SQL concatenado con reintento entre comillas | — | `FG.bas:300-311` |
| Ordinal del `IDEvento` | `TbEventos` | SQL dinámico con `Like` | no | `FG.bas:6225-6229` |
| Precarga | `TbEventos` | SQL dinámico, DAO | — | `:431-434` |
| Alta del evento | `TbEventos` | DAO `AddNew`/`Update` | no | `:694-771` |
| Alta de equipos de medida | `TbEventosEquipoMedida` | DAO `AddNew`/`Update` por equipo | no | `Evento.cls:1895-1914` |
| Combos de tipo, criticidad, originador, BUI | `TbTipoEvento`, `TbCriticidad`, `TbOriginador`, `TbNodoBUI` | consulta en `RowSource` | — | `.form.txt:770, 839, 880, 917` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `IDEvento` textual calculado en VBA con máximo+1 fuera de transacción (`FG.bas:6225-6242`); SQL concatenado con comillas simples en casi todas las lecturas; operador `Like` con `*` (`FG.bas:6227`); valores `Date` y `Now()` escritos desde VBA (`:767`); booleanos Access (`True`/`False`) mapeados a `Null` (`:740-765`); `IDEquipoMedida` insertado sin comillas (`Evento.cls:1898`); `DameID` para el id de `TbEventosEquipoMedida` (`Evento.cls:1906`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El `IDEvento` se calcula con máximo+1 sin transacción ni bloqueo (`FG.bas:6225-6242`; `:694-703`): dos altas simultáneas del mismo BUI y mes podrían coincidir; la segunda recibiría «ha devuelto un IDEvento ya existente» (`:699-702`).
- El alta del evento y el de sus equipos de medida no comparten transacción; el error de `RegistrarEquipoMedida` (`p_Error`) no se lee (`:782`), por lo que un fallo dejaría el evento sin sus equipos sin aviso.
- `m_ColEquiposMedidaPorEventoActivos` es global y solo se vacía en `Variables Globales.bas:291`; ni `Alta` ni `Form_Load` la reinician, así que equipos elegidos en un alta cancelada podrían heredarse en la siguiente.
- `ComandoAyudaFormulario` declara `OnClick` (`.form.txt:595`) y no existe handler en el `.cls`.
- Mensajes con nombre equivocado: «El método Eliminar ha devuelto un error» en `Registrar` y `ComandoAlta_Click` (`:79, 107`), y «El método Alta» en `RellenarComboContacto` (`:826`).
- En la precarga, el equipo se muestra por nombre (`:494-500`) y `Registrar` lo distingue del ID con `IsNumeric` (`:50-65`): un equipo cuyo nombre sea numérico se tomaría como ID.
- `EstablecerDatos` asigna `ComboTecnico` (`:473`) antes de que `Form_Load` rellene sus filas (`:383, 398`), y el combo tiene `LimitToList`: el técnico precargado podría quedar vacío.
- `Registrar`, `ComandoAlta_Click` y otros usan `flag`, `dato`, `pregunta` y `m_SQL` sin declarar localmente (`:58-60, 104, 431`); con `Option Explicit` (`:3`) deben ser globales de `VG.bas`.
- `ComboBUI_AfterUpdate` calcula `ListCount - 1` y selecciona `Selected(1)` si vale 1 (`:258-263`): el resultado depende de que el combo tenga fila de encabezado, que el diseño no confirma.
- `Numero_Actividad` es un control residual enlazado a `IDActividad` en un formulario no enlazado (`.form.txt:691-693`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador de evento generado de forma atómica por la base de datos, alta del evento y de sus equipos de medida en una sola transacción, y validación del contacto y de las fechas SLA (la fecha de recepción no se compara con la de alta).

## Criterios de aceptación

Propuestos, no aprobados:

1. Con `ComboBUI` vacío, `ComandoNuevoIDEquipo` y `ComandoEquiposDeMedida` permanecen deshabilitados; con BUI y subsistema se habilitan (`:224-232, 305-313`).
2. Cambiar de BUI vacía subsistema y equipo y recarga los subsistemas de ese BUI (`:242-257`).
3. Guardar sin cualquiera de los campos obligatorios muestra su mensaje y no crea el evento (`:595-635`).
4. Un equipo con fecha de obsolescencia anterior a la fecha de alta se rechaza (`:647-652`).
5. Un alta válida crea un `IDEvento` con formato iniciales + `aamm` + 3 cifras y lo muestra al usuario (`FG.bas:6221-6242`; `:131-142`).
6. Los equipos de medida elegidos quedan asociados al evento creado (`:777-788`).
7. Con `FormEventoGestion` abierto, el evento nuevo queda seleccionado en su buscador (`:116-120`).
8. `cmdSalir` cierra sin crear nada (`:358-361`).

## Pendientes

- Llamadores con `OpenArgs` (copia de evento): ninguno en `src/forms|modules|classes` salvo `Form_FormEventoGestion.cls:1133` sin argumento; posible apertura dinámica no resuelta.
- Origen y declaración de `flag`, `dato`, `pregunta`, `m_SQL` (`:58, 60, 104, 431`).
- Contenido de `Constructor.getEvento` (`:778`) y de `Constructor.getTecnicos` (`FG.bas:7239`).
- Efecto de `Me.Requery` en un formulario sin origen (`:364`).
- Si `ComandoEquiposDeMedida` queda habilitado en diseño (el `.form.txt` no lo deshabilita) antes de elegir BUI (`.form.txt:1916-1953`).
- Formato de `FechaAlta` (cadena `dd/mm/yyyy`, `:394`) frente a la configuración regional del cliente.
