# Form_FormMaterialIntervencionAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialIntervencionAlta.cls` (258 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de una intervención (seguimiento) sobre un material (`TbMaterialSeguimiento`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el ID del material en `OpenArgs`, muestra equipo, material, garantía y evento de origen en controles bloqueados, y delega la validación y la escritura en `MaterialIntervencionAlta` (`FG.bas:4013`). Abreviatura: `FG.bas` = `modules/Funciones Generales.bas`. Las citas sin fichero son `Form_FormMaterialIntervencionAlta.cls:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMaterialSeguimiento` | Cierra el formulario si estaba abierto y lo abre con el ID del material en `OpenArgs` | `Form_FormMaterialSeguimiento.cls:120-123` |
| `Form_FormMaterialesGestion` | Igual, con el ID del material seleccionado | `Form_FormMaterialesGestion.cls:523-526` |

Salidas: tras guardar, si `FormMaterialSeguimiento` está abierto llama a su `EstablecerDatos` (`:76-78`) y cierra (`:79`); `cmdSalir` cierra sin guardar (`:5-8`); `ComandoVerEvento` abre `FormEventoGestion` posicionado en el evento de origen (`:92-129`).

## Controles y eventos

Controles: 30 (38 `Begin` en el `.form.txt`: 1 formulario, 6 plantillas de estilo por defecto, 1 sección y 30 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaInicioIntervencion` | texto, fecha corta | — | Obligatoria y fecha válida | CAP-7 Materiales | `.form.txt:331-343`; `FG.bas:4074-4077` |
| `FechaFinIntervencion` | texto, fecha corta | — | Opcional; si se rellena ha de ser fecha | CAP-7 Materiales | `.form.txt:680-693`; `FG.bas:4089-4094` |
| `MOTIVO` | texto | — | Obligatorio | CAP-7 Materiales | `.form.txt:736-748`; `FG.bas:4085-4088` |
| `DESTINO` | combo (consulta) | — | Obligatorio; ofrece los destinos distintos ya registrados en `TbMaterialSeguimiento` | CAP-7 Materiales | `.form.txt:818-838`; `FG.bas:4081-4084` |
| `RESULTADO` | texto | — | Opcional | CAP-7 Materiales | `.form.txt:855-867`; `FG.bas:4179-4181` |
| `ComboUltimaIntervencion` | combo lista «Sí;No», valor por defecto «No» | — | Marca si cierra el seguimiento | CAP-7 Materiales | `.form.txt:937-962`; `FG.bas:4078-4080` |
| `Equipo`, `Material`, `Garantia`, `IDEVENTO` | texto, bloqueados (`Locked`) | — | Datos informativos del material | CAP-7 Materiales | `.form.txt:512-615, 1001-1015`; `:227-241` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Valida, guarda y cierra | CAP-7 Materiales | `.form.txt:359-369`; `:54-90` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-7 Materiales | `.form.txt:397-407`; `:5-8` |
| `ComandoVerEvento` («Ver Evento») | botón | `OnClick` | Abre la gestión de eventos en el evento de origen | CAP-3 Eventos | `.form.txt:1074-1086`; `:92-129` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «Alta de Intervención de Material» (`lblTitulo1` pasa a `Me.Caption`) | CAP-7 Materiales | `.form.txt:1031-1063`; `:140` |
| 10 etiquetas de campo (`Etiqueta83`, `84`, `91`, `94`, `96`, `98`, `102`, `107`, `108`, `110`) | etiquetas | — | Rótulos de los campos | CAP-7 Materiales | `.form.txt:458-986` |
| `Marco77`, `Marco93` | grupos de opciones | — | Marcos decorativos | sin asignar | `.form.txt:436-445, 631-640` |
| `CuadroPrincipal`, `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decorativos | sin asignar | `.form.txt:263-317` |

Cuadre: 6 campos editables, 4 informativos bloqueados, 3 botones, 2 títulos, 10 etiquetas, 2 marcos y 3 rectángulos = 30.

## Flujos

### Apertura (`Form_Open`, `:132-170`)

`AjustarTamaño` (`:138`), copia el título (`:140`) y lee `OpenArgs` en `strIDMaterial`; vacío ⇒ «Se ha abierto con parámetros inadecuados» (`:141-145`). `EstablecerDatos` (`:171-253`) pregunta con `MaterialProvieneDeActividad` (`FG.bas:3815-3860`) si el material cuelga de una actividad con evento. Si sí, une `TbMaterial`, `TbActividades`, `TbEventos` y `TbEquipos` (`:194-198`); si no, `TbMaterial` con `TbEquipos` por `IDEquipo` (`:211-213`). Rellena `Equipo`, `Material` y `Garantia` si no están vacíos; `IDEVENTO` y la habilitación de `ComandoVerEvento` solo en la primera rama (`:227-241`). Cualquier error: `MsgBox` crítico y cierre del formulario (`:163-169`).

### Guardado (`ComandoRegistrar_Click`, `:54-90`)

`Registrar` (`:9-53`) exige `strIDMaterial` (`:25-28`), recoge los seis valores con `Nz` (`:29-34`) y llama a `MaterialIntervencionAlta` (`:40`). Error ⇒ `MsgBox` y el formulario sigue abierto (`:84-89`). Éxito: refresca el seguimiento y cierra (`:76-79`).

### Ver evento (`:92-129`)

Sin `IDEVENTO` sale en silencio (`:95-98`). Con `FormEventoGestion` cerrado lo abre y fija `ComboIDEventoBusqueda` (`:107-110`); si ya está abierto y con otro evento llama a `ComandoActualizarTodo_Click` y cambia el combo (`:111-117`).

## Reglas, validaciones y cálculos

Las aplica `MaterialIntervencionAlta`, en este orden (`FG.bas:4040-4143`):

1. ID de material obligatorio (`4040-4043`).
2. El material no puede proceder de un evento franqueado, «Proviene de un evento franqueado» (`4049-4058`), ni estar facturado, «Proviene de un material ya facturado» (`4064-4073`).
3. Fecha de inicio válida; destino y motivo obligatorios; fecha de fin válida si se informa (`4074-4094`). Una «Última intervención» distinta de «Sí»/«No» se trata como «No» (`4078-4080`).
4. Sin seguimiento cerrado, «El proceso ya estaba finalizado previamente» (`4100-4109`), y sin intervención previa abierta (`UltimaIntervencion = "Sí"`), «Hay una intervención previa sin cerrar» (`4115-4124`).
5. La fecha de inicio no puede ser anterior a la última fecha de fin o a la de alta del material (`4130-4144`).
6. **Identificador:** `DameIDIntervencion` genera `<IDMaterial>_<NN>` con el máximo existente más uno (`FG.bas:4303-4328`).
7. **Efecto lateral:** si tras guardar el seguimiento queda cerrado, actualiza `TbMaterial.FechaEntrega` con la última fecha de fin (`FG.bas:4192-4215`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| ¿Material de actividad? | `TbEventos`, `TbMaterial`, `TbActividades` | SQL dinámico con el ID entre comillas | — | `FG.bas:3841-3845` |
| Datos del material (con evento) | `TbMaterial`, `TbActividades`, `TbEventos`, `TbEquipos` | SQL dinámico concatenado | — | `:194-199` |
| Datos del material (sin evento) | `TbMaterial`, `TbEquipos` | SQL dinámico concatenado | — | `:211-214` |
| Siguiente ID de seguimiento | `TbMaterialSeguimiento` | SQL dinámico concatenado | — | `FG.bas:4315-4318` |
| Alta | `TbMaterialSeguimiento` | DAO `AddNew`/`Update` | no | `FG.bas:4161-4183` |
| Actualizar entrega | `TbMaterial` | `UPDATE` con fecha literal `#mm/dd/yyyy#`, `DoCmd.RunSQL` con avisos desactivados | no | `FG.bas:4212-4216` |
| Origen del combo `DESTINO` | `TbMaterialSeguimiento` | consulta estática (`SELECT DISTINCT`) | — | `.form.txt:838` |

El SQL del formulario usa la variable global `m_SQL` (`VG.bas:203`), no declarada en el módulo del formulario.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: fecha literal `#mm/dd/yyyy#` y `DoCmd.RunSQL` (dialecto Access, a reescribir con parámetros); ID de texto con `Max + 1` calculado en VBA sin transacción; combinaciones de `INNER JOIN`/`RIGHT JOIN` anidadas con paréntesis (`FG.bas:3842-3843`, `:195-197`) que requieren reescritura para PostgreSQL; `SELECT DISTINCT` y `WHERE ... Is Null` portables.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El ID `<IDMaterial>_<NN>` se calcula como máximo más uno fuera de transacción (`FG.bas:4150-4170, 4315-4328`): dos altas simultáneas podrían colisionar; el formato `00` limita a 99 intervenciones por material (`FG.bas:4303`).
- El `UPDATE` de `FechaEntrega` usa `DoCmd.SetWarnings False` y, si falla, `SetWarnings True` no se ejecuta (`FG.bas:4214-4216`).
- La fecha del `UPDATE` se formatea `mm/dd/yyyy` (`FG.bas:4212`); asignada a `FechaInicioIntervencion` como texto (`FG.bas:4173`), la interpretación dependería de la configuración regional.
- `ComandoVerEvento_Click` sale en silencio con `IDEVENTO` vacío (`:95-98`), aunque el botón solo se habilita si hay evento (`:233-238`).
- `EstablecerDatos` no rellena `IDEVENTO` si el material no procede de actividad y no limpia valores previos (`:210-241`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: identificador de seguimiento generado por la base de datos y validaciones en una única transacción junto con la actualización de la fecha de entrega.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `OpenArgs` muestra «Se ha abierto con parámetros inadecuados» y cierra (`:141-145, 163-169`).
2. Abrir con un material de actividad muestra equipo, material, garantía y evento, y habilita «Ver Evento» (`:227-241`).
3. Guardar sin fecha de inicio, destino o motivo muestra el mensaje correspondiente y no escribe (`FG.bas:4074-4088`).
4. Un material facturado o de un evento franqueado no admite la intervención (`FG.bas:4049-4073`).
5. Con una intervención previa sin cerrar, se rechaza el alta (`FG.bas:4115-4124`).
6. Guardar con «Última intervención = Sí» cierra el seguimiento y fija `FechaEntrega` (`FG.bas:4192-4215`).
7. Tras guardar, el listado de seguimiento abierto se refresca (`:76-78`).

## Pendientes

- Cuerpo de `EstaFranqueado`, `EstaFacturado`, `SeguimientoCerrado`, `MaterialSeguimientoConItervencionAbierta` y `UltimaFechaFinSeguimiento` (`FG.bas:4049, 4064, 4100, 4115, 4130`).
- Si `TbMaterialSeguimiento` tiene clave o índice único sobre `IDSeguimiento` (ERD de origen).
- Efecto de `Locked` y `TabStop` en la navegación con teclado (`.form.txt:513-514`).
