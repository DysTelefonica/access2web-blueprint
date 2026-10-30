# Form_FormCalibracionEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormCalibracionEdicion.cls` (346 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición de una calibración existente: fecha de calibrado, fecha de fin y fecha de comunicación al cliente (esta última opcional). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el `IDCalibracion` en `OpenArgs`, carga la fila de `TbEquiposCalibrablesFechas` unida al equipo y la actualiza con DAO. Las citas son `Form_FormCalibracionEdicion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalibracionGestion` | `ComandoEditarCalibracion_Click`: exige una fila seleccionada con id numérico, cierra el formulario si estaba abierto y lo abre con el id en `OpenArgs` | `Form_FormCalibracionGestion.cls:148-175` |

`OpenArgs` es el `IDCalibracion` (`:194`). No devuelve datos; tras guardar refresca `ListaFiltrados` de `FormCalibracionGestion` si está abierto (`:129-131, 170-172`). Variables de módulo: `strIDCalibracion` y los tres valores iniciales de fecha (`:4-5`).

## Controles y eventos

Controles: 29 (37 `Begin` en el `.form.txt`: 1 `Form`, 6 plantillas de estilo por defecto en cabecera, 1 `Section` y 29 controles de la sección Detalle).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaCalibracion` | texto, fecha corta | — | Inicio de calibración; obligatoria | CAP-5 | `:36-38, 76`, `.form.txt:411-412` |
| `FechaFinCalibracion` | texto, fecha corta | — | Fin de calibración; obligatoria | CAP-5 | `:39-41, 77`, `.form.txt:749-750` |
| `FechaComunicacionCliente` | texto, fecha corta | — | Fecha de información al cliente; opcional, vaciarla pone `Null` | CAP-5 | `:42-44, 78`, `.form.txt:956-957` |
| `Equipo`, `Modelo`, `NS`, `BUI` | textos con `Enabled` y `Locked` distintos del valor por defecto | — | Datos del equipo de solo lectura | CAP-5 | `:45-56`, `.form.txt:587-588, 601, 618-619, 632, 649-650, 663, 707-708, 721` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | CAP-5 | `:99-144`, `.form.txt:332-334` |
| `cmdSalir` («Salir») | botón | `OnClick` | Si hay cambios pregunta si guardar; cierra | CAP-5 | `:145-186`, `.form.txt:358-360` |
| `ComandoAyudaFormulario` («Comando46») | botón | sin evento en el `.form.txt` | Sin handler | sin asignar | `.form.txt:869-870` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «EDICIÓN DE CALIBRACIÓN»; `lblTitulo1` pasa a `Caption` | CAP-5 | `:195`, `.form.txt:829-830, 849-850` |
| `Etiqueta538`, `Etiqueta190`, `Etiqueta191`, `Etiqueta192`, `Etiqueta193`, `Etiqueta197`, `Etiqueta243`, `Etiqueta245`, `Etiqueta99`, `Etiqueta0`, `Etiqueta1`, `Etiqueta3` | etiquetas | — | Rótulos («INICIAL CALIBRACIÓN», «FIN CALIBRACIÓN», «INF. CLIENTE»…) y asteriscos | CAP-5 | `.form.txt:439, 490, 516, 543, 570, 690, 777, 805, 908, 930, 984, 1008` |
| `Cuadro69`, `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:264, 287, 310` |
| `Marco175`, `Marco189` | grupos de opciones | — | Marcos decorativos sin eventos | sin asignar | `.form.txt:387, 465` |

## Flujos

### Apertura (`Form_Load`, `:187-224`)

1. `AjustarTamaño Me` (`FG.bas:261-288`), lee `OpenArgs` y copia el título (`:193-195`); no numérico ⇒ «No se han cargado todos los datos» (`:196-199`).
2. `EstablecerDatos` (`:6-68`) lee la calibración unida a su equipo (SQL concatenado con el id, `:15-20`), guarda las tres fechas iniciales y las vuelca a los controles con formato `dd/mm/yyyy` (`:26-44`), y copia los datos del equipo no vacíos (`:45-56`).
3. Cualquier error: `MsgBox` crítico y cierre (`:217-223`).

### Registro (`ComandoRegistrar_Click`, `:99-144`)

1. `HaHabidoCambios` falso ⇒ «No hay cambios que registrar» (`:110-113`).
2. `Registrar` (`:69-98`) lee los tres controles (vacíos como `#`) y llama a `EdicionCalibracion` (`:76-84`; `:264-341`).
3. Éxito: refresca `ListaFiltrados` si `FormCalibracionGestion` está abierto y cierra (`:129-132`). Error: `MsgBox` crítico y queda abierto (`:137-143`).

### Salida (`cmdSalir_Click`, `:145-186`)

Si hay cambios, pregunta «¿Desea guardar los cambios producidos?» (Sí por defecto); con Sí ejecuta `Registrar` y refresca la lista; con No descarta. En ambos casos cierra (`:149-178`). Si `Registrar` falla, el error se muestra y el formulario no se cierra (`:161-165, 180-185`).

## Reglas, validaciones y cálculos

- Calibrado y fin obligatorios y con fecha válida (`:289-296`); fin ≥ inicio (`:303-306`).
- Fecha de comunicación: si viene informada y no es `#`, debe ser fecha válida (`:297-302`); con `#` (control vacío) se guarda `Null` (`:322-324`); si no es fecha ni `#`, no se toca (`:319-325`).
- Detección de cambios: compara las tres fechas, normalizadas a `dd/mm/yyyy`, con las iniciales (`:243-259`); un error devuelve «hay cambios» (`:261-262`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga | `TbEquiposCalibrablesFechas` ⨝ `TbEquiposCalibrables` | SQL dinámico concatenado con el id, `INNER JOIN` | — | `:15-20` |
| Lectura para edición | `TbEquiposCalibrablesFechas` | SQL dinámico concatenado con el id, DAO | no | `:307-310` |
| Actualización | `TbEquiposCalibrablesFechas` | DAO `Edit`/`Update` (`FechaCalibrado`, `FechaFinCalibrado`, `FechaComunicacionCliente`) | no | `:316-326` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: fechas como texto en `dd/mm/yyyy` ida y vuelta (`:37-43, 247-253`) y asignación de cadena a campos de fecha (`:317-320`); `Null` explícito para la comunicación al cliente; SQL con el id concatenado.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `HaHabidoCambios` normaliza las fechas actuales a `dd/mm/yyyy` (`:246-254`) pero las iniciales se guardan sin normalizar (`:26-28`); si `Nz` devuelve un texto con otro formato (p. ej. con hora), siempre habría «cambios»: a verificar.
- Una fecha de comunicación vacía se envía como `#` y borra el valor (`:78, 322-324`); no hay forma de distinguir «sin tocar» de «vaciar» salvo por el valor del control.
- El comentario de `EdicionCalibracion` cita `Form_FormCalibracionEdicion.editar`, que no existe (`:277`).
- No se comprueba solape con otras calibraciones del mismo equipo (`:264-341`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: validar periodos sin solape y unificar la edición con el alta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con una calibración existente muestra las tres fechas y los datos del equipo no editables (`:26-56`; `.form.txt:587-588`).
2. Abrir sin id numérico muestra «No se han cargado todos los datos» y cierra (`:196-199, 217-223`).
3. «Registrar» sin cambios muestra «No hay cambios que registrar» (`:110-113`).
4. Fin anterior al inicio impide guardar con el mensaje de orden (`:303-306`).
5. Vaciar la fecha de comunicación y registrar la deja en blanco en el registro (`:322-324`).
6. Salir con cambios ofrece guardarlos; «No» cierra sin escribir (`:149-178`).

## Pendientes

- Tipo real de las columnas de fecha en `TbEquiposCalibrablesFechas` (ERD de origen).
- Si `FechaComunicacionCliente` interviene en informes o alertas: no resuelto con el código leído.
