# Form_FormActividadAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormActividadAlta.cls` (251 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de una actividad (`TbActividades`): técnico, ubicación, fecha, horas laborables y extras, tipo de actividad, evento generado y descripción. Es un formulario **no enlazado**: los controles se rellenan por código y el alta la hace `ActividadAlta` (`FG.bas:5358-5580`). Se abre en tres modos según `OpenArgs`: sin argumentos (actividad sin evento), con evento, o con evento y actividad de origen (se usa como plantilla de copia). Las citas son `Form_FormActividadAlta.cls:línea` (abreviadas `:N`) salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormActividadGestion` | Alta sin argumentos; cierra antes el formulario si estaba abierto | `Form_FormActividadGestion.cls:238-254` |
| `Form_FormEventoGestion` (alta con evento) | `OpenArgs = IDEvento & "\|" & "\|" & EstaFacturado` | `Form_FormEventoGestion.cls:1048, 1085-1090` |
| `Form_FormEventoGestion` (alta sin evento) | Sin argumentos | `Form_FormEventoGestion.cls:1102, 1114-1119` |
| `Form_FormEventoGestion` (copiar actividad) | `OpenArgs = IDEvento & "\|" & IDActividad & "\|" & EstaFacturado` | `Form_FormEventoGestion.cls:1450, 1494-1499` |

Formato de `OpenArgs`: tres partes separadas por `|`; `IDEvento` obligatorio y `EstaFacturado` ∈ {«Sí», «No»} (`:201-221`). Salida: tras guardar refresca `FormEventoGestion` (`Form_FormEventoGestion.ListaEventos_Click`, `:73-75`) y la lista `ListaFiltrados` de `FormActividadGestion` (`:80-82`), y se cierra (`:86`).

## Controles y eventos

Controles: 29 (37 `Begin` en el `.form.txt`: 1 `Form`, 6 plantillas por defecto —`Label`, `Rectangle`, `Image`, `CommandButton`, `TextBox`, `ComboBox`—, 1 `Section` y 29 controles con nombre).

| Control | Tipo | Evento y handler | Origen de datos, validación | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboTecnico` | combo lista de valores, 3 columnas | — | `RellenarComboTecnicos` con solo técnicos activos; `LimitToList`; `Tag = OBLIGATORIO` | CAP-9 | `:230`; `FG.bas:7224-7252`; `Form_FormActividadAlta.form.txt:1049-1073` |
| `ComboUbicacion` | combo lista de valores | `Enter` → requery | `SELECT` de `TbUbicacion`; `Tag = OBLIGATORIO` | CAP-12 | `:103-106, :121-123`; `Form_FormActividadAlta.form.txt:543-567` |
| `ComboActividad` | combo lista de valores | `Enter` → requery | `SELECT` de `TbCodActividad` (`CODIGO`, `ACTIVIDADNOPROGRAMADA`); sin `Tag` | CAP-4 | `:95-98, :118-120`; `Form_FormActividadAlta.form.txt:581-604` |
| `FechaAlta` | texto | — | Formato «Short Date»; `Tag = OBLIGATORIO`; hoy por defecto si no hay evento | CAP-4 | `:176`; `Form_FormActividadAlta.form.txt:482-496` |
| `HorasLaborables` | texto | — | «General Number»; `Tag = OBLIGATORIO` | CAP-4 | `Form_FormActividadAlta.form.txt:452-466` |
| `HorasExtras` | texto | — | «General Number» | CAP-4 | `Form_FormActividadAlta.form.txt:512-527` |
| `IDEventoGenerado` | texto (no combo) | `Enter` → `Requery` | Evento resultante; `Tag = OBLIGATORIO` | CAP-3 | `:99-102`; `Form_FormActividadAlta.form.txt:807-822` |
| `Descripcion` | texto | — | `Tag = OBLIGATORIO` | CAP-4 | `Form_FormActividadAlta.form.txt:263-276` |
| `IDEVENTO` | texto bloqueado | — | Recibe `strIDEvento` | CAP-3 | `:178`; `Form_FormActividadAlta.form.txt:991-1006` |
| `Numero_Actividad` | texto deshabilitado y bloqueado | — | `ControlSource = IDActividad`; valor por defecto con `DCount` sobre `Tabla_Parte_Actividad`; el código no lo usa | sin asignar | `Form_FormActividadAlta.form.txt:292-307` |
| `ComandoAlta` («Registrar») | botón | `Click` → `ComandoAlta_Click` | Valida horas y da de alta | CAP-4 | `:8-94`; `Form_FormActividadAlta.form.txt:375-385` |
| `cmdSalir` («Salir») | botón | `Click` | Cierra sin guardar | CAP-4 | `:107-110`; `Form_FormActividadAlta.form.txt:413-423` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título según haya evento | CAP-4 | `:173-175, :235`; `Form_FormActividadAlta.form.txt:849, 870` |
| 13 etiquetas rótulo (`Etiqueta241`…, `lbl1`) y 2 rectángulos | decoración | — | Rótulos («TÉCNICO», «UBICAC.», «F.ALTA», «H.LAB (x,y) decimal», «H.EXT (x,y) decimal», «EV. GRADO.», «ACTIVIDAD», «*») | sin asignar | `Form_FormActividadAlta.form.txt:618-780, 881-978, 1022` |
| Formulario | — | `Form_Load` | Valida argumentos, rellena combos y carga datos | CAP-4 | `:191-250` |

## Flujos

### Apertura (`Form_Load`, `:191-250`)

1. `AjustarTamaño` y lectura de `OpenArgs` (`:197-200`). Si hay argumento: ha de tener 3 partes, `IDEvento` no vacío y facturado «Sí»/«No»; si no, «No se ha abierto el formulario con datos suficientes» (`:201-221`).
2. `EstablecerDatos` (`:111-190`): rellena `ComboActividad` y `ComboUbicacion` (`:118-123`). Si hay actividad de origen, la lee de `TbActividades` y copia técnico, ubicación, fecha, horas, tipo, evento generado y descripción (`:124-170`); si no existe: «No se han cargado todos los datos» (`:130-133`). Título «ALTA DE ACTIVIDAD DE: <evento>» o «NO DEPENDIENTE DE EVENTO», con fecha de hoy solo en este segundo caso (`:172-177`).
3. `RellenarComboTecnicos` carga técnicos (`:230-233`) y se fijan títulos (`:235-236`).
4. Cualquier error: `MsgBox` y cierre del formulario (`:243-249`).

### Alta (`ComandoAlta_Click`, `:8-94`)

1. Lee los controles con `Nz`; horas no numéricas valen 0 (`:17-31`).
2. Si hay técnico y fecha, `DameHorasTecnico` suma las horas laborables ya registradas ese día; si el total con las laborables nuevas supera 7,5 se detiene con el mensaje de horas ya trabajadas y deja el foco en `HorasExtras` (`:33-53`; `FG.bas:509-575`).
3. `ActividadAlta(...)` con los datos y `strIDEvento` y `strEventoFacturado` de los argumentos (`:60-62`); si devuelve «#ERR\|texto» se muestra (`:63-67`).
4. Refresca los formularios abiertos y cierra (`:73-86`).
5. Error: `MsgBox` crítico; el formulario queda abierto (`:88-94`).

## Reglas, validaciones y cálculos

- **Alta con evento:** el evento ha de existir y la fecha no puede ser anterior a su alta (`FG.bas:5420-5434`); el evento no puede estar franqueado ni facturado (`FG.bas:5440-5468`).
- **Evento distinto del generado:** `IDEvento` ≠ `IDEventoGenerado` (`FG.bas:5404-5409`).
- **Fecha válida:** `IsDate` (`FG.bas:5410-5413`).
- **Tipo de actividad:** obligatorio (`FG.bas:5500-5503`). No se exigen técnico, ubicación ni descripción en el código.
- **Tope de 7,5 h laborables/día:** en el formulario, solo si hay técnico y fecha (`:33-53`); en `ActividadAlta` solo con `strAvisoMaximasHoras = "Sí"` (`FG.bas:5476-5499`), que el formulario no pasa (`:60-62`; valor por defecto «No», `FG.bas:5401-5403`).
- **Identificador:** `IDActividad = <alias>_<yyMMdd><hhmm><NN>`, con `NN` el siguiente al máximo de dos dígitos entre los registros con el mismo prefijo (`FG.bas:5504-5524`, `440-492`); se comprueba que no exista (`FG.bas:5529-5537`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Combo actividad | `TbCodActividad` | SQL fija vía `RellenaCombo` | — | `:118-120`; `FG.bas:1149-1180` |
| Combo ubicación | `TbUbicacion` | SQL fija vía `RellenaCombo` | — | `:121-123` |
| Combo técnicos | `TbTecnicos` | SQL fija (`FECHABAJA Is Null`) | — | `Constructor.bas:4263-4267` |
| Datos de actividad origen | `TbActividades` | SQL dinámico con el ID entre comillas | — | `:125-128` |
| Horas del día | `TbActividades` | SQL dinámico: fecha `#mm/dd/yyyy#` y alias concatenados | — | `FG.bas:549-561` |
| Fecha del evento | `TbEventos` | vía `Dame` | — | `FG.bas:5420` |
| Siguiente ID | `TbActividades` | recorrido completo de la tabla | no | `FG.bas:469-492` |
| Alta | `TbActividades` | DAO `AddNew`/`Update` | no | `FG.bas:5538-5565` |

## Coexistencia

- Fechas como literales `#mm/dd/yyyy#` (`FG.bas:551`) y fecha escrita desde cadena (`FG.bas:5546-5548`): revisar en PostgreSQL.
- Identificador textual calculado en VBA con recorrido completo y sin transacción (`FG.bas:469-492`); sin autonumérico.
- Valor por defecto con `DCount` (`Form_FormActividadAlta.form.txt:307`); `Nz`, `Format`, `Now()`.
- Pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Etiquetas `OBLIGATORIO` en `ComboTecnico`, `ComboUbicacion`, `FechaAlta`, `HorasLaborables`, `IDEventoGenerado` y `Descripcion` sin validación en código; `ActividadAlta` solo exige fecha y tipo (`FG.bas:5410, 5500`), y su comentario dice que la descripción es obligatoria (`FG.bas:5384`) sin comprobarlo (`FG.bas:5558`). Se podría dar de alta sin técnico.
- `EstablecerDatos` asigna `ComboTecnico` (`:148`) antes de que `RellenarComboTecnicos` cargue la lista (`:230`): el técnico copiado podría no aparecer.
- El tope de 7,5 h se calcula con las laborables (`:45`) pero el mensaje habla de la suma con extras (`:32, :47-48`); además ignora las horas extras ya registradas (`FG.bas:540-547`).
- El tope no se aplica en `ActividadAlta` porque el formulario no pasa `strAvisoMaximasHoras` (`:60-62`).
- `IDEventoGenerado` es un cuadro de texto con `Requery` en `Enter` (`:99-102`, bajo `On Error Resume Next`): sin efecto probable.
- ID con recorrido completo y contador de dos dígitos: con más de 99 altas en el mismo minuto y alias falla (`FG.bas:5518-5523, 489-492`), y dos usuarios simultáneos podrían generar el mismo ID.
- Con evento y sin actividad de origen, `FechaAlta` queda vacía (`:172-177`) y el alta falla con «fecha válida» (`FG.bas:5410-5413`).
- El error 1000 con `strTextoError` vacío muestra un `MsgBox` en blanco (`:89-93`).
- SQL concatenado con IDs y alias (`:125-127`; `FG.bas:549-559`): un apóstrofo rompería la consulta.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos, validación del tope de horas en servidor, campos obligatorios definidos en un único lugar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con argumentos incompletos muestra «No se ha abierto el formulario con datos suficientes» y se cierra (`:201-221, :243-249`).
2. En copia, los campos se precargan con los de la actividad de origen (`:124-170`).
3. Sin evento, la fecha inicial es hoy (`:176`).
4. Si técnico y fecha existen y las laborables del día pasarían de 7,5, no se guarda (`:46-51`).
5. Con evento, fecha anterior a su alta, evento franqueado o facturado impiden el alta con el mensaje correspondiente (`FG.bas:5431-5468`).
6. Sin tipo de actividad no se guarda (`FG.bas:5500-5503`).
7. Guardar crea un registro con el identificador calculado, cierra y refresca las listas abiertas (`FG.bas:5538-5565`; `:73-86`).
8. Un error deja el formulario abierto con el mensaje (`:88-94`).

## Pendientes

- Significado de `Tag = OBLIGATORIO` (consumidor no localizado en este formulario).
- Estructura de `TbActividades` y `TbCodActividad` (ERD de origen).
- `Dame` (`FG.bas:5420`), `EstaFranqueado`, `EstaFacturado` (`FG.bas:5440, 5457`) y `Tabla_Parte_Actividad` en `Form_FormActividadAlta.form.txt:307`: referencias dinámicas no resueltas.
