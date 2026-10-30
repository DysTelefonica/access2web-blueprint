# Form_FormPlanificacionGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionGestion.cls` (955 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Cuadro de mando de «Mantenimientos preventivos» (título en `Form_FormPlanificacionGestion.form.txt:875`) para un año elegido: diez contadores (equipos, programados, realizados, no realizados y sus subdivisiones) calculados por `CuadroDeMando`, que vive en este mismo módulo (`:734`). Al pulsar un contador se genera y abre un informe HTML con las filas de esa categoría, y `ComandoInforme` exporta a Excel las planificaciones abiertas con sus eventos posibles. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); solo escribe en la tabla auxiliar `TbAuxPlanificacion`. Las citas son `:línea` de `Form_FormPlanificacionGestion.cls` salvo indicación. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`, `VG.bas` = `modules/Variables Globales.bas`, `HTML.bas` = `modules/HTML.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | `lblMtoProg_Click` → `DoCmd.OpenForm "FormPlanificacionGestion"` | `Form_Form0BDOpcionesOtras.cls:34-37` |

Abre `FormPlanificacionEquipos` (`:69`) y `FormPlanificacionesBusqueda` (`:73`), ambos sin argumentos y con `On Error Resume Next`. No lanza eventos ni devuelve valores. Sale con `cmdSalir` (`:4-7`). Los resultados se dejan en las colecciones **globales** `colPreventivo*` (`VG.bas:143-153`), que sobreviven al cierre del formulario.

## Controles y eventos

Controles: 38 (45 `Begin` en el `.form.txt`: 38 controles + `Form`, la sección `Detalle` y 5 plantillas de estilo por defecto en `Form_FormPlanificacionGestion.form.txt:720-806`).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `Año` | texto no enlazado | `OnExit` declarado, **sin procedimiento** | Año a calcular; `Form_Open` lo rellena con el año actual | CAP-6 Planificación | `:117`, `Form_FormPlanificacionGestion.form.txt:1302-1315` |
| `ComandoActualizar` («Actualizar») | botón | `OnClick` | Vacía contadores, ejecuta `CuadroDeMando` y rellena los diez contadores | CAP-6 Planificación | `:8-66`, `Form_FormPlanificacionGestion.form.txt:1245-1258` |
| `ComandoInforme` («Informe») | botón | `OnClick` | Reconstruye `TbAuxPlanificacion` y la exporta a Excel | CAP-6 Planificación | `:75-108`, `Form_FormPlanificacionGestion.form.txt:2011-2023` |
| `ComandoBuscarPlanificaciones` («Actuaciones Planificadas ») | botón | `OnClick` | Abre `FormPlanificacionesBusqueda` | CAP-6 Planificación | `:71-74`, `Form_FormPlanificacionGestion.form.txt:1203-1215` |
| `ComandoBuscarEquiposPlanificables` («Equipos para Mto Prev.») | botón | `OnClick` | Abre `FormPlanificacionEquipos` | CAP-6 Planificación | `:67-70`, `Form_FormPlanificacionGestion.form.txt:1369-1381` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-6 Planificación | `:4-7`, `Form_FormPlanificacionGestion.form.txt:918-929` |
| `PreventivoEquipos` | texto bloqueado, `Tag="CONTADOR"` | `OnClick` | Contador `colPreventivoEquipos` (`:46`); informe «EQUIPOS», 4 columnas | CAP-6 Planificación | `:132-171`, `Form_FormPlanificacionGestion.form.txt:1015-1032` |
| `PreventivosEquiposProgramados` | ídem | `OnClick` | Contador de `colPreventivosEquiposProgramados` (`:47`); el informe usa **`colPreventivosProgramados`** (`:224`) | CAP-6 Planificación | `:213-252`, `Form_FormPlanificacionGestion.form.txt:1080-1097` |
| `PreventivosEquiposNoProgramados` | ídem | `OnClick` | Contador (`:48`); informe «EQUIPOS SIN PROGRAMACIÓN», 4 columnas | CAP-6 Planificación | `:172-212`, `Form_FormPlanificacionGestion.form.txt:1116-1133` |
| `PreventivosProgramados` | ídem | `OnClick` | Contador de `colPreventivosProgramados` (`:49`); el informe usa **`colPreventivosRealizados`** (`:384`), 12 columnas | CAP-6 Planificación | `:373-412`, `Form_FormPlanificacionGestion.form.txt:1911-1928` |
| `PreventivosRealizados` | ídem | `OnClick` | Contador (`:50`); informe «REALIZADOS» | CAP-6 Planificación | `:414-453`, `Form_FormPlanificacionGestion.form.txt:1520-1537` |
| `PreventivosRealizadosEnPlazo` | ídem | `OnClick` | Contador (`:54`); informe «REALIZADOS EN PLAZO» | CAP-6 Planificación | `:454-493`, `Form_FormPlanificacionGestion.form.txt:1875-1892` |
| `PreventivosRealizadosFueraDePlazo` | ídem | `OnClick` | Contador (`:52`); informe «REALIZADOS FUERA DE PLAZO» | CAP-6 Planificación | `:494-533`, `Form_FormPlanificacionGestion.form.txt:1592-1609` |
| `PreventivosNoRealizados` | ídem | `OnClick` | Contador (`:51`); informe «NO REALIZADOS» | CAP-6 Planificación | `:253-292`, `Form_FormPlanificacionGestion.form.txt:1556-1573` |
| `PreventivosNoRealizadosProximos` | ídem | `OnClick` | Contador (`:53`); informe «NO REALIZADOS PRÓXIMOS» | CAP-6 Planificación | `:293-332`, `Form_FormPlanificacionGestion.form.txt:1628-1645` |
| `PreventivosNoRealizadosVencidos` | ídem | `OnClick` | Contador (`:55`); informe «NO REALIZADOS VENCIDOS» | CAP-6 Planificación | `:333-372`, `Form_FormPlanificacionGestion.form.txt:1664-1681` |
| `lblTitulo` | etiqueta | — | «MANTENIMIENTOS PREVENTIVOS»; pasa a `Caption` del formulario | CAP-6 Planificación | `:116`, `Form_FormPlanificacionGestion.form.txt:864-875` |
| 15 etiquetas `Etiqueta208`, `210`, `214`, `221`, `190`-`192`, `206`, `1`-`5`, `8`, `181` | etiquetas | — | Rótulos fijos (TOTALES, PROGRAMADOS, AÑO, REALIZADOS, SIN PROGRAMAR, Fuera de plazo…, «Pulse sobre los resultados para ver el Informe») | CAP-6 Planificación | `Form_FormPlanificacionGestion.form.txt:981-2000` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2` (rectángulos) y `Marco207`, `Marco215`, `Marco189` (grupos de opciones) | decorativos | — | Marcos visuales; sin evento ni origen de datos | CAP-6 Planificación | `Form_FormPlanificacionGestion.form.txt:818, 842, 894, 960, 1182, 1411` |

## Flujos

### Apertura (`Form_Open`, `:109-131`)

1. Cursor de espera, `AjustarTamaño Me` (`FG.bas:261`), título desde `lblTitulo` y `Año` = año actual (`:113-117`). **No calcula nada**: los contadores quedan vacíos hasta pulsar «Actualizar».
2. Cualquier error: `MsgBox` crítico y cierre del formulario (`:124-130`).

### Actualizar (`ComandoActualizar_Click`, `:8-66`)

1. Pone a `Null` los controles con `Tag = "CONTADOR"` (`:18-20`) y toma `Año` (`:21`).
2. `CuadroDeMando` devuelve un número de registros o `#ERR|texto` (`:27-32, 953`). Con `|` muestra «El método Eliminar ha devuelto un error…» (`:34-38`); con `"0"`, «No hay registros para ese año» (`:39-42`).
3. Asigna `.count` de cada colección a su contador (`:46-55`). Error: `MsgBox` crítico (`:60-65`).

### Cálculo (`CuadroDeMando`, `:734-954`)

1. Año no numérico o `0` ⇒ año actual (`:766-768`). Rango «01/01/año»–«31/12/año» y rango de consulta con fecha final en formato mes/día (`:769-772`).
2. Recorre `TbPlanificacionEquipos` ⨝ `TbEquipos` (`:787-815`). Un equipo entra en `colPreventivoEquipos` si no tiene `FechabajaParaPlanificacion` o esta es posterior al 31/12 del año (`:801-809`). Sin equipos ⇒ `"0"` (`:818-821`).
3. Por cada equipo consulta si tiene alguna `FechaPrevistaCierre` del año ⇒ «equipos programados» o «no programados» (`:826-843`).
4. Por cada equipo programado lee todas sus planificaciones (`:848-934`). Si la fecha prevista cae en el año: programado (`:874-875`); con `FechaCierre` válida: realizado (`:879-884`), fuera de plazo si cierre > prevista (`:885-890`) o en plazo (`:896`); sin cierre: no realizado (`:904`) con `DateDiff` a hoy: 0-30 días ⇒ próximo (`:905-911`), > 30 ⇒ no se añade a ninguna subcolección (`:912-917`), < 0 ⇒ vencido (`:923`).

### Informe de un contador (`*_Click`, `:132-533`)

`DameHTML(colección, "Sí"/"No", título)` compone la tabla HTML (`HTML.bas:116-244`): con «Sí» 4 columnas (BUI, subsistema, equipo, periodicidad), con «No» 12 (`HTML.bas:169-182, 203`). `HTMLENTXT` borra los `.html`/`.htm` no abiertos del directorio local (`HTML.bas:11-17`), elige `HTMLn.txt` libre entre 1 y 50 (`HTML.bas:45-60`), escribe el texto, lo renombra a `.html` y lo abre con `Ejecutar` (`HTML.bas:99-104`).

### Informe Excel (`ComandoInforme_Click`, `:75-108`)

1. `DameListadoPlanificacionesPendientesConIDEventosPosibles` (`:534-647`): equipos con alguna planificación sin `FechaCierre` (`:551-554`); sin ellos devuelve «Sin Planificaciones» y sale (`:568-571`). Si no, vacía `TbAuxPlanificacion` con `RunSQL` (`:572-575`) y añade una fila por planificación abierta con las cadenas de eventos posibles y usados (`:578-629`).
2. `DameCadenaIDEventos` (`:648-733`): «posibles» = eventos `MANT. PROGRAMADO` del equipo, terminados, no ligados a una planificación y con fecha de última actividad (`:667-693`); «usados» = los ligados (`:700-715`). Formato `id(fecha)|id(fecha)#id(planif)|…` (`:722`).
3. `GenerarConsultasSQL "Select * from TbAuxPlanificacion;"` crea un `.xlsx` en el directorio local y lo abre (`:92`; `FG.bas:1584, 1606-1608, 1670`). Error de este paso: `p_Error` ⇒ `MsgBox` (`:94-96, 107`).

## Reglas, validaciones y cálculos

- **Año**: sin validación en el control; `CuadroDeMando` lo sustituye por el actual si no es numérico (`:766-768`).
- **Equipo planificable**: sin fecha de baja o con baja posterior al fin de año (`:801-809`).
- **Plazo**: realizado «en plazo» si `FechaCierre ≤ FechaPrevistaCierre` (`:885-896`). «Próximo»: de 0 a 30 días desde `Now()` (`:905-906`).
- Los contadores **no se recalculan** al pulsarlos: el informe usa la colección global de la última ejecución (`VG.bas:143-153`).
- `colPreventivosNoRealizadosEnPlazo` se declara y se crea pero nunca se rellena ni se muestra (`:773-783, 912-917`).
- No hay control de rol ni de usuario en este módulo; cualquier usuario que abra el formulario puede reconstruir `TbAuxPlanificacion`.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Equipos planificables | `TbPlanificacionEquipos`, `TbEquipos` | SQL dinámico, `OpenRecordset` | — | `:787-790` |
| Planificaciones del año por equipo | `TbPlanificacion` | SQL dinámico con `IDEquipo=` y fechas entre `#…#` | — | `:829-833` |
| Todas las planificaciones del equipo | `TbPlanificacion` | SQL dinámico concatenado | — | `:855-858` |
| Equipos con planificación abierta | `TbPlanificacion`, `TbPlanificacionEquipos`, `TbEquipos` | SQL dinámico | — | `:551-555` |
| Planificaciones abiertas por equipo | ídem | SQL dinámico con `DateDiff('d',Now(),…)` y `IDEquipo=` | — | `:580-585` |
| Vaciado de la auxiliar | `TbAuxPlanificacion` | `DoCmd.RunSQL` con `delete *` y avisos desactivados | no | `:572-575` |
| Alta de filas auxiliares | `TbAuxPlanificacion` | DAO `AddNew`/`Update` | no | `:577, 613-622` |
| Eventos posibles / usados | `TbEventos`, `TbPlanificacion` | SQL dinámico, `RIGHT JOIN`, literal `'MANT. PROGRAMADO'` | — | `:667-672, 700-704` |
| Fecha de última actividad | `TbActividades` | SQL dinámico concatenado | — | `FG.bas:798-831` |
| Exportación a Excel | `TbAuxPlanificacion` | `OpenRecordset` + automatización de Excel | — | `FG.bas:1619-1670` |
| Ficheros HTML / `.xlsx` | directorio local de `m_ObjEntorno` | `FileSystemObject`, `ShellExecute` | — | `HTML.bas:11-17, 99-104`; `FG.bas:1606-1608` |

## Coexistencia

Adaptar a PostgreSQL vinculado: `delete * from` y `RunSQL` (`:572`); literales de fecha `#…#` en formato mes/día (`:772, 832`) y `DateDiff('d',Now(),…)` (`:581, 905`); `RIGHT JOIN` con `Is Null` (`:668`); `Nz` sobre campos de recordset; recuentos de colecciones en memoria que pasarían a ser agregaciones SQL; `IDEquipo` concatenado sin comillas, válido solo si es numérico (`:584, 669`); la tabla auxiliar compartida entre usuarios (`:572-577`) y la exportación a Excel por automatización COM (`FG.bas:1625`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `strIDEquipo` se declara (`:758`) y se concatena en el `WHERE` (`:831, 857`) pero no se le asigna nada en `CuadroDeMando`; además `dato(0)` es el nombre del equipo, no su ID (`:800, 828`). La consulta quedaría como `IDEquipo=` y fallaría, o filtraría un valor vacío; entonces «Actualizar» daría error siempre que haya equipos.
- `Año` declara `OnExit` (`Form_FormPlanificacionGestion.form.txt:1315`) sin procedimiento en el `.cls`: manejador huérfano.
- `PreventivosProgramados_Click` muestra `colPreventivosRealizados` (`:384`) y `PreventivosEquiposProgramados_Click` muestra `colPreventivosProgramados` (`:224`), mientras que sus contadores cuentan otras colecciones (`:47, 49`): el número y el informe podrían no coincidir.
- `lngNumeroRegistros` suma dos veces `colPreventivosProgramados` y no suma `colPreventivosEquiposProgramados` (`:935-942`): con equipos sin planificaciones podría devolver `"0"` y mostrar «No hay registros para ese año» pese a haber equipos.
- Fechas: `CDate("31/12/"&año)` y `CDate(strFechaPrevistaCierre)` dependen de la configuración regional (`:770, 807, 874`); una `FechaPrevistaCierre` nula daría error 13.
- `ComandoInforme_Click` no distingue «Sin Planificaciones» de «OK» (`:86-92, 569`): exportaría el contenido anterior de `TbAuxPlanificacion`.
- `DameListadoPlanificacionesPendientesConIDEventosPosibles` une `TbPlanificacion` y `TbPlanificacionEquipos` solo por `IDEquipo` (`:552, 582`): podría duplicar filas si hay varias filas por equipo.
- Mensajes de error con «El método Eliminar» en rutas que no eliminan (`:36, 606`).
- `HTMLENTXT` borra todos los `.html`/`.htm` no abiertos del directorio local (`HTML.bas:11-17`), y `DameHTML` inserta los valores sin escapar (`HTML.bas:203-235`).
- Las colecciones globales persisten: tras reabrir el formulario los contadores están vacíos pero el clic abre los datos del cálculo anterior (`VG.bas:143-153`; `:18-20`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: cálculo de los contadores en una consulta agregada por año; la exportación sin tabla auxiliar compartida; el informe HTML como vista web.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, `Año` muestra el año actual y los diez contadores están vacíos (`:117`; `:18-20`).
2. «Actualizar» con un año sin equipos planificables muestra «No hay registros para ese año» (`:39-42, 818-821`).
3. Un equipo con baja de planificación anterior o dentro del año no se cuenta en `PreventivoEquipos` (`:801-809`).
4. Una planificación con cierre posterior a la fecha prevista se cuenta en «Fuera de plazo» y una cerrada a tiempo en «En plazo» (`:885-896`).
5. Una planificación sin cierre con 0-30 días hasta su fecha prevista se cuenta en «Próx venc.»; con fecha pasada, en «Vencidos» (`:905-923`).
6. Pulsar cada contador abre un informe cuyas filas coinciden con su contador (`:132-533`).
7. «Informe» genera un Excel con una fila por planificación abierta, con eventos posibles y usados (`:578-629`).
8. «Actuaciones Planificadas» y «Equipos para Mto Prev.» abren sus formularios (`:67-74`).

## Pendientes

- Estructura y relaciones de `TbPlanificacion`, `TbPlanificacionEquipos` y `TbAuxPlanificacion` (ERD de origen).
- Contenido de `m_ObjEntorno.CSS` y de `m_ObjEntorno.URLDirectorioLocal` (`HTML.bas:158, 11`).
- Si el comentario de `FG.bas:2593` (llamada de `NPlanificaciones` desde este formulario) sigue vigente: este módulo no la invoca.
