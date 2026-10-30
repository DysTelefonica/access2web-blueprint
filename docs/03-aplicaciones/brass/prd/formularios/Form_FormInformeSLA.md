# Form_FormInformeSLA

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormInformeSLA.cls` (753 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestor de informes de cumplimiento SLA de eventos franqueados (`Form_FormInformeSLA.cls:5-8`). Es un formulario **no enlazado** (sin `RecordSource`): el usuario indica un rango de fechas, el formulario carga en una lista de valores los eventos franqueados con alta en ese rango y permite abrir su detalle o generar un informe HTML. Las citas sin fichero son `Form_FormInformeSLA.cls:línea`. Capacidad de todos los controles: CAP-10 Informes (el origen de datos son eventos y SLA de CAP-3).

Controles: 12 (24 `Begin` en el `.form.txt`: 12 controles, 3 secciones —`EncabezadoDelFormulario`, `Detalle`, `PieDelFormulario`—, el propio `Form` y 8 plantillas de estilo por defecto de tipo Label, Rectangle, Image, CommandButton, OptionGroup, TextBox, ListBox y ComboBox, `.form.txt:5, 421-545`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDInformes` | Clic en `lblInformeSLA`: abre `FormInformeSLA` sin argumentos y cierra el formulario de informes | `Form_Form0BDInformes.cls:98-101` |

Sin `OpenArgs` ni eventos propios. Abre `FormEventoEdicion` (`:246`) y `FormEventoSLA` (`:298`); `cmdSalir` cierra sin guardar (`:335-338`). La salida persistente es un fichero HTML en disco (ver Lecturas y escrituras).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Rótulo «Informe de cumplimiento SLA» (cabecera) | CAP-10 | `.form.txt:577-586` |
| `lblFechaInicio`, `lblFechaFin` | etiquetas | — | Rótulos «Fecha Inicio» y «Fecha Fin» | CAP-10 | `.form.txt:723-735, 778-790` |
| `txtFechaInicioInforme`, `txtFechaFInInforme` | texto no enlazado, formato `Short Date` | — | Fechas del filtro; se vacían al cargar y al limpiar | CAP-10 | `.form.txt:750-763, 805-818`; `:687-688, 322-323` |
| `cmdObtgenerEventos` (rótulo «Obtener Eventos») | botón | `OnClick` → `cmdObtgenerEventos_Click` | Valida fechas, consulta y rellena la lista | CAP-10 | `:72-191`; `.form.txt:698-709` |
| `ListaFiltrados` | cuadro de lista, 6 columnas, lista de valores | `OnClick` → `ListaFiltrados_Click` | Resultados; la fila 0 es una cabecera simulada | CAP-10 | `:17-18, 59-65, 699-718`; `.form.txt:662-681` |
| `cmdVerEvento` («Ver Detalle Evento») | botón, inicia deshabilitado | `OnClick` | Abre `FormEventoEdicion` con el evento seleccionado | CAP-3 | `:197-261`; `.form.txt:609-620` |
| `cmdDetalleSLA` («Ver Detalle SLA») | botón, inicia deshabilitado | `OnClick` | Abre `FormEventoSLA` en solo lectura | CAP-3 | `:267-313`; `.form.txt:635-647` |
| `cmdLimpiarEventos` («Limpiar») | botón | `OnClick` | Vacía fechas, lista y variables de módulo | CAP-10 | `:319-329`; `.form.txt:833-844` |
| `cmdExportarHTML` (rótulo «Genear Informe») | botón, inicia deshabilitado (pie) | `OnClick` | Genera el informe HTML del periodo | CAP-10 | `:621-678`; `.form.txt:871-882` |
| `cmdSalir` («Salir») | botón (pie) | `OnClick` | Cierra el formulario sin guardar | CAP-10 | `:335-338`; `.form.txt:897-906` |

Evento de formulario: `OnLoad` → `Form_Load` (`.form.txt:222`; `:684-693`). Variables de módulo: `m_strIDEventoSeleccionado`, `m_strFechaInicioFiltro`, `m_strFechaFinFiltro` (`:13-15`).

## Flujos

### Carga (`Form_Load`, `:684-693`)

Con `On Error Resume Next` vacía ambas fechas, fija la lista con solo la cabecera (`:59-65`) y reinicia las variables de módulo.

### Obtener eventos (`cmdObtgenerEventos_Click`, `:72-191`)

1. Cursor de espera y validaciones en orden: fecha inicio vacía (`:92-95`), fecha fin vacía (`:97-100`), inicio no válida (`:102-105`), fin no válida (`:107-110`) e inicio posterior a fin (`:112-115`). Cada fallo muestra su mensaje con `MsgBox` crítico (`:182-190`).
2. Guarda las fechas en `m_strFechaInicioFiltro`/`m_strFechaFinFiltro` para la exportación (`:118-119`).
3. Pide la SQL a `SLAReportService.bas:93-162` y abre un `Recordset` DAO con `CurrentDb()` (`:123-128`). Un error de construcción se muestra como «Error al construir consulta» (`:124-127`).
4. Compone la lista de valores: cabecera más seis columnas por evento (`IDEVENTO`, `FECHAALTAEVENTO` en `dd/mm/yyyy`, `BUI`, `SUBSISTEMA`, `CRITICIDAD`, `FechaRecepcionNotificacion`); esta última queda vacía si es nula o `30/12/1899` (`:131-159`).
5. Habilita `cmdExportarHTML` si hay filas y lo deshabilita si no (`:134-168`); asigna el `RowSource` (`:175-176`).

### Selección y detalle

- `ListaFiltrados_Click`: fila 0 ⇒ deselecciona; fila >0 ⇒ guarda el ID y habilita `cmdVerEvento` y `cmdDetalleSLA`; sin selección ⇒ los deshabilita (`:699-718`).
- `cmdVerEvento_Click`: exige selección y rechaza la cabecera (`:212-221`); lee el ID de la columna 0 (`:224-228`); carga el evento con `Constructor.getEvento` (`:232-236`); calcula «facturado» con `DCount` sobre `TbFacturaEventosInvolucrados` (`:24-45`, `:238`); fija `strFranqueado = "Sí"` por defecto (`:239`); cierra `FormEventoEdicion` si estaba abierto (`:241-243`), asigna la global `m_ObjEventoActivo` y lo abre con `OpenArgs` `ID|facturado|franqueado` (`:245-246`), formato que ese formulario exige de tres partes (`Form_FormEventoEdicion.cls:567-575`).
- `cmdDetalleSLA_Click`: mismas comprobaciones (`:279-295`); abre `FormEventoSLA` con `acFormReadOnly` y el ID como `OpenArgs` (`:298`; `Form_FormEventoSLA.cls:53`).

### Generar informe (`cmdExportarHTML_Click`, `:621-678`)

Usa las fechas guardadas o, si están vacías, las de los cuadros (`:634-642`); sin fechas ⇒ «No hay datos filtrados para exportar» (`:644-647`); con `ListCount = 0` ⇒ «No hay datos en la lista para exportar» (`:650-653`). Llama a `SLAHTMLService.bas:738` (`GenerarInformeHTML_Corporate`, `:656`); si devuelve el prefijo `#ERR|` muestra el resto como error (`:658-661`).

## Reglas, validaciones y cálculos

- **Rango:** inicio ≤ fin, con fechas válidas (`:92-115`); `SLAReportService.bas:107-127` repite la validación.
- **Eventos incluidos:** `Franqueado = True` y `FECHAALTAEVENTO` entre ambas fechas inclusive, orden descendente por fecha e ID (`SLAReportService.bas:149-154`). No se exige ningún campo SLA informado, pese a que el comentario de la rutina lo indica (`:69-70`; `SLAReportService.bas:152-153`).
- **Cabecera simulada:** `ColumnHeads = False` y primera fila de la lista con los rótulos, por lo que el índice de datos empieza en 1 (`:52, 61, 218-221`).
- **Estado «facturado»:** «Sí» si existe al menos una fila en `TbFacturaEventosInvolucrados` con el ID; ante cualquier error, «No» (`:34-44`).
- **Cumplimiento SLA (TRES/TRCM/TRSS/global):** no se calcula en este formulario; el bloque equivalente de este fichero está comentado (`:349-614`). Lo resuelve `SLAHTMLService.bas` (`CalcularCumplimiento`) y `SLAReportService.bas:207`; su detalle queda fuera de esta ficha.

## Lecturas y escrituras

| Operación | Origen | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Eventos franqueados del rango | `TbEventos` (alias `E`) | SQL dinámico concatenado con fechas literales `#aaaa-mm-dd#`; `OpenRecordset` DAO | no | `SLAReportService.bas:129-156`; `:128` |
| Estado facturado | `TbFacturaEventosInvolucrados` | `DCount` con el ID entre comillas (apóstrofos duplicados) | no | `:34` |
| Carga del evento | `TbEventos` | `SELECT *` con el ID concatenado sin escapar | no | `Constructor.bas:2253-2254` |
| Informe HTML | sistema de ficheros | `CreateTextFile` (Unicode, sobrescribe) en el directorio local configurado o, si falta, `Documents` del usuario | no | `SLAHTMLService.bas:1099-1121` |
| Apertura del informe | proceso del sistema | `Shell "cmd /c start …"` en modo oculto | no | `SLAHTMLService.bas:1123` |

El formulario no escribe en tablas. `SLAHTMLService.bas` escribe el HTML y lo lanza con el visor predeterminado del sistema.

## Coexistencia

- **Dialecto:** literales de fecha `#aaaa-mm-dd#` y `Franqueado = True` (`SLAReportService.bas:149-151`); en PostgreSQL requieren parámetros de fecha y booleanos nativos. Alias `AS E` y `ORDER BY` son compatibles.
- **Funciones Jet/ACE:** `DCount` (`:34`), `Format` (`:142, 155`) y `Nz`; hay que sustituirlas por consultas o lógica de aplicación.
- **Tipos:** la columna `FechaRecepcionNotificacion` se compara con la fecha cero `30/12/1899` (`:156`), artefacto de `Nz`/`Format` sobre nulos de Access.
- **Transacciones:** solo lectura; no hay transacciones que trasladar.
- **Efectos fuera de la base:** escritura de fichero y `Shell` (`SLAHTMLService.bas:1115-1123`) no tienen equivalente directo en una aplicación web (descarga o visor en línea).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `FECHAALTAEVENTO <= #fin#` compara con la fecha sin hora (`SLAReportService.bas:151`): si el campo guarda hora, los eventos del último día podrían quedar fuera.
- `strFranqueado = "Sí"` se fija sin comprobar el dato (`:239`), coherente solo mientras la consulta filtre `Franqueado = True` (`SLAReportService.bas:149`).
- `getEvento` concatena el ID sin escapar apóstrofos (`Constructor.bas:2253-2254`), a diferencia de `:34`.
- `Form_Load`, `cmdLimpiarEventos_Click` y `cmdSalir_Click` usan `On Error Resume Next` (`:320, 336, 685`) y ocultan fallos.
- Si `OpenRecordset` falla tras el cursor de espera, el flujo de error lo restablece; pero el recordset no se cierra en la rama de error (`:128, 182-190`).
- Tras «Limpiar» o una búsqueda sin resultados, `cmdVerEvento`/`cmdDetalleSLA` no se deshabilitan explícitamente (`:319-329, 164-168`); solo lo hace `ListaFiltrados_Click` (`:713-716`).
- Si se cambian las fechas sin pulsar «Obtener Eventos», el informe usa las fechas guardadas de la última búsqueda, no las visibles (`:634-642`).
- Rótulo «Genear Informe» con errata (`.form.txt:881`) y nombre `cmdObtgenerEventos` (`:72`). Rutina `cmdExportarExcel_Click` comentada (`:349-614`) y funciones `ExisteCampoEnRecordset`/`ObtenerValorCampo` sin uso aparente (`:725-753`; a verificar con el índice).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtro de fecha con hora inclusiva; informe descargable desde el servidor; tabla paginada en lugar de lista de valores; mostrar el cumplimiento de cada evento en la lista.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con fechas vacías, no válidas o inicio > fin, «Obtener Eventos» muestra el mensaje correspondiente y no consulta (`:92-115`).
2. Con un rango válido, la lista muestra solo eventos franqueados con alta en el rango, por fecha descendente, con las seis columnas (`SLAReportService.bas:149-154`; `:131-159`).
3. Un rango sin eventos deja `cmdExportarHTML` deshabilitado (`:164-168`).
4. Seleccionar una fila habilita los dos botones de detalle; la cabecera no es seleccionable (`:699-718`).
5. «Ver Detalle Evento» abre el evento con su estado de facturación (`:238-246`); «Ver Detalle SLA» lo abre en solo lectura (`:298`).
6. «Limpiar» vacía fechas y lista (`:322-328`).
7. «Genear Informe» genera el HTML del periodo de la última búsqueda y lo abre (`:634-656`; `SLAHTMLService.bas:1121-1123`).

## Pendientes

- Estructura y tipos de `TbEventos` (incluida la parte horaria de `FECHAALTAEVENTO`) y de `TbFacturaEventosInvolucrados` (ERD de origen).
- Reglas de cálculo de TRES, TRCM y TRSS en el HTML (`SLAHTMLService.bas`, `CalcularCumplimiento`) y contenido del informe: pendiente de la ficha de módulos.
- Ruta que devuelve `m_ObjEntorno.URLDirectorioLocal` (`SLAHTMLService.bas:1102-1107`; `Entorno.cls:191`).
- Referencias dinámicas: `DoCmd.OpenForm "FormEventoEdicion"` (`:246`) y `"FormEventoSLA"` (`:298`) por nombre literal; `FormularioAbierto` (`:241`; `Funciones Generales.bas:171`).
