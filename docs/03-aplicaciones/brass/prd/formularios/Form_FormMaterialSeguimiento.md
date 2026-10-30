# Form_FormMaterialSeguimiento

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialSeguimiento.cls` (802 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla de seguimiento de las intervenciones (envíos a reparación o análisis con fechas de inicio y fin y destino) de un material de la tabla `TbMaterial`. Es un formulario **no enlazado**: sin `RecordSource` en el `.form.txt`; los datos se cargan por código con `CurrentDb().OpenRecordset` (`Form_FormMaterialSeguimiento.cls:338, 360`) y la lista de intervenciones tiene su propio `RowSource` (`:413`). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; las citas sin fichero son `Form_FormMaterialSeguimiento.cls:línea`. Capacidad: CAP-7 Materiales, reparaciones y repuestos.

Controles: 43 (53 `Begin` en el `.form.txt`: 1 formulario, 8 plantillas de control por defecto en `.form.txt:738-862`, 1 sección y 43 controles).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMaterialesGestion` (`ComandoSeguimiento_Click`) | Cierra el seguimiento si estaba abierto. Si el material no tiene seguimientos y no está facturado ni franqueado, abre `FormMaterialIntervencionAlta`; si no, abre este formulario con `OpenArgs` = `IDMaterial\|nº intervenciones\|IDEvento\|facturado\|franqueado` | `Form_FormMaterialesGestion.cls:509-531` |
| `Form_FormMaterialIntervencionAlta` / `Form_FormMaterialIntervencionEdicion` | Tras guardar, si este formulario está abierto llaman a `EstablecerDatos` para refrescarlo | `Form_FormMaterialIntervencionAlta.cls:76-77`; `Form_FormMaterialIntervencionEdicion.cls:162-163` |

`OpenArgs` (`:264-283`): sin `\|` se toma como `IDMaterial`; con `\|` exige exactamente 5 partes (`UBound = 4`) o aborta con «No se ha abierto el formulario con los datos necesarios». `IDMaterial` vacío también aborta. Hacia otros formularios abre `FormEventoGestion`, `FormAnexos`, `FormMaterialIntervencionAlta` y `FormMaterialIntervencionEdicion` (ver controles).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaIntervenciones` | lista (5 columnas, la 0 oculta) | `OnClick` → `ListaIntervenciones_Click` | Muestra `IDSeguimiento`, `F Inicio`, `Destino`, `F Fin` y `Última` del material, de más reciente a más antigua. Sin selección deshabilita Edición, Eliminar y Adjuntos; con selección habilita Edición y Adjuntos, y Eliminar solo si `strBloqueado = "No"` | CAP-7 | `.form.txt:942-964`; `:408-413, 490-524` |
| `ComandoAlta` | botón «Alta» | `OnClick` | Rechaza si el seguimiento ya está cerrado o hay una intervención sin fecha de fin; cierra `FormMaterialIntervencionAlta` si estaba abierto y lo abre con `IDMaterial` | CAP-7 | `:79-134`; `.form.txt:982-992` |
| `ComandoEdicion` | botón «Editar» (deshabilitado de diseño) | `OnClick` | Exige selección; cierra y reabre `FormMaterialIntervencionEdicion` con el `IDSeguimiento` | CAP-7 | `:136-166`; `.form.txt:1155-1167` |
| `ComandoEliminar` | botón «Eliminar» (deshabilitado de diseño) | `OnClick` | Pide confirmación (por defecto «No»), ejecuta `SeguimientoMaterialEliminar` dos veces (validar y borrar) y refresca | CAP-7 | `:167-231`; `.form.txt:1111-1124` |
| `ComandoAdjuntos` | botón «Adjuntar» (deshabilitado de diseño) | `OnClick` | Exige selección, carga `MaterialSeguimiento` con `Constructor.getMaterialSeguimiento`, la deja en `m_ObjEntidadParaAnexoActiva` y abre `FormAnexos` | CAP-11 | `:44-78`; `Constructor.bas:1897-1939`; `.form.txt:2010-2022` |
| `ComandoVerEvento` | botón «Ver Evento» (deshabilitado de diseño) | `OnClick` | Abre `FormEventoGestion` o, si ya está abierto con otro evento, ejecuta `ComandoActualizarTodo_Click` y lo reposiciona en `ComboIDEventoBusqueda` | CAP-3 | `:6-43`; `.form.txt:2121-2133` |
| `ComandoExportarExcel` | botón (rótulo «Comando227», deshabilitado de diseño) | `OnClick` | Pasa el `RowSource` de la lista a `GenerarConsultasSQL`, que vuelca la consulta a un `.xlsx` temporal y lo abre | CAP-10 | `:232-254`; `FG.bas:1584-1688`; `.form.txt:1195-1206` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra el formulario sin guardar | CAP-7 | `:486-489`; `.form.txt:1020-1030` |
| `IDEVENTO`, `IDActividad`, `EQUIPO`, `NODO`, `BUI`, `SUBSISTEMA`, `MATERIAL`, `GARANTÍA` (8 cuadros de texto bloqueados) | texto | — | Datos del material, rellenados solo si el valor no está vacío; `GARANTÍA` solo admite «Sí»/«No» y en otro caso queda «No» | CAP-7 | `:378-406`; `.form.txt:1269-1409, 1586-1633, 1878-1892` |
| `UltimaFechaInicial`, `UltimaFechaFinal`, `UltimoDestino` (3 cuadros bloqueados) | texto | — | Datos de la última intervención (`Null` si no son fecha / vacío) | CAP-7 | `:430-444` |
| `FINTERVENCION` | texto bloqueado («Acabado») | — | «Sí»/«No» según `SeguimientoCerrado` | CAP-7 | `:456-472` |
| `IDMATERIAL` | texto bloqueado | — | Sin escritura desde el código (ningún `Me.IDMATERIAL`); solo figura en el diseño | CAP-7 | `.form.txt:2090-2104` |
| `lblFacturado`, `lblFranqueado` | etiquetas (ocultas de diseño) | — | Visibles si `OpenArgs` marca facturado / evento franqueado | CAP-8 | `:316-325`; `.form.txt:2174-2218` |
| `lblTitulo1`, `lblTitulo2`, `lblListaIntervenciones` | etiquetas | — | `lblTitulo1` se reescribe a «SEGUIMIENTO DEL MATERIAL <id>» y da el `Caption` del formulario; `lblListaIntervenciones` queda con el texto fijo «Intervenciones ( 1)» | CAP-7 | `:263, 407`; `.form.txt:1081-1092, 1967-2000` |
| `Etiqueta2`, `Etiqueta142`, `235`, `236`, `237`, `240`, `252`, `261`, `263`, `264`, `266`, `268`, `270` (13 etiquetas) | etiquetas | — | Rótulos de los campos | sin asignar | `.form.txt:1424-1936, 2063-2075` |
| `CuadroSuperior`, `Cuadro2`, `Marco120`, `Marco246` | 2 rectángulos y 2 grupos de opciones | — | Decoración | sin asignar | `.form.txt:894, 918, 1059, 1247` |

## Flujos

### Apertura (`Form_Open`, `:255-308`)

1. `AjustarTamaño Me` (`FG.bas:261`) y `Caption`; interpreta `OpenArgs` (arriba).
2. `EstablecerDatos` (`:309-485`): según `strIDEventoAlInicio` lee el material por `TbMaterial`→`TbActividades`→`TbEventos`→`TbEquipos` (`:331-354`) o solo `TbMaterial`→`TbEquipos` (`:355-377`, sin evento ni actividad). Sin filas: «No se ha obtenido ningún dato con el IDMaterial».
3. Calcula `strBloqueado` = «Sí» si facturado o franqueado (`:326-330`), carga la lista (`:408-413`), la última intervención (`IntervencionUltimosDatos`, `:525-599`) y `FINTERVENCION`.
4. `ComandoAlta` queda habilitado si el seguimiento no está cerrado y `strBloqueado = "No"`; si `SeguimientoCerrado` devuelve «No» se habilita siempre (`:456-472`).
5. Cualquier error: `MsgBox` crítico y cierre del formulario (`:301-307`).

### Alta (`:79-134`)

`SeguimientoCerrado` = «Sí» ⇒ «El proceso ya estaba finalizado previamente» (`:96-99`; `FG.bas:3872-3929`). `MaterialSeguimientoConItervencionAbierta` = «Sí» (algún seguimiento con `FechaFinIntervencion` nula) ⇒ «Hay una intervención previa sin cerrar» (`:111-114`; `FG.bas:4231-4288`). Si no, abre el formulario de alta; este refresca el seguimiento al guardar.

### Eliminación (`:167-231`; `SeguimientoMaterialEliminar`, `:601-802`)

Primera pasada `"No"` (solo valida) y segunda `"Sí"` (borra). Validaciones: material no franqueado (`EstaFranqueado("MA", …)`, `:658-667`), no facturado (`EstaFacturado("MA", …)`, `:673-682`) y, con más de una intervención, solo la «Última» (`:714-719`). Con `"Sí"` elimina cada anexo de la intervención (`Anexo.Eliminar`, `:724-740`; `Anexo.cls:912-938`), borra la fila de `TbMaterialSeguimiento` (`:742-754`) y, si el seguimiento queda abierto, pone `TbMaterial.FechaEntrega = Null` (`:769-775`). Errores: `MsgBox` con el texto de `strTextoError`; el código 1000 es el error de negocio.

### Alternativos y errores

Todos los handlers abren `Hourglass`, capturan errores y muestran `MsgBox` crítico (p. ej. `:37-43, 72-78`); `cmdSalir` ignora errores (`:487`).

## Reglas, validaciones y cálculos

- **Intervención abierta:** la que no tiene `FechaFinIntervencion` (`FG.bas:4267`); solo puede haber una antes de dar de alta otra.
- **Seguimiento cerrado:** alguna fila con `UltimaIntervencion = "Sí"` (`FG.bas:3905-3907`).
- **Posición de la intervención:** «Primera» / «Intermedia» / «Última» por los dos últimos caracteres de `IDSeguimiento` (`FG.bas:4666-4672`). Con una única intervención exige `TbMaterial.FechaIncio` con fecha (`FG.bas:4598-4614`).
- **Última intervención mostrada:** si la última no tiene fecha de fin y hay más de una (según `OpenArgs`), toma la fecha de fin de la anterior (`:565-568`).
- **Bloqueo:** facturado o evento franqueado impide Eliminar y, en la apertura, Alta (`:326-330, 458-461, 508-512`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Datos del material (con o sin evento) | `TbMaterial`, `TbActividades`, `TbEventos`, `TbEquipos` | SQL dinámico concatenado con `IDMaterial` entre comillas | — | `:332-337, 356-359` |
| Lista de intervenciones | `TbMaterialSeguimiento` | SQL dinámico en `RowSource` | — | `:408-413` |
| Última intervención, cierre, abierta | `TbMaterialSeguimiento` | SQL dinámico concatenado | — | `:554-558`; `FG.bas:3883-3886, 4252-4262` |
| Posición y fechas vecinas | `TbMaterialSeguimiento`, `TbMaterial` | SQL dinámico con `First/Last(Right(...))`, `Dame` | — | `FG.bas:4628-4750`; `FG.bas:290-300` |
| Facturado / franqueado | `TbFacturaMaterialesInvolucrados`, `TbEventos` | SQL dinámico | — | `FG.bas:5794, 5809-5815, 6033, 6086` |
| Baja de intervención | `TbMaterialSeguimiento` | DAO `Delete` sobre un recordset | no | `:742-754` |
| Baja de anexos | `TbAnexos` + ficheros | `DELETE` por ejecución directa y borrado de fichero | no | `Anexo.cls:925-931` |
| `FechaEntrega = Null` | `TbMaterial` | `UPDATE` con `DoCmd.RunSQL` y avisos desactivados | no | `:770-774` |
| Exportación | consulta de la lista | recordset + Excel (automatización) + fichero local temporal | — | `FG.bas:1619-1670` |

## Coexistencia

Adaptar a PostgreSQL vinculado: SQL concatenado con `IDMaterial` entre comillas (`:337, 359, 411, 744`); `Dame` prueba primero sin comillas y reintenta con ellas capturando el error (`FG.bas:303-312`); `First/Last(Right([IDSeguimiento],2))` y comparaciones `Right(...) >` (`FG.bas:4628, 4649, 4718`); `DELETE *` de Jet (`Anexo.cls:925`); `DoCmd.RunSQL`; el booleano `Franqueado` se compara con «Verdadero»/«True» (`FG.bas:6086`); identificador `IDSeguimiento` con sufijo de orden (`DameIDIntervencion`, `FG.bas:4289`); `Right`, `Nz`, `Null` y valores «Sí»/«No» como texto. La exportación depende de Excel y de un directorio local de entorno.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `IntervencionDatosDefechasIntermedias` no filtra por `IDMaterial` al buscar la anterior/posterior (`FG.bas:4718, 4732`) y, para la «Última», usa `Not IDMaterial = …` (`FG.bas:4697`): podría tomar intervenciones de otros materiales.
- Si `SeguimientoCerrado` devuelve «No», `ComandoAlta` se habilita aunque el material esté facturado o franqueado (`:468-469` frente a `:458-461`); la validación de bloqueo no existe en `ComandoAlta_Click` (`:79-134`).
- `CLng(strNumeroIntervencionesEnSeguimiento)` en `IntervencionUltimosDatos` (`:565`) falla si `OpenArgs` llegó sin `\|` (cadena vacía, `:266-267`); `And` no hace cortocircuito.
- `colAnexosAEliminar` nunca se rellena (`:777-784`); el borrado de ficheros de esa colección es código inerte. Los anexos se borran antes que la fila de la intervención y sin transacción (`:724-754`; `Anexo.cls:925-931`).
- Tras eliminar, `EstablecerDatos` reasigna `RowSource` pero no reevalúa los botones: Edición/Adjuntos podrían quedar habilitados sin selección (`:413`, `:490-524`).
- `ComandoExportarExcel` está deshabilitado en el diseño (`.form.txt:1196`) y ningún código lo habilita; `IDMATERIAL` nunca se rellena; `lblListaIntervenciones` es texto fijo (`.form.txt:1092`).
- Mensajes de error con nombre equivocado: «El método Eliminar…» en `IntervencionUltimosDatos` (`:421`) y en `EstablecerDatos` tras eliminar (`:218`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador de intervención generado por la base y orden explícito en lugar de sufijo textual; baja de intervención y anexos en una sola transacción; recalcular el contador y los botones desde el servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `IDMaterial` o con `OpenArgs` de partes distintas de 5 muestra «No se ha abierto el formulario con los datos necesarios» y cierra (`:266-283`).
2. Un material facturado o de evento franqueado muestra la etiqueta correspondiente y no permite eliminar (`:316-330, 508-512`).
3. Alta con seguimiento cerrado o con intervención sin fecha de fin muestra el mensaje y no abre el alta (`:96-114`).
4. Eliminar una intervención intermedia con posteriores muestra «No se puede eliminar una intervención intermedia…» (`:714-719`).
5. Eliminar la última con seguimiento resultante abierto deja `FechaEntrega` vacía (`:769-775`).
6. `Ver Evento` abre la gestión del evento con ese evento seleccionado (`:21-31`).
7. Tras guardar un alta o una edición, la lista y la última intervención se refrescan (`Form_FormMaterialIntervencionAlta.cls:76-77`).

## Pendientes

- Contenido de `AjustarTamaño` más allá de la firma (`FG.bas:261`) y de la API `Ejecutar` (`Variables Globales.bas:15`).
- Construcción de `MaterialSeguimiento.ColAnexos` (`MaterialSeguimiento.cls:38`) y destino de `m_ObjEntidadParaAnexoActiva` en `FormAnexos` (`:65`).
- Si el `.xlsx` temporal se elimina tras abrirlo (`FG.bas:1670`) y la ruta de `m_ObjEntorno.URLDirectorioLocal` (`FG.bas:1608`).
- Formato de `IDSeguimiento` (sufijo de dos dígitos) y reglas de `FormMaterialIntervencionAlta`/`Edicion`: otras fichas.
