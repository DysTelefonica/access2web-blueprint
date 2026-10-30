# Form_FormMaterialesGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialesGestion.cls` (1393 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla «Gestión de los Materiales»: filtra `TbMaterial` con once criterios (BUI, subsistema, equipo, PN, NS, reparación, entrega, material, ID, irreparable) y muestra el resultado en `ListaFiltrados`, desde la que se edita o consulta el material, se abre su seguimiento y sus anexos, se elimina o se exporta a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena con `AddItem` desde un `OpenRecordset` (`:1121, 1206, 1254`). Las citas son `Form_FormMaterialesGestion.cls:línea` (o `:línea`) salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Capacidad de todos los controles: CAP-7 Materiales, reparaciones y repuestos.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | Etiqueta `lblGestionDeLosMateriales`: `DoCmd.OpenForm "FormMaterialesGestion"` sin argumentos | `Form_Form0BDOpciones.cls:130` |
| `Form_FormEventoGestion` | `ComandoDetalleMaterial_Click`: cierra el formulario si está abierto, lo abre, escribe el ID en `ComboIDMaterialBusqueda` y llama a `ActualizarLista` (pública) | `Form_FormEventoGestion.cls:1630, 1656-1661` |

Hacia otros formularios (todos con el ID del material como `OpenArgs`, cerrando antes el que ya estuviera abierto): `FormMaterialReparacionEdicion` (`:153`), `FormMaterialRepuestoEdicion` con `ID|franqueado|facturado|URL anexo` (`:192`), `FormMaterialIntervencionAlta` (`:526`), `FormMaterialSeguimiento` con `ID|nº seguimientos|IDEvento|facturado|franqueado` (`:530-531`) y `FormAnexos` tras fijar `m_ObjEntidadParaAnexoActiva` (`:107-112`). Sale con `cmdSalir` (`:1011-1014`). Exporta con `GenerarConsultas` (`:296`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista 9 columnas con encabezado | `OnClick` → `ListaFiltrados_Click` | Guarda ID, reparación, evento; habilita/rotula botones; calcula nº de seguimientos | `:1015-1088`; `.form.txt:604-621, 907` |
| `ComandoEdicionMaterialSinActividad` | botón («Ver») | `OnClick` | Abre edición de reparación o de repuesto según la columna 6 | `:125-204`, `.form.txt:1353-1355` |
| `ComandoEliminar` | botón | `OnClick` | Confirma y elimina (con o sin actividad) | `:205-286` |
| `ComandoSeguimiento` («Seguimiento(0)») | botón | `OnClick` | Abre alta de seguimiento o su lista | `:468-543` |
| `ComandoAdjuntos` («Adjuntar») | botón | `OnClick` | Abre `FormAnexos` del material | `:90-123` |
| `ComandoExportarExcel` | botón | `OnClick` | Exporta la lista a Excel | `:288-311` |
| `ComandoActualizarTodo` | botón | `OnClick` (público) | Ejecuta los diez `ComandoLimpiar*` | `:62-89` |
| `ComandoLimpiarBUI`, `…SubSistema` | botones | `OnClick` | Vacían su combo, relistan y recargan el combo dependiente | `:313-328, 452-467` |
| `ComandoLimpiarEquipo`, `…PN`, `…NS`, `…EsReparacion`, `…Entrega`, `…Material`, `…IDMaterial`, `…Irreparable` (8) | botones | `OnClick` | Vacían su combo y relistan (ver defectos sobre `…EsReparacion`) | `:330-450`; `.form.txt:1697, 1946, 2112, 2278, 2444, 2610, 2973, 3735` |
| `ComboBUIBusqueda` | combo | `OnChange`/`OnExit` | Si cambia, recarga subsistema y relista | `:546-592` |
| `ComboSubsistemaBusqueda` | combo | `OnChange`/`OnExit` | Si cambia, recarga equipo y relista | `:949-994` |
| `ComboEquipoBusqueda` | combo | `OnChange`/`OnExit` | Relista si cambia la columna 0 | `:682-725` |
| `ComboPNBusqueda`, `ComboNSBusqueda`, `ComboMaterialBusqueda`, `ComboIDMaterialBusqueda`, `ComboEsReparacion`, `ComboEntrega`, `ComboIrreparable` (7) | combos | `OnChange`/`OnExit` | Relistan si cambia el valor respecto al anterior | `:594-681, 726-948`; `.form.txt:1322, 1496, 1535, 1602, 1668, 2943, 3706` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:1011-1014` |
| `ComandoAyudaFormulario` | botón | `OnClick="[Event Procedure]"` sin handler en el `.cls` | Sin efecto aparente | `.form.txt:1244-1246` |
| `Etiqueta248`, `251`, `254`, `258`, `260`, `8`, `267`, `268`, `273`, `13`, `lblTitulo1`, `lblTitulo2` (12) | etiquetas | — | Rótulos (MAT, PN, NS, Entrega, Rep., ID, BUI, SUBS., Equipo., Irreparable); `lblTitulo1` pasa a `Caption` | `:1001`; `.form.txt:1406-3672` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | `.form.txt:566, 590` |
| `Detalle` | sección | — | Contenedor | `.form.txt:546` |

Controles: 50 (50 Begin en el .form.txt): el formulario (1), cinco plantillas de control sin nombre (etiqueta, rectángulo, botón, lista y combo; `Form_FormMaterialesGestion.form.txt:436-523`), la sección (1) y 43 nombrados: 18 botones, 10 combos, 12 etiquetas, 1 lista y 2 rectángulos.

## Flujos

### Apertura (`Form_Open`, `:996-1010`)
Con `On Error Resume Next`: copia el título, `AjustarTamaño Me` (`FG.bas:261`) y `ActualizarLista`. Sin ningún filtro, `FiltrarMateriales` sale tras dejar la lista solo con el encabezado (`:1121-1125`): la pantalla se abre vacía.

### Filtrado (`ActualizarLista`, `:7-61`)
1. Deshabilita editar, eliminar, seguimiento y adjuntos y llama a `FiltrarMateriales` con los once combos (`:16-29`).
2. `FiltrarMateriales` fija el encabezado, construye el SQL con `WHERE` por PN, NS, reparación, entrega, material, irreparable e ID (si hay equipo, dos ramas `OR`: equipo del evento o del material) y lo abre con `CurrentDb().OpenRecordset` (`:1127-1206`).
3. Recorre el resultado y descarta en VBA por ID de equipo, BUI y subsistema (el del evento o, si vacío, el del equipo), deduce «Fact» («Sí» si hay `IDFactura`) y añade fila `ID;Material;Inicio;Entrega;Equipo;IDEvento;Rep.;NS;Fact` (`:1211-1256`). Habilita exportar si hay filas (`:1209`).
4. Con una sola fila la selecciona y lanza el clic; si no, relanza el clic sobre la fila seleccionada (`:36-50`).
Error: `FiltrarMateriales` devuelve `#ERR|texto` y `ActualizarLista` lo muestra con `MsgBox` (`:30-34, 55-60`).

### Selección (`ListaFiltrados_Click`, `:1015-1088`)
Habilita adjuntos y seguimiento (`:1038-1039`). Si la columna 7 es «Sí», editar pasa a «Ver» y eliminar se deshabilita; si no, «Editar» y eliminar habilitado (`:1046-1056`). `IntervencionNumeroPorMaterial` cuenta los seguimientos y el botón se rotula «Seguimiento(n)» (`:1063-1076, 1284-1333`).

### Editar/Ver (`:125-204`)
Exige ID y columna 6 en «Sí»/«No» (`:133-140`). Si «Sí», abre `FormMaterialReparacionEdicion` (`:153`). Si «No», calcula el franqueo con `EstaFranqueado("MA", ID)` (`:172-178`), arma la URL del anexo con `m_ObjEntorno.URLDirectorioDocumentacion` y la columna 9 (`:179-182`) y abre `FormMaterialRepuestoEdicion` (`:192`).

### Seguimiento (`:468-543`)
Sin nº numérico de seguimientos, error (`:483-486`). Si el material tiene evento, calcula el franqueo del evento (`:489-503`). Con 0 seguimientos y material facturado o evento franqueado, error «no se puede generar el primero» (`:513-517`); si no, abre `FormMaterialIntervencionAlta` (`:526`). Con 1 o más abre `FormMaterialSeguimiento` (`:530-531`).

### Eliminar (`:205-286`)
Pide confirmación (`:220-224`). Sin evento llama dos veces a `EliminarMaterialSinActividad`, primero con «No» (simulación) y luego con «Sí» (`:228-250`); con evento hace lo mismo con `EliminarMaterialEnActividad` (`:251-274`). Recarga la lista (`:275`). Errores con `MsgBox` (`:280-285`).

### Exportar (`:288-311`)
`GenerarConsultas("Lista de Materiales", Me.hWnd, , Me.ListaFiltrados, "No")` vuelca la lista a un `.xlsx` en el directorio local de `m_ObjEntorno` (`FG.bas:1434-1435`); sin filas, «La consulta no contiene registros» (`FG.bas:1427-1429`).

## Reglas, validaciones y cálculos

- **Combos dependientes:** `EstablecerCombo` rellena subsistema desde `TbSubsistemaBui` (BUI elegido o `-1`) y equipo desde `TbEquipos` (subsistema y BUI); con un único elemento lo preselecciona (`:1341-1378`).
- **Irreparable:** «Sí» ⇒ `TipoAccion = 'IDENTIFICADO_COMO_IRREPARABLE_OBSOLETO'`; «No» no filtra (`:1165-1173`).
- **Facturado:** existencia de `IDFactura` en `TbFacturaMaterialesInvolucrados` (`:1133-1134, 1246-1251`).
- **Eliminar sin actividad:** rechaza si proviene de actividad, si está facturado o si un anexo está abierto; borra los seguimientos con `DoCmd.RunSQL`, el material con `.Delete` y después los ficheros de anexo (`FG.bas:4975-4984, 5011-5024, 5051-5054, 5064-5095`).
- **Eliminar en actividad:** rechaza si está facturado o anexo abierto; el franqueo solo se comprueba si se pide, y por defecto no (`FG.bas:4817-4819, 4858-4861, 4882-4885`).
- **Anti-repetición:** las variables `str…Anterior` evitan relistar si el valor no cambia; `LimpiarUltimos` las reinicia (`:553, 1273-1283`).

## Lecturas y escrituras

| Operación | Tabla/objeto | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista filtrada | `TbMaterial`, `TbActividades`, `TbEventos`, `TbEquipos` (×2), `TbFacturaMaterialesInvolucrados` | SQL dinámico concatenado, `SELECT DISTINCT` con `LEFT/RIGHT JOIN` y `Like '*'` | — | `:1127-1206` |
| Nº de seguimientos | `TbMaterialSeguimiento` | SQL concatenado | — | `:1303-1306` |
| Combos de equipo y subsistema | `TbEquipos`, `TbSubsistemaBui` | SQL concatenado asignado a `RowSource` | — | `:1349-1375` |
| Combos de diseño | `TbMaterial`, `TbBUI` | `RowSource` estático | — | `.form.txt:1324, 1498, 1537, 2945, 3190` |
| Baja de seguimientos | `TbMaterialSeguimiento` | `DELETE` concatenado con `DoCmd.RunSQL` y avisos apagados | no | `FG.bas:5065-5070` |
| Baja de material | `TbMaterial` | DAO `.Delete` | no | `FG.bas:5072-5084` |
| Franqueo / facturación | `TbEventos`, `TbFactura…Involucrad…` | SQL concatenado | — | `FG.bas:5758-5830, 6111-6115` |
| Borrado de anexos | ficheros del directorio de anexos | `fso.DeleteFile` | no | `FG.bas:5089-5093` |

## Coexistencia

Pendiente de `../coexistencia.md`. A adaptar: `Like '*'` y `Is Null` (Jet/ACE) sobre columnas de fecha y texto (`:1136-1166`); `RIGHT/LEFT JOIN` anidados con paréntesis (`:1131-1134`); `SELECT DISTINCT` con campos de varios orígenes; `DELETE tabla.campo FROM` (`FG.bas:5065`); booleano `Franqueado` leído como «Verdadero»/`True` (`FG.bas:6086`); comparación `FechaEntrega='Sí'` sobre un campo de fecha (`:1153`); ausencia de transacción en baja de dos tablas más ficheros.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `strFacturado` se lee de la columna 7 (`:488, 1030`), pero la fila añadida pone `NS` en la 7 y «Fact» en la 8 (`:1254-1256`; `Column(8)` en `:156`): «Ver»/eliminar deshabilitado y el bloqueo del primer seguimiento por facturado podrían no aplicarse.
- La columna 9 (`:179`) no existe en una lista de nueve columnas (`.form.txt:611`): `strURLAnexo` quedaría siempre vacío.
- El filtro «Entrega» ofrece «Entregado;Sin Entregar» (`.form.txt:1604`) pero compara `FechaEntrega` con «Sí»/«No» (`:1150-1153`): nunca filtra.
- `ComandoLimpiarEsReparacion_Click` vacía `ComboIrreparable` (`:363, 369`) y no `ComboEsReparacion`, que además no se limpia desde `ComandoActualizarTodo` (`:62-77`).
- `ComboEquipoBusqueda.Column(3)` (`:27`) sobre un combo de una columna (`.form.txt:3281`): `strIDEquipo` probablemente vacío y el filtro por ID de equipo (`:1219-1227`) inoperante.
- Con todos los filtros vacíos se sale sin deshabilitar exportar (`:1122-1126`); el botón podría seguir activo con la lista vacía.
- `ComandoEliminar` llama a `EliminarMaterialEnActividad` sin respetar franqueo (`:257, 268`; `FG.bas:4817-4819`), mientras seguimiento sí lo comprueba (`:496-502`).
- Baja de seguimientos, material y ficheros sin transacción; los ficheros se borran después de la base (`FG.bas:5064-5095`).
- En `EliminarMaterialSinActividad` el bloque tras `If flag = "Sí"` es inalcanzable: antes se rechaza todo `flag <> "No"` (`FG.bas:4981-4991`).
- SQL concatenado con valores de combos sin escapar apóstrofos (`:1138-1163, 1184-1185, 1305, 1351`).
- Errores silenciados: `Form_Open` y `ComandoLimpiar*` con `On Error Resume Next` (`:997, 314`); `ComandoExportarExcel` no muestra `m_Error` (`:306-310`); `ComandoActualizarTodo` cierra el formulario al fallar (`:88`).
- `ComandoAyudaFormulario` declara `OnClick` sin handler (`.form.txt:1246`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrado y facturado calculados en servidor; baja transaccional con borrado de anexos diferido; consolidar filtros «Entrega» e «Irreparable» con valores explícitos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abierto sin filtros, el listado se muestra vacío (`:1122-1125`).
2. Cada filtro acota el listado; combinados se aplican con «Y» (`:1175-1182`).
3. Al cambiar BUI se recargan subsistemas y se vacía el equipo (`:1341-1357`).
4. Un material con factura se muestra con «Ver» y sin eliminar (`:1046-1049`).
5. Eliminar pide confirmación y no actúa si se responde «No» (`:220-224`).
6. Un material con actividad, facturado o con anexo abierto no se elimina (`FG.bas:4858-4861, 4882-4885`).
7. El rótulo «Seguimiento(n)» refleja los seguimientos del material (`:1076`).
8. Exportar sin filas muestra «La consulta no contiene registros» (`FG.bas:1427-1429`).
9. Abrir desde un evento deja el material filtrado por su ID (`Form_FormEventoGestion.cls:1660-1661`).

## Pendientes

- Contenido de `Dame` y `FicheroAbierto` (`FG.bas:6080, 4882`).
- Contenido de `AjustarTamaño` más allá de `:261-275`.
- Qué columna de `TbMaterial` es `PN`/`TipoAccion` y su tipo (ERD de origen).
- Resto de `GenerarConsultas` tras `FG.bas:1470` (escritura del `.xlsx`).
