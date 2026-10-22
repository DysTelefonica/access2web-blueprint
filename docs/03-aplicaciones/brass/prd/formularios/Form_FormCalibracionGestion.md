# Form_FormCalibracionGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormCalibracionGestion.cls` (625 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de las calibraciones de un equipo calibrable: cabecera de solo lectura con equipo, modelo, número de serie y BUI, una lista de calibraciones (`TbEquiposCalibrablesFechas`) y botones para dar de alta, editar, borrar, adjuntar o quitar el documento anexo y exportar la lista a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista tiene `RowSource` propio. Escribe en la base con SQL ejecutado por `DoCmd.RunSQL` y copia o borra ficheros en el directorio de anexos de calibraciones. Las citas son `Form_FormCalibracionGestion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas` y `VG.bas` = `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalibracionEquipoGestion` | cierra el formulario si estaba abierto y lo abre con `OpenArgs` = «IDEquipo\|texto de calibraciones» | `Form_FormCalibracionEquipoGestion.cls:111-117` |
| `Form_FormCalibracionEquipoEdicion` | lo reabre con `OpenArgs` = ID del equipo | `Form_FormCalibracionEquipoEdicion.cls:113-116` |

`Form_Open` acepta `OpenArgs` con barra (ID y texto) o solo el ID; si el ID no es numérico, error «No se ha abierto con los parámetros adecuados» y cierre (`:252-263, 327-337`). Hacia otros formularios: abre `FormCalibracionAlta` con el ID del equipo (`:143`) y `FormCalibracionEdicion` con el ID de la calibración (`:175`), cerrándolos antes si estaban abiertos (`:140-142, 172-174`). Quienes guardan después refrescan la lista con `Requery` (`Form_FormCalibracionAlta.cls:43-44`; `Form_FormCalibracionEdicion.cls:129-131, 170-172`). Al borrar, llama a `Form_FormCalibracionEquipoGestion.ActualizarLista` si está abierto (`:225-227`). Sale con `cmdSalir` (`:339-342`).

Estado de módulo: ID de la calibración seleccionada, fecha de baja para calibración, ID del equipo y URL del anexo (`:4-5`). Usa las globales `m_SQL`, `pregunta`, `flag` y `dato` (`VG.bas:203-206`).

## Controles y eventos

Controles: 35 (35 Begin en el .form.txt: 25 controles de la sección Detalle, 8 plantillas de estilo por defecto, el formulario y la sección Detalle).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | lista 5 columnas con encabezados (anchos «0;0;2268;2268;0») | `OnClick` → `ListaFiltrados_Click` (público) | Lee ID (columna 0) y URL del anexo (columna 4); sin selección deshabilita Editar, Eliminar y Adjuntar; con selección habilita Adjuntar y, si hay anexo, muestra `ImagenAnexo` y rotula «Quitar Adjunto» | CAP-5 | `:367-409`, `Form_FormCalibracionGestion.form.txt:960-983` |
| `ComandoAltaCalibracion` («Alta Calibración») | botón | `OnClick` | Abre `FormCalibracionAlta` con el ID del equipo | CAP-5 | `:125-154`, `Form_FormCalibracionGestion.form.txt:1029-1040` |
| `ComandoEditarCalibracion` («Editar Calibración») | botón, deshabilitado en diseño | `OnClick` | Abre `FormCalibracionEdicion` con la calibración elegida | CAP-5 | `:155-186`, `Form_FormCalibracionGestion.form.txt:1001-1013` |
| `ComandoEliminarCalibracion` («Eliminar Calibración») | botón, deshabilitado en diseño | `OnClick` | Pide confirmación, borra anexo y calibración | CAP-5 | `:187-240`, `Form_FormCalibracionGestion.form.txt:1056-1069` |
| `ComandoAdjuntarDocumento` («Adjuntar» / «Quitar Adjunto») | botón, deshabilitado en diseño | **sin `OnClick` en el diseño**; existe `ComandoAdjuntarDocumento_Click` | Según el rótulo, adjunta o quita el documento | CAP-11 | `:29-124`, `Form_FormCalibracionGestion.form.txt:1699-1710` |
| `ComandoExportarExcel` | botón (icono), deshabilitado en diseño | `OnClick` | Pasa el `RowSource` de la lista a `GenerarConsultasSQL` | CAP-10 | `:6-28`, `Form_FormCalibracionGestion.form.txt:1658-1670` |
| `ImagenAnexo` | imagen, oculta en diseño | `OnClick` | `AbrirEnLocal` con la URL del anexo y el `hWnd` del formulario | CAP-11 | `:343-366`, `Form_FormCalibracionGestion.form.txt:1388-1400` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-5 | `:339-342`, `Form_FormCalibracionGestion.form.txt:1088-1098` |
| `ComandoAyudaFormulario` | botón (icono) | `OnClick` en el diseño, **sin handler en el `.cls`** | Sin código asociado | sin asignar | `Form_FormCalibracionGestion.form.txt:1578-1588` |
| `EQUIPO`, `MODELO`, `NS`, `BUI` | 4 cuadros de texto bloqueados | — | Se rellenan en `Form_Open` solo si el valor leído no está vacío | CAP-5 | `:274-293`, `Form_FormCalibracionGestion.form.txt:1242-1255, 1271-1285, 1301-1315, 1358-1372` |
| `lblCalibraciones` | etiqueta | — | El texto de diseño se sustituye por el de `OpenArgs` si llega | CAP-5 | `:294-296`, `Form_FormCalibracionGestion.form.txt:1548-1559` |
| `lblTitulo1`, `lblTitulo2` | 2 etiquetas | — | `lblTitulo1` pasa a ser el `Caption` del formulario | CAP-5 | `:250, 302`, `Form_FormCalibracionGestion.form.txt:1617-1647` |
| `Etiqueta191`, `Etiqueta192`, `Etiqueta193`, `Etiqueta197` | 4 etiquetas fijas | — | Rótulos EQUIPO, MODELO, NS, BUI | CAP-5 | `Form_FormCalibracionGestion.form.txt:1161-1173, 1188-1200, 1215-1227, 1331-1343` |
| `Marco120`, `Marco189` | 2 grupos de opciones | — | Marcos decorativos sin código | sin asignar | `Form_FormCalibracionGestion.form.txt:1115-1125, 1138-1148` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2` | 3 rectángulos | — | Decorativos | sin asignar | `Form_FormCalibracionGestion.form.txt:891-946` |

Recuento: 1 lista + 4 botones de gestión + `ComandoAdjuntarDocumento` + `ComandoExportarExcel` + `ImagenAnexo` + `cmdSalir` + `ComandoAyudaFormulario` (=10) + 4 textos + 5 etiquetas de texto propio (`lblCalibraciones`, 2 títulos… ver tabla) + etiquetas fijas + marcos + rectángulos = 25 en Detalle.

## Flujos

### Apertura (`Form_Open`, `:242-338`)

1. `AjustarTamaño Me` y título (`:249-250`); interpreta `OpenArgs` (`:252-263`).
2. Construye `m_SQL` con un `SELECT` sobre `TbEquiposCalibrables` (equipo, modelo, NS, BUI y fecha de baja) concatenando el ID (`:264-267`) y lo abre con `CurrentDb().OpenRecordset`; sin fila, error de parámetros (`:268-273`).
3. Vuelca los campos no vacíos en `EQUIPO`, `MODELO`, `NS` y `BUI` (`:282-293`). Si la fecha de baja es una fecha válida, deshabilita `ComandoAltaCalibracion` (`:297-301`).
4. Asigna **ese mismo `m_SQL`** como `RowSource` de `ListaFiltrados` y la refresca (`:303-305`). Sin elementos deshabilita la exportación; con uno, lo selecciona y ejecuta `ListaFiltrados_Click`; con varios, deselecciona todo (`:306-320`).

### Selección en la lista (`ListaFiltrados_Click`, `:367-409`)

Si el anexo guardado contiene separador de unidad o de red, se queda con el nombre de fichero (`:387-392`); en cualquier caso reconstruye la URL con `m_ObjEntorno.URLDirectorioCalibraciones` (`:395`; `Entorno.cls:451-473`, que exige que la carpeta sea alcanzable).

### Adjuntar (`:29-124`, `AnexarDocumento` `:410-510`)

1. Exige calibración seleccionada (`:35-39`). Con el rótulo «Adjuntar», abre el selector de ficheros (`Seleccionar`, `FG.bas:1117`); si el fichero no existe sale en silencio (`:50-62`).
2. `AnexarDocumento`: rechaza si el origen no existe o está abierto (`:436-449`), o si el anexo previo o el destino están abiertos (`:452-484`). Nombre destino = ID con formato de cuatro cifras + extensión original (`:471-474`). Borra el anterior, copia el fichero y ejecuta `UPDATE` de `URLDocumentosAnexados` con el nombre (`:485-495`).
3. Muestra `ImagenAnexo` y rotula «Quitar Adjunto» (`:77-79`).

### Quitar adjunto y borrar

- `QuitarAnexo` (`:511-577`): primero `UPDATE ... = Null` (`:535-539`) y después borra el fichero si había URL (`:540-561`).
- Con el rótulo «Quitar Adjunto», pregunta «¿Desea eliminar el adjunto?» (por defecto No) y llama a `QuitarAnexo` (`:80-103`).
- `ComandoEliminarCalibracion_Click`: pregunta «¿Desea borrar la calibración seleccionada junto con el anexo que pudiera tener?» (por defecto No) (`:201-205`), llama a `EliminarCalibracion` (`:578-624`), que ejecuta `QuitarAnexo` y luego `DELETE` (`:604-616`); refresca la lista (`:225-229`).

### Errores

Todos los handlers usan el error interno 1000 para mensajes propios y `MsgBox` crítico; se quita el cursor de espera antes (p. ej. `:115-124`). Las funciones devuelven «#ERR\|texto» y el llamador lo divide (`:72-76, 509`).

## Reglas, validaciones y cálculos

- **Selección obligatoria** para adjuntar, editar, borrar y anexar: la columna 0 ha de ser numérica (`:35-39, 162-165, 194-197`).
- **Equipo de baja:** una fecha de baja para calibración válida impide el alta de calibraciones (`:297-301`).
- **Nombre del anexo:** ID de calibración con cuatro dígitos mínimo y la extensión del original (`:472-473`).
- **Ficheros abiertos:** `FicheroAbierto` abre el fichero en exclusiva y trata cualquier error como «abierto» (`FG.bas:195-217`).
- **Confirmaciones** con «No» por defecto en borrado y quitar adjunto (`:81, 201-202`).

## Lecturas y escrituras

| Operación | Objeto | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Datos del equipo | `TbEquiposCalibrables` | SQL dinámico concatenado (ID) con `OpenRecordset` | no | `:264-268` |
| Lista de la pantalla | `ListaFiltrados` | `RowSource` = la consulta anterior | — | `:303-305` |
| Anexar | `TbEquiposCalibrablesFechas.URLDocumentosAnexados` | SQL dinámico concatenado (`UPDATE`) con `RunSQL` y avisos desactivados | no | `:491-495` |
| Quitar anexo | ídem | SQL dinámico concatenado (`UPDATE ... Null`) | no | `:535-539` |
| Borrar calibración | `TbEquiposCalibrablesFechas` | SQL dinámico concatenado (`DELETE`) | no | `:610-615` |
| Ficheros | directorio de anexos de calibraciones | `FileSystemObject`: copia y borrado | no | `:485-490, 560` |
| Exportar | `GenerarConsultasSQL` | SQL de la lista | — | `:12-13`, `FG.bas:1584` |

## Coexistencia

Adaptar a PostgreSQL vinculado: `DELETE tabla.campo FROM tabla` es sintaxis Jet (`:610-612`); corchetes y alias con espacios en la consulta de diseño (`Form_FormCalibracionGestion.form.txt:977-981`); `DoCmd.RunSQL` con `SetWarnings` (`:493-495, 537-539, 613-615`); fechas tratadas como texto con `Nz`/`IsDate` (`:278, 297`); el ID de calibración llega como texto y se concatena sin delimitadores (`:267, 492, 536, 612`). La gestión de ficheros (`FileSystemObject`, rutas de red configuradas) queda fuera de la base.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La lista recibe el `SELECT` del equipo (`:264-267, 304`), no el de calibraciones del diseño (`Form_FormCalibracionGestion.form.txt:977-981`); si es así, `Column(0)` no sería el ID de calibración y editar, borrar y anexar fallarían con «Seleccione…».
- `ComandoEditarCalibracion` y `ComandoEliminarCalibracion` nacen deshabilitados (`Form_FormCalibracionGestion.form.txt:1002, 1057`) y el código solo los deshabilita (`:376-377`); no hay `Enabled = True` para ellos.
- `ComandoAdjuntarDocumento` no tiene `OnClick` en el diseño (`Form_FormCalibracionGestion.form.txt:1699-1710`), por lo que `ComandoAdjuntarDocumento_Click` (`:29`) podría no dispararse.
- `QuitarAnexo` evalúa `If (strURLAnexo) Then` (`:555`), probablemente intención de `FicheroAbierto`; con una cadena produciría error de tipo, tras haber anulado ya el campo (`:535-539`) y antes de borrar el fichero.
- Orden en el borrado: se anula el anexo y se borra el fichero antes del `DELETE` sin transacción (`:604-615`); un fallo intermedio deja datos y fichero inconsistentes.
- `AnexarDocumento` ejecuta `DeleteFile` sobre `strURLAnexo` aunque pueda estar vacío (`:485-489`).
- `ComandoAyudaFormulario` tiene evento sin handler (`Form_FormCalibracionGestion.form.txt:1588`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: transacción única de base y fichero; consultas parametrizadas; almacén de ficheros con identificador en lugar de nombre derivado del ID.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con `OpenArgs` no válido se muestra «No se ha abierto con los parámetros adecuados» y se cierra (`:260-263, 327-337`).
2. Con fecha de baja para calibración válida, `Alta Calibración` queda deshabilitado (`:297-301`).
3. Sin selección no se puede editar, borrar ni adjuntar, con los mensajes de `:163, 195, 37`.
4. Borrar pide confirmación, elimina el anexo y la calibración y refresca la lista (`:201-229`).
5. Adjuntar guarda el fichero con nombre de cuatro cifras y actualiza `URLDocumentosAnexados` (`:471-495`).
6. Quitar adjunto deja el campo a `Null` y borra el fichero (`:535-560`).
7. Abrir el anexo (`ImagenAnexo`) lo copia a local y lo abre (`:352`; `FG.bas:218-258`).

## Pendientes

- `GenerarConsultasSQL` (`:13`; `FG.bas:1584`): destino y formato exactos de la exportación.
- `Seleccionar` (`:50`; `FG.bas:1117`): filtros del selector.
- Contenido de `AjustarTamaño` (`:249`; `FG.bas:261`).
- Esquema de `TbEquiposCalibrablesFechas` y tipo de `FechaBajaParaCalibracion` (ERD de origen).
- Quién habilita Editar y Eliminar, si alguien (`:376-377`).
