# Form_FormNCProyectoAcciones

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoAcciones.cls` (1458 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Acciones» de la NC de proyecto: dos listas maestro-detalle, acciones correctivas (AC) arriba y sus tareas (AR) abajo, con altas, ediciones, borrados, historial, anexos y recálculo de estados. Se incrusta como subformulario de `FormNCProyecto` (`Form_FormNCProyectoAC.cls:213`). Es **no enlazado**: las listas se rellenan por código con `AddItem` sobre una lista de valores (`:680,746`). Trabaja con `m_ObjNCProyectoActiva` y, si la NC tiene NC vinculada, muestra los datos de esta (`:630-633`). Las citas son `Form_FormNCProyectoAcciones.cls:línea` salvo indicación; `ACProyectoOperaciones.cls` y `ARProyectoOperaciones.cls` abrevian `classes/ACProyectoOperaciones.cls` y `ARProyectoOperaciones.cls`. Fichas relacionadas: [AC](Form_FormNCProyectoAC.md) y [AR](Form_FormNCProyectoAR.md).

## Entrada y salida

Se carga con `Form_Load` ⇒ `EstablecerDatos` (`:570-597`); exige NC activa («… parámetros insuficentes», `:607-608`). Abre `FormNCProyectoAC` (alta `:37`, edición `:216`), `FormNCProyectoAR` (alta `:379`, edición `:506`) y `FormARProyectoDocumentos` (`:179`), cerrando antes cualquier instancia abierta. Escucha sus eventos con `WithEvents` (`:12-14`, handlers `:1009-1238`).

## Controles y eventos

`.form.txt` declara 32 bloques `Begin`: el formulario, 8 valores por defecto (`.form.txt:141-278`), la sección `Detalle` (`:299`) y los 22 controles de la tabla.

| Control | Tipo | Evento | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaAC` | lista de 7 columnas | `OnClick` | Selecciona la AC, habilita botones según permisos y recarga las tareas | `:866-934`, `.form.txt:358` |
| `ComandoACAlta`, `ComandoACEdicion`, `ComandoACBorrar`, `ComandoACHistorial` | botones | `OnClick` | Alta, edición, borrado e historial de la AC | `:17,195,58,235` |
| `OrdenAC`, `OrdenAR` | combos de 2 columnas | `AfterUpdate` | Orden de las listas; conservan la selección | `:1238,1395`; `.form.txt:1192,1267` |
| `ListaAR` | lista de 7 columnas | `OnClick` | Selecciona la tarea y habilita edición, borrado, historial y anexos | `:936-1000`, `.form.txt:571` |
| `ComandoARAlta`, `ComandoAREdicion`, `ComandoARBorrar`, `ComandoARHistorial`, `ComandoARDocumentos` | botones | `OnClick` | Alta, edición, borrado, historial y anexos de la tarea | `:349,473,400,525,157` |
| `ComandoActualizarACs`, `ComandoActualizarARs` (rótulo «Contador») | botones | `OnClick` | Recalculan estado y fechas de las AC / tareas | `:276,306` |
| `ImagenAnexo` | imagen | `OnClick` | Abre el anexo único o la ficha de anexos | `:793-850` |
| `ImagenNotaAC`, `ImagenNotaAR` | imágenes | `OnClick` | Indican que hay notas y llevan a la edición | `:852-864`, `:912-913,975-976` |
| `Etiqueta64` «Tareas», `Etiqueta246`, `Etiqueta75` «Ordenado por», etiqueta de `.form.txt:311` | etiquetas | — | Rótulos | `.form.txt:311,524,1141,1215` |

## Flujos

### Carga y selección
`EstablecerDatos` habilita «Actualizar» solo si la NC no tiene NC vinculada y el alta de AC si además el usuario no es técnico (`:616-626`); fija el orden por defecto (`:637-638`) y rellena ambas listas. `EstablecerListaAC` consulta `constructor.getACsProyecto` (`:689`), añade filas con `;` sustituido por `:` en el texto (`:706`), y selecciona la primera (`:716-718`). `ListaAC_Click` relee la AC y habilita alta de tarea, edición, historial y borrado según `m_PermitidoEditar`, NC vinculada y rol (`:895-908`). `ListaAR_Click` hace lo propio con la tarea (`:962-990`).

### Alta y edición
Alta de AC: rechaza técnicos (`:24`) y exige `DatosGeneralesOK` («Debe regularizar primero los datos que faltan en la ficha general», `:28-29`). Alta de tarea: exige AC seleccionada y los mismos controles (`:356-372`). Tras `Alta`/`Editado` de los formularios hijos se recalcula, se recarga la lista y se avisa «La NC de proyecto pasa a estar cerrada» o «… ha dejado de estar cerrada» según cambie el estado (`:1025-1036,1145-1156,1206-1217`).

### Borrado
- AC: confirma con «No» por defecto; si tiene tareas pregunta «… junto con sus N tareas asociadas?» (`:86-103`); `ACProyectoOperaciones.Eliminar` borra tareas y la AC y recalcula la NC (`ACProyectoOperaciones.cls:318-343`).
- Tarea: confirma «¿Desea borrar la Acción Realizada Seleccionada?» (`:425-428`); `ARProyectoOperaciones.Eliminar` ejecuta `DELETE`, invalida la caché y recalcula AC y NC (`ARProyectoOperaciones.cls:351-366`).

### Otros
- Historial: `HTMHistorialProyecto` (`Funciones Generales.bas:1551`) y `HTMLENTXT` (`:1693`) generan y muestran el historial (`:252-256,542-546`).
- Actualizar: `ActualizarDatosACsProyecto` / `ActualizarDatosARsProyecto` (`FG:2535,2661`) y refresco de la gestión (`:283-290,323-332`).
- Errores: 1000 ⇒ aviso con el motivo; otro ⇒ correo al administrador y mensaje crítico (patrón en cada handler, p. ej. `:49-56`).

## Reglas, validaciones y cálculos

- Técnicos: sin alta ni borrado (`:24,68,356,408,619,897,983`).
- Edición de los hijos: bloqueada con NC vinculada o `m_PermitidoEditar` = No (`:681-684,895-909,980-984`).
- Mensaje «Seleccione un elemento de la lista superior» si no hay selección (`:78,171`).
- Cierre/reapertura de la NC detectado comparando estados (`:119-130`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de AC y tareas | vía `constructor.getACsProyecto` / `getARsDeACProyecto` | no resuelto | `:689,754` |
| Borrar tarea | `TbNCAccionesRealizadas` | SQL dinámico concatenado, `Execute` | `ARProyectoOperaciones.cls:351-354` |
| Borrar AC | `TbNCAccionCorrectivas` | SQL dinámico concatenado, `Execute`; antes borra cada tarea | `ACProyectoOperaciones.cls:318-334` |
| Altas y ediciones | en las fichas AC y AR | DAO | ver fichas hijas |

Sin transacción: el borrado de una AC elimina tareas una a una y luego la AC (`ACProyectoOperaciones.cls:320-334`).

## Coexistencia

`DELETE *` (`ARProyectoOperaciones.cls:351`, `ACProyectoOperaciones.cls:331`) es dialecto Jet: en PostgreSQL, `DELETE FROM`. Borrado en cascada manual ⇒ transacción única o `ON DELETE CASCADE`. Las listas se montan como cadenas delimitadas por `;`: en la web, datos tipados. Fechas concatenadas como texto (`:706-707,773-774`).

## Comportamiento actual · Posibles defectos · Propuestas

**Comportamiento actual.** Las actualizaciones de línea reemplazan en el `RowSource` la fila seleccionada por la nueva (`:1316-1329`), sin recargar la lista.

**Posibles defectos (a verificar).**
- `ComandoARBorrar_Click` relee la AC con `Me.ListaAR.Column(0)`, que es el id de la **tarea**, no de la AC (`:444`).
- `ActualizarLineaLista*` copia solo 6 columnas (`j = 0 To 5`) de una fila de 7 (`:1318,1376`), y reemplaza texto por coincidencia.
- Un nombre de AC con `;` se altera en pantalla pero no en base (`:706`).
- Los botones «Contador» no describen su efecto (`.form.txt:841,885`).

**Propuestas web pendientes de decisión.** Una sola vista maestro-detalle con recarga de datos tipados; confirmación de borrado unificada.

## Criterios de aceptación

1. Sin NC activa, aviso de parámetros insuficientes (`:607-608`).
2. Un técnico no puede dar de alta ni borrar AC/tareas (`:24,68,356,408`).
3. Alta de AC o tarea sin datos generales completos se rechaza con el mensaje indicado (`:28,370`).
4. Borrar una AC con tareas avisa del número de tareas y las elimina (`:97-103`; `ACProyectoOperaciones.cls:318-328`).
5. Con NC vinculada o no editable, no hay altas ni bajas (`:616-626,895-909`).
6. Al cambiar el estado a/desde «Cerrada» se muestra el aviso correspondiente (`:119-130`).
7. Cambiar el orden mantiene la selección (`:1249-1261`).

## Pendientes

- Cuerpos de `getACsProyecto`, `getARsDeACProyecto`, `DatosGeneralesOK`, `EstablecerColorPestañasNCProyecto` (`:1144`) y `EnumOrden`.
- Tramo posterior de `ACProyectoOperaciones.Eliminar` (`ACProyectoOperaciones.cls:343`).
