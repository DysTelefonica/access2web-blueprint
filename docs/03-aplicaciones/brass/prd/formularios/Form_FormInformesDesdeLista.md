# Form_FormInformesDesdeLista

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormInformesDesdeLista.cls` (272 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Generador del informe de partes de trabajo (eventos) para el RAC. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): muestra en una lista de valores (`ListaFiltrados`) los eventos acumulados en la global `m_ColEventosParaInformeDeRAC` (`VG.bas:218`) y, con ellos, genera un libro Excel. Abreviaturas: `VG.bas` = `modules/Variables Globales.bas`; `FG.bas` = `modules/Funciones Generales.bas`. Las citas sin fichero son `Form_FormInformesDesdeLista.cls:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDInformes` | `lblInformeParaRACDesdeLista_Click` abre el formulario y cierra el menú | `Form_Form0BDInformes.cls:87-90` |
| `Form_FormEventoGestion` | Añade el evento seleccionado a la colección global y llama a `RellenarLista`; exige el formulario abierto y rechaza duplicados | `Form_FormEventoGestion.cls:436-461` |
| `Form_FormEventoGestion` | Muestra `ComandoParaInforme` solo si este formulario está abierto | `Form_FormEventoGestion.cls:2077-2081` |

Salidas: abre `FormEventoGestion` filtrado por un evento (`:19-45`), un libro de informe (`:152-185`) y una exportación de la lista a Excel (`:83-111`). Cierra sin guardar con `cmdSalir` (`:8-11`).

## Controles y eventos

Controles: 14 (25 `Begin` en el `.form.txt`: 1 formulario, 9 plantillas de estilo por defecto, 1 sección y 14 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | lista de valores | `OnClick`, `OnDblClick` | Cabecera `P.TRABAJO;DESCRIPCIÓN;FINICIAL;FFINAL`; el clic habilita o no `ComandoEliminar` y `ComandoDetalleEvento` según la fila; el doble clic invoca el detalle | CAP-10 Informes | `.form.txt:937-953, 1142-1143`; `:253-272` |
| `ComandoLlenarDesdeLista` («Cargar Datos») | botón | `OnClick` | Carga los eventos de `TbAuxEventosParaInforme` en la colección y rellena la lista | CAP-10 Informes | `.form.txt:1654-1666`; `:188-223` |
| `ComandoAbrirLista` («Leer Lista») | botón | `OnClick` | Abre la tabla `TbAuxEventosParaInforme` | CAP-10 Informes | `.form.txt:1623-1635`; `:14-16` |
| `ComandoEliminar` | botón | `OnClick` | Quita de la colección los eventos seleccionados, previa confirmación | CAP-10 Informes | `.form.txt:1495-1508`; `:47-81` |
| `ComandoDetalleEvento` | botón | `OnClick` | Abre `FormEventoGestion` limpiando filtros y fijando el ID del evento | CAP-3 Eventos | `.form.txt:1527-1540`; `:19-45` |
| `ComandoGenerarInforme` | botón | `OnClick` | Genera el informe con `Informe.EventosRAC` y lo abre | CAP-10 Informes | `.form.txt:1559-1572`; `:152-185` |
| `ComandoExportarExcel` | botón | `OnClick` | Exporta la lista tal cual a Excel con `GenerarConsultas` | CAP-10 Informes | `.form.txt:1413-1425`; `:83-111` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-10 Informes | `.form.txt:1347-1357`; `:8-11` |
| `ComandoAyudaFormulario` | botón | `OnClick` declarado en el `.form.txt` | Sin procedimiento en el `.cls` | sin asignar | `.form.txt:1374-1384` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título; `lblTitulo1.Caption` se copia a `Me.Caption` | CAP-10 Informes | `.form.txt:1475-1484`; `:242` |
| `lblEstado` | etiqueta | — | Se oculta al cargar | CAP-10 Informes | `.form.txt:1591-1605`; `:245` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decorativos | sin asignar | `.form.txt:891-923` |

## Flujos

### Carga (`Form_Load`, `:225-251`)

Con `On Error Resume Next` (`:226`): `AjustarTamaño` (`FG.bas:261`), crea `m_ObjEntorno` si no existe (`:234-236`), copia el título, **crea una colección vacía** (`:243`, descarta la anterior), fija la cabecera de la lista y oculta `lblEstado` (`:244-245`).

### Cargar datos (`:188-223`)

Deshabilita Eliminar, Detalle y Exportar (`:196-198`); `getListaEventosDeTabla` (`Constructor.bas:1347-1400`) une `TbAuxEventosParaInforme` con `TbEventos` por `IDEvento`, ordenado por ID, y construye un `Evento` por fila; si no hay filas devuelve «No existe ningún registro con ese ID» (`Constructor.bas:1367-1370`). `RellenarLista` (`:114-150`) vuelca cada evento como `ID;Descripcion;FECHAALTAEVENTO;FechaFinal` (`:135`) y habilita Exportar y Generar (`:143-144`); con colección vacía las deja deshabilitadas (`:123-130`).

### Quitar de la lista (`:47-81`)

Sin colección: sale en silencio (`:52-54`). Sin selección: «Seleccione al menos un registro» (`:55-58`). Pregunta con «No» por defecto (`:59-62`), elimina cada clave seleccionada y recarga la lista (`:63-69`).

### Detalle (`:19-45`)

Sale en silencio si no hay exactamente una fila seleccionada (`:22-27`). Abre `FormEventoGestion` si no lo está (`:29-31`), anula seis filtros de búsqueda, pone `ComboIDEventoBusqueda` con la columna 0 (`:34-40`), llama a `ActualizarEventos` (`:42`) y le da el foco (`:43`).

### Generar informe (`:152-185`)

`Informe.EventosRAC` con hoja resumen (`:163-164`; `Informe.cls:23-28`) exige colección no vacía (`Informe.cls:46-53`), abre un libro nuevo desde la plantilla configurada (`Informe.cls:301-330`), copia una hoja por evento (`Informe.cls:67-80`) y guarda (`Informe.cls:154`). Después comprueba el fichero (`:169-172`) y lo abre con `Ejecutar` (`:174`).

### Errores

Todos los handlers usan `Err.Raise 1000` como error controlado con `m_Error` y un `MsgBox` crítico (`:75-80`, `:105-110`, `:179-184`, `:216-221`). `RellenarLista` no muestra mensaje: devuelve el texto por parámetro (`:146-149`).

## Reglas, validaciones y cálculos

- La lista de valores es **la única vista**; la fuente real es la colección global, indexada por ID de evento (`:132-135`).
- El libro de exportación se nombra con hora, minuto y segundo y se borra si existe (`FG.bas:1431-1441`); la exportación falla con lista vacía: «La consulta no contiene registros» (`FG.bas:1421-1431`).
- El informe usa un nombre fijo y falla si el anterior está abierto: «Tiene abierto otro informe previo» (`Informe.cls:322-328`).
- Se mezclan en la hoja de materiales dos claves distintas (`Informe.cls:111-114`).

## Lecturas y escrituras

| Operación | Origen | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga de eventos | `TbAuxEventosParaInforme`, `TbEventos` | SQL estática en VBA (`INNER JOIN`, `ORDER BY`), DAO `OpenRecordset` | — | `Constructor.bas:1362-1365` |
| Apertura de la tabla auxiliar | `TbAuxEventosParaInforme` | `DoCmd.OpenTable` | — | `:15` |
| Datos del informe | actividades y materiales del evento | vía `Evento.ColActividades` y `Actividad.ColMateriales` | — | `Informe.cls:84-99` |
| Exportación | contenido de la lista | sin SQL (celdas desde `lst.Column`) | — | `FG.bas:1443-1448` |
| Ficheros | informe y exportación Excel en el directorio local configurado; plantilla configurada | `FileSystemObject`, Excel por automatización | — | `Informe.cls:301-330`; `FG.bas:1535` |

No hay escrituras en tablas.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `INNER JOIN` y `ORDER BY` estándar (portables); dependencia de Excel y de `ShellExecuteA` (`VG.bas:15, 39`), fuera de la base; la tabla auxiliar debe seguir poblándose desde fuera de este formulario (no hay código que la escriba aquí).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoAyudaFormulario` tiene `OnClick` en el `.form.txt` (`.form.txt:1384`) sin procedimiento en el `.cls`.
- `ListaFiltrados.RowSource` está volcada con datos de ejemplo en el `.form.txt` (`.form.txt:955`), que `Form_Load` y `RellenarLista` sobrescriben (`:122, 244`).
- En `RellenarLista`, la clave de la lista usa `m_IDEvento` pero un evento nulo se omite sin aviso (`:134-138`).
- `ComandoDetalleEvento_Click` tiene `On Error Resume Next` (`:21`): un fallo al abrir `FormEventoGestion` pasaría inadvertido.
- `Informe.cls:113` comprueba `Exists(m_IDMaterial)` pero la clave añadida es `m_IDMaterial & "|" & IDActividad` (`Informe.cls:111, 114`): la condición nunca detecta duplicados.
- `ComandoGenerarInforme_Click` pone `Hourglass False` al empezar (`:160`), por lo que no hay indicador de espera durante la generación.
- `Application.Echo False` en el doble clic sin `On Error GoTo` real (`:266-271`): con error, `Echo` podría quedar desactivado.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir la tabla auxiliar y la colección global por una selección persistida por usuario.

## Criterios de aceptación

Propuestos, no aprobados:

1. «Cargar Datos» muestra todos los eventos de la tabla auxiliar ordenados por ID, con las cuatro columnas (`:114-150`).
2. Con tabla auxiliar vacía aparece «No existe ningún registro con ese ID» y no se habilitan Exportar ni Generar.
3. «Eliminar del Informe» sin selección muestra «Seleccione al menos un registro»; con selección y confirmación los quita.
4. «Detalle» con una sola fila abre la gestión de eventos filtrada por ese evento; con cero o varias no hace nada.
5. «Generar Informe» con colección vacía muestra «No se ha seleccionado ningún Parte de Trabajo para el Informe».
6. «Exportar» con la lista vacía muestra «La consulta no contiene registros».
7. Añadir desde la gestión de eventos un ID ya incluido muestra el aviso de duplicado (`Form_FormEventoGestion.cls:445-448`).

## Pendientes

- Contenido de `RellenarDatosEvento` y `RellenarActividades` (`Informe.cls:80, 91`).
- Origen de la propia tabla `TbAuxEventosParaInforme` (quién la rellena; ERD de origen).
- Valor de `URLDirectorioLocal` y de la plantilla de informe en `Entorno` (`Informe.cls:305-322`).
