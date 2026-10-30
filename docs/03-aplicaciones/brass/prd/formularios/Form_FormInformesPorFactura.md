# Form_FormInformesPorFactura

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormInformesPorFactura.cls` (312 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Generador de informes de partes de trabajo (eventos) de una factura: se elige una factura, se cargan los eventos que tiene asociados, se pueden quitar algunos del informe, exportar la lista a Excel o generar el informe en Excel con `Informe.EventosRAC`. Es un formulario **no enlazado** (sin `RecordSource`); el conjunto de eventos vive en la global `m_ColEventosParaInformeDeRAC` (`VG.bas:218`) y la lista `ListaFiltrados` es una lista de valores. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son `Form_FormInformesPorFactura.cls:línea`. Capacidad de todos los controles: CAP-10 Informes.

## Entrada y salida

Lo abre `Form_Form0BDInformes` con `DoCmd.OpenForm` (`Form_Form0BDInformes.cls:79`). Hacia otros formularios: `FormEventoGestion` desde `ComandoDetalleEvento` (`:25-39`) y `FormParteGestion` desde `ComandoGenerarInformesPorCentroParaFirma` (`:237`). `cmdSalir` cierra sin guardar (`:8-11`).

## Controles y eventos

Controles: 25 (27 `Begin` en el `.form.txt`: 25 controles más `Form` y `Section`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComboIDFactura` | combo | — (sin evento) | Facturas (`IDFactura`, `Titulo1`) de más reciente a más antigua | `:266-268`, `.form.txt:1739-1772` |
| `ComandoFiltrar` | botón «Filtrar» | `OnClick` | Carga los eventos de la factura elegida | `:109-153`, `.form.txt:1561-1572` |
| `ListaFiltrados` | lista de valores, 4 columnas, cabecera, selección múltiple (`MultiSelect =2`) | `OnClick`, `OnDblClick` | `P.TRABAJO;DESCRIPCIÓN;FINICIAL;FFINAL`; el clic habilita Eliminar y Detalle; el doble clic abre el detalle | `:154-195, 292-311`, `.form.txt:937-955, 1142-1143` |
| `ComandoEliminar` («Eliminar del Informe») | botón, inicia deshabilitado | `OnClick` | Quita de la colección los eventos seleccionados, no de la factura | `:43-77`, `.form.txt:1521-1533` |
| `ComandoDetalleEvento` («Detalle») | botón, inicia deshabilitado | `OnClick` | Abre `FormEventoGestion` filtrado por el evento seleccionado | `:15-41`, `.form.txt:1600-1612` |
| `ComandoExportarExcel` | botón, inicia deshabilitado | `OnClick` | Exporta la lista a Excel | `:79-107`, `.form.txt:1425-1437` |
| `ComandoGenerarInforme` («Generar Informe») | botón, inicia deshabilitado | `OnClick` | Genera el informe en Excel y lo abre | `:197-232`, `.form.txt:1640-1652` |
| `ComandoGenerarInformesPorCentroParaFirma` | botón, inicia deshabilitado | `OnClick` | Abre `FormParteGestion` | `:234-246`, `.form.txt:1774-1786` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra | `:8-11`, `.form.txt:1347-1357` |
| `ComandoAyudaFormulario` | botón con imagen | `OnClick` declarado | Sin procedimiento en el `.cls` | `.form.txt:1386-1396` |
| `lblEstado` | etiqueta, invisible | — | Se oculta en la carga | `:272, 278`, `.form.txt:1680-1695` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta16` | 3 etiquetas | — | Títulos y rótulo «IDFactura»; `lblTitulo1` pasa a `Caption` | `:265`, `.form.txt:1478-1510, 1712-1724` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | `.form.txt:891-935` |
| Definiciones por defecto | 9 `Begin` sin `Name` (etiqueta, rectángulo, imagen, botón, casilla, grupo de opciones, cuadro de texto, lista, combo) | — | Plantillas de estilo | `.form.txt:740-878` |

## Flujos

### Carga (`Form_Load`, `:249-290`)

1. `AjustarTamaño Me` (`FG.bas:261`); crea `m_ObjEntorno` si no existe (`:260-263`).
2. Rellena `ComboIDFactura` con `RellenaCombo` (`:266-268`; `FG.bas:1149`), crea la colección global vacía y llama a `RellenarLista` (`:270-271`).
3. Error: `MsgBox` crítico; el formulario queda abierto (`:284-289`).

### Filtrar (`:109-153`)

1. Deshabilita Eliminar, Detalle y Exportar (`:119-121`). Sin factura elegida: «Se ha de elegir la Factura» (`:124-127`).
2. `Constructor.getEventosEnFactura` devuelve un diccionario por `IDEvento` (`:131`; `Constructor.bas:2752-2799`); `RellenarLista` pinta una fila por evento con `IDEvento;Descripcion;FECHAALTAEVENTO;FechaFinal` (`:172-175`).
3. Con eventos, habilita Exportar y Generar Informe; «Informes por Centro Para Firma» solo si `ListCount > 2`, es decir, más de un evento por la cabecera (`:183-189`).

### Quitar, detalle y exportación

- **Quitar:** exige selección; confirma con «No» por defecto; borra de la colección las claves seleccionadas y repinta (`:51-68`).
- **Detalle:** requiere exactamente una fila (`:18-23`); abre `FormEventoGestion` si no estaba abierto, pone a `Null` los siete filtros salvo `ComboIDEventoBusqueda` (= evento) y llama a `Form_FormEventoGestion.ActualizarEventos` (`:25-39`). `On Error Resume Next` (`:17`).
- **Exportar:** `GenerarConsultas("Eventos para Informe", Me.hWnd, , Me.ListaFiltrados, "No")`; un resultado con «|» es error (`:91-96`; `FG.bas:1384`).
- **Generar informe:** `Informe.EventosRAC` con la colección y hoja resumen (`:209-211`; `Informe.cls:23-28`); comprueba que el fichero resultante existe («No se ha podido obtener la URL del informe», `:216-219`) y lo abre con `Ejecutar` (`:221`; `VG.bas:15, 39`).

## Reglas, validaciones y cálculos

- **Factura obligatoria** para filtrar (`:124-127`).
- **Informe sin eventos:** `EventosRAC` rechaza colección vacía o `Nothing` con «No se ha seleccionado ningún Parte de Trabajo para el Informe» (`Informe.cls:42-49`).
- **Una hoja por evento**, con actividades y materiales agregados por actividad (`Informe.cls:59-117`).
- **Unicidad:** el diccionario se indexa por `CStr(IDEVENTO)` y descarta duplicados (`Constructor.bas:2794-2796`).

## Lecturas y escrituras

| Operación | Origen | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Combo de facturas | `TbFacturaPrincipal` | SQL fija vía `RellenaCombo` | no | `:266-268` |
| Eventos de la factura | `TbFacturaEventosInvolucrados` + `TbEventos` | SQL dinámica concatenando `IDFactura` | no | `Constructor.bas:2770-2774` |
| Detalle de evento | formulario `FormEventoGestion` | asignación a controles de búsqueda | no | `:29-38` |
| Informe | libro de Excel nuevo y plantilla | automatización de Excel; sin escrituras en la base | no | `Informe.cls:50-58` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: join de Jet/ACE entre `TbFacturaEventosInvolucrados` y `TbEventos` con `IDFactura` concatenado sin comillas, es decir, asumido numérico (`Constructor.bas:2771-2772`); `TbFacturaPrincipal.IDFactura` ordenado como orden de combo (`:266`). El resto de la funcionalidad depende de Excel (automatización COM) y de `ShellExecute`, no de la base.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `m_Error` solo se reinicia en `Form_Load` (`:258`); `ComandoFiltrar_Click`, `ComandoEliminar_Click` y `ComandoGenerarInforme_Click` lo pasan a funciones que comprueban `m_Error <> ""` (`:131-133`, `:65-67`, `:212`) y `getEventosEnFactura` no lo reinicia (`Constructor.bas:2765-2767`): tras un primer error, las operaciones siguientes podrían fallar con el mensaje antiguo.
- `ComandoGenerarInforme_Click` llama a `Hourglass False` al empezar en lugar de `True` (`:205`), y la generación no muestra cursor de espera.
- Si la factura no tiene eventos, `getEventosEnFactura` devuelve `Nothing` sin error (`Constructor.bas:2776-2780`) y la lista queda vacía sin aviso.
- `ListCount > 2` (`:185`) cuenta la cabecera: el botón «por Centro» se habilita con dos eventos, no con uno; a verificar si es la intención.
- `ComandoAyudaFormulario` declara `OnClick` (`.form.txt:1396`) sin manejador en el `.cls`.
- `ComandoDetalleEvento_Click` usa `On Error Resume Next` (`:17`) y puede ocultar fallos al abrir `FormEventoGestion`.
- La lista guarda en `RowSource` de diseño valores de ejemplo (`.form.txt:955`) que `RellenarLista` sustituye (`:162`).
- Rótulos de diseño sin sentido: «Comando46», «Comando227» (`.form.txt:1395, 1436`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: aviso explícito de factura sin eventos, informe descargable sin Excel instalado y trazabilidad de qué eventos se excluyeron.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, el combo lista las facturas de más reciente a más antigua y la lista está vacía (`:266-271`).
2. Filtrar sin factura muestra «Se ha de elegir la Factura» (`:124-127`).
3. Filtrar con factura lista sus eventos con código, descripción y fechas inicial y final (`:172-175`).
4. Eliminar, Detalle, Exportar e Informe están deshabilitados hasta que corresponde (`:119-121, 163-164, 183-184, 292-301`).
5. Quitar eventos pide confirmación y los excluye del informe sin tocar la factura (`:55-65`).
6. El doble clic o «Detalle» con un solo evento seleccionado abre su gestión (`:18-39, 304-311`).
7. Generar informe crea un libro con una hoja por evento y lo abre (`Informe.cls:70-76`; `:221`).

## Pendientes

- Si hay otras aperturas por menú o macro, no visibles en `src`.
- Contenido de `Informe.EventosRAC` más allá de `Informe.cls:23-117` y de la plantilla de Excel que usa.
- Comportamiento de `ActualizarEventos` en `Form_FormEventoGestion` (`:38`) y si respeta el filtro por evento.
- Formato y ubicación del fichero de `GenerarConsultas` con `strRemota = "No"` (`:91`).
