# Form_FormInformesParaRAC

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormInformesParaRAC.cls` (329 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Lista de las averías (eventos) franqueadas que todavía no figuran en ningún informe al RAC, con acceso al detalle del evento, exportación a Excel y generación del informe, que además registra la fecha del día como fecha de informe en cada evento incluido. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos llegan de `Constructor.getEventosSinFechaInformeRAC`. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas son `Form_FormInformesParaRAC.cls:línea` salvo indicación; aquí se abrevian a `:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDInformes` | `lblInformeParaRAC_Click`: abre el formulario y cierra el de informes | `Form_Form0BDInformes.cls:70-74` |

Abre `FormEventoEdicion` con `OpenArgs` = `IDEVENTO|facturado|franqueado` (`:49`) y escucha su evento `FechaInformeRACambiada` (`:10, 307`; declarado en `Form_FormEventoEdicion.cls:12`, lanzado en `Form_FormEventoEdicion.cls:507`). Abre el informe generado con `Ejecutar` (`ShellExecuteA`, `VG.bas:15, 39`; `:126`). Sale con `cmdSalir` (`:11-14`) o al terminar de generar el informe (`:135`).

## Controles y eventos

Controles: 23 (23 `Begin` con tipo en el `.form.txt`: formulario, 9 definiciones por defecto `.form.txt:143-278`, sección `Detalle`, 2 rectángulos, 1 lista, 5 botones y 4 etiquetas).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | lista 4 columnas con encabezados (`P.TRABAJO`, `DESCRIPCIÓN`, `FINICIAL`, `FFINAL`) | `OnClick` → `ListaFiltrados_Click`; `OnDblClick` | Carga el evento de la fila y habilita Detalle; el doble clic ejecuta Detalle si está habilitado | CAP-10 | `:211-263`, `.form.txt:356-394` |
| `ComandoDetalleEvento` («Detalle») | botón, deshabilitado en diseño | `OnClick` | Abre `FormEventoEdicion` con el evento seleccionado | CAP-3 | `:17-63`, `.form.txt:506-536` |
| `ComandoGenerarInforme` («Generar Informe») | botón, deshabilitado en diseño | `OnClick` | Genera el informe de todos los eventos de la lista y registra la fecha | CAP-10 | `:100-146`, `.form.txt:537-568` |
| `ComandoExportarExcel` (rótulo «Comando227») | botón, deshabilitado en diseño | `OnClick` | Exporta la lista con `GenerarConsultas` | CAP-10 | `:67-95`, `.form.txt:422-462` |
| `ComandoRecargar` (rótulo «Comando227») | botón | `OnClick` | Relee la lista | CAP-10 | `:148-174`, `.form.txt:628-660` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar, con `On Error Resume Next` | CAP-10 | `:11-14`, `.form.txt:395-421` |
| `lblFiltrados` | etiqueta | — | «Sin  Informe Previo (n)» con el recuento | CAP-10 | `:293`, `.form.txt:569-595` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «AVERÍAS FRANQUEADAS NO INCLUIDAS EN OTROS INFORMES AL RAC»; `lblTitulo1` pasa a `Caption` | CAP-10 | `:191`, `.form.txt:463-505` |
| `lblEstado` | etiqueta | — | Texto de progreso; se oculta al abrir | CAP-10 | `:197`, `.form.txt:596-627` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:310-355` |

Además: `Form_Load` (`:176-209`) y `m_FormEvento_FechaInformeRACambiada` (`:307-330`). El `Caption` de diseño del formulario («PARTES DE TRABAJO EN UN PERIODO», `.form.txt:39`) se sustituye por el de `lblTitulo1` (`:191`).

## Flujos

### Apertura (`Form_Load`, `:176-209`)

1. `AjustarTamaño Me`; si `m_ObjEntorno` es `Nothing`, llama a `EVE` (`:184-189`; `VG.bas:264`).
2. `EstablecerLista` (`:265-305`): fija los encabezados, vacía la selección y deshabilita Detalle, Exportar y Generar (`:274-278`); lee los eventos pendientes (`:279`). Sin eventos sale con todo deshabilitado (`:283-287`). Con eventos añade una fila `IDEvento;Descripcion;FECHAALTAEVENTO;FechaFinal` (`:288-292`), rotula el recuento (`:293`), selecciona la fila única si solo hay una (`:294-297`) y habilita Exportar y Generar (`:298-299`).
3. Oculta `lblEstado` (`:197`). Error: `MsgBox` crítico y cierre (`:202-208`).

### Selección y detalle

`ListaFiltrados_Click` (`:211-254`) deshabilita Detalle, vacía la selección si no hay ID (`:222-228`), relee el evento con `Constructor.getEvento` si cambia (`:229-241`) y habilita Detalle (`:242`). `ComandoDetalleEvento_Click` (`:17-63`) calcula «Sí/No» de facturado y franqueado con `IsDate` sobre `FechaFacturado` y `FechaFinal` (`:36-45`), cierra `FormEventoEdicion` si estaba abierto y lo abre pasando los tres datos (`:46-52`). Cuando éste cambia la fecha de informe al RAC, se repite `EstablecerLista` (`:315`).

### Generar informe (`ComandoGenerarInforme_Click`, `:100-146`)

1. Sin colección: «No hay elementos en la lista» (`:111-114`).
2. `Informe.EventosRAC m_Col, m_URLInforme, True` (`:115-117`; `Informe.cls:23-60`) genera el libro con hoja resumen. Si el fichero no existe: «No se ha podido obtener la URL del informe» (`:121-124`).
3. Abre el informe (`:126`), llama a `GrabarFechaInformeRACACole m_Col` (`:127`), que por cada evento ejecuta `InformeRACRegistrarFecha` con la fecha del día (`FG.bas:54-76`; `Evento.cls:1442-1458`).
4. Relee la lista y cierra el formulario (`:131-135`).

### Exportación (`:67-95`)

`GenerarConsultas("Eventos para Informe", Me.hWnd, , Me.ListaFiltrados, "No")` (`:79`); si el resultado contiene `|`, muestra el detalle como error (`:80-84`).

### Errores

Error 1000 = mensaje preparado; otro número se compone con número y descripción; `MsgBox` crítico (`:57-62` y análogos).

## Reglas, validaciones y cálculos

- **Selección de eventos:** `FechaEnInformeRAC Is Null` y `Franqueado=true` (`Constructor.bas:2664-2667`); se indexan por `IDEVENTO` sin duplicados (`Constructor.bas:2689-2691`).
- **Informe sobre toda la lista**, no sobre la selección (`:116`).
- **Fecha de informe:** la del día del sistema (`FG.bas:71`).
- Botones: Exportar y Generar solo con eventos (`:298-299`); Detalle solo con fila seleccionada (`:220, 242`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Eventos sin fecha de informe | `TbEventos` | SQL literal en VBA (sin parámetros) | — | `Constructor.bas:2664-2669` |
| Lectura de un evento | `TbEventos` | vía `Constructor.getEvento` | — | `:27, 230` |
| Registro de la fecha de informe | `TbEventos` | SQL dinámico concatenado (`UPDATE`, fecha `#mm/dd/yyyy#`, ID entre comillas), `Execute` | no | `Evento.cls:1450-1452` |
| Exportación | — | `GenerarConsultas` sobre la lista | — | `:79` |
| Informe | libro Excel nuevo en ruta que decide `AbrirNuevoExcel` | — | — | `Informe.cls:59-60` |

## Coexistencia

A adaptar a PostgreSQL vinculado: literal de fecha Jet `#mm/dd/yyyy#` y `Format` (`Evento.cls:1450`); `Franqueado=true` (booleano Jet, `Constructor.bas:2667`); `Nz(..., "")` aplicado a todos los campos (`Constructor.bas:2680`); `CurrentDb().Execute` en `Evento.cls:1452`; el diccionario usa `TextCompare` (`Constructor.bas:2687`). Las actualizaciones se hacen evento a evento, sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `GrabarFechaInformeRACACole m_Col` se llama sin `p_Error` (`:127`): un fallo al registrar la fecha no llegaría al formulario, que lo trataría como éxito y cerraría; y un fallo intermedio dejaría parte de los eventos marcados (`FG.bas:67-76`).
- El informe se abre antes de grabar las fechas (`:126-127`): si falla el registro, el informe ya existe pero los eventos seguirían como pendientes.
- Tras generar se ejecuta `EstablecerLista` y se cierra el formulario (`:131-135`); la relectura no tiene efecto visible.
- `ListaFiltrados_DblClick` usa `On Error Resume Next` y `Application.Echo False` (`:257-262`): un error intermedio dejaría silenciados los fallos.
- El texto de error de `ListaFiltrados_Click` carece de espacio («Clickse», `:250`).
- El `RowSource` y rótulos de diseño contienen datos residuales («Copiando plantila de excel…», `.form.txt:610`); no se reproducen. `ComandoRecargar` y `ComandoExportarExcel` comparten rótulo «Comando227» (`.form.txt:433, 638`).
- Si `FormEventoEdicion` ya estaba abierto, se cierra con `acSaveNo` (`:46-48`): se descartarían los cambios sin guardar que tuviera.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: generar el informe solo sobre filas elegidas; registrar las fechas en una transacción única antes de entregar el informe; consulta con parámetros.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir se listan solo eventos franqueados sin fecha de informe, con recuento en `lblFiltrados` (`Constructor.bas:2664-2667`; `:293`).
2. Sin eventos pendientes, Exportar y Generar quedan deshabilitados (`:283-287`).
3. Detalle abre el evento seleccionado y, si cambia su fecha de informe, la lista se actualiza (`:49, 307-330`).
4. Generar informe incluye todos los eventos de la lista, abre el fichero y deja de listarlos (`:116, 126-127`).
5. Si el informe no se genera, se muestra «No se ha podido obtener la URL del informe» y no se registra fecha (`:121-124`).
6. Exportar genera el Excel «Eventos para Informe» con el contenido de la lista (`:79`).

## Pendientes

- Contenido de `Informe.EventosRAC` más allá de `Informe.cls:60`: plantilla, ruta del libro y columnas (`:116`).
- Contenido de `AbrirNuevoExcel`, `AjustarTamaño`, `Avance` y `GenerarConsultas` (`:184`, `:130`, `:79`).
- Si `FechaEnInformeRAC` admite valores nulos con restricciones en `TbEventos` (ERD de origen).
