# Form_FormNCProyectoSeguimientoTareas

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoSeguimientoTareas.cls` (599 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Tareas» del seguimiento de proyecto: lista las tareas (acciones realizadas) de NC de proyecto activas o pendientes de replanificar, filtrables por responsable de Calidad, responsable de la tarea y expediente; abre la NC de la tarea elegida en `FormNCProyecto` y exporta la lista filtrada a Excel. Se aloja en `FrmDetalle` de [Form_FormNCProyectoSeguimiento](Form_FormNCProyectoSeguimiento.md). Es un formulario **no enlazado** (sin `RecordSource`); la lista sale de las colecciones `ColSegsTareasProyecto…` del entorno. Las citas son `Form_FormNCProyectoSeguimientoTareas.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCProyectoSeguimiento | Botón de navegación `NavTareas` carga este formulario en `FrmDetalle` | `Form_FormNCProyectoSeguimiento.form.txt:1011` |
| `PintarIndicadores` | Llama a `Filtrar` si `FrmDetalle.SourceObject` es este formulario | `FG.bas:328-329` |

Sale hacia `FormNCProyecto`: fija `m_ObjNCProyectoActiva` con la NC de la tarea, cierra la instancia abierta sin guardar y la abre (`:236-241`); hacia `FormExpedientesBusqueda` (`:85-88, 165-168`); y hacia Excel, que abre el informe con `Ejecutar … "open"` (`:137`). No se cierra a sí mismo salvo error en `Form_Load` (`:323`).

## Controles y eventos

`.form.txt` declara 33 bloques `Begin`: el formulario (`.form.txt:4`), 9 de valores por defecto (`:713-772`), la sección `Detalle` (`:778`) y los 22 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista (valores, 7 columnas) | `OnClick`, `OnDblClick` | Muestra IDAccionRealizada (oculta), NºTRA, Tarea, Responsable, Estado, F.prev.cierre, Tipo NC; el clic fija la tarea seleccionada y habilita «Detalle»; el doble clic abre el detalle si está habilitado | `:430-490`, `.form.txt:785-850` |
| `lblListaDeResultados` | etiqueta | — | Contador «Registros filtrados (n)» | `:390, 417`, `.form.txt:921-932` |
| `ComandoVerDetalle` («Detalle») | botón | `OnClick` | Abre `FormNCProyecto` con la NC de la tarea | `:219-253`, `.form.txt:1062-1074` |
| `ComandoExportarAExcel` («Exp. Excel») | botón (deshabilitado de diseño) | `OnClick` | Exporta la lista filtrada a un libro Excel y lo abre | `:109-154`, `.form.txt:2262-2275` |
| `RESPONSABLECALIDAD` + `Etiqueta57` («Responsable calidad») | combinado (2 columnas) + etiqueta | `AfterUpdate` | Filtra por responsable de Calidad | `:572-598`, `.form.txt:942-985` |
| `Estado` + `Etiqueta261` («Estado tarea») | combinado (2 columnas) + etiqueta | `AfterUpdate` | Filtra por estado de tarea (2 valores) | `:255-281, 355-372`, `.form.txt:1002-1045` |
| `Responsable` + `Etiqueta263` («Responsable tarea») | combinado (2 columnas) + etiqueta | `AfterUpdate` | Filtra por responsable (jefes de proyecto) | `:544-570`, `.form.txt:1477-1487, 1680-1714` |
| `ComandoLimpiarResposableCalidad` | botón | `OnClick` | Pone `RESPONSABLECALIDAD` a Null y filtra | `:213-217`, `.form.txt:1110-1122` |
| `ComandoLimpiarEstado` | botón | `OnClick` | Pone `Estado` a Null y filtra | `:193-197`, `.form.txt:1293-1305` |
| `ComandoLimpiarResposable` | botón | `OnClick` | Pone `Responsable` a Null y filtra | `:207-211`, `.form.txt:1497-1509` |
| `PalabraClave` | texto | `OnKeyDown` | Con Intro o Tab lanza la búsqueda de expediente | `:530-542`, `.form.txt:1732-1749` |
| `ComandoBuscarExpediente` | botón | `OnClick` | Busca expediente por palabra clave | `:15-107`, `.form.txt:1767-1779` |
| `ComandoIrABusquedaExpedientes` | botón | `OnClick` | Abre `FormExpedientesBusqueda` con la palabra clave | `:156-191`, `.form.txt:1851-1863` |
| `ComandoActualizararExpedientes` («Contador») | botón | sin handler | Solo se habilita o deshabilita desde el código | `:47-49`, `.form.txt:1916-1928` |
| `IDExpediente`, `Titulo` | texto bloqueado | — | Expediente elegido (id y nemotécnico, código o id) | `:63-71`, `.form.txt:1989-1990, 2024-2025` |
| `ComandoLimpiarIDExpediente` | botón | `OnClick` | Vacía `IDExpediente`, `PalabraClave` y `Titulo`; filtra | `:199-205`, `.form.txt:2058-2070` |
| `Etiqueta237`, `Etiqueta268` («Expedientes») | etiquetas | — | Decorativas, sin evento | `.form.txt:912-914, 2242-2252` |

Los 22 controles son: `ListaFiltrados`, `Etiqueta237`, `lblListaDeResultados`, `Etiqueta57`, `RESPONSABLECALIDAD`, `Etiqueta261`, `Estado`, `ComandoVerDetalle`, tres botones «limpiar» de filtros, `Etiqueta263`, `Responsable`, `PalabraClave`, `ComandoBuscarExpediente`, `ComandoIrABusquedaExpedientes`, `ComandoActualizararExpedientes`, `IDExpediente`, `Titulo`, `ComandoLimpiarIDExpediente`, `Etiqueta268` y `ComandoExportarAExcel` (`.form.txt:785-2275`).

## Flujos

### Carga (`Form_Load`, `:283-324`)

1. Cursor de espera, `Ajustar Me`, cabecera de la lista (`:289-294`).
2. `EstablecerCombos` rellena responsables de Calidad (`FG.bas:1921`), los 2 estados (`:355-364`) y los jefes de proyecto (`FG.bas:1840`) (`:326-346`).
3. Si `EsTecnico` es `Sí`, `Responsable` = usuario conectado; si no, y el usuario está en `ColUsuariosCalidad`, `RESPONSABLECALIDAD` = usuario (`:300-306`).
4. Fija `Estado` al título del estado `ACTIVA` y ejecuta `Filtrar` (`:308-309`).

### Filtrado (`Filtrar`, `:374-428`)

Deshabilita «Detalle» y «Exp. Excel», pone cabecera y contador a 0 (`:387-392`) y guarda en la global `m_ColFiltradoTareasNCProyectos` el resultado de `constructor.getARsDeProyectoBusqueda` con `Column(0)` de cada combinado e `IDExpediente` (`:395-397`). Si hay resultado habilita «Exp. Excel» (`:404`) y añade una fila por tarea: nº con dos dígitos, tarea con `;` sustituido por `:`, técnico, estado y fecha prevista de cierre (`:412-413`). El estado mostrado es el de la tarea si el filtro es `ACTIVA` o vacío; en otro caso, el texto del filtro (`:408-411`). Después actualiza el contador, selecciona la fila 1 y dispara `ListaFiltrados_Click` (`:417-419`).

### Búsqueda de expediente (`ComandoBuscarExpediente_Click`, `:15-107`)

Sin palabra clave no hace nada (`:29-34`). `getExpedientesBusqueda` (`:35`): sin resultado, vacía campos y abre la búsqueda (`:42-57`); con un único expediente lo asigna y filtra (`:59-84`); con varios abre `FormExpedientesBusqueda` y enlaza el evento `Seleccionado` (`:85-93`). `m_FormExpedientes_Seleccionado` rellena `IDExpediente` y `Titulo` y filtra (`:502-513`).

### Selección y detalle

`ListaFiltrados_Click` carga con `constructor.getARProyecto` la tarea de `Column(0)` si cambia (`:449-461`) y habilita «Detalle» (`:468`). «Detalle» reutiliza la tarea seleccionada o la carga; sin selección avisa «Seleccione un elemento de la lista» (`:226-235`) y abre la NC `m_ARProyectoSeleccionado.AC.nc` (`:236`).

### Exportar a Excel (`ComandoExportarAExcel_Click`, `:109-154`)

Deduce `m_Tipo` del texto de `Estado` (`:119-125`), toma `lblEstado` del formulario padre para mostrar el avance (`:127-128`) y llama a `TareasAExcel` con la colección filtrada (`:129`; `FG.bas:4080-4366`). `TareasAExcel` crea un libro con 9 columnas (Nemotécnico, NC, Nº Tarea, Tarea, Responsable, Estado, Fecha inicio, Fecha prev. cierre, Tipo NC) (`FG.bas:4134-4142`), lo guarda en la subcarpeta `Informes` del directorio local del entorno (`FG.bas:4110-4111, 4341`) y devuelve la ruta (`FG.bas:4348`). Si el fichero no existe, avisa «No se ha podido obtener el documento excel» (`:133-136`).

### Errores

Error 1000 ⇒ advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (p. ej. `:100-106, 147-153`).

## Reglas, validaciones y cálculos

- Estados del combinado: ACTIVA y PTEREPLANIFICAR (`:363-364`); el diseño trae además «ESTADO IRREGULAR» en su `RowSource` (`.form.txt:1043`), que el código reemplaza al cargar (`:362`).
- `getARsDeProyectoBusqueda` elige la colección por título de estado (`ACTIVA`, `PENDIENTE DE REPLANIFICAR`, resto ⇒ todas) y filtra por `RespCalidad`, `Tecnico` e `IDExpediente` (`constructor.bas:2480-2504`).
- Los comentarios del código indican que `RespCalidad` y `Tecnico` son «usuariored» en el filtro (`constructor.bas:2495, 2500`), mientras que `SegTareasProyecto.Tecnico` se documenta como nombre (`SegTareasProyecto.cls:17`) y el combinado entrega `Column(0)` (nombre).
- Fechas del informe con formato `mm/dd/yyyy` (`FG.bas:4198, 4204`).
- No hay comprobación de rol más allá de la preselección de filtros (`:300-306`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado y combos | Colecciones del entorno (`ColSegsTareasProyecto…`, `ColUsuariosCalidad`, `ColJefesProyecto`, `ColEstadosARTitulo`) | Memoria; carga en `Entorno`, fuera de esta ficha | `constructor.bas:2466-2530`, `Entorno.cls:1677, 1703` |
| Cargar la tarea seleccionada | `TbNCAccionesRealizadas` | SQL dinámico concatenado con `p_IDAR` | `constructor.bas:918-921` |
| Búsqueda de expedientes | Fuera de esta ficha | `constructor.getExpedientesBusqueda` | `:35` |
| Informe Excel | Fichero `.xlsx` en el directorio local de informes | Escritura en disco con automatización de Excel; sin escritura en base de datos | `FG.bas:4110-4123, 4341` |

## Coexistencia

La concatenación de `p_IDAR` en `getARProyecto` debe revisarse al pasar a PostgreSQL vinculado (tipo del id, comillas). Las comparaciones de responsable y estado son en memoria, sin dialecto. La exportación depende de Excel instalado en el puesto y de una carpeta local; en web requiere otro mecanismo de generación. `Format(…, "00")` y `Nz` son de VBA, no del motor.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Estado inicial «ACTIVA», con responsable o responsable de Calidad preseleccionado según el rol (`:300-309`).
- La exportación usa la colección filtrada guardada en la variable global, no la lista visible (`:129`, `:395`).

**Posibles defectos (a verificar)**

- `Form_Load` tiene la etiqueta `errores:` pero ningún `On Error GoTo errores`; el manejador y el cierre de `:314-323` no se ejecutan.
- `ComandoExportarAExcel_Click` compara `Estado` con «ACTIVAS» (`:119`), pero el título del estado es «ACTIVA» (`Entorno.cls:236`); en ese caso `m_Tipo` queda vacío y `TareasAExcel` usa «Activas» (`FG.bas:4105-4106`). «ESTADO IRREGULAR» (`:123`) ya no figura en el combinado.
- `TareasAExcel` compone el nombre del fichero con `Format(Month(Now()), "yyyy")` (`FG.bas:4110`), que da el año en lugar del mes.
- `ComandoVerDetalle_Click` mantiene la tarea ya seleccionada sin recargar (`:226`), igual que `ListaFiltrados_Click` solo recarga si cambia el id (`:455`).
- `ComandoIrABusquedaExpedientes_Click` declara un `m_FormExpedientes` local que oculta al `WithEvents` del módulo (`:9, 157, 170`).
- `ComandoActualizararExpedientes` carece de handler (`.form.txt:1916-1928`).
- `Filtrar` sale sin limpiar la selección cuando la colección es `Nothing` (`:401-403`) y hace `lst.Selected(1) = True` sin comprobar que hay filas (`:418`).
- `ListaFiltrados_Click` vuelve a poner el cursor de espera al salir con `m_ARProyectoSeleccionado Is Nothing` (`:462-466`).
- Cabecera de la lista distinta entre diseño («Tarea; Responsable; …; NºAR», `.form.txt:804`) y código (`:294`).
- Los mensajes de error de `ComandoVerDetalle_Click` y `Form_Load` nombran «ComandoEditar_Click» y «Form_Open» (`:247, 317`).

**Propuestas web pendientes de decisión**

- Sustituir las listas de valores por una tabla paginada con filtros en servidor.
- Generar el informe en servidor y ofrecerlo como descarga, en lugar de automatizar Excel en el cliente.
- Fijar el filtro inicial por rol de forma explícita.

## Criterios de aceptación

1. Al abrir, la lista muestra las tareas activas, filtradas por el usuario si es técnico o de Calidad (`:300-309`).
2. Cambiar responsable de Calidad, responsable o estado refiltra y actualiza el contador (`:255-281, 544-598, 417`).
3. Los botones «limpiar» vacían su filtro y refiltran (`:193-217`).
4. Buscar con una única coincidencia asigna el expediente y filtra; con varias, abre la búsqueda (`:59-93`).
5. «Detalle» (o doble clic) abre `FormNCProyecto` con la NC de la tarea; sin selección avisa (`:226-241, 485-490`).
6. «Exp. Excel» solo está habilitado con resultados y genera un libro con las 9 columnas indicadas (`:404`; `FG.bas:4134-4142`).
7. Cada error muestra el mensaje y avisa al administrador salvo error 1000 (`:100-106`).

## Pendientes

- Carga de `ColSegsTareasProyecto…` en `Entorno` (`Entorno.cls:1677, 1703`), a detallar en la ficha de la clase.
- Construcción de `ARProyecto` en `getARProyecto` tras la consulta (`constructor.bas:900-947`): fuera de esta ficha.
- Valor real de `BoundColumn` de `Estado` en el binario, que determina si `Me.Estado` devuelve título o id (`:119-125`).
