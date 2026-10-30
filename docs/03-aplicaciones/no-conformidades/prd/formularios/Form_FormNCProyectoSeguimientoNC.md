# Form_FormNCProyectoSeguimientoNC

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoSeguimientoNC.cls` (550 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Rev. Estados» del seguimiento de proyecto: lista las NC de proyecto pendientes de revisión, filtrables por responsable de Calidad, responsable (técnico), estado y expediente, y abre la NC elegida en `FormNCProyecto`. Se aloja en `FrmDetalle` de [Form_FormNCProyectoSeguimiento](Form_FormNCProyectoSeguimiento.md). Es un formulario **no enlazado** (sin `RecordSource`); la lista no sale de una consulta sino de las colecciones `ColSegsNCProyecto…` del entorno. Las citas son `Form_FormNCProyectoSeguimientoNC.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCProyectoSeguimiento | Botón de navegación `NavNC` carga este formulario en `FrmDetalle` | `Form_FormNCProyectoSeguimiento.form.txt:1063` |
| `PintarIndicadores` | Llama a `Filtrar` si `FrmDetalle.SourceObject` es este formulario | `FG.bas:325-327` |

Sale hacia `FormNCProyecto`: fija `m_ObjNCProyectoActiva`, cierra la instancia abierta sin guardar y la abre con `DoCmd.OpenForm` (`:188-193`). Hacia `FormExpedientesBusqueda` para elegir expediente (`:84-87, 117-120`). No se cierra a sí mismo salvo error en `Form_Load` (`:278`).

## Controles y eventos

`.form.txt` declara 32 bloques `Begin`: el formulario (`.form.txt:4`), 9 de valores por defecto (`:713-772`), la sección `Detalle` (`:778`) y los 21 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista (valores, 5 columnas) | `OnClick`, `OnDblClick` | Muestra IDNoconformidad, Código, Descripción, Exp., Estado; el clic fija la NC seleccionada y habilita «Detalle»; el doble clic abre el detalle si está habilitado | `:383-443`, `.form.txt:785-808` |
| `lblListaDeResultados` | etiqueta | — | Contador «Registros filtrados (n)» | `:345, 370`, `.form.txt:836-847` |
| `ComandoVerDetalle` («Detalle») | botón (deshabilitado de diseño) | `OnClick` | Abre `FormNCProyecto` con la NC seleccionada | `:171-205`, `.form.txt:975-988` |
| `RESPONSABLECALIDAD` + `Etiqueta57` («Responsable calidad») | combinado (valores, 2 columnas) + etiqueta | `AfterUpdate` | Filtra por responsable de Calidad | `:523-549`, `.form.txt:857-899` |
| `Estado` + `Etiqueta261` («Estado tarea») | combinado (valores, 2 columnas) + etiqueta | `AfterUpdate` | Filtra por estado (5 valores) | `:207-233, 310-328`, `.form.txt:916-958` |
| `Responsable` + `Etiqueta263` («Responsable NC») | combinado (valores, 2 columnas) + etiqueta | `AfterUpdate` | Filtra por técnico (jefes de proyecto) | `:495-521`, `.form.txt:1391-1401, 1594-1642` |
| `ComandoLimpiarResposableCalidad` | botón | `OnClick` | Pone `RESPONSABLECALIDAD` a Null y filtra | `:165-169`, `.form.txt:1024-1036` |
| `ComandoLimpiarEstado` | botón | `OnClick` | Pone `Estado` a Null y filtra | `:145-149`, `.form.txt:1207-1219` |
| `ComandoLimpiarResposable` | botón | `OnClick` | Pone `Responsable` a Null y filtra | `:159-163`, `.form.txt:1411-1423` |
| `PalabraClave` | texto | `OnKeyDown` | Con Intro o Tab lanza la búsqueda de expediente | `:481-493`, `.form.txt:1660-1677` |
| `ComandoBuscarExpediente` | botón | `OnClick` | Busca expediente por palabra clave (ver Flujos) | `:14-106`, `.form.txt:1695-1707` |
| `ComandoIrABusquedaExpedientes` | botón | `OnClick` | Abre `FormExpedientesBusqueda` con la palabra clave | `:108-143`, `.form.txt:1779-1791` |
| `ComandoActualizararExpedientes` («Contador») | botón | sin handler | Solo se habilita o deshabilita desde el código | `:46-48`, `.form.txt:1844-1856` |
| `IDExpediente`, `Titulo` | texto bloqueado | — | Expediente elegido (id y nemotécnico, código o id) | `:62-70`, `.form.txt:1917-1918, 1952-1953` |
| `ComandoLimpiarIDExpediente` | botón | `OnClick` | Vacía `IDExpediente`, `PalabraClave` y `Titulo`; filtra | `:151-157`, `.form.txt:1986-1998` |
| `Etiqueta237`, `Etiqueta268` («Expedientes») | etiquetas | — | Decorativas, sin evento | `.form.txt:827, 2170-2180` |

Los 21 controles son: `ListaFiltrados`, `Etiqueta237`, `lblListaDeResultados`, `Etiqueta57`, `RESPONSABLECALIDAD`, `Etiqueta261`, `Estado`, `ComandoVerDetalle`, tres botones «limpiar» de filtros, `Etiqueta263`, `Responsable`, `PalabraClave`, `ComandoBuscarExpediente`, `ComandoIrABusquedaExpedientes`, `ComandoActualizararExpedientes`, `IDExpediente`, `Titulo`, `ComandoLimpiarIDExpediente`, `Etiqueta268`.

## Flujos

### Carga (`Form_Load`, `:235-279`)

1. Cursor de espera, `Ajustar Me`, cabecera de la lista (`:238-246`).
2. `EstablecerCombos` rellena responsables de Calidad (`FG.bas:1921`), los 5 estados (`:310-322`) y los jefes de proyecto (`FG.bas:1840`) (`:281-301`).
3. Si `EsTecnico` es `Sí`: `Responsable` = usuario conectado y `Estado` = estado «en ejecución fuera de plazo»; si no, y el usuario está en `ColUsuariosCalidad`, `RESPONSABLECALIDAD` = usuario, y `Estado` = «Registrada» (`:252-260`).
4. Ejecuta `Filtrar` (`:264`).

### Filtrado (`Filtrar`, `:330-381`)

Deshabilita «Detalle», pone cabecera y contador a 0, llama a `constructor.getNCsDeProyectoBusqueda` con `Column(0)` de responsable de Calidad, responsable y estado, e `IDExpediente` (`:349-351`); añade una fila por NC (`:363-365`), actualiza el contador, selecciona la fila 1 y dispara `ListaFiltrados_Click` (`:370-372`).

### Búsqueda de expediente (`ComandoBuscarExpediente_Click`, `:14-106`)

Sin palabra clave, no hace nada (`:28-33`). `getExpedientesBusqueda` devuelve un diccionario (`:34`): si es `Nothing`, vacía los campos y abre la ventana de búsqueda (`:41-56`); si hay un único expediente, lo asigna y filtra (`:58-82`); si hay varios, abre `FormExpedientesBusqueda`, copia la palabra clave y enlaza el evento `Seleccionado` (`:84-92`). `m_FormExpedientes_Seleccionado` rellena `IDExpediente` y `Titulo` (`:445-463`).

### Selección y detalle

`ListaFiltrados_Click` carga con `constructor.getNCProyecto` la NC de `Column(0)` si cambia (`:402-414`) y habilita «Detalle» (`:421`). «Detalle» reutiliza la NC seleccionada o la carga; sin selección avisa «Seleccione un elemento de la lista» (`:178-186`).

### Errores

Error 1000 ⇒ advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (p. ej. `:99-105, 196-204, 269-278`).

## Reglas, validaciones y cálculos

- Estados del combinado: CERRADAPTECE, CERRADAPTECECADUCADA, CERRADACENOCONFORME, ACSSINTAREAS y REGISTRADA (`:318-322`).
- `getNCsDeProyectoBusqueda` elige la colección por título de estado y, sin estado reconocido, usa `ColSegsNCProyectoTotales`; filtra por `NombreCalidad`, `Tecnico` e `IDExpediente` (`constructor.bas:2531-2587`).
- Las descripciones sustituyen `;` por `:` porque la lista es de valores (`:364`).
- No hay comprobación de rol más allá de la preselección de filtros (`:252-260`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado y combos | Colecciones del entorno (`ColSegsNCProyecto…`, `ColUsuariosCalidad`, `ColJefesProyecto`, `ColEstadosNCTitulo`) | Memoria; carga en `Entorno`, fuera de esta ficha | `constructor.bas:2531-2561`, `FG.bas:1840-1870, 1921-1945` |
| Cargar la NC seleccionada | `TbNoConformidades` | SQL dinámico concatenado con `p_IDNC`; antes, caché si `AplicarCache` | `constructor.bas:432-445`, `Variables Globales.bas:140` |
| Búsqueda de expedientes | Fuera de esta ficha | `constructor.getExpedientesBusqueda` | `:34` |
| Escrituras | — | Ninguna | `:1-550` |

## Coexistencia

La concatenación de `p_IDNC` en `getNCProyecto` debe revisarse al pasar a PostgreSQL vinculado (tipo del id, comillas). `SELECT *` más `SetPropiedad` por campo depende del orden y los tipos de `TbNoConformidades` (`constructor.bas:470-478`). Las comparaciones de responsable y estado se hacen por texto en memoria, sin dialecto. `Nz` es de VBA, no del motor.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Filtros combinados sobre una única colección elegida por estado (`constructor.bas:2546-2587`).
- Tras el filtrado se selecciona la fila 1 y se dispara el clic (`:371-372`).
- Los combinados se rellenan con `AddItem` y se leen por `Column(0)`.

**Posibles defectos (a verificar)**

- `Form_Load` tiene la etiqueta `errores:` pero ningún `On Error GoTo errores`; el manejador y el cierre de `:269-278` no se ejecutan.
- Para técnicos, `Form_Load` asigna a `Estado` el título del estado «en ejecución fuera de plazo» (`:254`), que no figura entre los 5 valores del combinado (`:318-322`) ni tiene colección propia en `getNCsDeProyectoBusqueda` (`constructor.bas:2546-2557`).
- `ComandoVerDetalle_Click` activa el cursor de espera y sale sin restaurarlo en el camino correcto (`:174, 195`).
- `ComandoIrABusquedaExpedientes_Click` declara un `m_FormExpedientes` local que oculta al `WithEvents` del módulo (`:8, 109, 122`); el evento `Seleccionado` solo queda enlazado desde `ComandoBuscarExpediente_Click` (`:89`).
- `m_FormExpedientes_Seleccionado` rellena los campos pero no llama a `Filtrar` (`:445-463`).
- `ComandoActualizararExpedientes` carece de handler (`.form.txt:1844-1856`).
- `Filtrar` sale sin limpiar `m_NCProyectoSeleccionado` cuando `getNCsDeProyectoBusqueda` devuelve `Nothing` (`:355-357`) y hace `lst.Selected(1) = True` sin comprobar que hay filas (`:371`).
- `ListaFiltrados_Click` vuelve a poner el cursor de espera al salir con `m_NCProyectoSeleccionado Is Nothing` (`:415-419`).
- La cabecera de la lista en diseño tiene columnas de tareas (`.form.txt:804`) distintas de las del código (`:246`).
- `NombreCalidad` está comentado como «usuariored» (`SegNCProyecto.cls:19`) pero se compara con el nombre del combinado (`constructor.bas:2565-2567`).
- Las etiquetas «Estado tarea» y «Responsable NC» no corresponden a lo que filtran (`.form.txt:926, 1401`).

**Propuestas web pendientes de decisión**

- Sustituir las listas de valores por una tabla paginada con filtros en servidor.
- Unificar la selección de expediente en un selector con búsqueda incremental.
- Fijar el filtro inicial por rol de forma explícita.

## Criterios de aceptación

1. Al abrir, un técnico ve su nombre en `Responsable`; un usuario de Calidad, en `RESPONSABLECALIDAD`, y la lista se filtra por el estado inicial (`:252-264`).
2. Cambiar responsable de Calidad, responsable o estado refiltra y actualiza el contador (`:207-233, 495-549, 370`).
3. Los botones «limpiar» vacían su filtro y refiltran (`:145-169`).
4. Buscar con una única coincidencia asigna el expediente y filtra; con varias, abre la búsqueda (`:58-92`).
5. «Detalle» (o doble clic) abre `FormNCProyecto` con la NC elegida; sin selección avisa (`:178-193, 438-443`).
6. Cada error muestra el mensaje y avisa al administrador salvo error 1000 (`:99-105`).

## Pendientes

- Carga de `ColSegsNCProyecto…` en `Entorno` (`Entorno.cls:1757`), a detallar en la ficha de la clase.
- Valor del estado «en ejecución fuera de plazo» en el combinado: confirmar en el binario si se muestra.
- `CacheNCProyecto.ObtenerNCConCache` (`constructor.bas:435`): no resuelta en esta ficha; `AplicarCache` queda en `False` en `Variables Globales.bas:286, 328`.
