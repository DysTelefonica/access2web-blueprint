# Form_frmBuscarSolicitudes

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmBuscarSolicitudes.cls` (359 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Buscador de solicitudes: filtra por palabra clave, «solo mis solicitudes», «solo tareas pendientes» y filtros avanzados, y abre la solicitud elegida en `frmGestionSolicitud`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista se rellena como `Value List` (`:30-35`). Las citas son `Form_frmBuscarSolicitudes.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas`, `SolSrv.cls` abrevia `classes/SolicitudServicio.cls` y `Ppal`/`PpalTec`, `Form_frm0Ppal.cls`/`Form_frm0PpalTecnico.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú principal (botón buscar) | `DoCmd.OpenForm "frmBuscarSolicitudes"` | `Ppal:55` |
| Menú principal y menú técnico (tareas pendientes) | `OpenArgs:="PENDIENTES"` | `Ppal:91`, `PpalTec:19` |
| Menú técnico (buscar) | Abre, inyecta `FiltroInicial` con `idResponsableTecnico` del usuario activo y relanza `RealizarBusqueda` | `PpalTec:86-99` |

Abre `frmGestionSolicitud` con el `idSolicitud` como `OpenArgs` (`:309`) y escucha sus eventos `SolicitudEliminada` y `SolicitudModificada` (`:6, 310, 328-356`; declarados en `Form_frmGestionSolicitud.cls:7-8`). Abre `frmFiltrosAvanzadosSolicitudes` en modo diálogo (`:252`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `txtPalabraClave` | texto | `AfterUpdate` | Relanza la búsqueda | `:198-200`, `.form.txt:947-948` |
| `cmdLimpiarPalabraClave` | botón | `OnClick` | Vacía la palabra clave y busca | `:15-18`, `.form.txt:1002-1004` |
| `chkSoloMisSolicitudes` | casilla | `OnClick` | Busca; para el técnico queda marcada y deshabilitada | `:202-204, 39-43`, `.form.txt:747-748` |
| `chkSoloPendientes` | casilla | `OnClick` | Si se marca, pone `idEstado = "0"` en los filtros y busca | `:127-142`, `.form.txt:779-780` |
| `cmdFiltrosAvanzados` | botón | `OnClick` | Abre los filtros avanzados; oculto para el técnico | `:245-273, 40`, `.form.txt:800-802` |
| `cmdQuitarFiltros` («X») | botón oculto | `OnClick` | Limpia filtros simples y avanzados y busca | `:206-243`, `.form.txt:888-901` |
| `cmdBuscar` | botón | `OnClick` | Busca | `:192-194`, `.form.txt:849-851` |
| `lstResultados` | lista, 9 columnas | `OnDblClick` | Abre la solicitud seleccionada | `:275-277, 30-35`, `.form.txt:658-682` |
| `cmdVerDetalleSolicitud` («Ver Solicitud») | botón | `OnClick` | Igual que el doble clic | `:279-281`, `.form.txt:1272-1274` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:283-285`, `.form.txt:1195-1197` |
| `lblBloqueo` | etiqueta oculta | — | «Buscando, por favor espere...» durante las operaciones | `.form.txt:632-646` |

## Flujos

**Normal (apertura y búsqueda):**

1. `Form_Load` muestra `lblBloqueo`, ajusta el formulario y configura la lista: 9 columnas con `idSolicitud`, `idExpediente` y `Cód. Exp` a 0 cm (`:23-35`).
2. Crea un `FiltrosSolicitud` vacío y aplica el rol: técnico ⇒ oculta filtros avanzados y «X», fuerza «solo mis solicitudes»; resto ⇒ filtros avanzados visibles (`:37-49`; enumerado `rol` en `modules/modEnumeradores.bas:11-15`, global `rolUsuario` en `modules/Variables Globales.bas:43`).
3. Con `OpenArgs = "PENDIENTES"` marca «solo pendientes» (`:52-54`) y llama a `RealizarBusqueda` (`:55`).
4. `RealizarBusqueda` lee los tres filtros simples y llama a `SolicitudServicio.getSolicitudesViewModel` (`:79-82`; `SolSrv.cls:67-172`).
5. Compone la cadena de la lista: cabecera con «Estado», «Tipo», «Cód. Solicitud», «Nemotécnico», «Cód. Exp», «Resp. Técnico», «Resp. Calidad» y una fila por solicitud; responsables nulos ⇒ «N/A» (`:84-105`).
6. `ActualizarEstadoFiltrosVisual` pinta el botón de filtros avanzados en amarillo si hay filtros avanzados y en azul si no; muestra «X» si hay algún filtro y el rol no es técnico (`:144-188`; `classes/FiltrosSolicitud.cls:24-34`).

**Alternativos:**

- **Filtros avanzados:** abre el diálogo; si al volver su `Tag` es `"OK"`, toma su objeto `filtros` y busca; después cierra el diálogo, oculto pero abierto (`:252-265`).
- **Quitar filtros:** vacía la palabra clave, desmarca «solo mis» si está habilitada y «solo pendientes», y pone a cero los cinco campos del filtro (`:210-229`).
- **Abrir solicitud:** sin selección ⇒ «Por favor, seleccione una solicitud de la lista.» (`:292-296`). Si `frmGestionSolicitud` ya está abierto con otra solicitud, lo cierra con `acSaveNo`; si es la misma, solo le da el foco (`:301-308`).
- **Refresco reactivo:** cualquier `SolicitudModificada` o `SolicitudEliminada` del formulario de gestión relanza la búsqueda (`:328-356`).

**Errores:** todos los handlers con control de errores delegan en `ManejadorErroresFormulario` (`:60, 118, 186, 241, 271, 318`), que oculta `lblBloqueo`, muestra el mensaje según el número (513 validación, 514 operativo, resto técnico) y registra con `LogErrorUI` (`FU.bas:588-677`). Los errores del servicio llegan como `CondorError` (`SolSrv.cls:168-171`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Tipo visible: `PC`, `CD_CA`, `CD_CA_SUB` traducidos con `IIf`; otro código se muestra tal cual | consulta | `SolSrv.cls:94` |
| Responsable técnico mostrado: `FIRST` de los responsables con `EsJefeProyecto = 'SÍ'` | subconsulta | `SolSrv.cls:86-89` |
| Pendientes del técnico: estado interno `estadoDesarrolloTecnico` (3) | `WHERE` | `SolSrv.cls:110-112`, `modEnumeradores.bas:39` |
| Pendientes de calidad o administrador: estado distinto de 3 y de `estadoAprobada` (8), sin filtrar por usuario | `WHERE` | `SolSrv.cls:113-116`, `modEnumeradores.bas:44` |
| Con «solo pendientes», el filtro de estado avanzado no se aplica (`ElseIf`) | `WHERE` | `SolSrv.cls:109-122` |
| Solo mis solicitudes: técnico ⇒ cualquier fila de `TbExpedientesResponsables` del usuario; resto ⇒ `IDResponsableCalidad` del expediente | `WHERE` | `SolSrv.cls:123-131` |
| Palabra clave: `LIKE *texto*` sobre código de solicitud, nemotécnico y código de expediente | `WHERE` | `SolSrv.cls:133-136` |
| Con «solo mis solicitudes», los filtros avanzados de responsable técnico y de calidad se ignoran | `WHERE` | `SolSrv.cls:139, 144` |
| Suministrador: existe relación en `TbExpedientesSuministradores`, sin distinguir contratista principal | `WHERE` | `SolSrv.cls:149-152` |
| Orden: `fechaCreacion` descendente; la colección conserva el orden de inserción | consulta | `SolSrv.cls:162`, `modules/RepositorioComun.bas:82-93` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Búsqueda de solicitudes | `tbSolicitudes`, `tbEstados`, `TbExpedientes`, `TbUsuariosAplicaciones`, `TbExpedientesResponsables`, `TbExpedientesSuministradores` | SQL dinámico con parámetros `[p_*]` en `QueryDef` temporal, `dbOpenSnapshot` | no | `SolSrv.cls:86-164`, `RepositorioComun.bas:39-61` |

Solo lectura. La conexión es `getdb()`, el backend de Condor (`SolSrv.cls:77`, `FU.bas:84-108`); el estado se concatena en `IN (...)` sin parámetro (`SolSrv.cls:120`).

## Coexistencia

Pendiente del resumen de coexistencia (`../coexistencia.md`). Ya se ve:

- Una sola consulta cruza tablas de solicitudes y de expedientes y usuarios sobre la conexión de Condor (`SolSrv.cls:77, 103-107`), mientras otros servicios leen expedientes con `getdbExpedientes()` y usuarios con `getdbLanzadera()` (`modules/ExpedienteRepositorio.bas:40`, `classes/UsuarioServicio.cls:141`): se infiere que son tablas vinculadas en el backend de Condor (a verificar).
- Sintaxis propia de Jet/ACE: `FIRST()`, `IIf`, comodín `*` en `LIKE`, parámetros entre corchetes (`SolSrv.cls:86, 94, 134-135`).
- Literal `'SÍ'` en mayúsculas frente a `'Sí'` en otras consultas del mismo dato (`SolSrv.cls:88`; `ExpedienteRepositorio.bas:22`): Jet compara sin distinguir mayúsculas; PostgreSQL sí distingue.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La lista se compone como `Value List` concatenando con «;» (`:92-100`): un valor con «;» o comillas desplazaría columnas, y el límite de longitud de `RowSource` truncaría listas largas.
- El filtro de responsable técnico usa cualquier responsable (`SolSrv.cls:140`), pero la columna muestra solo el jefe de proyecto (`SolSrv.cls:88`) y el desplegable solo lista jefes de proyecto (`classes/UsuarioServicio.cls:116-119`).
- El filtro inyectado desde el menú técnico (`PpalTec:94-96`) no restringe la consulta, porque «solo mis solicitudes» está forzada y anula ese filtro (`:42`; `SolSrv.cls:139`).
- Al reutilizar un `frmGestionSolicitud` abierto desde otro sitio, `m_frmGestion` no se asigna y no llegan sus eventos (`:305-306, 310`).
- `ActualizarEstadoFiltrosVisual` hace `cmdBuscar.SetFocus` también durante `Form_Load` (`:178, 55`); en Access puede fallar si el formulario aún no es visible.
- El comodín de `LIKE` no escapa `*`, `?` ni `[` tecleados por el usuario (`SolSrv.cls:135`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: lista paginada con orden por columna; filtros reflejados en la URL; refresco tras modificar una solicitud sin depender de eventos de formulario.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, se listan las solicitudes por fecha de creación descendente con las siete columnas visibles del legacy, con «N/A» en responsables vacíos.
2. Un técnico ve solo solicitudes de expedientes donde es responsable, sin acceso a filtros avanzados.
3. «Solo pendientes» devuelve, por rol, el conjunto de estados de la tabla de reglas.
4. La palabra clave encuentra coincidencias parciales en código de solicitud, nemotécnico o código de expediente.
5. «Quitar filtros» restablece todos los filtros y la lista completa permitida al rol.
6. Abrir sin selección muestra el aviso; con selección abre la solicitud.
7. Tras modificar o eliminar una solicitud, la lista refleja el cambio.

## Pendientes

- Qué estados existen en `tbEstados` y cuáles son finales además de 8 (`modules/EstadoRepositorio.bas:18`).
- Si `TbExpedientes` y `TbUsuariosAplicaciones` son vínculos en el backend de Condor o copias locales.
- Contenido de `RellenarObjetoDesdeRecordset` y de las propiedades de `SolicitudBusquedaViewModel` (`classes/SolicitudBusquedaViewModel.cls`).
