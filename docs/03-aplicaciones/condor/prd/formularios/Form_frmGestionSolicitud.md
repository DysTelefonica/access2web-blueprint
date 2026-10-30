# Form_frmGestionSolicitud

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmGestionSolicitud.cls` (2416 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Orquestador de una solicitud: muestra en un control Edge (`webInfo`) la ficha resumen, el visor de cambio de estado o el gráfico de tiempos, todos generados como HTML por `WorkflowServicio`, y ejecuta las acciones de flujo que el HTML solicita (avanzar, rechazar, revertir, ciclo RAC, PDF final, cierre). Es **no enlazado**: el `.form.txt` no declara `RecordSource` y los datos llegan por `SolicitudServicio.getSolicitudViewModelPorID` (`:91`). Las citas son `:línea` de `Form_frmGestionSolicitud.cls` salvo indicación. Abreviaturas: `WS.cls` = `classes/WorkflowServicio.cls`, `SS.cls` = `classes/SolicitudServicio.cls`, `VRS.cls` = `classes/ValidacionRevisionServicio.cls`, `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `ft` = `Form_frmGestionSolicitud.form.txt`.

## Entrada y salida

| Desde | Modo | Evidencia |
|---|---|---|
| `Form_frmBuscarSolicitudes` (abrir solicitud) | `OpenArgs` = id; si ya está abierto con otro id lo cierra; se suscribe con `WithEvents` a `SolicitudEliminada` y `SolicitudModificada` | `Form_frmBuscarSolicitudes.cls:6, 301-310, 328, 343` |
| `Form_frm0Ppal` (tras crear una solicitud) | `OpenArgs` = id nuevo; después llama a `AbrirFormularioDeDatos` | `Form_frm0Ppal.cls:289-296`; `:2328-2333` |
| `Form_frmDatosPC`, `Form_frmDatosCDCA`, `Form_frmDatosCDCASUB` | Lo muestran si está abierto o lo abren con el id | `Form_frmDatosPC.cls:32-37`, `Form_frmDatosCDCA.cls:73-78`, `Form_frmDatosCDCASUB.cls:155-160` |

Abre `frmDatosPC`, `frmDatosCDCA` o `frmDatosCDCASUB` según `tipoSolicitud` (`:302-305, 335`), `frmGestionAdjuntos` en diálogo (`:371, 763`), `frmAltaSolicitud` en edición y diálogo (`:1787`) y `frmWebVisor` para capturas y visores (`:598, 605, 612, 671, 678, 790`). Publica `SolicitudModificada`, `SolicitudEliminada` y `SolicitudCambiada(tipo)` (`:7-9`).

Métodos públicos que invocan otros formularios: `CargarDatosSolicitud` (`:89`; `Form_frmDatosPC.cls:260`), `NotificarCambioExterno` (`:1943`; `Form_frmGestionAdjuntos.cls:215, 342`), `RecibirDatosRAC`, `RecibirDatosRechazoCalidad`, `RecibirImportacionWordRAC` (`:883-899`; `Form_frmWebVisor.cls:340, 346, 370`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Valida `OpenArgs`, carga la solicitud y activa el sondeo si procede | `:60-87`, `ft:223` |
| Formulario | — | `OnTimer` → `Form_Timer` | Lee la cola de comandos JS (modo Edge) o ejecuta la acción pendiente (modo heredado) | `:419-493`, `ft:222` |
| `webInfo` | Edge | `OnBeforeNavigation` → `webInfo_BeforeNavigate` | Intercepta URL `vba:<ACCION>:<param>`, cancela la navegación y arranca el timer a 50 ms | `:1550-1597`, `ft:1053, 1062` |
| `webInfo` | Edge | `OnDocumentComplete` → `webInfo_DocumentComplete` | En modo heredado inyecta el HTML sobre `about:blank` | `:2214-2230`, `ft:1063` |
| `cmdCambiarEstado` | botón | `Click` | Alterna ficha resumen ↔ visor de cambio de estado; el rótulo cambia | `:233-246, 142` |
| `cmdVerHistorial` | botón | `Click` | Alterna ficha resumen ↔ gráfico de tiempos | `:248-261, 143` |
| `cmdVerDatos` | botón | `Click` | Abre o reactiva el formulario de datos del tipo y se suscribe a `SolicitudHaCambiado` | `:295-365` |
| `cmdGestionarAdjuntos` | botón | `Click` | Abre `frmGestionAdjuntos` en diálogo | `:367-379` |
| `cmdEditarPreregistro` | botón | `Click` | Abre `frmAltaSolicitud` en edición y recarga | `:1784-1794` |
| `cmdEliminarSolicitud` | botón (pie) | `Click` | Borrado completo con confirmación y cierre del formulario | `:1796-1832`, `ft:1086-1088` |
| `cmdRefrescarWeb` | botón oculto | `Click` → `ForzarResetWebView` | Navega a `about:blank` y recarga | `:263-290`, `ft:1127-1140` |
| `cmdSalir` | botón | `Click` | Cierra el formulario | `:381-383` |
| `txtEstado` | cuadro de texto | — | Nombre del estado actual | `:1737` |
| `lblBloqueo` | etiqueta oculta | — | «Procesando documento...» durante las operaciones | `ft:633-645` |

Acciones que el HTML envía (`EjecutarAccionPendiente`, `:495-811`); el emisor está en `WS.cls` salvo indicación:

| Acción | Efecto | Evidencia |
|---|---|---|
| `AVANZAR_ESTADO:<id>` | Confirmación y `AccionAvanzarEstado`; si el destino es Rechazada, inyecta motivo automático; de Aprobada a Formalización, reversión con borrado | `:508-528`; emisor `WS.cls:311` |
| `APROBAR_TECNICA:<id>` | `AccionAprobarYAvanzar` | `:530-531` |
| `REVERTIR_TECNICA`, `_DESARROLLO`, `_VALIDACION`, `_FORMALIZACION`, `RECHAZO` | Confirmación «Reset de Fase» y `AccionRevertirTecnica` o `AccionRevertirGenerica` | `:533-564` |
| `RECHAZAR_TECNICA` | Abre la captura del motivo en `frmWebVisor` (modo `CAPTURA`) | `:785-791` |
| `VER_MOTIVO_CALIDAD`, `VER_MOTIVOS_RECHAZO` | Visor del motivo; `EDICION` solo Calidad o Administrador sin cambios tras el rechazo | `:566-599`; emisor `WS.cls:199` |
| `VER_CAMBIOS`, `VER_DELTA`, `VER_HISTORIAL_RAC` | Visores HTML en `frmWebVisor` | `:601-613`; `WS.cls:203, 229` |
| `ABRIR_ADJUNTO:<id>` | Abre el fichero bajo `URLDirectorioDocumentacion` | `:615-632` |
| `EXPORTAR_BORRADOR` | `ExportarBorradorSingleton`; vacío = sin cambios; si hay ruta, abre y marca la casilla del checklist por JS | `:634-661`; `WS.cls:2026` |
| `NOTIFICAR_ENVIO_RAC`, `COMENTARIOS_RAC`, `IMPORTAR_WORD_RAC_DESDE_BOTON`, `APROBAR_RAC` | Ciclo RAC (ver flujos) | `:663-697` |
| `GENERAR_PDF_FINAL`, `ADJUNTAR_PDF_FIRMADO` | PDF para firma y cierre | `:699-768` |
| `APROBAR_FINAL`, `GET_HAY_CAMBIOS` | Sin emisor identificado en `src` | `:770-771, 793-797` |

## Flujos

### Apertura

1. Muestra `lblBloqueo`, ajusta el formulario y crea `SolicitudServicio` (`:62-66`).
2. `OpenArgs` ≤ 0 lanza el error 513 «No se proporcion un ID de Solicitud válido.» (`:68-69`).
3. Carga el ViewModel; si es `Nothing`, error 513 «No se pudo cargar la información.» (`:91-92`). Devuelve `Nothing` si falta la solicitud o el expediente (`SS.cls:184-190`).
4. Título «Gestión de Solicitud: » más el código (`:93`); vista resumen (`:100-102`); habilita botones (`:106`).
5. Con `TempVars("UsarRenderizadoFichero")` = «Sí» activa el sondeo a 300 ms (`:73-75`, `:41-51`).
6. Cualquier error: `ErrorPresenter.MostrarError` y cierre del formulario (`:81-84`).

### Renderizado y puente HTML↔VBA

- **Modo fichero (Edge):** escribe el HTML en UTF-8 sin BOM en `%TEMP%\CONDOR_Vista_<id>_<marca>.html` y navega a `https://msaccess/<ruta>` (`:154-205`). El HTML escribe `window.colaComandos`; `Form_Timer` lo lee con un tope de 2 s, lo vacía, lo parte por `:` y ejecuta (`:391-417, 426-479`).
- **Modo heredado:** navega a `about:blank` e inyecta `innerHTML` en `DocumentComplete` (`:215-219, 2218-2221`); los enlaces `vba:` se capturan en `BeforeNavigate` y el timer dispara una sola vez (`:1556-1588, 484-489`).

### Avance, aprobación y rechazo

- **Avanzar:** de Desarrollo técnico (3) a Modificación (4) por Calidad con rechazo activo, pide confirmación y desactiva los rechazos (`:1609-1626`). Después `EjecutarTransicion`, refresco de tableros, aviso, eventos `SolicitudModificada` y `SolicitudCambiada("Avanzar")`, contadores y recarga manteniendo la vista (`:1630-1647`).
- **Aprobar técnica:** en una transacción guarda la decisión de revisión «APROBADO» y ejecuta la transición; actualiza antes el objeto en memoria (`:1985-2004`).
- **Rechazo técnico:** el motivo capturado vuelve por `RecibirDatosRechazoCalidad` (`:889-893`); en `CAPTURA` llama a `RechazarFaseTecnica`, que devuelve a Desarrollo técnico (`:1458-1463, 1669`; `WS.cls:3453-3528`). En `EDICION` solo actualiza el motivo activo (`:1458-1460`).
- **Rechazo rápido** (destino 9): escribe `racDecision = RECHAZADO` en los datos PC o CD_CA antes de transicionar (`:2167-2207`).

### Reversiones

`AccionRevertirGenerica` llama a `RevertirAFaseAnterior` (`:2288-2299`), que exige Administrador o Calidad (`WS.cls:2389-2390`) y purga según el estado de origen: datos técnicos, adjuntos de etapa, historial de rechazos y de cambios (3→2, `WS.cls:2413-2426`), dictamen RAC (`WS.cls:2432-2439, 2494-2506, 2525-2542`) y decisión final (8→7, `WS.cls:2469-2491`). `AccionRevertirTecnica` deja antes `revisionCalidadEstado = "PENDIENTE"` y lo guarda (`:2049-2059`).

### Ciclo RAC

1. **Notificar envío:** exige adjunto del ciclo actual (`:666-669`); la fecha capturada vuelve como `FECHA|<fecha>`, se valida, se registra el envío y se avanza a Revisión (6) (`:937-956`).
2. **Comentarios:** registra respuesta «COMENTARIOS» sobre el último ciclo pendiente y retrocede a Validación (5) (`:957-971`).
3. **Importar Word:** selector de `.docx/.doc`; actualiza `observacionesRAC`, `racCodigo` y `racNombre` del tipo, con alta si `tbDatosCDCASUB` no tenía fila; registra la respuesta, restaura el hash previo de la revisión y retrocede a Validación. Hay dos copias casi iguales: desde `frmWebVisor` (`:1007-1222`) y desde el botón del HTML (`:1224-1433`).
4. **Aprobar RAC:** registra «APROBADO» y avanza a Formalización (7) (`:684-697`).

### PDF final y aprobación

- `GENERAR_PDF_FINAL` exige «Aprobación Suministrador» completa del tipo, pide confirmación y abre el PDF generado (`:709-744`).
- `ADJUNTAR_PDF_FIRMADO` exige el PDF de firma generado y la aprobación completa, abre `frmGestionAdjuntos` con `TempVars("AutoUploadEtapa") = "Cierre"` y, si existe adjunto de «Cierre» o «Documento Final Firmado», aprueba (`:752-768`).
- `AccionAprobarFinal`: en una transacción guarda `decisionFinal = "APROBADO"` del tipo y transiciona a Aprobada (8) (`:813-861`).

### Eliminación

Avisa si hay una NC vinculada al código (`:1805-1811`), confirma con «No» por defecto (`:1813`) y llama a `EliminarSolicitudCompleta` (`:1818`): exige Administrador o Calidad (`SS.cls:475`), notifica por correo antes de borrar (`SS.cls:509`), borra en transacción datos, adjuntos, logs, ciclos, rechazos y la solicitud, borra los ficheros y confirma (`SS.cls:533-562`), y limpia la NC fuera de la transacción (`SS.cls:579`). Después lanza `SolicitudEliminada` y cierra (`:1823-1827`).

### Errores

La mayoría de handlers terminan en `ErrorPresenter.MostrarError` (`:259, 363, 809`). Las importaciones Word muestran «Error al importar...» (`:1219-1221, 1430-1432`). `PoblarControlesDesdeViewModel` reanuda tras cualquier error (`:1746-1748`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Editar pre-registro: estado < Validación, rol Calidad o Administrador y estado no final | `PoblarControlesDesdeViewModel` | `:1726-1727` |
| Eliminar: Administrador o Calidad, estado no final o Rechazada | Idem | `:1730-1731` |
| Transición permitida según `rolUsuario` global y precondiciones | `EjecutarTransicion` | `WS.cls:1244-1250`; `VG.bas:43, 117` |
| De Modificación a Validación desactiva rechazos; llegar a Validación crea ciclo | `EjecutarTransicion` | `WS.cls:1266-1276` |
| Estados 1 a 9 | Enumerado | `modules/modEnumeradores.bas:36-46` |
| Tipos con formulario de datos: `PC`, `CD_CA`, `CD_CA_SUB` | `cmdVerDatos_Click` | `:302-308` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Cargar solicitud y expediente | `tbSolicitudes`; expedientes en su backend | DAO por servicios | No | `SS.cls:184-190`; `modules/ExpedienteRepositorio.bas:40` |
| Transición de estado | `tbSolicitudes`, log de estados y de cambios, ciclo de validación | DAO | Propia o heredada | `WS.cls:1233-1301` |
| Aprobación técnica | Decisión de revisión + transición | DAO | Sí, workspace propio | `:1986-2004` |
| Aprobación final | `decisionFinal` del tipo + transición | DAO | Sí | `:834-861` |
| Desactivar rechazos (3→4) | Rechazos | DAO | No, fuera de la transición | `:1625` |
| Importar Word RAC | `tbDatosPC`, `tbDatosCDCA`, `tbDatosCDCASUB` | SQL dinámico `UPDATE` | No | `:1067-1120, 1280-1333` |
| Eliminar solicitud | Todas las dependientes y `tbSolicitudes`; NC externa | DAO | Sí; NC fuera | `SS.cls:533-579` |
| Vista HTML | `%TEMP%` | Fichero | — | `:158, 188` |

## Coexistencia

Detalle común en `../coexistencia.md`. Debe adaptarse:

- Las transacciones abren una segunda conexión con `ws.OpenDatabase(getdb().Name, …, "MS Access;PWD=…")` sobre el workspace `CondorWS` (`:835, 1987`; `FU.bas:56-62`); con tablas vinculadas a PostgreSQL esa conexión no abarca el backend remoto sin rediseño.
- El `UPDATE` dinámico con escapado de comillas y `RecordsAffected` (`:1062-1068, 1124`) debe pasar a sentencia parametrizada.
- La eliminación cruza dos backends sin transacción común (`SS.cls:562, 579`) y borra ficheros antes del `CommitTrans` (`SS.cls:555-562`).
- Las notificaciones de transición son posteriores al commit y silenciosas (`WS.cls:1311`).

Debe preservarse: la atomicidad de aprobación técnica y final, la purga por estado de origen y el orden «notificar, borrar» de la eliminación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Form_Timer` compara la marca con «Sí» literal (`:426`) y `EsRenderizadoFicheroActivo` acepta también «SI» (`:50`): con «SI», el sondeo se activa (`:74`) pero el timer entra en la rama heredada y se apaga (`:486`).
- `AccionAprobarFinal` revierte sin mostrar el error al usuario (`:869-880`), y no cierra la conexión abierta en `:835`.
- `IMPORTAR_WORD_RAC_DESDE_BOTON` aparece dos veces en el `Select Case`; la segunda rama es inalcanzable (`:681, 779`).
- `CAPTURA_ENVIO_RAC` sale con `Exit Sub` sin limpiar la acción pendiente (`:668`, frente a `:799-802`).
- En las importaciones Word, `db` no está asignada al calcular el hash previo (`:1043, 1256`); `comentariosIncompletos` se calcula y no se usa (`:1192`).
- `AccionCerrar` y `AccionReabrir` no tienen llamadas en `src` (`:1834, 1865`).
- La desactivación de rechazos en 3→4 queda fuera de la transacción de la transición (`:1624-1630`).
- El permiso de `EjecutarTransicion` usa `rolUsuario` (`WS.cls:1244`) y el formulario usa `m_ObjUsuarioActivo.rol` (`:576, 1612`); `frm0Ppal` suplanta sin cambiar `rolUsuario` (`Form_frm0Ppal.cls:199`).
- El rechazo rápido solo inyecta motivo para `PC`; `CD_CA` solo la decisión y `CD_CA_SUB` nada (`:2171-2207`).
- Cinco procedimientos `Debug_*` públicos accesibles desde la ventana Inmediato (`:907, 2089, 2109, 2145, 2236`).
- El título de diseño del `.form.txt` contiene un código real de solicitud (`ft:41`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el puente `colaComandos`/`vba:` por acciones de API con confirmación en la propia página; vistas resumen, estado y tiempos como rutas hijas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin id válido muestra un error y no deja la pantalla abierta (`:68-69, 81-84`).
2. Las tres vistas se alternan y el rótulo del botón refleja la vista activa (`:142-143`).
3. Editar pre-registro y Eliminar se habilitan según rol y estado (`:1726-1731`).
4. Cada transición exige rol permitido y precondiciones, registra log de estado y de cambio, y es atómica con la decisión asociada (`WS.cls:1244-1301`; `:1985-2004, 834-861`).
5. Revertir purga exactamente los datos y adjuntos de la fase de origen (`WS.cls:2405-2542`).
6. Notificar envío RAC exige borrador exportado y fecha válida (`:937-956`).
7. Eliminar avisa de la NC vinculada, borra todo lo dependiente y notifica (`:1805-1818`; `SS.cls:509-579`).
8. Tras cualquier acción, los buscadores y contadores de tareas abiertos se refrescan (`VG.bas:165-177`; `:2398-2412`).

## Pendientes

- **Variante `Form_frmGestionSolicitud.form-admCasa.txt` (D-CON-3):** diseño idéntico al `.form.txt` hasta la línea 1234, salvo bloques binarios; incrusta tras `CodeBehindForm` (`form-admCasa.txt:1235`) una versión anterior del código (136 bloques de diferencia, 113 `Debug.Print`). Contiene una contraseña literal de base de datos (`form-admCasa.txt:2113, 3320`) que esta ficha no reproduce. El `.form.txt` ordinario incrusta el mismo código que el `.cls`. Queda por confirmar cuál se importa al binario.
- `src/forms/frmGestionSolicitud.form.txt` (3112 líneas, sin prefijo `Form_`) no figura en el inventario; también contiene una contraseña literal (`frmGestionSolicitud.form.txt:2022, 2679`).
- Emisor de `APROBAR_FINAL` y `GET_HAY_CAMBIOS` (`:770, 793`).
- Si `RecibirImportacionWordRAC_Boton` (`:901`) tiene llamadores fuera de `src`.
