# Clase del flujo de trabajo

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/WorkflowServicio.cls` (4160 líneas, 82 procedimientos) · **PRD:** [README](../README.md)

La clase más grande de Condor y el corazón del producto: decide qué transiciones existen, si un usuario puede hacerlas, qué precondiciones debe cumplir la solicitud, ejecuta la transición con su transacción, genera el HTML del visor de estado, la ficha resumen y el gráfico de tiempos, y administra el ciclo de borradores que va al RAC. Todo lo demás orbita alrededor de este objeto.

Abreviaturas: `WF`, `classes/WorkflowServicio.cls`; `SOL`, `classes/Solicitud.cls`; `SS`, `classes/SolicitudServicio.cls`; `LS`, `classes/LogServicio.cls`; `WR`, `modules/WorkflowRepositorio.bas`; `LER`, `modules/LogEstadoRepositorio.bas`; `ENUM`, `modules/modEnumeradores.bas`; `WVC`, `classes/WebVisorCacheServicio.cls`; `CH`, `modules/ChecklistHelper.bas`; `AS`, `classes/AdjuntosServicio.cls`; `NS`, `classes/NotificacionServicio.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `WorkflowServicio` | clase de servicio | treinta y tres procedimientos públicos (`WF:25-4160`) | Transiciones, precondiciones, permisos por bloque, visor de estado, borradores y documentos |

## Procedimientos y llamadas

| Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|
| `getTransicionesValidas` | Devuelve las transiciones posibles desde el estado actual para el rol | `WorkflowRepositorio.getPosiblesDestinos`, `UsuarioPuedeTransicionarDesde` | Lectura | `WF:820-873` |
| `getProximaTransicionValida` | Devuelve la siguiente transición aplicable | `getTransicionesValidas`, `getProximaTransicionTeorica` | Lectura | `WF:2218-2282`, `WF:2718-2779` |
| `EsTransicionPermitida`, `UsuarioPuedeTransicionarDesde` | Comprueban la transición contra el catálogo y el rol | `WR.getPosiblesDestinos` | Lectura | `WF:1672-1700`, `WF:1701-1733` |
| `PermisoSuficiente` | Traduce el rol requerido por la transición en permiso, con jerarquía | — | Sin escritura | `WF:1410-1418`; administrador en `WF:1413`, calidad en `WF:1414`, ingeniería en `WF:1415`, resto permitido en `WF:1416` |
| `PrecondicionesCumplidas` y sus cuatro variantes por tipo | Evalúan las precondiciones de la transición para cada tipo de solicitud | `CumplePasoA*`, servicios de datos | Lectura | `WF:874-980`; variante CD_CA en `WF:895`, sub en `WF:926`, PC en `WF:954` |
| `CumplePasoARegistro`, `CumplePasoAModificacion`, `CumplePasoAValidacion`, `CumplePasoARevision`, `CumplePasoAFormalizacion`, `CumplePasoAAprobada` | Un procedimiento por destino de transición | servicios de datos y de flujo | Lectura | `WF:981-1210`; validación en `WF:1072`, revisión en `WF:1123`, formalización en `WF:1147` |
| `GetPrecondicionesPendientes` | Devuelve la colección de precondiciones que faltan | `PrecondicionesCumplidas*` | Lectura | `WF:2645-2717` |
| `EjecutarTransicion` | Aplica la transición: valida, guarda la solicitud, escribe el log de estado y genera los adjuntos que correspondan | `PrecondicionesCumplidas`, `PermisoSuficiente`, `LogEstadoRepositorio`, servicios de datos | `UPDATE` de estado más escritura en varias tablas, con transacción propia si el llamador no pasa la suya | `WF:1211-1409`; transacción local declarada en `WF:1221` |
| `ReabrirSolicitudCerrada` | Reabre una solicitud formalizada | servicios de flujo y de estado | `UPDATE` de estado | `WF:1793-1843` |
| `RevertirFlujo`, `RevertirAFaseAnterior` | Devuelven la solicitud a una fase anterior con motivo | servicios de flujo y de estado | `UPDATE` de estado | `WF:2283-2365`, `WF:2366-2644` |
| `PuedeEditarBloque` | Indica si el rol y el estado permiten editar un bloque | `PermiteEdicion_CDCA`, `PermiteEdicion_CDCASUB`, `PermiteEdicion_PC` | Lectura | `WF:1419-1446`; variantes en `WF:1478`, `WF:1557`, `WF:1599` |
| `EsAprobacionSuministradorCompleta` | Indica si el bloque de aprobación está completo | servicios de datos | Lectura | `WF:1447-1477` |
| `getEtapasDocumentalesAlcanzadas` | Devuelve qué etapas documentales admite el estado actual | catálogos | Lectura | `WF:1734-1792` |
| `getPaginaActivaGenerica`, `getPaginaActivaPC`, `getPaginaActivaCDCA`, `getPaginaActivaCDCASUB` | Devuelven la pestaña que debe mostrarse para el estado | catálogos | Sin escritura | `WF:1844-1903` |
| `GenerarHTML_VisualizadorDeEstado` | Genera el HTML del visor de estado y lo cachea | `ConstruirListadoRechazos`, `ConstruirListadoRAC`, `ConstruirResumenCambiosTecnico`, `WVC` | Lectura; escritura de caché | `WF:95-539`; listados en `WF:684-819` |
| `GenerarHTML_FichaResumen`, `GenerarHTML_GraficoTiempos` | Generan la ficha resumen y el gráfico de tiempos | `GetTiemposPorEstado`, `FormatearDuracion` | Lectura | `WF:2780-2865`, `WF:2887-3035` |
| `GenerarYAdjuntarDocumentoBorrador`, `GenerarYAdjuntarPDFfinal`, `ValidarYExportarDocumentoParaRAC` | Generan documentos y los adjunta a la solicitud | `DocumentoServicio`, `AS` | Escritura de fichero y de adjunto | `WF:73-94`, `WF:3180-3207`, `WF:3105-3179` |
| `GetProximaVersionBorrador`, `RegistrarBorradorRAC`, `TieneAdjuntoEnCicloActual`, `ObtenerHayCambiosValidacion`, `ExportarBorradorSingleton` | Administran el ciclo de borradores y su versionado | servicios de adjuntos y de validación | Lectura y escritura de adjuntos | `WF:1904-1923`, `WF:1984-2000`, `WF:2001-2025`, `WF:1924-1983`, `WF:2026-2217` |
| `InvalidarCacheTimeline`, `LimpiarCacheTimeline` | Descartan el HTML cacheado del visor | `WVC` | Borra caché | `WF:25-39` |

## Flujos de trabajo

- **Normal — transición:** la interfaz pide las transiciones válidas (`WF:820`), el usuario elige una y `EjecutarTransicion` comprueba el permiso (`WF:1410-1418`), evalúa las precondiciones del tipo de solicitud (`WF:874-980`), guarda la solicitud, escribe el log de estado y genera los adjuntos derivados (`WF:1211-1409`). Si el llamador no pasó conexión, la transacción es propia (`WF:1221`).
- **Normal — visor:** el visor de estado se genera desde el servicio y se cachea por solicitud (`WF:95-539`); la caché se invalida cuando cambia algo (`WF:25-39`).
- **Alternativos — vuelta atrás:** `ReabrirSolicitudCerrada` (`WF:1793`) y `RevertirAFaseAnterior` (`WF:2366`) recorren el flujo en sentido contrario, con motivo, y ambos escriben estado.
- **Errores:** las precondiciones devuelven el motivo por parámetro de salida (`WF:874`), de modo que la interfaz puede mostrarlo sin que sea un error.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| El rol requerido por una transición se comprueba con jerarquía y, si no se indica, se permite | `PermisoSuficiente` | `WF:1410-1418` |
| Cada tipo de solicitud tiene su propio conjunto de precondiciones | cuatro variantes | `WF:895`, `WF:926`, `WF:954` |
| Una transición exige que todas sus precondiciones se cumplan | `EjecutarTransicion` | `WF:1211-1409` |
| El editor de un bloque depende del rol y del estado | `PuedeEditarBloque` | `WF:1419-1671` |
| El HTML del visor se cachea por solicitud y se invalida al cambiar | caché del timeline | `WF:25-39`, `WF:40-72` |
| Los documentos del ciclo se versionan con un hash de datos | ciclo de borradores | `WF:1904-2217` |

## Lecturas y escrituras

| Operación | Tabla o recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Transición de estado | `tbSolicitudes`, `tbLogEstados`, bloques de datos | varias | Propia si el llamador no pasa la suya | `WF:1211-1409` |
| Catálogo de transiciones | `tbTransiciones`, `tbEstados` | lectura | Ninguna | `WF:820-873`, `WF:1701-1733` |
| Visor, ficha y gráfico | solicitudes, rechazos, RAC, logs, tiempos | lectura | Ninguna | `WF:95-539`, `WF:2780-3035` |
| Caché del visor | ficheros de caché | escritura en disco | — | `WF:25-72` |
| Documentos y adjuntos | documentos y `tbAdjuntos` | escritura de fichero y de fila | La del ciclo | `WF:73-94`, `WF:3105-3207` |

## Coexistencia

Esta clase decide la arquitectura del destino, no solo su código. `EjecutarTransicion` (`WF:1211-1409`) es la operación que en la versión web debe ser una transacción del servidor con una tabla de transiciones como fuente de verdad, y las precondiciones (`WF:874-1210`) son la especificación que hay que preservar literalmente: son la diferencia observable entre una versión migrada correcta y una que deja pasar transiciones que antes no se permitían. La caché del visor (`WF:25-72`) depende de ficheros locales y desaparece: en el destino, el HTML o su equivalente se genera bajo demanda o se sirve como recurso. Y el versionado de borradores (`WF:1904-2217`) implica un hash por versión que debe seguir siendo comparable con los ya almacenados, igual que en el bloque de hash.

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Un único servicio reúne transiciones, permisos, precondiciones, generación de HTML, gestión de documentos y ciclo de borradores, con transacción propia cuando nadie le pasa conexión.

**Posibles defectos, a verificar.**
1. Una sola clase concentra reglas de dominio, generación de HTML y operaciones de fichero (`WF:18-4160`): cualquier cambio obliga a tocar el mismo fichero.
2. `EjecutarTransicion` abre su propia transacción si no recibe conexión (`WF:1221`).
3. Las precondiciones están replicadas por tipo de solicitud (`WF:895`, `WF:926`, `WF:954`) y pueden divergir.
4. El visor genera HTML con estilos incrustados (`WF:540-638`, `WF:2866-2886`, `WF:3016-3035`): el contenido visual está en el código.
5. `PermisoSuficiente` permite cualquier transición sin rol requerido (`WF:1416`).

**Propuestas web pendientes de decisión.** Que la tabla de transiciones y las precondiciones sean la fuente de verdad del servidor y que el cliente solo las consulte. Que la caché del visor desaparezca o se sustituya por un recurso del servidor. Que el HTML y sus estilos se generen en el destino a partir de una plantilla versionable. Que el versionado de borradores conserve el algoritmo de hash.

## Criterios de aceptación

1. Solo las transiciones permitidas por el catálogo y el rol se ofrecen y se ejecutan. Trazado a `WF:820-873`, `WF:1410-1418`.
2. Una transición no se ejecuta si alguna precondición del tipo de solicitud falla. Trazado a `WF:874-1210`.
3. El estado de la solicitud y su historial se escriben en la misma unidad. Trazado a `WF:1211-1409`.
4. Revertir a una fase anterior exige motivo y queda registrado. Trazado a `WF:2283-2644`.
5. El ciclo de borradores mantiene su versión y su hash. Trazado a `WF:1904-2217`.

## Pendientes

- Confirmar contra el binario el catálogo completo de transiciones de `tbTransiciones`, porque es la especificación del flujo del destino (`WF:1701-1733`).
- Confirmar qué precondiciones devuelve `GetPrecondicionesPendientes` en cada transición, para que los mensajes de la interfaz se conserven (`WF:2645-2717`).
- Confirmar si el visor de estado, la ficha resumen y el gráfico se muestran en la versión web o si se rediseñan (`WF:95-539`, `WF:2780-3035`).
- Confirmar si `ExportarBorradorSingleton` sustituye a un servicio de exportación que exista fuera de Condor (`WF:2026-2217`).