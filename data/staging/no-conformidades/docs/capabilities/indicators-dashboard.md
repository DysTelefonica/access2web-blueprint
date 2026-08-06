# Capacidad: cuadro de mando de indicadores

## Â§0 Identidad
- **ID de capacidad**: `CAP-INDICATORS-DASHBOARD`
- **Nivel**: critical
- **Estado**: active / documentaciÃ³n alineada con v2; evidencia runtime reciente de indicadores recogida por slices
- **Fuente**: hybrid (documentos de funcionalidad Issue #18 + inventario de fuente + documentos adyacentes de cachÃ©)
- **Responsable / autoridad de producto**: ConfirmaciÃ³n pendiente â€” informes de Calidad / gestiÃ³n
- **Ãšltima verificaciÃ³n**: 2026-06-15, evidencia aportada de ejecuciones Dysflow por filtros/slices; esta actualizaciÃ³n documental no ejecutÃ³ Dysflow/Access
- **Confianza global**: mixed â€” mÃºltiples slices pasan como `Verified-runtime`; el manifest completo timeoutea; la reconstrucciÃ³n completa idempotente y la propagaciÃ³n de fallo post-escritura tienen focused PASS del 2026-06-15; los 3 contratos divergentes previos (Issue18_GlobalCache + Issue38 + Issue50) se resolvieron funcionalmente el 2026-06-15

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Proporcionar visibilidad de gestiÃ³n sobre trabajo de NC/acciones/tareas mediante buckets de indicadores, recuentos y filas de detalle en Proyecto y AuditorÃ­a.
- **Usuarios / personas**: Equipo de calidad, managers/supervisores, usuarios de Proyecto/AuditorÃ­a, soporte/desarrolladores.
- **Problema que resuelve**: Los usuarios necesitan visibilidad actual de carga/estado sin abrir manualmente cada NC.
- **Valor de negocio / por quÃ© existe**: Indicadores fiables impulsan seguimiento, priorizaciÃ³n y diagnÃ³stico de regresiones tras escrituras de negocio.
- **No objetivos**: No define cada ciclo de vida subyacente de acciones; eso vive en capacidades de dominio.
- **Fuente de intenciÃ³n**: Borrador de capacidad existente + documentos de funcionalidad Issue #18; nombres/umbrales de buckets pendientes de confirmaciÃ³n de producto/UAT.
- **Referencia tracker de origen**: Issue #18, Issue #39, Issue #67.

## Â§2 Contrato de comportamiento

### Escenarios (Given / When / Then)
- **GIVEN** un usuario abre el cuadro de mando **WHEN** la cachÃ© de indicadores estÃ¡ actualizada **THEN** los recuentos de buckets y filas de detalle se filtran por usuario/responsable y dominio.
- **GIVEN** cambian datos de NC/AC/AR/tarea **WHEN** se ejecuta sincronizaciÃ³n post-escritura **THEN** solo se refresca el alcance de `IDNoConformidad` afectado, no una reconstrucciÃ³n amplia.
- **GIVEN** falla la sincronizaciÃ³n post-escritura **WHEN** la escritura tuvo Ã©xito **THEN** el fallo es visible y no se afirma que la cachÃ© estÃ© actualizada.
- **GIVEN** filas de Proyecto y AuditorÃ­a comparten infraestructura **WHEN** se ejecutan lecturas del cuadro de mando **THEN** las filas nunca se fugan entre dominios.

### Reglas de negocio
| ID de regla | Enunciado (previsto) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba (evidencia) | Confianza |
|---|---|---|---|---|---|
| BR-IND-1 | La cachÃ© de indicadores es estado materializado compartido en backend, no estado de sesiÃ³n frontend. | Docs Issue #18 | SÃ­, con evidencia por slices | `Issue18_BackendCacheSchema` 2/2, `CacheIndicadoresMaterializado` 8/8, `CacheIndicadoresAuditoriaMaterializado` 3/3 tras retry | Verified-runtime |
| BR-IND-2 | La cachÃ© incluye filas de detalle necesarias por el cuadro de mando, no solo recuentos agregados. | Docs Issue #18 | SÃ­, con deuda de rendimiento | `Issue18_CargarDetalle` 2/2 (~122s en un test), `Issue18_DetalleCompleto` 1/1 | Verified-runtime |
| BR-IND-3 | Las lecturas runtime filtran por usuario conectado/responsable. | Docs Issue #18 | SÃ­ | `Issue18_CargarBucket` 2/2 (~120s en un test); `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic` 1/1 tras arreglar la query global del propio test para sumar recuentos por dominio (los recuentos per-domain evitan un quirk de cachÃ© DAO/Jet con `IN (1, 2)` en `COUNT(*)` sin mÃ¡s predicados); las aserciones por usuario ya pasaban | Verified-runtime (resuelto 2026-06-15) |
| BR-IND-4 | Las filas de Proyecto y AuditorÃ­a permanecen separadas por dominio. | Docs Issue #18 | SÃ­ | `CacheIndicadoresAuditoriaMaterializado` 3/3 y `Issue18_CargarDetalle` 2/2 pasan; `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic` 1/1 tras el mismo arreglo de query per-domain; el test afirma ahora 5 filas globales sumando Proyecto (1) + AuditorÃ­a (2) y mantiene `QA_User` 3 + `Otro_User` 2 | Verified-runtime (resuelto 2026-06-15) |
| BR-IND-5 | Los cambios correctos de NC/AC/AR/tarea sincronizan solo el alcance de la NC afectada. | Docs Issue #18 | SÃ­, con hooks lentos | `Issue18_SincronizarNC` 3/3, `Issue18_NCWriteHook` 1/1 (~120s), `Issue18_ACWriteHook` 1/1 (~114s), `Issue18_ARWriteHook` 4/4 con dos hooks ~117â€“123s | Verified-runtime |
| BR-IND-6 | Los cambios de AC resuelven ACâ†’NC; los de AR/tarea resuelven AR/tareaâ†’ACâ†’NC. | Docs Issue #18 | SÃ­ | `Issue18_ResolverNCDesde` 3/3, `Issue18_ACWriteHook` 1/1, `Issue18_ARWriteHook` 4/4 | Verified-runtime |
| BR-IND-7 | La sincronizaciÃ³n post-escritura fallida es visible e impide afirmaciones falsas de cachÃ© actual. | Docs Issue #18 | SÃ­, con evidencia focused | `Test_Issue18_NCWriteHook_InvalidarCache_FailedSync_ReturnsError_Atomic` PASS el 2026-06-15 (`4788` ms, `issue18_nc_write_hook_failed_sync_ok`): `NC 992099` inexistente, `InvalidarCache` devuelve `False` y propaga `pError` explÃ­cito | Verified-runtime focused |
| BR-IND-8 `#73` | Los nombres, umbrales y salidas de gestiÃ³n de buckets del cuadro de mando estÃ¡n aprobados por producto. | Autoridad de producto pendiente | Desconocido | FALTA â†’ crear mediante access-vba-tdd tras confirmar escenario UAT | Intended |

### Validaciones
- El esquema de cachÃ© tiene campos de dominio/detalle/responsable.
- Las lecturas filtran por usuario/responsable y dominio.
- La resoluciÃ³n de NC padre para escrituras AC/AR/tarea es obligatoria para refrescar el alcance afectado.
- Los nombres/umbrales de buckets estÃ¡n pendientes de confirmaciÃ³n.

### Transiciones de estado
- `Sin cachÃ© materializada` --(`ReconstrucciÃ³n completa`)--> `CachÃ© de indicadores construida`.
- `CachÃ© actual` --(`Escritura afectada`)--> `Alcance de NC afectada sincronizado`.
- `Escritura correcta + fallo de sincronizaciÃ³n` --(`PropagaciÃ³n de fallo`)--> `Fallo visible / cachÃ© no actual`.
- `Filas de cachÃ© compartida` --(`Lectura de cuadro de mando`)--> `Vista filtrada por usuario/dominio`.

### Caminos lÃ­mite y de error
- Los resultados vacÃ­os de bucket/detalle pueden ser vÃ¡lidos; no inventar fallback vivo sin regla.
- Una reconstrucciÃ³n completa amplia como ruta normal de escritura es un olor de diseÃ±o.
- La fuga de dominio entre Proyecto y AuditorÃ­a es crÃ­tica.

### SeÃ±ales de aceptaciÃ³n / presencia
- Los slices de indicadores pasan en staging actual con datos fixture controlados; no afirmar manifest completo verde.
- Las pruebas demuestran filtrado de dominio/responsable, completitud de detalle, sincronizaciÃ³n de alcance afectado y propagaciÃ³n explÃ­cita de fallo post-escritura donde los slices/focused runs estÃ¡n verdes; no afirmar manifest completo verde.
- Las definiciones de buckets aprobadas por producto estÃ¡n documentadas.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada UI**: `Form_FormIndicadores`; vistas bucket/detalle pendientes de mapeo exacto de controles; consumidores de seguimiento/listado de Proyecto/AuditorÃ­a.
- **Puntos de entrada de fuente**: `ModuloCacheIndicadoresIssue18`, `ModuloCacheIndicadores`, `IndicadorRepositorio`, `IndicadorServicio`, `Test_IndicadoresCaracterizacion`, `Test_IndicadoresTelemetry`.
- **Datos tocados**: tablas compartidas de cabecera/configuraciÃ³n/detalle de cachÃ© de indicadores, NC Proyecto/NC AuditorÃ­a, datos AC/AR/tarea, campos responsable/usuario/dominio, `TbConfiguracionBackends` para enrutamiento de entorno.
- **Salidas**: recuentos de buckets del cuadro de mando, filas de detalle, diagnÃ³sticos/logs, posibles informes de gestiÃ³n.
- **Dependencias e integraciones**: acciones/seguimiento de Proyecto, acciones/seguimiento de AuditorÃ­a, soporte transversal de cachÃ©/preparaciÃ³n.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada en esta tarea solo documental; no se importÃ³, compilÃ³ ni ejecutÃ³ Access/VBA durante esta actualizaciÃ³n.
- **EvaluaciÃ³n de diseÃ±o (as-built vs ideal)**: el modelo de lectura materializado en backend tiene evidencia runtime amplia por slices, pero la confianza sigue siendo mixta: tres contratos divergentes previos (Issue18_GlobalCache + Issue38 + Issue50) se resolvieron el 2026-06-15; la reconstrucciÃ³n completa idempotente y la propagaciÃ³n de fallo post-escritura tienen focused PASS; quedan la cautela de manifest completo no ejecutado/verde y deuda de rendimiento en hooks/lecturas lentas.

## Â§4 Receta de reconstrucciÃ³n
1. Confirmar nombres de buckets, umbrales, reglas de visibilidad y SLA de frescura.
2. Inspeccionar esquema y escribir pruebas fixture-first para filas de Proyecto y AuditorÃ­a, filtros de responsable, filas de detalle y resoluciÃ³n de padres.
3. Probar sincronizaciÃ³n de mutaciones mediante costuras helper/servicio; los formularios siguen siendo consumidores finos.
4. Mantener pruebas de fallo de sincronizaciÃ³n que afirmen fallo visible/no afirmaciÃ³n falsa de cachÃ© actual; evidencia focused actual: `Test_Issue18_NCWriteHook_InvalidarCache_FailedSync_ReturnsError_Atomic` PASS.
5. Cambios futuros: importaciÃ³n Dysflow â†’ compilaciÃ³n manual del usuario â†’ pruebas Dysflow.

## Â§5 Evidencia y trazabilidad
- **Pruebas**: evidencia aportada de ejecuciones Dysflow por filtros/slices. El manifest completo `tests/tests.vba.indicadores-caracterizacion.json` contiene 55 procedimientos y timeoutea como conjunto (`MCP error -32001: Request timed out`), por lo que no se debe declarar verde completo. Esta actualizaciÃ³n documental no ejecutÃ³ Dysflow/Access; los `Verified-runtime` de BR-IND-3/BR-IND-4 y de los contratos Issue #38/Issue #50 se basan en la misma evidencia runtime ya recogida y aportada por el usuario.

### Evidencia runtime reciente (Dysflow por slices)

| Ãmbito | Manifest / filtro | Resultado | Confianza | Nota |
|---|---|---:|---|---|
| Recuentos rÃ¡pidos | `tests/tests.vba.indicator-fast-counts.json` | 5/5 | Verified-runtime | Cobertura runtime de conteos rÃ¡pidos. |
| CachÃ© materializada | `tests/tests.vba.cache-materialized.json` | 13/13 | Verified-runtime | Cobertura relacionada de cachÃ© materializada. |
| Helper de gestiÃ³n/auditorÃ­a | `tests/tests.vba.audit-gestion-helper.json` | 11/11 | Verified-runtime | PasÃ³ tras el arreglo de selecciÃ³n de informe de auditorÃ­a. |
| CaracterizaciÃ³n indicadores | `tests/tests.vba.indicadores-caracterizacion.json` + filtro `Indicadores_` | 11/11 | Verified-runtime | Slice de indicadores. |
| CachÃ© proyecto | filtro `CacheIndicadoresMaterializado` | 8/8 | Verified-runtime | Slice de cachÃ© materializada. |
| CachÃ© auditorÃ­a | filtro `CacheIndicadoresAuditoriaMaterializado` | 3/3 | Verified-runtime | PasÃ³ tras retry; la cancelaciÃ³n anterior fue accidental. |
| Esquema backend Issue #18 | filtro `Issue18_BackendCacheSchema` | 2/2 | Verified-runtime | Evidencia de estructura de cachÃ©. |
| Fixtures Issue #18 | filtro `Fixture` | 2/2 | Verified-runtime | PreparaciÃ³n de datos fixture. |
| SincronizaciÃ³n NC | filtro `Issue18_SincronizarNC` | 3/3 | Verified-runtime | SincronizaciÃ³n por NC afectada. |
| ResoluciÃ³n AC/ARâ†’NC | filtro `Issue18_ResolverNCDesde` | 3/3 | Verified-runtime | ResoluciÃ³n de padres. |
| Buckets | filtro `Issue18_CargarBucket` | 2/2 | Verified-runtime | Un test tardÃ³ ~120s; comportamiento verificado con deuda de rendimiento/diagnÃ³stico. |
| Detalle | filtro `Issue18_CargarDetalle` | 2/2 | Verified-runtime | Un test tardÃ³ ~122s; comportamiento verificado con deuda de rendimiento/diagnÃ³stico. |
| Detalle completo | filtro `Issue18_DetalleCompleto` | 1/1 | Verified-runtime | Campos requeridos de detalle. |
| Hook NC | filtro `Issue18_NCWriteHook` | 1/1 | Verified-runtime | ~120s; hook verificado pero lento. |
| Hook NC â€” fallo de sincronizaciÃ³n explÃ­cito | `Test_Issue18_NCWriteHook_InvalidarCache_FailedSync_ReturnsError_Atomic` | 1/1 | Verified-runtime focused | PASS 2026-06-15, `4788` ms, valor `issue18_nc_write_hook_failed_sync_ok`; `NC 992099` inexistente fuerza fallo, `InvalidarCache` devuelve `False` y propaga `pError` explÃ­cito. |
| Hook AC | filtro `Issue18_ACWriteHook` | 1/1 | Verified-runtime | ~114s; hook verificado pero lento. |
| Hook AR | filtro `Issue18_ARWriteHook` | 4/4 | Verified-runtime | Dos hooks ~117â€“123s; verificado con deuda de rendimiento. |
| Tests de cachÃ© especÃ­ficos | `Cache_Proyecto_Delegacion_Y_Reset_Atomic`, `Cache_InvalidarTodo_SeparaProyectosYAuditorias_Atomic`, `Cache_InvalidacionSelectiva_Atomic`, `Cache_ConsistenciaConEntorno_Atomic` | 4/4 | Verified-runtime | RegresiÃ³n de cache/reset/invalidation. |
| Formulario | filtro `Formulario` | 2/2 | Verified-runtime | Cobertura de formulario relacionada. |
| Seguimiento auditorÃ­a | filtro `Issue38_SeguimientoAuditoria` | 1/1 | Verified-runtime | Contrato de auditorÃ­a. |
| Reset colecciÃ³n tareas | filtro `Issue38_ResetearColTareas` | 1/1 | Verified-runtime | Contrato de reset de tareas. |

### Evidencia focused reciente (2026-06-15)

| Procedimiento | Estado | Evidencia |
|---|---|---|
| `Test_Issue18_ReconstruirTodo_Idempotent_Atomic` | Verified-runtime focused PASS | DuraciÃ³n `4862` ms, valor `issue18_rebuild_idempotent_ok`; fixture-first con schema Proyecto/AuditorÃ­a inspeccionado, Proyecto `NC=992001` / `AC=992011` / `AR=992021` y AuditorÃ­a `Auditoria=992201` / `NC=992202` / `AC=992211` / `AR=992221`; primera y segunda reconstrucciÃ³n con `pError=''`, resultado `ok true`, `Proyecto=true`, `Auditoria=true`; cabeceras estables y teardown OK. |
| `Test_Issue18_NCWriteHook_InvalidarCache_FailedSync_ReturnsError_Atomic` | Verified-runtime focused PASS | DuraciÃ³n `4788` ms, valor `issue18_nc_write_hook_failed_sync_ok`; confirmado que no existe `NC 992099`; `InvalidarCache` devuelve `False` con `pError` `CacheNCProyecto.InvalidarCache no pudo sincronizar indicador de Proyecto para NC 992099: Cache_IndicadoresProyectoMaterializado_SincronizarNC: no se encontraron filas para NC 992099.`; asserts OK. |

### Divergencias resueltas (2026-06-15)

| Procedimiento | Estado previo | ResoluciÃ³n | Evidencia |
|---|---|---|---|
| `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic` | Divergent (recuento global) | Arreglo del propio test en `src/modules/Test_IndicadoresCaracterizacion.bas` (`Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic`): la query global ahora suma recuentos per-domain (`IDCacheIndicadorProyecto=1` + `IDCacheIndicadorProyecto=2`) para esquivar el quirk de cachÃ© DAO/Jet con `IN (1, 2)` en `COUNT(*)` sin mÃ¡s predicados. Aserciones por usuario ya pasaban. | 1/1 |
| `Test_Issue38_SeguimientoProyecto_ActualizarModoProyecto_Contract` | Divergent (no delegaba al helper) | Refactor de `src/forms/Form_FormNCProyectoSeguimiento.cls` para replicar el patrÃ³n de `Form_FormNCAuditoriaSeguimiento.cls`: `ComandoActualizar_Click` ya no llama a `PintarIndicadores` directamente; ahora fija `m_CargaInicialIndicadoresPendiente = True` y `Me.TimerInterval = 100` y el `.cls` referencia `NCProyectoSeguimientoHelper.CargarIndicadoresSeguimientoProyecto`. `OnTimer = "[Event Procedure]"` en el `.form.txt` (lÃ­nea 369) ya estaba enlazado. | 1/1 |
| `Test_Issue50_SeguimientoProyecto_CargaDiferidaHelper_Contract` | Divergent (faltan flags, guard, `Form_Timer`, programaciÃ³n, delegaciÃ³n, duraciÃ³n) | Mismo refactor de `src/forms/Form_FormNCProyectoSeguimiento.cls` aÃ±adiÃ³ los flags privados `m_CargaInicialIndicadoresPendiente`, `m_CargandoIndicadores`, `m_UltimaDuracionIndicadores`; `Form_Timer`; programaciÃ³n del timer en `Form_Load`; delegaciÃ³n al helper y llamada al helper usando `p_DuracionSegundos:=m_UltimaDuracionIndicadores` para casar con la firma `NCProyectoSeguimientoHelper.CargarIndicadoresSeguimientoProyecto`. | 1/1 |

### Caveats del runner/MCP

- `proceduresJson` como array shorthand fallÃ³ con `VBA_INVALID_TEST_PLAN: Test #1 must be an object`; workaround efectivo: `testsPath` + `filter`.
- El manifest completo `tests/tests.vba.indicadores-caracterizacion.json` timeoutea (`MCP error -32001: Request timed out`); usar slices/filtros.
- Las lecturas/hooks lentos estÃ¡n verificados funcionalmente, pero son deuda de rendimiento y riesgo diagnÃ³stico.
- La evidencia focused PASS de reconstrucciÃ³n completa y fallo post-escritura no equivale a manifest completo verde; el manifest completo `tests/tests.vba.indicadores-caracterizacion.json` sigue sin evidencia de ejecuciÃ³n completa satisfactoria.

| Elemento (funcionalidad o arreglo) | Ref. tracker | VersiÃ³n staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en prod | Nota |
|---|---|---|---|---|---|---|
| CachÃ© compartida de indicadores en backend | Issue #18 / Issue #67 | Pendiente | pending | Pendiente | Pendiente | Evidencia runtime reciente por slices y focused PASS para reconstrucciÃ³n completa/fallo post-escritura; no manifest completo verde. |
| Datos relacionados cache-first | Issue #39 / Issue #67 | Pendiente | pending | Pendiente | Pendiente | Los documentos existentes citan 7/7 cache-e2e en `20b71f64`, evidencia adyacente no cobertura completa de cuadro de mando. |
| Indicadores diferidos de proyecto | Cambio de helper de seguimiento | Pendiente | pending | Pendiente | Pendiente | Contratos `Issue38_SeguimientoProyecto_ActualizarModoProyecto` e `Issue50_SeguimientoProyecto_CargaDiferidaHelper` resueltos funcionalmente el 2026-06-15 (refactor de `Form_FormNCProyectoSeguimiento.cls` alineado con el patrÃ³n `Form_FormNCAuditoriaSeguimiento.cls`). |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla documental |
|---|---|---|---|
| Recuentos obsoletos tras escrituras | RegresiÃ³n de sincronizaciÃ³n de alcance afectado | Reejecutar pruebas de mutaciÃ³n/sincronizaciÃ³n de indicadores | BR-IND-5..7 |
| Fuga de dominio | Falta filtro de dominio | Reejecutar/aÃ±adir pruebas cross-domain | BR-IND-4 |
| Vista de detalle incompleta | RegresiÃ³n de esquema/detalle de cachÃ© | Reejecutar pruebas de detalle completo | BR-IND-2 |
| Se discute el significado de buckets | Falta definiciÃ³n de producto | Confirmar definiciones UAT | BR-IND-8 `#73` |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- La cachÃ© de indicadores como estado materializado en backend, no estado de sesiÃ³n frontend (BR-IND-1): la web debe seguir resolviendo recuentos y detalle desde un modelo de lectura backend, no recalcularlos en cada request del navegador.
- La inclusiÃ³n de filas de detalle en la cachÃ©, no solo recuentos agregados (BR-IND-2): las pruebas `Issue18_CargarDetalle` y `Issue18_DetalleCompleto` asÃ­ lo documentan; la API REST debe seguir devolviendo las filas de detalle requeridas por el cuadro de mando, con el mismo esquema.
- El filtrado runtime por usuario conectado/responsable y dominio (BR-IND-3): cada endpoint del cuadro de mando debe aplicar el filtro de responsable del usuario actual, sin permitir que un usuario vea filas de otro responsable o de otro dominio. La regla de per-domain-counts para esquivar el quirk DAO/Jet con `IN (1, 2)` debe replicarse en el backend SQL.
- La separaciÃ³n de filas de Proyecto y AuditorÃ­a en cada lectura (BR-IND-4): la web debe seguir separando dominios en la respuesta, no devolver una uniÃ³n cruzada. La query global debe sumar recuentos per-domain, como ya se documenta para `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic`.
- El refresco incremental por NC afectada tras mutaciones de NC/AC/AR/tarea (BR-IND-5): un cambio en `NC 992099` solo debe refrescar el alcance de esa NC, no disparar una reconstrucciÃ³n completa.
- La resoluciÃ³n de padres ACâ†’NC y AR/tareaâ†’ACâ†’NC (BR-IND-6): cualquier hook de escritura debe resolver la NC padre antes de sincronizar, replicando `Issue18_ResolverNCDesde`.
- La propagaciÃ³n explÃ­cita del fallo de sincronizaciÃ³n post-escritura (BR-IND-7): un fallo de `InvalidarCache` debe propagar `pError` y no afirmar que la cachÃ© estÃ¡ actualizada. La web debe hacer lo propio: devolver `5xx` o un payload con `sincronizado=false`.
- La pertenencia al cuadro de mando como un recurso de lectura, no como UI: la web debe poder consumir los mismos buckets/detalle desde un panel web, una API externa o un job, sin pasar por la UI.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir `ModuloCacheIndicadoresIssue18` y `ModuloCacheIndicadores` por un servicio backend de modelo de lectura con su propio SLA de frescura, no por cÃ³digo VBA in-process.
- Convertir `IndicadorRepositorio` y `IndicadorServicio` en dos servicios REST diferenciados (lectura vs. sincronizaciÃ³n), con autenticaciÃ³n y autorizaciÃ³n declarativas.
- Mover la lÃ³gica de `Issue18_NCWriteHook`, `Issue18_ACWriteHook` y `Issue18_ARWriteHook` a eventos del backend (cola/worker) que se disparen tras las mutaciones de dominio, no a un hook de formulario.
- Reemplazar `Test_IndicadoresCaracterizacion` y `Test_IndicadoresTelemetry` por suites automatizadas de la API REST del cuadro de mando, con tiempos de respuesta esperados documentados.
- Mover el contrato de buckets del cuadro de mando (umbrales, nombres, reglas de visibilidad) a un archivo de configuraciÃ³n versionado y revisable por producto, en lugar de estar embebido en el cÃ³digo.
- Sustituir el patrÃ³n de `Form_FormIndicadores` por una SPA con componentes finos que consuman la API REST, no un formulario Access con estado mutable.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar la carga de indicadores en `Form_Load` con `TimerInterval = 100` y `m_CargaInicialIndicadoresPendiente` como patrÃ³n de la web: la API REST debe poder invocarse de forma sÃ­ncrona o asÃ­ncrona real (cola/worker), no con un timer del cliente.
- No migrar la query global `IN (1, 2)` con `COUNT(*)` sin predicados: la web debe usar SQL parametrizado con `GROUP BY` per-domain, no reproducir el quirk de DAO/Jet.
- No usar la reconstrucciÃ³n completa como ruta normal de escritura: en la web, una "reconstrucciÃ³n completa" es un job explÃ­cito de operaciÃ³n con auditorÃ­a, no un side effect de `Save`.
- No acoplar la lectura de la cachÃ© de indicadores al evento de un formulario concreto: la lectura debe ser un servicio reutilizable por cualquier consumidor.
- No propagar el resultado de un hook de sincronizaciÃ³n fallido como Ã©xito: la web debe devolver error explÃ­cito y, si el cliente lo ignora, no debe reescribir el estado de la cachÃ©.

### Â§6.4 Preguntas abiertas al product owner
- Â¿Los nombres de buckets, umbrales y reglas de visibilidad del cuadro de mando (BR-IND-8) estÃ¡n aprobados? Confirmar lista canÃ³nica y SLA de actualizaciÃ³n por bucket.
- Â¿La resoluciÃ³n ACâ†’NC y AR/tareaâ†’ACâ†’NC (BR-IND-6) puede tener AC o AR sin NC padre en algÃºn caso especial? Hoy la regla es "toda AC/AR pertenece a una NC"; Â¿se mantiene?
- Â¿El fallo de sincronizaciÃ³n post-escritura (BR-IND-7) debe notificar al usuario o basta con log? Hoy se propaga `pError`; Â¿la web debe mostrar un toast o devolver `5xx`?
- Â¿La deuda de rendimiento de los hooks (~120s) se aborda antes de la migraciÃ³n o se acepta como coste? Confirmar presupuesto de tiempo.
- Â¿El manifest completo `tests/tests.vba.indicadores-caracterizacion.json` debe dejarse como timeout histÃ³rico o se reescribe en formato no-aggregate para la web? Hoy timeoutea como conjunto; la web debe poder ejecutarse de forma atÃ³mica.
- Â¿El cuadro de mando debe exponer un endpoint "what-if" para simulaciones o solo los recuentos reales? Confirmar alcance.

## Â§7 Libro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-IND-1 â€” La cachÃ© de indicadores es estado materializado compartido en backend, no estado de sesiÃ³n frontend. | Verified-runtime | `Issue18_BackendCacheSchema` 2/2, `CacheIndicadoresMaterializado` 8/8, `CacheIndicadoresAuditoriaMaterializado` 3/3 tras retry | 2026-06-15 |
| BR-IND-2 â€” La cachÃ© incluye filas de detalle necesarias por el cuadro de mando, no solo recuentos agregados. | Verified-runtime | `Issue18_CargarDetalle` 2/2 (~122s en un test), `Issue18_DetalleCompleto` 1/1 | 2026-06-15 |
| BR-IND-3 â€” Las lecturas runtime filtran por usuario conectado/responsable. | Verified-runtime (resuelto 2026-06-15) | `Issue18_CargarBucket` 2/2 (~120s en un test); `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic` 1/1 tras arreglar la query global del propio test para sumar recuentos por dominio | 2026-06-15 |
| BR-IND-4 â€” Las filas de Proyecto y AuditorÃ­a permanecen separadas por dominio. | Verified-runtime (resuelto 2026-06-15) | `CacheIndicadoresAuditoriaMaterializado` 3/3, `Issue18_CargarDetalle` 2/2; `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic` 1/1 tras el mismo arreglo de query per-domain; el test afirma 5 filas globales sumando Proyecto (1) + AuditorÃ­a (2) y mantiene `QA_User` 3 + `Otro_User` 2 | 2026-06-15 |
| BR-IND-5 â€” Los cambios correctos de NC/AC/AR/tarea sincronizan solo el alcance de la NC afectada. | Verified-runtime | `Issue18_SincronizarNC` 3/3, `Issue18_NCWriteHook` 1/1 (~120s), `Issue18_ACWriteHook` 1/1 (~114s), `Issue18_ARWriteHook` 4/4 con dos hooks ~117â€“123s | 2026-06-15 |
| BR-IND-6 â€” Los cambios de AC resuelven ACâ†’NC; los de AR/tarea resuelven AR/tareaâ†’ACâ†’NC. | Verified-runtime | `Issue18_ResolverNCDesde` 3/3, `Issue18_ACWriteHook` 1/1, `Issue18_ARWriteHook` 4/4 | 2026-06-15 |
| BR-IND-7 â€” La sincronizaciÃ³n post-escritura fallida es visible e impide afirmaciones falsas de cachÃ© actual. | Verified-runtime focused | `Test_Issue18_NCWriteHook_InvalidarCache_FailedSync_ReturnsError_Atomic` PASS el 2026-06-15 (`4788` ms, `issue18_nc_write_hook_failed_sync_ok`): `NC 992099` inexistente, `InvalidarCache` devuelve `False` y propaga `pError` explÃ­cito | 2026-06-15 |
| BR-IND-8 â€” Los nombres, umbrales y salidas de gestiÃ³n de buckets del cuadro de mando estÃ¡n aprobados por producto. | Intended | FALTA â†’ crear mediante access-vba-tdd tras confirmar escenario UAT | 2026-06-15 |
| La cachÃ© materializada de indicadores tiene evidencia runtime reciente por slices. | Verified-runtime | `indicator-fast-counts` 5/5, `cache-materialized` 13/13, filtros Issue #18 descritos en Â§5 | 2026-06-15 |
| El manifest completo de caracterizaciÃ³n de indicadores no puede tratarse como verde completo. | Pending | `tests/tests.vba.indicadores-caracterizacion.json` tiene 55 procedimientos y timeoutea como conjunto; usar slices | 2026-06-15 |
| La reconstrucciÃ³n completa idempotente y la propagaciÃ³n de fallo post-escritura tienen evidencia focused PASS. | Verified-runtime focused | `Test_Issue18_ReconstruirTodo_Idempotent_Atomic` 1/1 (`4862` ms, `issue18_rebuild_idempotent_ok`) y `Test_Issue18_NCWriteHook_InvalidarCache_FailedSync_ReturnsError_Atomic` 1/1 (`4788` ms, `issue18_nc_write_hook_failed_sync_ok`) | 2026-06-15 |
| Las reglas de filtrado por responsable y separaciÃ³n de dominio (BR-IND-3, BR-IND-4) son `Verified-runtime` tras el arreglo del propio test Issue #18 GlobalCache. | Verified-runtime | `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic` 1/1 con query per-domain; recuentos `QA_User` 3 + `Otro_User` 2, total global 5 | 2026-06-15 |
| Los contratos de seguimiento de proyecto (Issue #38, Issue #50) son `Verified-runtime` tras alinear `Form_FormNCProyectoSeguimiento.cls` con el patrÃ³n de `Form_FormNCAuditoriaSeguimiento.cls`. | Verified-runtime | `Test_Issue38_SeguimientoProyecto_ActualizarModoProyecto_Contract` 1/1, `Test_Issue50_SeguimientoProyecto_CargaDiferidaHelper_Contract` 1/1; flags privados, `Form_Timer`, programaciÃ³n en `Form_Load`, delegaciÃ³n al helper y duraciÃ³n del helper presentes | 2026-06-15 |
| El comportamiento adyacente de cache-trust tiene evidencia reciente en documentos de funcionalidad. | Verified-static | Documento existente de funcionalidad cache-e2e; no reejecutado en esta tarea documental | 2026-06-15 |
| Las definiciones de buckets del cuadro de mando estÃ¡n aprobadas por producto. | Intended | ConfirmaciÃ³n pendiente | 2026-06-15 |

**âš ï¸ Divergencias activas (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- No hay divergencias funcionales activas documentadas a 2026-06-15 en la evidencia aportada. Mantener cautela: no hay manifest completo verde de `tests/tests.vba.indicadores-caracterizacion.json`, no hay UAT tag/release y quedan deudas de reachability/rendimiento.

**âœ… Divergencias resueltas el 2026-06-15**
- `Test_Issue18_GlobalCache_DosResponsables_DosDominios_Atomic` â€” query global del test ajustada para sumar recuentos per-domain y esquivar el quirk DAO/Jet con `IN (1, 2)` en `COUNT(*)` sin mÃ¡s predicados.
- `Test_Issue38_SeguimientoProyecto_ActualizarModoProyecto_Contract` â€” `ComandoActualizar_Click` delega ahora al helper vÃ­a `Form_Timer` (`m_CargaInicialIndicadoresPendiente = True`, `Me.TimerInterval = 100`).
- `Test_Issue50_SeguimientoProyecto_CargaDiferidaHelper_Contract` â€” el mismo refactor aÃ±adiÃ³ flags privados, `Form_Timer`, programaciÃ³n en `Form_Load`, delegaciÃ³n y duraciÃ³n del helper con `p_DuracionSegundos:=m_UltimaDuracionIndicadores`.
