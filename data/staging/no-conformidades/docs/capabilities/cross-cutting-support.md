# Capacidad: soporte transversal

## Â§0 Identidad
- **ID de capacidad**: `CAP-CROSS-CUTTING-SUPPORT`
- **Nivel**: standard
- **Estado**: active / documentaciÃ³n alineada con v2; evidencia mixta
- **Fuente**: hybrid (documentos de funcionalidad + inventario de fuente + convenciones operativas)
- **Responsable / autoridad de producto**: ConfirmaciÃ³n pendiente â€” soporte de aplicaciÃ³n / servicios de plataforma
- **Ãšltima verificaciÃ³n**: 2026-06-15 migraciÃ³n solo documental; no se ejecutÃ³ Dysflow/Access
- **Confianza global**: mixed â€” partes de cachÃ©/listado tienen evidencia estÃ¡tica; runbooks de permisos/correo/enrutamiento backend estÃ¡n mayoritariamente sin probar

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Mantener fiable la aplicaciÃ³n Access entre capacidades mediante acceso a datos compartido, frescura de cachÃ©, permisos, diagnÃ³sticos, configuraciÃ³n, correo, logging, serializaciÃ³n y evidencia de regresiÃ³n.
- **Usuarios / personas**: Usuarios de negocio, operadores de soporte, revisores UAT, desarrolladores/agentes IA.
- **Problema que resuelve**: Las mecÃ¡nicas compartidas pueden romper varios flujos de negocio; necesitan trazabilidad y pruebas explÃ­citas en lugar de acoplamiento oculto.
- **Valor de negocio / por quÃ© existe**: Servicios compartidos fiables reducen regresiones entre Proyecto, AuditorÃ­a, documentos, indicadores y releases.
- **No objetivos**: Esta pÃ¡gina no sustituye a los documentos de capacidad de dominio; enlaza mecÃ¡nicas de apoyo.
- **Fuente de intenciÃ³n**: Documentos de funcionalidad/inventario de fuente existentes; runbooks de permisos/correo/operador pendientes de confirmaciÃ³n.
- **Referencia tracker de origen**: Issue #67, Issue #39, Issue #18, cambios de cachÃ© de auditorÃ­a.

## Â§2 Contrato de comportamiento

### Escenarios (Given / When / Then)
- **GIVEN** un formulario de negocio necesita datos cacheados **WHEN** la cachÃ© estÃ¡ vacÃ­a/desactivada/obsoleta **THEN** el comportamiento de soporte debe evitar resultados falsamente vacÃ­os/obsoletos.
- **GIVEN** se reconstruye la cachÃ© de auditorÃ­a **WHEN** se ejecuta la operaciÃ³n **THEN** los lÃ­mites transaccionales impiden datos parciales.
- **GIVEN** falla la sincronizaciÃ³n de indicadores **WHEN** una escritura de negocio tuvo Ã©xito **THEN** el fallo es visible y no se afirma una cachÃ© actual falsa.
- **GIVEN** el enrutamiento/configuraciÃ³n backend es incorrecto **WHEN** se ejecutan pruebas/UAT **THEN** los diagnÃ³sticos identifican el desacuerdo antes de confiar en los datos.
- **GIVEN** se ejecuta una acciÃ³n/correo/informe sensible **WHEN** aplican permisos o notificaciones **THEN** las reglas deben ser explÃ­citas y estar probadas antes de afirmar release.

### Reglas de negocio
| ID de regla | Enunciado (previsto) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba (evidencia) | Confianza |
|---|---|---|---|---|---|
| BR-XCUT-1 | El soporte de cachÃ©/repositorio debe preservar el comportamiento de dominio y no cambiar el significado de negocio. | Contrato de capacidad | Parcial | Documentos de funcionalidad existentes; hace falta reejecutar | Verified-static |
| BR-XCUT-2 | Los resultados de cachÃ© cargada-vacÃ­a son vÃ¡lidos y no fallos de cachÃ©. | Documento de funcionalidad cache-trust | SÃ­ segÃºn docs | `tests/tests.vba.cache-e2e.json` existente; hace falta reejecutar | Verified-static |
| BR-XCUT-3 | La reconstrucciÃ³n de cachÃ© de auditorÃ­a es atÃ³mica. | Documento de funcionalidad de auditorÃ­a | SÃ­ segÃºn docs | Manifest de auditorÃ­a existente; hace falta reejecutar | Verified-static |
| BR-XCUT-4 | La cachÃ© compartida de indicadores filtra por usuario/responsable/dominio y expone fallos de sincronizaciÃ³n. | Docs Issue #18 | HistÃ³rico / reciente pendiente | FALTA â†’ crear mediante access-vba-tdd; reejecutar/aÃ±adir pruebas de staging actual | Intended |
| BR-XCUT-5 | La configuraciÃ³n backend enruta lecturas/escrituras al entorno previsto; las pruebas no dependen de datos accidentales. | Reglas de seguridad del proyecto | Parcial/nombres de fuente | FALTA â†’ crear mediante access-vba-tdd; definir/ejecutar diagnÃ³sticos de configuraciÃ³n backend | Intended |
| BR-XCUT-6 `#82` | La matriz de permisos para cerrar/eliminar/rehabilitar/documento/acciÃ³n/informe/configuraciÃ³n es explÃ­cita. | Autoridad de producto pendiente | Desconocido | FALTA â†’ crear mediante `access-vba-tdd` tras confirmar matriz; aÃ±adir UAT cuando proceda. Cross-link: misma matriz pretendida por `users-permissions-navigation` BR-UPN-7 y `master-data-catalogues` BR-CAT-6 (cobertura CRUD de catÃ¡logos) | Intended |
| BR-XCUT-7 | El comportamiento de correo/logging requerido por negocio es explÃ­cito y comprobable. | Autoridad de producto pendiente | Desconocido | FALTA â†’ crear mediante access-vba-tdd tras confirmar notificaciÃ³n/log; probar costuras | Intended |

### Validaciones
- Validez de esquema/frescura de cachÃ© antes de vistas respaldadas por cachÃ©.
- Filtros de dominio/usuario en cachÃ© de indicadores.
- Seguridad del destino backend antes de confiar en evidencia de pruebas/UAT.
- Comprobaciones de permisos para acciones sensibles pendientes.

### Transiciones de estado
- `CachÃ© de proyecto vacÃ­a/desactivada` --(`Listado/lectura`)--> `Listado/lectura fallback legacy`.
- `CachÃ© de auditorÃ­a obsoleta/ausente` --(`Reconstruir`)--> `CachÃ© de auditorÃ­a reconstruida`.
- `CachÃ© de indicadores actual` --(`Escritura de negocio`)--> `Alcance de NC afectada sincronizado`.
- `Backend/configuraciÃ³n incorrectos` --(`DiagnÃ³stico`)--> `Problema de enrutamiento diagnosticado`.
- `Usuario sin permiso` --(`AcciÃ³n sensible`)--> `AcciÃ³n bloqueada/limitada` â€” pendiente de confirmaciÃ³n.

### Caminos lÃ­mite y de error
- Datos de backend incorrecto invalidan conclusiones de pruebas/UAT.
- El comportamiento de permisos/correo/logging no puede inferirse de nombres de clases.
- El comportamiento de soporte visible desde formularios debe probarse mediante costuras helper/servicio.

### SeÃ±ales de aceptaciÃ³n / presencia
- Las pruebas de cachÃ©/listado/auditorÃ­a/indicadores pasan en staging actual.
- Los diagnÃ³sticos de enrutamiento backend estÃ¡n documentados y son deterministas.
- La matriz de permisos y los eventos obligatorios de correo/log tienen pruebas o evidencia UAT.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada UI**: formularios de negocio en Proyecto/AuditorÃ­a; `Form_FormIndicadores`; `Form_FormCorreo`; rutas operativas de cachÃ©/preparaciÃ³n/warm-up pendientes.
- **Puntos de entrada de fuente**: `CacheNCService`, `CacheNCCrud`, `CacheNCCacheRepositorio`, `CacheNCProyecto`, `CacheTrustDiagnostics`, `NCAuditoriaListadoCache`, `NCAuditoriaGestionListadoHelper`, `NCProyectoGestionListadoHelper`, `ModuloCacheIndicadoresIssue18`, `ModuloCacheIndicadores`, `IndicadorRepositorio`, `IndicadorServicio`, `NCRepository`, `ACRepository`, `ARRepository`, `RiesgoRepositorio`, `Usuario`, `UsuarioAplicacionPermisos`, `Entorno`, `Variables Globales`, `Test_BackendConfigPaths`, `Correo`, `LogNCProyecto`, `LogNCAuditoria`, `JSONHelper`, `JsonConverter`.
- **Datos tocados**: cachÃ©s de proyecto/auditorÃ­a/listado, cachÃ© de indicadores, `TbConfiguracionBackends`, usuarios/permisos, registros de logs/correo.
- **Salidas**: diagnÃ³sticos/logs de cachÃ©, evidencia de preparaciÃ³n/warm-up, correos, libros de funcionalidad de regresiÃ³n.
- **Dependencias e integraciones**: todas las capacidades de negocio.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada en esta tarea solo documental.
- **EvaluaciÃ³n de diseÃ±o (as-built vs ideal)**: las costuras de cachÃ© son cada vez mÃ¡s explÃ­citas. Permisos/correo/enrutamiento backend necesitan pruebas/runbooks de primer nivel antes de la migraciÃ³n.

## Â§4 Receta de reconstrucciÃ³n
1. Mantener documentos de soporte enlazados a las capacidades de negocio afectadas; no tratar las mecÃ¡nicas de soporte como comportamiento de negocio por sÃ­ mismas.
2. Para cualquier comportamiento de soporte que toque datos, inspeccionar primero el esquema y crear fixtures sandbox deterministas.
3. Crear pruebas de enrutamiento backend, permisos, costuras de correo/log y soporte de indicadores en staging actual.
4. Las reglas de soporte visibles desde formularios deben probarse mediante costuras helper/servicio; pruebas directas de formulario solo para cableado.
5. Cambios futuros de cÃ³digo: importaciÃ³n Dysflow â†’ compilaciÃ³n manual del usuario â†’ pruebas Dysflow.

## Â§5 Evidencia y trazabilidad
- **Pruebas**: los documentos existentes citan form-helper, project-cache, cache-e2e, audit-helper, manifests histÃ³ricos de indicadores/preparaciÃ³n/warm-up/fast-count. No hubo ejecuciÃ³n reciente en esta tarea. **Tampoco existe manifest dedicado a permisos/roles/navegaciÃ³n**; ver `users-permissions-navigation` Â§5 para el detalle de bÃºsqueda y BR-UPN-1..6. La precondiciÃ³n comÃºn para ejecutar pruebas de cualquiera de estas Ã¡reas es `configuration-backends-runtime` BR-CFG-5 (`AssertSafeBackendForCatalogBootstrap`) y BR-CFG-6 (auditorÃ­a de routing/kill-switch/indicadores).

| Elemento (funcionalidad o arreglo) | Ref. tracker | VersiÃ³n staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en prod | Nota |
|---|---|---|---|---|---|---|
| Soporte de listado/cachÃ© de proyecto | Issue #67 | Pendiente | pending | Pendiente | Pendiente | Los documentos existentes citan `20b71f64`. |
| Confianza cache-first AC/AR/Riesgo | Issue #39 / Issue #67 | Pendiente | pending | Pendiente | Pendiente | Los documentos existentes citan `23af345` / `20b71f64`. |
| Soporte de cachÃ©/informe de auditorÃ­a | Cambio de cachÃ© de auditorÃ­a | Pendiente | pending | Pendiente | Pendiente | Los documentos existentes citan varios SHAs; traza de arreglo de informe pendiente. |
| Soporte compartido de indicadores | Issue #18 | Pendiente | pending | Pendiente | Pendiente | Evidencia histÃ³rica; ejecuciÃ³n reciente pendiente. |
| Matriz de permisos (cerrar/eliminar/rehabilitar/documento/acciÃ³n/informe/configuraciÃ³n) | Pendiente | Pendiente | pending | Pendiente | Pendiente | Falta evidencia dedicada. Cross-link: `users-permissions-navigation` BR-UPN-7 (misma matriz). |
| Permisos/correo/enrutamiento backend | Pendiente | Pendiente | pending | Pendiente | Pendiente | Falta evidencia dedicada. |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla documental |
|---|---|---|---|
| Listas de proyecto/auditorÃ­a obsoletas/en blanco | RegresiÃ³n de soporte de cachÃ© | Reejecutar manifests helper/cachÃ© | BR-XCUT-1..3 |
| Recuentos del cuadro de mando obsoletos | RegresiÃ³n de soporte de indicadores | Reejecutar manifests actuales de indicadores | BR-XCUT-4 |
| Pruebas/UAT ven datos incorrectos | Desacuerdo de enrutamiento backend | Ejecutar/crear diagnÃ³sticos de configuraciÃ³n backend | BR-XCUT-5 |
| AcciÃ³n no autorizada permitida/bloqueada | Falta/regresiÃ³n de matriz de permisos | Crear pruebas/UAT de permisos (cross-link `users-permissions-navigation` BR-UPN-7) | BR-XCUT-6 `#82` |
| Falta notificaciÃ³n requerida | Costura de correo indefinida/regresada | Confirmar regla + crear prueba de costura | BR-XCUT-7 |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- La semÃ¡ntica cargado-vacÃ­o de las cachÃ©s de proyecto, auditorÃ­a e indicadores como resultado vÃ¡lido, no como fallo (BR-XCUT-1, BR-XCUT-2): la web debe distinguir "cachÃ© cargada con cero filas" de "cachÃ© no disponible", igual que `cache-e2e` y los diagnÃ³sticos de cache-trust ya lo prueban.
- La atomicidad de la reconstrucciÃ³n de la cachÃ© de auditorÃ­a (BR-XCUT-3): si el sistema hace `Reconstruir` y falla a mitad, no debe quedar un dataset parcial: la web debe replicar el patrÃ³n transaccional de borrado-y-regeneraciÃ³n o un job con compensaciÃ³n.
- El filtrado de la cachÃ© compartida de indicadores por usuario/responsable y dominio (BR-XCUT-4): el backend de la web debe seguir aplicando filtros de autorizaciÃ³n/dominio en cada lectura, no permitir que un usuario vea filas de otro responsable o de otro dominio.
- La regla de routing de backend expuesta al diagnÃ³stico (BR-XCUT-5): si un test o UAT ejecuta contra el backend equivocado, el sistema debe detectarlo y reportarlo en vez de seguir adelante, igual que `AssertSafeBackendForCatalogBootstrap` ya hace para catÃ¡logos.
- La matriz de permisos por acciÃ³n sensible (cerrar/eliminar/rehabilitar/documento/acciÃ³n/informe/configuraciÃ³n) como contrato Ãºnico (BR-XCUT-6): la web debe tener una sola fuente de verdad que aplique esa matriz; no debe haber reglas duplicadas por formulario.
- La obligatoriedad de notificar/loggear eventos sensibles (BR-XCUT-7): cualquier acciÃ³n de cerrar/eliminar/rehabilitar debe dejar traza de quiÃ©n, cuÃ¡ndo, desde quÃ© IP, en quÃ© estado quedÃ³.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Tratar `CacheNCService`, `CacheNCCrud`, `CacheNCCacheRepositorio`, `CacheNCProyecto`, `ModuloCacheIndicadoresIssue18`, `ModuloCacheIndicadores` como una sola familia de servicios: API REST de modelos de lectura con versiÃ³n, invalidaciÃ³n explÃ­cita y reintento, no como mÃ³dulos VBA acoplados.
- Reemplazar `NCAuditoriaListadoCache`, `NCAuditoriaGestionListadoHelper`, `NCProyectoGestionListadoHelper` por una capa de aplicaciÃ³n con la misma responsabilidad (lista/cachÃ©) pero expuesta como endpoints y con reintento en vez de formularios.
- Convertir `IndicadorRepositorio` y `IndicadorServicio` en servicios de lectura materializada (modelos de lectura backend) con su propio SLA de frescura y observabilidad, en lugar de llamadas in-process desde VBA.
- Sustituir `LogNCProyecto` / `LogNCAuditoria` (escritura de logs) por un appender de logs estructurados (JSON) a un bus de eventos, no por una tabla de Access.
- Mover `JSONHelper` / `JsonConverter` al backend de la web (no debe quedar cÃ³digo de serializaciÃ³n dentro de la UI ni de los servicios de lectura).
- Convertir `Usuario` y `UsuarioAplicacionPermisos` en un servicio de autorizaciÃ³n con guardas declarativas, no como flags booleanos consultados en cada formulario.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No migrar globals (`m_ObjEntorno`, `m_ObjUsuarioConectado`) ni TempVars como estado compartido entre request en la web: el estado de runtime vive en el request scope del framework, no en variables globales.
- No acoplar lÃ³gica de soporte a eventos de formulario (`Form_Load`, `Form_Open`): la inicializaciÃ³n de cachÃ©, de permisos y de diagnÃ³stico debe ocurrir en el arranque del servicio o del request, no en eventos UI.
- No usar la cinta (Ribbon) ni la visibilidad de menÃºs como control de seguridad real: la web debe aplicar permisos en el servidor y devolver `403` cuando corresponda, no ocultar UI como "control de acceso".
- No replicar el patrÃ³n "leer config desde `Variables Globales` cada vez" en la web: la configuraciÃ³n se lee una vez al arranque del proceso y se cachea de forma inmutable.
- No portar `TempVars` como mecanismo de comunicaciÃ³n entre formularios: la web usa rutas, query params y body de request, no variables globales.

### Â§6.4 Preguntas abiertas al product owner
- Â¿La matriz de permisos (BR-XCUT-6) es la misma para Proyecto y AuditorÃ­a o se diferencia? Confirmar si las acciones sensibles (cerrar/eliminar/rehabilitar) tienen la misma matriz en ambos dominios.
- Â¿QuÃ© eventos generan notificaciÃ³n obligatoria (BR-XCUT-7)? Â¿Basta con loggear o hay que enviar correo/Slack? Confirmar umbral y destinatarios.
- Â¿La antigÃ¼edad de cachÃ© tolerable tiene un SLA formal? Hoy la web no expone un SLA explÃ­cito; Â¿debe ser configurable por tipo de dato?
- Â¿El `IndicadorServicio` debe seguir exponiendo las mismas filas de detalle que `Issue18_CargarDetalle` o se redefinen los campos de detalle? (BR-IND-2 adyacente)
- Â¿La `matriz de permisos` debe ser declarativa (roles â†’ acciones) o seguir siendo checks booleanos (`EsTecnico`, `EsAdministrador`)? Â¿QuÃ© cobertura quiere el equipo de seguridad?

## Â§7 Libro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-XCUT-1 â€” El soporte de cachÃ©/repositorio debe preservar el comportamiento de dominio y no cambiar el significado de negocio. | Verified-static | Documentos de funcionalidad existentes; FALTA â†’ reejecutar antes de promover | 2026-06-15 |
| BR-XCUT-2 â€” Los resultados de cachÃ© cargada-vacÃ­a son vÃ¡lidos y no fallos de cachÃ©. | Verified-static | `tests/tests.vba.cache-e2e.json` existente; FALTA â†’ reejecutar | 2026-06-15 |
| BR-XCUT-3 â€” La reconstrucciÃ³n de cachÃ© de auditorÃ­a es atÃ³mica. | Verified-static | Manifest de auditorÃ­a existente; FALTA â†’ reejecutar | 2026-06-15 |
| BR-XCUT-4 â€” La cachÃ© compartida de indicadores filtra por usuario/responsable/dominio y expone fallos de sincronizaciÃ³n. | Intended | FALTA â†’ crear mediante access-vba-tdd; reejecutar/aÃ±adir pruebas de staging actual | 2026-06-15 |
| BR-XCUT-5 â€” La configuraciÃ³n backend enruta lecturas/escrituras al entorno previsto; las pruebas no dependen de datos accidentales. | Intended | FALTA â†’ crear mediante access-vba-tdd; definir/ejecutar diagnÃ³sticos de configuraciÃ³n backend | 2026-06-15 |
| BR-XCUT-6 â€” La matriz de permisos para cerrar/eliminar/rehabilitar/documento/acciÃ³n/informe/configuraciÃ³n es explÃ­cita. | Intended | FALTA â†’ crear mediante `access-vba-tdd` tras confirmar matriz; cross-link `users-permissions-navigation` BR-UPN-7 y `master-data-catalogues` BR-CAT-6 | 2026-06-15 |
| BR-XCUT-7 â€” El comportamiento de correo/logging requerido por negocio es explÃ­cito y comprobable. | Intended | FALTA â†’ crear mediante access-vba-tdd tras confirmar notificaciÃ³n/log; probar costuras | 2026-06-15 |
| El soporte de cachÃ©/listado de proyecto tiene pruebas documentadas. | Verified-static | Documentos de funcionalidad existentes; sin reejecuciÃ³n | 2026-06-15 |
| El soporte de cachÃ© de auditorÃ­a tiene pruebas documentadas. | Verified-static | Documento de funcionalidad de auditorÃ­a existente; sin reejecuciÃ³n | 2026-06-15 |
| El soporte de indicadores estÃ¡ vigente en staging. | Intended | Solo evidencia histÃ³rica | 2026-06-15 |
| El comportamiento de permisos/correo/enrutamiento backend estÃ¡ completamente especificado. | Intended | Faltan reglas/pruebas | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Sin divergencia confirmada. Hueco sospechado: el inventario de clases de soporte sugiere capacidades de permisos/correo/enrutamiento backend que aÃºn no tienen contrato de negocio ni pruebas runtime. La pieza de permisos es la misma matriz que `users-permissions-navigation` declara como intenciÃ³n en BR-UPN-7; mientras esa matriz no estÃ© aprobada y probada, BR-XCUT-6 no puede ascender de `Intended`.
