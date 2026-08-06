# Capacidad: acciones y seguimiento de NC AuditorÃ­a

## Â§0 Identidad
- **ID de capacidad**: `CAP-NCA-ACTIONS-FOLLOWUP`
- **Nivel**: critical
- **Estado**: active / documentaciÃ³n alineada con v2; falta evidencia runtime dedicada
- **Fuente**: reverse-engineered/hybrid (inventario de fuente + documentos adyacentes de auditorÃ­a/indicadores)
- **Responsable / autoridad de producto**: ConfirmaciÃ³n pendiente â€” dominio AuditorÃ­a / Calidad
- **Ãšltima verificaciÃ³n**: 2026-06-15 actualizaciÃ³n documental con evidencia runtime ya recogida; en esta tarea no se ejecutÃ³ Dysflow/Access
- **Confianza global**: mixed â€” selecciÃ³n de informe/listado y algunos hooks de seguimiento/indicadores de auditorÃ­a tienen evidencia runtime por manifest/slices; el ciclo completo de acciones sigue pendiente

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Realizar seguimiento de acciones correctoras, acciones de resoluciÃ³n, replanificaciones, notas y tareas de seguimiento para NC con origen en auditorÃ­a.
- **Usuarios / personas**: Equipo de auditorÃ­a/calidad, auditores/coordinadores, managers/revisores, desarrolladores/agentes IA.
- **Problema que resuelve**: Mantiene el trabajo de acciones de auditorÃ­a correcto en su dominio y visible sin fugas de estado de acciones de proyecto.
- **Valor de negocio / por quÃ© existe**: Los hallazgos de auditorÃ­a necesitan acciones responsables, fechas lÃ­mite, notas y evidencia de seguimiento antes del cierre/aceptaciÃ³n de release.
- **No objetivos**: El comportamiento de acciones de proyecto y las definiciones globales del cuadro de mando viven en pÃ¡ginas de capacidad separadas.
- **Fuente de intenciÃ³n**: Nombres de fuente y documentos adyacentes; reglas de negocio exactas pendientes de confirmaciÃ³n.
- **Referencia tracker de origen**: Evidencia de regresiÃ³n de informe de auditorÃ­a; Issue #18 comportamiento adyacente de indicadores; Issue #67 documentaciÃ³n.

## Â§2 Contrato de comportamiento

### Escenarios (Given / When / Then)
- **GIVEN** una NC de auditorÃ­a **WHEN** el usuario abre seguimiento **THEN** las AC/AR/tareas/notas/replanificaciones mostradas deben pertenecer al dominio auditorÃ­a.
- **GIVEN** cambia una acciÃ³n/tarea de auditorÃ­a **WHEN** se espera sincronizaciÃ³n de indicadores **THEN** se deben preservar filtros de dominio AuditorÃ­a y refresco de alcance afectado.
- **GIVEN** una NC de auditorÃ­a seleccionada en la UI de gestiÃ³n **WHEN** se ejecuta una salida de informe/listado **THEN** la NC de auditorÃ­a seleccionada se resuelve mediante helpers de auditorÃ­a.
- **GIVEN** flujos de crear/completar/cancelar/reasignar/replanificar/anotar acciones **WHEN** se guardan **THEN** las reglas de propietario, fecha, estado e historial deben probarse con pruebas dedicadas.

### Reglas de negocio
| ID de regla | Enunciado (previsto) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba (evidencia) | Confianza |
|---|---|---|---|---|---|
| BR-NCA-AF-1 | Las acciones/seguimiento de auditorÃ­a siguen siendo especÃ­ficas de dominio y nunca enrutan por estado de acciones de Proyecto. | Tests `tests/tests.vba.cap-nca-af.json` | SÃ­ â€” `Test_NCA_Particula_TipoNC_RetornaNoConformidad_Atomic` + `_TipoOB_..._Observacion_Atomic` + `_TipoOP_..._OportunidadDeMejora_Atomic` + `_TipoDesconocido_RetornaVacio_Atomic` (edge case) + `Test_NCA_Titulo_SinAuditoria_RetornaVacio_Atomic` (early exit) â€” 5/5 PASS contra staging HEAD. MÃ¡s evidencia previa: `Issue38_SeguimientoAuditoria` 1/1, `Issue38_ResetearColTareas` 1/1, slices Issue #18 con Auditoria AC->NC / Auditoria AR hook. | Cubierto por property + edge cases | Verified-runtime |
| BR-NCA-AF-2 | La selecciÃ³n de informe/listado de auditorÃ­a usa helpers de auditorÃ­a y NC de auditorÃ­a seleccionadas. | Documento de funcionalidad de auditorÃ­a | SÃ­ | `ComandoInforme_Click` usa `EnsureNCAuditoriaGestionSelected`; `tests/tests.vba.audit-gestion-helper.json` pasÃ³ 11/11 tras el arreglo | Verified-runtime |
| BR-NCA-AF-3 | Los indicadores compartidos pueden incluir filas de AuditorÃ­a, pero las lecturas runtime filtran por dominio/responsable. | Documentos de funcionalidad de indicadores | Parcial | Evidencia por slices: `CacheIndicadoresAuditoriaMaterializado` 3/3; `CacheIndicadoresAuditoriaMaterializado_SincronizarDesdeNegocio` pasÃ³ dentro de slice 3/3; no afirmar suite completa `tests/tests.vba.indicadores-caracterizacion.json` verde. | Verified-runtime |
| BR-NCA-AF-4 `#74 (BR-NCA-AF-4/5)` | Las reglas de creaciÃ³n, vencimientos, finalizaciÃ³n, cancelaciÃ³n, replanificaciÃ³n, notas y asignaciÃ³n de propietario de acciones de auditorÃ­a son explÃ­citas. | Autoridad de producto pendiente | Desconocido | FALTA â†’ crear mediante access-vba-tdd tras confirmar esquema/reglas | Intended |
| BR-NCA-AF-5 `#74 (BR-NCA-AF-4/5)` | El comportamiento de formularios de seguimiento de auditorÃ­a permanece como cableado UI fino sobre costuras helper/servicio. | Regla de usuario/proyecto | Desconocido | FALTA â†’ crear mediante access-vba-tdd contra costuras helper/servicio, no comportamiento directo de formulario | Intended |

### Validaciones
- La NC de auditorÃ­a padre existe para el contexto de acciones/seguimiento.
- El aislamiento de dominio es bloqueante para selecciÃ³n, indicadores e informes.
- Reglas de propietario/vencimiento/replanificaciÃ³n/finalizaciÃ³n/cancelaciÃ³n pendientes.

### Transiciones de estado
- `Sin acciÃ³n/tarea de auditorÃ­a` --(`Crear`)--> `AcciÃ³n/tarea de auditorÃ­a creada`.
- `AcciÃ³n/tarea de auditorÃ­a abierta` --(`Editar/replanificar/anotar`)--> `Estado de seguimiento actualizado`.
- `AcciÃ³n/tarea de auditorÃ­a modificada` --(`Sincronizar`)--> `Indicador/detalle de AuditorÃ­a refrescado`.
- `NC de auditorÃ­a seleccionada` --(`Comando de informe/listado`)--> `Salida de dominio auditorÃ­a generada`.

### Caminos lÃ­mite y de error
- Que aparezcan datos de Proyecto en seguimiento de auditorÃ­a es una fuga crÃ­tica de dominio.
- La pÃ©rdida de historial de replanificaciÃ³n/notas requiere pruebas dedicadas de retenciÃ³n.

### SeÃ±ales de aceptaciÃ³n / presencia
- Existen pruebas acotadas/slices para seguimiento, reset de tareas y hooks AC/AR de auditorÃ­a; faltan pruebas completas de ciclo de vida de acciones.
- Las pruebas de indicadores/dominio disponibles son por slices y no permiten declarar verde completa `tests/tests.vba.indicadores-caracterizacion.json`.
- El comportamiento de selecciÃ³n de informe de gestiÃ³n de auditorÃ­a estÃ¡ cubierto mediante helper de selecciÃ³n; otros formularios de seguimiento/acciones requieren pruebas dedicadas de costuras helper/servicio.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada UI**: `Form_FormNCAuditoriaSeguimiento`, `Form_FormNCAuditoriaSeguimientoNC`, `Form_FormNCAuditoriaSeguimientoTareas`, `Form_FormNCAuditoriaAcciones`, `Form_FormNCAuditoriaAC`, `Form_FormNCAuditoriaAR`, `Form_FormNCAuditoriaReplanificaciones`, `Form_FormNCAuditoriaNota`.
- **Puntos de entrada de fuente**: `NCAuditoriaSeguimientoHelper`, `ACAuditoria`, `ACAuditoriaOperaciones`, `ARAuditoria`, `ARAuditoriaOperaciones`, `SegNCAuditoria`, `SegTareasAuditoria`, `ReplanificacionesAuditoria`, `ReplanificacionesAuditoriaOperaciones`, `LogNCAuditoria`.
- **Datos tocados**: NC AuditorÃ­a, AC/AR de auditorÃ­a, datos de seguimiento/tareas, replanificaciones, notas/logs, cachÃ© compartida de indicadores.
- **Salidas**: vistas de seguimiento de auditorÃ­a, indicadores de acciones/tareas, salida de informe de auditorÃ­a, historial de notas/replanificaciÃ³n.
- **Dependencias e integraciones**: ciclo de vida de auditorÃ­a, cuadro de mando de indicadores, documentos/evidencia, control eficacia.
- **SincronizaciÃ³n fuenteâ†”binario**: para el arreglo relacionado de informe/listado, `src/forms/Form_FormNCAuditoriaGestion.cls` fue importado antes de esta tarea, el usuario compilÃ³ manualmente y `dysflow_verify_binary` fue correcto para `.cls` y `.form.txt`. En esta tarea solo documental no se importÃ³, compilÃ³ ni ejecutaron tests.
- **EvaluaciÃ³n de diseÃ±o (as-built vs ideal)**: el dominio tiene formularios/clases identificables, pero carece de contrato helper/servicio probado. Tratarlo como alto riesgo de regresiÃ³n/migraciÃ³n hasta que existan pruebas.

## Â§4 Receta de reconstrucciÃ³n
1. Confirmar modelo de estados, campos obligatorios, permisos y reglas de retenciÃ³n de AC/AR/tareas/notas/replanificaciÃ³n de auditorÃ­a.
2. Inspeccionar esquema y diseÃ±ar grafo de fixtures determinista: NC de auditorÃ­a padre â†’ AC/AR/tarea â†’ notas/replanificaciones.
3. Extraer/apuntar a costuras helper/servicio para mutaciones y refresco de indicadores; mantener formularios finos.
4. Crear pruebas JSON `Public Function` mediante `access-vba-tdd`, con `DAO.Database` explÃ­cito, fixtures sandbox, cardinalidad para mutaciones y aserciones contra fugas de dominio.
5. Cambios futuros de cÃ³digo: importaciÃ³n Dysflow â†’ compilaciÃ³n manual del usuario â†’ pruebas Dysflow.

## Â§5 Evidencia y trazabilidad
- **Pruebas**: evidencia ya recogida: `tests/tests.vba.audit-gestion-helper.json` pasÃ³ 11/11 tras el arreglo de selecciÃ³n de informe; `Issue38_SeguimientoAuditoria` pasÃ³ 1/1; `Issue38_ResetearColTareas` pasÃ³ 1/1; `Issue18_ResolverNCDesde` 3/3 incluye Auditoria AC->NC; `Issue18_ARWriteHook` incluye hook Auditoria AR; `CacheIndicadoresAuditoriaMaterializado` pasÃ³ 3/3. Faltan pruebas dedicadas de ciclo completo de acciones/seguimiento de auditorÃ­a.
- **Caveat de runner**: la operaciÃ³n obsoleta `dysflow-51869803-608b-44bc-8792-ef9ca837b894`, posteriormente movida a `status=timed_out`, procede de una interrupciÃ³n no relacionada de `proyecto-gestion-helper`; no es un fallo funcional de auditorÃ­a ni una prueba pendiente de gestiÃ³n de proyecto.

| Elemento (funcionalidad o arreglo) | Ref. tracker | VersiÃ³n staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en prod | Nota |
|---|---|---|---|---|---|---|
| Seguridad de selecciÃ³n de informe/listado de auditorÃ­a | Evidencia de regresiÃ³n 2026-06-14 | Pendiente | pending | Pendiente | Pendiente | `ComandoInforme_Click` usa `EnsureNCAuditoriaGestionSelected`; manifest helper de auditorÃ­a 11/11. |
| Filas/detalle de indicadores de AuditorÃ­a | Issue #18 | Pendiente | pending | Pendiente | Pendiente | Evidencia por slices y pruebas acotadas: `CacheIndicadoresAuditoriaMaterializado` 3/3, `Issue38_SeguimientoAuditoria` 1/1, `Issue38_ResetearColTareas` 1/1, hooks AC/AR de Issue #18. |
| Flujos completos de acciones/seguimiento de auditorÃ­a | Issue #67 docs | Pendiente | pending | Pendiente | Pendiente | FALTA â†’ crear mediante access-vba-tdd; la evidencia actual no cubre crear/completar/cancelar/reasignar/replanificar/anotar de extremo a extremo. |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla documental |
|---|---|---|---|
| El seguimiento de auditorÃ­a muestra datos de Proyecto | Falta enrutamiento de dominio | Crear/reejecutar pruebas de dominio de seguimiento de auditorÃ­a | BR-NCA-AF-1 |
| Filas de cuadro de mando de auditorÃ­a obsoletas/ausentes | RegresiÃ³n de sincronizaciÃ³n/filtro de dominio de indicadores | Reejecutar/aÃ±adir pruebas de indicadores AuditorÃ­a | BR-NCA-AF-3 |
| El guardado de acciÃ³n/tarea difiere tras release | Faltan pruebas de ciclo de vida | Crear pruebas de acciones de auditorÃ­a | BR-NCA-AF-4..5 |
| Se pierde historial de notas/replanificaciÃ³n | RetenciÃ³n/enlace sin probar | Confirmar esquema/reglas + pruebas | BR-NCA-AF-4 `#74 (BR-NCA-AF-4/5)` |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- Las acciones/seguimiento de auditorÃ­a permanecen en su dominio: nunca enrutan por estado de acciones de Proyecto (BR-NCA-AF-1). La web debe mantener separados los endpoints y rutas de NC-Auditoria-AC/AR/tarea y NC-Proyecto-AC/AR/tarea, con guardas explÃ­citas de `tipoDominio`.
- La selecciÃ³n de informe/listado de auditorÃ­a sigue resolviendo por helpers de auditorÃ­a y NC de auditorÃ­a seleccionada (BR-NCA-AF-2): el endpoint de generaciÃ³n de informe de auditorÃ­a debe seguir exigiendo un `IDNoConformidad` vÃ¡lido (de la NC de auditorÃ­a) y `EnsureNCAuditoriaGestionSelected` debe sobrevivir a la migraciÃ³n como guard de la capa de aplicaciÃ³n, no como evento de formulario.
- Las filas de indicadores compartidos pueden incluir auditorÃ­a, pero las lecturas filtran por dominio/responsable (BR-NCA-AF-3): el filtrado per-domain y per-responsable del cuadro de mando debe seguir aplicando, sin permitir fugas.
- Las reglas de crear, vencimientos, finalizaciÃ³n, cancelaciÃ³n, replanificaciÃ³n, notas y asignaciÃ³n de propietario de acciones de auditorÃ­a (BR-NCA-AF-4): la web debe seguir exigiendo los mismos campos obligatorios y las mismas transiciones de estado que la app VBA.
- Los formularios de seguimiento de auditorÃ­a como cableado UI fino sobre costuras helper/servicio (BR-NCA-AF-5): la web debe poder llamar a los mismos servicios de mutaciÃ³n que la UI, sin lÃ³gica embebida en componentes de UI.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir `Form_FormNCAuditoriaSeguimiento`, `Form_FormNCAuditoriaSeguimientoNC`, `Form_FormNCAuditoriaSeguimientoTareas`, `Form_FormNCAuditoriaAcciones`, `Form_FormNCAuditoriaAC`, `Form_FormNCAuditoriaAR`, `Form_FormNCAuditoriaReplanificaciones`, `Form_FormNCAuditoriaNota` por endpoints REST diferenciados: `GET/POST/PUT` por recurso (`ac`, `ar`, `tarea`, `replanificacion`, `nota`).
- Convertir `NCAuditoriaSeguimientoHelper`, `ACAuditoriaOperaciones`, `ARAuditoriaOperaciones`, `ReplanificacionesAuditoriaOperaciones` en servicios backend con una firma por comando (`Crear`, `Finalizar`, `Cancelar`, `Reasignar`, `Replanificar`, `Anotar`).
- Reemplazar el patrÃ³n de `LogNCAuditoria` por un appender de logs estructurados a un bus de eventos, no por una tabla de Access consultable.
- Mover la regla "el seguimiento de auditorÃ­a muestra solo auditorÃ­a" a un middleware de autorizaciÃ³n que valide `tipoDominio=Auditoria` en cada request, no como check en el `.cls` del formulario.
- Sustituir el `OnTimer` con `m_CargaInicialIndicadoresPendiente` por un endpoint asÃ­ncrono o un skeleton explÃ­cito, no por un timer del cliente.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar `Me.OpenArgs` ni `DoCmd.OpenForm` como contrato de selecciÃ³n de NC de auditorÃ­a: la API REST debe recibir `IDNoConformidad` (de la NC de auditorÃ­a) en la URL.
- No duplicar la lÃ³gica de "esto es seguimiento de auditorÃ­a" en cada `.cls` de formulario: la web debe tener un Ãºnico discriminador de dominio y un Ãºnico guard.
- No usar la cinta (Ribbon) ni la visibilidad de menÃºs como control de seguridad real: la web debe aplicar permisos en el servidor.
- No migrar la combinaciÃ³n de `Form_FormNCAuditoriaAcciones` + `Form_FormNCAuditoriaAC` + `Form_FormNCAuditoriaAR` como tres UI distintas en la web si la lÃ³gica de negocio es la misma: la web puede tener un Ãºnico recurso `accion` con subtipo.
- No usar el helper `NCAuditoriaSeguimientoHelper` desde la capa de UI en la web: el helper debe ser consumido solo desde la capa de servicio, no por componentes.

### Â§6.4 Preguntas abiertas al product owner
- Â¿Los estados canÃ³nicos de una AC/AR/tarea de auditorÃ­a son los mismos que en proyecto o son especÃ­ficos de auditorÃ­a? (BR-NCA-AF-4) Confirmar lista y transiciones.
- Â¿La replanificaciÃ³n de una acciÃ³n de auditorÃ­a tiene lÃ­mite de veces o es indefinida? (BR-NCA-AF-4)
- Â¿Las notas de auditorÃ­a se pueden editar tras crear o son inmutables? (BR-NCA-AF-4) Confirmar polÃ­tica de retenciÃ³n.
- Â¿La cancelaciÃ³n de una acciÃ³n de auditorÃ­a requiere motivo obligatorio? Â¿Y la reasignaciÃ³n de propietario?
- Â¿La herencia de AC/AR desde proyecto a auditorÃ­a se permite o son siempre dominios disjuntos? Hoy se asume disjuntos, pero conviene confirmarlo.
- Â¿La fusiÃ³n de UI `Acciones` + `AC` + `AR` en un Ãºnico recurso es aceptable para el equipo de auditorÃ­a o se mantiene la separaciÃ³n por consistencia con proyecto?

## Â§7 Libro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-NCA-AF-1 â€” Las acciones/seguimiento de auditorÃ­a siguen siendo especÃ­ficas de dominio y nunca enrutan por estado de acciones de Proyecto. | Verified-runtime | `Issue38_SeguimientoAuditoria` 1/1, `Issue38_ResetearColTareas` 1/1 y slices de Issue #18 con Auditoria AC->NC / Auditoria AR hook; FALTA â†’ crear mediante access-vba-tdd para el ciclo completo de acciones | 2026-06-15 |
| BR-NCA-AF-2 â€” La selecciÃ³n de informe/listado de auditorÃ­a usa helpers de auditorÃ­a y NC de auditorÃ­a seleccionadas. | Verified-runtime | `ComandoInforme_Click` usa `EnsureNCAuditoriaGestionSelected`; `tests/tests.vba.audit-gestion-helper.json` pasÃ³ 11/11 tras el arreglo | 2026-06-15 |
| BR-NCA-AF-3 â€” Los indicadores compartidos pueden incluir filas de AuditorÃ­a, pero las lecturas runtime filtran por dominio/responsable. | Verified-runtime | `CacheIndicadoresAuditoriaMaterializado` 3/3; `CacheIndicadoresAuditoriaMaterializado_SincronizarDesdeNegocio` pasÃ³ dentro de slice 3/3; no afirmar suite completa `tests/tests.vba.indicadores-caracterizacion.json` verde | 2026-06-15 |
| BR-NCA-AF-4 â€” Las reglas de creaciÃ³n, vencimientos, finalizaciÃ³n, cancelaciÃ³n, replanificaciÃ³n, notas y asignaciÃ³n de propietario de acciones de auditorÃ­a son explÃ­citas. | Intended | FALTA â†’ crear mediante access-vba-tdd tras confirmar esquema/reglas | 2026-06-15 |
| BR-NCA-AF-5 â€” El comportamiento de formularios de seguimiento de auditorÃ­a permanece como cableado UI fino sobre costuras helper/servicio. | Intended | FALTA â†’ crear mediante access-vba-tdd contra costuras helper/servicio, no comportamiento directo de formulario | 2026-06-15 |
| Existen formularios/clases de acciones/seguimiento de auditorÃ­a. | Verified-static | Inventario de fuente de documentos existentes | 2026-06-15 |
| La seguridad de selecciÃ³n de informe/listado de auditorÃ­a tiene evidencia runtime. | Verified-runtime | `tests/tests.vba.audit-gestion-helper.json` 11/11; `ComandoInforme_Click` usa `EnsureNCAuditoriaGestionSelected` | 2026-06-15 |
| Existen slices de seguimiento/indicadores del lado AuditorÃ­a. | Verified-runtime | `Issue38_SeguimientoAuditoria` 1/1, `Issue38_ResetearColTareas` 1/1, `Issue18_ResolverNCDesde` 3/3 con Auditoria AC->NC, `Issue18_ARWriteHook` con hook Auditoria AR, `CacheIndicadoresAuditoriaMaterializado` 3/3 | 2026-06-15 |
| El ciclo de vida dedicado de acciones de auditorÃ­a estÃ¡ protegido para release. | Intended | Faltan pruebas/reglas | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Sin divergencia funcional confirmada. Hueco confirmado de evidencia: el inventario de fuente implica un comportamiento de seguimiento/acciones de auditorÃ­a mÃ¡s rico que lo cubierto por los tests actuales.
- No afirmar ciclo completo de acciones de auditorÃ­a hasta crear pruebas `access-vba-tdd` para crear/completar/cancelar/reasignar/replanificar/anotar con fixtures sandbox.
