# Índice de capacidades

> Fuente de verdad del comportamiento: código + verificación por Dysflow. Este índice es una guía de navegación; la evidencia vive en cada documento de capacidad, en los tests y en los artefactos SDD vinculados.

> **ÉPICA ACTIVA**: la cobertura strict TDD v2.4.2 de las 10 capabilities está en marcha. Roadmap y avance en [ÉPICA: Cobertura strict TDD v2.4.2](../epics/strict-tdd-coverage.md). Documento de capacidad = nivel de contrato de negocio; documento de épica = nivel de plan operacional.

## Política de confianza

- `Verified-runtime`: probado mediante `dysflow.test_vba` o verificación Dysflow vigente.
- `Verified-static`: confirmado por lectura de código o documentación técnica, pero sin prueba ejecutada en esta sesión.
- `Intended`: declarado por SDD, PRD o autoridad de producto, pendiente de confirmar en código.
- `Likely`: inferido por nombres, UI o estructura, pendiente de confirmar.
- `Divergent`: SDD/PRD y código no coinciden.

## Contrato TDD vigente

Las pruebas nuevas o migradas que sustenten capacidades deben cumplir `access-vba-tdd` v2.4.2: `Public Function` con retorno JSON canónico; fixture propio en sandbox y sin datos existentes por azar; revisión schema-first antes de sembrar; inyección explícita de `DAO.Database`; cardinalidad antes/después para mutaciones; manifests Dysflow con procedimiento global único, no calificado; cero UI/`Debug.Print`; cero mutación de `TbConfiguracionBackends`. Si se prueba comportamiento de formulario, la lógica debe extraerse detrás de helper, servicio o ViewModel para que el test no dependa de controles Access salvo una verificación UI explícita y justificada.

La deuda transversal y el orden de migración están centralizados en [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). Este informe prevalece sobre descripciones resumidas en documentos de capacidad.

## Referencia Phase 0 — formularios finos E2E

- Auditoría viva: [Auditoría Phase 0 — formularios finos E2E](./audit-e2e-thin-forms-phase-0.md).
- Cambio SDD: `phase-0-e2e-thin-forms-consolidation`.
- Estado de PCSUB helper/form coverage en el `staging` actual: `Divergent`. Los commits `b946c15^..59b4438` de `feature/pcsub-workflow-helper-coverage` no son ancestros de `staging`, por lo que sus helpers, formularios, tests y manifest son candidatos de adopción futura, no evidencia runtime vigente.
- Deuda Tier 3: los subformularios `*_Propuesta` y `*_Impacto` se aceptan solo como adaptadores finos documentados; no cierran cobertura por sí mismos y dependen de extraer el formulario padre a helper testeable.
- Gaps de harness: CDCA mantiene harness legacy; PC no tiene manifest atómico propio; CDCASUB y Workflow tienen evidencia runtime parcial; PCSUB tiene servicio/repositorio con evidencia histórica, pero la cobertura de helper/formulario actual está divergente.
- Bloqueo UAT: no se debe generar HTML UAT para los casos de la matriz hasta que cada caso tenga átomo verde en `staging`, manifest focal y `ref` firmado.
- Contrato de extracción futura: antes de tocar `src/` debe existir RED en `src/modules/Test_*.bas` y manifest focal `tests/*.json`; módulos/clases cierran con importación, compilación Dysflow, `test_vba`, `verify_code` y ledger; formularios/reportes requieren además compilación manual en VBE confirmada por el usuario antes de ejecutar `test_vba`.

## Registro maestro

| ID capacidad | Nombre | Dominio | Tier | Estado | Source | Confianza global | ¿Pruebas en verde? | Última release de producción | Documento |
|---|---|---|---|---|---|---|---|---|---|
| CAP-001 | Gestión de Propuestas de Cambio de Subcontratista (PCSUB) | PCSUB / workflow técnico | critical | active | hybrid | `Verified-runtime` histórico para BR-001..BR-011 servicio/repositorio; `Divergent` para helper/form coverage del `staging` actual según Phase 0 WU2 | 29 atómicos strict TDD v2.4.2 históricos; helper/form candidates no integrados en `staging` | Pendiente de confirmación | [CAP-001-pcsub.md](./CAP-001-pcsub.md) |
| CAP-002 | Generación de documentos Word y validación de mapeos de plantilla | Documentos Word / mapeos | critical | active | hybrid | mixta: `Verified-runtime` para validación estricta de campo ausente; barrido completo bloqueado por timeout Word | Parcial: `Test_DTM_StrictMissingFieldValidation` OK; `Test_DTM_AllMappedFields` timeout 300 s | Pendiente de confirmación | [CAP-002-documentos-word-mapeo.md](./CAP-002-documentos-word-mapeo.md) |
| CAP-003 | Captura y ciclo de vida de CD/CA | CDCA / workflow técnico | critical | active | hybrid | mayoritariamente `Verified-static`; smoke runtime diagnóstico 3/3; batería strict y harness v2.4.2 pendientes; Phase 0 registra Tier 1/2/3 y bloqueo UAT | Smoke `tests.cdca.smoke.json` 3/3 OK; `tests.cdca.json` no reclamado como strict | Pendiente de confirmación | [CAP-003-cdca.md](./CAP-003-cdca.md) |
| CAP-004 | Gestión de Propuestas de Cambio (PC) | PC / workflow técnico | critical | active | hybrid | `Verified-runtime` parcial con 3 átomos; code-behind sincronizado, layout de seis subformularios con deriva accionable; Phase 0 registra Tier 1/2/3, deuda de seam y bloqueo UAT | 3 átomos strict TDD v2.4.2 en verde por átomo individual; falta manifest PC de formulario/helper | Pendiente de confirmación | [CAP-004-pc.md](./CAP-004-pc.md) |
| CAP-005 | Gestión de Cambios de Diseño/Alcance de Subcontratista (CDCASUB) | CDCASUB / workflow técnico | critical | active | hybrid | `Verified-runtime` parcial para 3 átomos; `Verified-static` para guardados/transiciones; Phase 0 registra Tier 1/2/3 y bloqueo UAT | `tests/testsCdcasub.json` 3/3 parcial; faltan guardados, transiciones y helper/form | Pendiente de confirmación | [CAP-005-cdcasub.md](./CAP-005-cdcasub.md) |
| CAP-006 | Ciclo de vida, búsqueda y navegación de solicitudes | Solicitudes / búsqueda | critical | active | hybrid | mayoritariamente `Verified-static`; code-behind sincronizado, 9 layouts de formularios con deriva accionable | No reclamada; no hay manifest atómico propio aún | Pendiente de confirmación | [CAP-006-ciclo-vida-busqueda-navegacion.md](./CAP-006-ciclo-vida-busqueda-navegacion.md) |
| CAP-007 | Workflow y precondiciones por estado | Workflow transversal | critical | active | hybrid | `Verified-runtime` parcial para repositorio de log; `Verified-static` para motor de transición; Phase 0 bloquea UAT hasta átomo verde por transición/precondición | `tests/testsWorkflow.json` 3/3 parcial; faltan `EjecutarTransicion` y `ReabrirSolicitudCerrada` | Pendiente de confirmación | [CAP-007-workflow.md](./CAP-007-workflow.md) |
| CAP-008 | Gestión de adjuntos y archivo de documentación | Adjuntos | critical | active | hybrid | `Verified-runtime` parcial para la slice de servicio/repositorio de `AdjuntosServicio`: `GuardarAdjuntoDesdeArchivo` (validación + dedup + sad PDF-only + sad idSolicitud<=0 + sad ruta vacía), `ActualizarFicheroAdjunto` (reemplazo físico preservando fila + sad paths idAdjunto<=0/id inexistente), `SubirYCerrar` (seam de callback atómica BR-009 extraído en Slice A5 2026-06-15), `EliminarAdjunto`, `EliminarAdjuntosPorEtapas`; quedan BR-004 reglas 3-5 (256+ chars, etapaWF, usuarioSubida, fechaSubida), BR-005 (transacción), BR-006 (rollback persistencia) y el happy path de la callback atómica (A5b). | 12 átomos strict TDD v2.4.2 en verde por átomo individual (`fc133ad`, `60ea93a`, `80be457`, `?A5`); batch único ≥ 9 cae en `VBA_MANAGER_TIMEOUT` (~29 s) por suma de `durationMs` ≈ 28 s | Pendiente de confirmación | [CAP-008-adjuntos.md](./CAP-008-adjuntos.md) |
| CAP-009 | Logs, auditoría y snapshots | Auditoría | critical | active | hybrid | mayoritariamente `Verified-static`; código sincronizado, `SnapshotServicio` con `getdb()` directo y formularios de logs con deriva accionable | No reclamada; no hay manifest atómico propio aún | Pendiente de confirmación | [CAP-009-logs-auditoria.md](./CAP-009-logs-auditoria.md) |
| CAP-010 | Seguridad, administración e infraestructura | Seguridad / infra | critical | active | hybrid | mayoritariamente `Verified-static`; fuente↔binario sincronizado, identidad local sin IdP, `TbConfiguracionBackends` auditada en frontend | No reclamada; no hay manifest atómico propio aún | Pendiente de confirmación | [CAP-010-seguridad-administracion.md](./CAP-010-seguridad-administracion.md) |

## Taxonomía inicial pendiente de documentación

> Estado tras el cierre del slice 2026-06-15: las 10 capacidades candidatas de la taxonomía original ya tienen documento de capacidad dedicado (CAP-001 a CAP-010). La tabla se conserva como registro histórico; no quedan dominios pendientes de primera documentación en este pase. Futuros refinamientos, extensiones o sub-capacidades nuevas deben abrir su fila aquí antes de generar un nuevo `CAP-XXX`.

## Lagunas de cobertura abiertas

| Capacidad | Regla / área | Confianza actual | Acción |
|---|---|---|---|
| CAP-001 | Reglas UI/formulario de PCSUB no cubiertas por los 29 atómicos de servicio/repositorio | Verified-static / deuda UI | Extraer seams para apertura, navegación y cierre formalización; después ejecutar pruebas Dysflow focales sin depender de controles Access como contrato de negocio. |
| CAP-001 | Helper/form coverage PCSUB de `feature/pcsub-workflow-helper-coverage` | Divergent / Phase 0 | No usar como evidencia runtime de `staging`; adoptar en slice futura con RED, importación, compile y pruebas focales. |
| CAP-001/003/004/005 | Subformularios Tier 3 (`Propuesta`/`Impacto`) | Verified-static / deuda Phase 0 | Mantener como adaptadores finos solo si el formulario padre tiene helper y átomo verde; no generar UAT propio sin `ref` firmado. |
| CAP-001 | Reglas de workflow de UI en `Form_frmDatosPCSUB.GuardarDesdeSubform` | Verified-static | Crear pruebas o procedimiento de verificación Dysflow/UI que cubra promoción a Desarrollo Técnico, navegación técnica y envío a Calidad. |
| CAP-001 | Gaps caracterizados de completitud RAC, aprobación suministrador y decisión final | Verified-static / divergencia funcional documentada como gap | Decidir con producto si el comportamiento actual es válido; si no, crear SDD/fix + pruebas fixture-first. |
| CAP-001 | Mapeos ViewModel↔Entidad y dependencias de expediente/proveedor | Verified-static parcial | Crear pruebas puras de mapeo y pruebas con fixture ampliada de expediente/proveedor. |
| CAP-002 | Contrato de plantilla Word vs `tbMapeoCampos` | Verified-static / runtime parcial | `Test_DTM_StrictMissingFieldValidation` está en verde; dividir/instrumentar `Test_DTM_AllMappedFields` por plantilla porque agotó 300 s y dejó Word huérfano. |
| CAP-002 | Generación con datos de negocio y PDF | Intended / deuda | Crear E2E con fixture legal que pruebe valores sembrados en DOCX y cierre de Word/PDF sin residuos. |
| CAP-003 | Harness CDCA legacy (`SuiteSetup`/`SuiteTeardown`, `ForceLocalBackend`, llamadas con `Nothing`) | Divergent / deuda v2.4.2 | Migrar a `BeginTestSession`/`EndTestSession`, `DAO.Database` explícito, precheck sandbox endurecido y cardinalidad de mutaciones. |
| CAP-003 | Contrato de completitud RAC y decisión final | Divergent | Confirmar con producto si deben exigirse `racDecision` y `NombreFirmanteFinal`; después alinear código y tests. |
| CAP-004 | Gaps de completitud de Decisión Final y Dictamen RAC (omiten `NombreFirmanteFinal`/`racDecision`) | Divergent / deuda | Mismo hallazgo que CAP-003; cross-ref con BR-007/BR-008. Confirmar con producto y aplicar fix o documentar como comportamiento aceptado. |
| CAP-004 | `EsSolicitudEnValidacion` (helper privado) rompe la inyección de `DAO.Database` | Verified-static / deuda arquitectónica | Mover a helper inyectable antes de cualquier prueba focal de la rama transaccional de Impacto. |
| CAP-004 | `Form_frmDatosPC.GuardarDesdeSubform` mezcla UI y transacciones | Verified-static / deuda de seam | Extraer la decisión de navegación y la mini-transacción del RAC a un helper/servicio testeable sin controles Access. |
| CAP-004 | Layout fuente↔binario de seis subformularios PC | Divergent / blocker UI | Revisar `Form_subfrmDatosPC_*.form.txt` (`bothChanged`) antes de importar/exportar; code-behind `.cls` ya está sincronizado. |
| CAP-005 | RAC delegado (campos `racDelegado*`) no implementado; cross-ref con CAP-001 §7 | Divergent | Decidir con producto el contrato real: el esquema `tbDatosCDCASUB` ya tiene `racNombreDelegador` y `observacionesRACDelegador`; ampliar cobertura o documentar como fuera de alcance. |
| CAP-005 | `EsMotivosCompleto` exige `descripcionImpactoNCCont` (overflow manual) | Verified-static / alinear con CAP-002 | Revisar si debe armonizarse con la política de overflow `_extN`/`Cont` de la generación documental. |
| CAP-005 | Layout fuente↔binario de seis subformularios CDCASUB | Divergent / blocker UI | Revisar `Form_subfrmDatosCDCASUB_*.form.txt` (`bothChanged`) antes de importar/exportar; code-behind `.cls` ya está sincronizado. |
| CAP-006 | Contraseña `dpddpd` embebida en `frmAltaSolicitud.cls` línea 475 | Verified-static / deuda | Migrar a `GetPasswordDB()` para alinearse con el patrón canónico. |
| CAP-006 | Subquery de responsable técnico usa `FIRST` (Access-only) | Verified-static / riesgo de portabilidad | Considerar `MIN`/`MAX` con `TOP 1` ordenado, o factorizar el responsable en una capa de servicio. |
| CAP-006 | Sin manifest atómico para búsqueda/alta con fixture propio | Divergent / pendiente | Crear `tests/tests.lifecycle.json` con `BeginTestSession`/`EndTestSession`, `DAO.Database` explícito y cardinalidad de mutaciones. |
| CAP-006 | Layout fuente↔binario de formularios de navegación/alta/búsqueda | Divergent / blocker UI | Revisar 9 `.form.txt` con `bothChanged` antes de importar/exportar; code-behind `.cls` ya está sincronizado. |
| CAP-007 | `WorkflowServicio.EjecutarTransicion` activa `On Error Resume Next` para notificaciones y enmascara el error tras `CommitTrans` | Verified-static / política a revisar | Confirmar con producto si la transición debe fallar si el correo falla o si la política actual es aceptable. |
| CAP-007 | `ReabrirSolicitudCerrada` hace `If db Is Nothing Then Set db = getdb()` | Verified-static / deuda | Migrar a `db` siempre inyectado para que las pruebas puedan inyectar sandbox. |
| CAP-007 | Sin manifest atómico de workflow | Divergent / pendiente | Crear `tests/tests.workflow.json` con cardinalidad sobre `tbLogEstados`/`tbSolicitudes`. |
| CAP-007 | Casos UAT de transición/precondición sin átomo de `EjecutarTransicion` | Verified-static / Phase 0 | Bloquear HTML UAT hasta que cada caso tenga átomo verde, manifest focal y `ref` firmado. |
| CAP-007 | Layout fuente↔binario de formularios de log workflow | Divergent / blocker UI | Revisar `Form_frmVerLogCambios.form.txt` y `Form_frmDetalleLogCambio.form.txt` (`bothChanged`) antes de importar/exportar. |
| CAP-008 | Callback atómica de aprobación al subir PDF de cierre acoplada a `frmGestionAdjuntos` | Verified-static / deuda UI | Extraer a `AdjuntosServicio.SubirYCerrar(idSolicitud, rutaPDF)` para seam testeable. |
| CAP-008 | Lógica de eliminación de `Documento Final Firmado` mezcla formulario y servicios | Verified-static / deuda UI | Extraer a un helper. |
| CAP-008 | Sin manifest atómico de adjuntos | Divergent / pendiente | Crear `tests/tests.adjuntos.json` con cardinalidad sobre `tbAdjuntos` y verificación de rollback de archivo. |
| CAP-008 | Layout fuente↔binario de formularios de adjuntos | Divergent / blocker UI | Revisar `Form_frmGestionAdjuntos.form.txt` y `Form_frmElegirEtapaAdjunto.form.txt` (`bothChanged`) antes de importar/exportar. |
| CAP-009 | `SnapshotServicio.GenerarSnapshotGlobal` lee con `getdb()` directamente (seam roto) | Verified-static / deuda | Mover a un repositorio inyectable. |
| CAP-009 | `LogServicio.RegistrarError` cae a `Debug.Print` si la BD falla | Verified-static / política a revisar | Confirmar con producto si la app debe abortar o continuar sin log persistente. |
| CAP-009 | Sin manifest atómico de logs | Divergent / pendiente | Crear `tests/tests.logs.json` con cardinalidad sobre `tbLogCambios`/`tbLogErrores`/`tbLogEstados`. |
| CAP-009 | Layout fuente↔binario de formularios de logs | Divergent / blocker UI | Revisar cuatro `.form.txt` de logs (`bothChanged`) antes de importar/exportar. |
| CAP-009 | `LogEstadoServicio.cls` listado en documentación previa | Divergent / documentación corregida | No existe en `src/` ni en el proyecto VBA; la capacidad usa `LogEstado` + `LogEstadoRepositorio`. |
| CAP-010 | `TbConfiguracionBackends` vive en el frontend, no en el backend activo | Verified-static / deuda de testabilidad | Ya auditado con Dysflow; crear prueba focal de conmutación sin mutar configuración productiva. |
| CAP-010 | Identidad local basada en `Wscript.Network.UserName` | Verified-static / no apta para web | Planificar IdP corporativo (Kerberos/SAML/OAuth) en la migración. |
| CAP-010 | `Factoria.CreateEntity` con lista cerrada | Verified-static / acoplamiento central | Considerar un registro declarativo. |
| CAP-010 | Sin manifest atómico de seguridad | Divergent / pendiente | Crear `tests/tests.security.json` con cardinalidad sobre `TbUsuariosAplicacionesPermisos` y la matriz de roles. |

## Deuda de pruebas v2.4.2 transversal

Ver el informe central [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). Regla operativa: las suites existentes no deben asumirse como nueva evidencia `Verified-runtime` hasta completar la migración indicada; el runner probe es diagnóstico; el E2E documental es candidato read-only, no evidencia de mutación.

## Divergencias pendientes de revisión humana

| Capacidad | Hallazgo | Detectada | Estado |
|---|---|---|---|
| CAP-001 | La spec `pcsub-nueva-solicitud` pretendía RAC delegado con campos `racDelegado*`; el código actual de persistencia usa campos `racCodigo`, `racDecision`, `racNombre` y campos delegador parciales (`observacionesRACDelegador`, `racNombreDelegador`). | 2026-06-15 | **Pendiente** (no resuelto en esta sesión) |
| CAP-001 | ~~Las pruebas de caracterización documentan funciones de completitud que pueden devolver `True` con campos relevantes en blanco (`racDecision`, `NombreFirmanteFinal`, u otros según caso).~~ | 2026-06-15 | **Resuelto 2026-06-18** (commit pendiente): `EsDictamenRACCompleto` ahora exige `racDecision`; `EsAprobacionSuministradorCompleta` ahora exige `decisionFinal`; `EsDecisionFinalCompleta` ahora exige `NombreFirmanteFinal`. Atomos `Verified-runtime` con tests `Test_PCSUB_Es*Blank*_ReturnsFalse`. |
| CAP-001 | `verify_binary` detectó deriva accionable: `Form_frmDatosPCSUB.cls` está `bothChanged` y `Form_subfrmDatosPCSUB_DecisionFinal.cls` está `sourceNewer`; por tanto las afirmaciones dependientes de formulario no deben tratarse como cierre runtime completo. | 2026-06-15 | **Pendiente** (no resuelto en esta sesión) |
| CAP-003 | ~~Los tests candidatos CDCA esperan que `EsDictamenRACCompleto` falle sin `racDecision`, pero el código actual solo exige `racCodigo` y `racNombre`.~~ | 2026-06-15 | **Resuelto 2026-06-18** (commit pendiente): `DatosCDCAServicio.EsDictamenRACCompleto` ahora exige `racDecision`. |
| CAP-003 | ~~Los tests candidatos CDCA esperan que `EsDecisionFinalCompleta` falle sin `NombreFirmanteFinal`, pero el código actual solo exige `decisionFinal`.~~ | 2026-06-15 | **Resuelto 2026-06-18** (commit pendiente): `DatosCDCAServicio.EsDecisionFinalCompleta` ahora exige `NombreFirmanteFinal`. Idem `EsAprobacionSuministradorCompleta` ahora exige `decisionFinal`. |
| CAP-004 | ~~`EsDecisionFinalCompleta` (gemelo PC) y `EsDictamenRACCompleto` (gemelo PC) omiten `NombreFirmanteFinal` y `racDecision` respectivamente. Mismo patrón que CAP-003.~~ | 2026-06-15 | **Resuelto 2026-06-18** (commit pendiente): `DatosPCServicio.EsDictamenRACCompleto` ahora exige `racDecision`; `DatosPCServicio.EsDecisionFinalCompleta` ahora exige `NombreFirmanteFinal`; `DatosPCServicio.EsAprobacionSuministradorCompleta` ahora exige `decisionFinal`. |
| CAP-004 | `DatosPCServicio.EsSolicitudEnValidacion` (helper privado) salta a `getdb()` cuando el llamador no pasa `db`, rompiendo la inyección de `DAO.Database` para pruebas. | 2026-06-15 | **Pendiente** (no resuelto en esta sesión) |
| CAP-005 | `DatosCDCASUBServicio.EsDecisionFinalCompleta` y `EsDictamenRACCompleto` replican la omisión de `NombreFirmanteFinal` y `racDecision`. Mismo patrón que CAP-001/003/004. | 2026-06-15 | **Resuelto 2026-06-18** (commit pendiente): `DatosCDCASUBServicio.EsDictamenRACCompleto` ahora exige `racDecision`; `DatosCDCASUBServicio.EsDecisionFinalCompleta` ahora exige `NombreFirmanteFinal`; `DatosCDCASUBServicio.EsAprobacionSuministradorCompleta` ahora exige `decisionFinal`. |
| CAP-005 | El esquema `tbDatosCDCASUB` tiene `racNombreDelegador` y `observacionesRACDelegador` pero el resto del contrato RAC delegado (campos `racDelegado*`) no se persiste. Cross-ref con CAP-001. | 2026-06-15 |
| CAP-006 | `frmAltaSolicitud.cls` línea 475 embebe la contraseña `dpddpd` en lugar de leerla de `GetPasswordDB()`. Migrar al patrón canónico. | 2026-06-15 |
| CAP-006 | `SolicitudServicio.getSolicitudesViewModel` usa `FIRST(...)` en la subquery de responsable técnico, que es Access-only. Riesgo de portabilidad a web/SQL estándar. | 2026-06-15 |
| CAP-007 | `WorkflowServicio.ReabrirSolicitudCerrada` hace `If db Is Nothing Then Set db = getdb()` (línea 1954), rompiendo la inyección de `db` para pruebas. | 2026-06-15 |
| CAP-008 | `frmGestionAdjuntos.EjecutarFlujoSubida` contiene una callback atómica que aprueba la solicitud al subir el PDF de cierre; debería extraerse a `AdjuntosServicio.SubirYCerrar`. | 2026-06-15 |
| CAP-009 | `SnapshotServicio.ObtenerDatosTabla` lee con `getdb()` directamente (línea 68), rompiendo la inyección de `db` para pruebas. | 2026-06-15 |
| CAP-010 | El esquema de `TbConfiguracionBackends` no se ha auditado con `dysflow.list_tables` en esta tarea. Pendiente confirmar contrato. | 2026-06-15 |
| CAP-010 | `Factoria.CreateEntity` con lista cerrada de tipos: añadir un nuevo tipo de entidad requiere editar `Factoria.bas`. Considerar registro declarativo. | 2026-06-15 |
