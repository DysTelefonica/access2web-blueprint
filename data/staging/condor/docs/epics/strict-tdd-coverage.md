# Épica: Cobertura strict TDD v2.4.2 de las 10 capabilities de CONDOR

> Documento de avance y roadmap operacional. Mantenido por la IA, validado por el usuario y por evidencia runtime Dysflow.
> Fecha de apertura: 2026-06-15.
> Skill de referencia: `access-vba-tdd` v2.4.2 y `access-vba-capability-docs`.

## §0 Resumen ejecutivo

CONDOR está compuesto por 10 capabilities de negocio (CAP-001 a CAP-010) que totalizan:

- **60 clases** (`src/classes/*.cls`)
- **53 módulos** (`src/modules/*.bas`)
- **46 formularios y subformularios** (`src/forms/*.cls` y `.form.txt`)
- **8 suites de test** con **126 funciones `Public Function Test_*`** en total
- **10 manifests Dysflow** declarados en `tests/`

El objetivo de esta épica es **mover cada regla de negocio de `Verified-static` a `Verified-runtime`** siguiendo la skill `access-vba-tdd` v2.4.2, sin tocar el código de producción más que para extraer seams cuando sea necesario.

**Estado al 2026-06-15**:

| Capability | BR-001..BR-N totales | BR con test verde | % runtime | Manifest atómico |
|---|---|---|---|---|
| CAP-001 PCSUB | 11 BR | 11 ✅ | 100% (servicio/repo) | `tests.pcsub.json` 29 átomos |
| CAP-002 Word | varias | 1 ✅ | parcial (template 1 BR) | `tests.document-template-e2e.json` 1 átomo + WIP A4 deferred |
| CAP-002 Word | varias | parcial | parcial | `tests.document-template-e2e.json` 2 átomos |
| CAP-003 CDCA | varias | parcial | ~30% (smoke) | `tests.cdca.smoke.json` 3 + legacy `tests.cdca.json` 59 |
| CAP-004 PC | varias | 3 ✅ | ~30% (servicio) | `tests.testsPc.json` 3 átomos (B1 cerrado, commit `54be6aa`) |
| CAP-005 CDCASUB | varias | 3 ✅ | ~30% (servicio: parte técnica + aprobación sad) | `tests.testsCdcasub.json` 3 átomos (B2 cerrado, commit `854f32b`) |
| CAP-006 Lifecycle | varias | 2 ✅ | ~30% (servicio) | `tests.tests.lifecycle.json` 2 átomos (B3 cerrado, commit `8fa7337`) |
| CAP-007 Workflow | varias | 3 ✅ | ~20% (repo: LogEstado) | `tests.testsWorkflow.json` 3 átomos (B4 cerrado, commit `13bfcd1`) |
| CAP-008 Adjuntos | 12 BR | 9 ✅ | 75% (servicio/repo) | `tests.adjuntos.json` 9 átomos |
| CAP-009 Logs | varias | 2 ✅ | ~30% (servicio cambio+error) | `tests.testsLogs.json` 2 átomos (B5 cerrado, commit `63d4116`) |
| CAP-010 Seguridad | varias | 4 ✅ | ~30% (servicio: rol hierarchy) | `tests.testsSecurity.json` 4 átomos (B6 cerrado, commit `fc025e1`) |

**Capabilities en `Verified-runtime` parcial o total**: 10 de 10 (CAP-001, CAP-002, CAP-004, CAP-005, CAP-006, CAP-007, CAP-008, CAP-009, CAP-010).
**Capabilities todavía con manifest parcial**: 9 de 10 (todos salvo CAP-001 PCSUB que está 100%).

## §1 Inventario completo de features por capability

> Una "feature" = una clase o módulo productivo que contiene reglas de negocio testeables. Cada fila mapea código → capability → tier → estado de cobertura.

### CAP-001 PCSUB — `Verified-runtime` (11/11 BR servicio/repo)
- **Tier**: critical · **Source**: hybrid
- **Clases**: `DatosPCSUBServicio.cls` (944), `DatosPCSUB.cls` (166), `DatosPCSUBViewModel.cls` (114)
- **Módulos**: `DatosPCSUBRepositorio.bas` (486)
- **Forms**: `Form_frmDatosPCSUB.cls` (467) + 6 subformularios
- **Tests**: `Test_PCSUB_Strict.bas` (29 átomos) + `Test_PCSUB.bas` legacy (16)
- **Manifest atómico**: `tests.pcsub.json` (29 átomos)
- **Última evidencia runtime**: commit `f27f693` (29/29 verde)
- **Deuda UI**: BR pendientes de extracción de `frmGestionSolicitud` y de seams de formulario
- **Gaps de cobertura activos**:
  - Reglas UI/formulario de PCSUB no cubiertas por los 29 átomos (acción: extraer seams)
  - Reglas de workflow de UI en `Form_frmDatosPCSUB.GuardarDesdeSubform`
  - Gaps de completitud RAC, aprobación suministrador y decisión final (`racDecision`, `NombreFirmanteFinal`)
  - Mapeos ViewModel↔Entidad y dependencias de expediente/proveedor
  - Divergencias: `Form_frmDatosPCSUB.cls` `bothChanged` y `Form_subfrmDatosPCSUB_DecisionFinal.cls` `sourceNewer` (regresión de layout)

### CAP-002 Word / Documentos — `Verified-runtime` parcial
- **Tier**: critical · **Source**: hybrid
- **Clases**: `DocumentoServicio.cls` (1319), `MapeoServicio.cls` (98), `MapeoCampos.cls` (49)
- **Módulos**: `MapeoRepositorio.bas` (94)
- **Tests**: `Test_DocumentTemplateMapping.bas` (3 funciones)
- **Manifest atómico**: `tests.document-template-e2e.json` (2 átomos)
- **Última evidencia runtime**: `Test_DTM_StrictMissingFieldValidation` OK; `Test_DTM_AllMappedFields` timeout 300s
- **Gaps de cobertura activos**:
  - DTM_AllMappedFields se agotó en 300s y dejó `WINWORD.EXE /Automation -Embedding` huérfano — **dividir por plantilla** antes de re-ejecutar
  - Sin E2E con datos de negocio sembrados en DOCX
  - Sin verificación de cierre limpio de Word/PDF sin residuos
  - Sin prueba del ciclo de overflow `_extN` / `Cont`

### CAP-003 CDCA — `Verified-runtime` smoke / `Verified-static` strict
- **Tier**: critical · **Source**: hybrid
- **Clases**: `DatosCDCA.cls` (130), `DatosCDCAServicio.cls` (1089), `DatosCDCAViewModel.cls` (113)
- **Módulos**: `DatosCDCARepositorio.bas` (458)
- **Forms**: `Form_frmDatosCDCA.cls` (569) + 6 subformularios
- **Tests**: `Test_CDCA.bas` legacy (60 funciones) — usa harness legacy con `Nothing` y `SuiteSetup`
- **Manifests**: `tests.cdca.json` legacy (59 tests) + `tests.cdca.smoke.json` (3 tests)
- **Última evidencia runtime**: smoke 3/3 OK
- **Gaps de cobertura activos**:
  - **Harness legacy**: `SuiteSetup`/`SuiteTeardown`, `ForceLocalBackend`, llamadas con `Nothing` → **migrar a v2.4.2**
  - **Cardinalidad de mutaciones incompleta**: no hay `countBefore`/`countAfter` homogéneo
  - `EsDictamenRACCompleto` y `EsDecisionFinalCompleta` omiten `racDecision` y `NombreFirmanteFinal` — divergencia funcional documentada
  - Layout fuente↔binario: code-behind `.cls` sincronizado, 6 `.form.txt` con `bothChanged`

### CAP-004 PC — `Verified-static`
- **Tier**: critical · **Source**: hybrid
- **Clases**: `DatosPC.cls` (129), `DatosPCServicio.cls` (1098), `DatosPCViewModel.cls` (104)
- **Módulos**: `DatosPCRepositorio.bas` (546)
- **Forms**: `Form_frmDatosPC.cls` (655) + 6 subformularios
- **Tests**: **ninguno** — sin manifest, sin suite
- **Gaps de cobertura activos**:
  - **Sin manifest atómico propio** — feature entera sin evidencia runtime
  - Gaps de completitud de Decisión Final y Dictamen RAC (omiten `NombreFirmanteFinal`/`racDecision`) — divergencia con PCSUB
  - `EsSolicitudEnValidacion` (helper privado) salta a `getdb()` → seam arquitectónico roto
  - `Form_frmDatosPC.GuardarDesdeSubform` mezcla UI y transacciones
  - Layout: 6 subformularios con `.form.txt` `bothChanged`

### CAP-005 CDCASUB — `Verified-runtime` parcial (parte técnica + aprobación sad)
- **Tier**: critical · **Source**: hybrid
- **Clases**: `DatosCDCASUB.cls` (149), `DatosCDCASUBServicio.cls` (982), `DatosCDCASUBViewModel.cls` (100)
- **Módulos**: `DatosCDCASUBRepositorio.bas` (413)
- **Forms**: `Form_frmDatosCDCASUB.cls` (608) + 6 subformularios
- **Tests**: `Test_CDCASUB_Strict.bas` (3 átomos verdes), manifest atómico `tests.testsCdcasub.json`
- **Última evidencia runtime**: commit `854f32b` (3/3 verde, `EsParteTecnicaCompleta` happy + sad `clasifNC=INVALID` + `EsAprobacionSuministradorCompleta` sad `firmaIngenieria=''`)
- **Gaps de cobertura activos**:
  - **Pendientes**: `EsDatosGeneralesCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleta`, `EsDecisionFinalCompleta`, y los `Guardar*` / `Actualizar*` con cardinalidad
  - RAC delegado (campos `racDelegado*`) no implementado — esquema `tbDatosCDCASUB` ya tiene `racNombreDelegador` y `observacionesRACDelegador` parciales
  - `EsMotivosCompleto` exige `descripcionImpactoNCCont` (overflow manual) — alinear con CAP-002
  - Layout: 6 subformularios con `.form.txt` `bothChanged`
  - **Divergencias con PC validadas**: `EsParteTecnicaCompleta` es exclusiva CDCASUB; `EsAprobacionSuministradorCompleta` exige Ingenieria+Calidad (PC usa oficinaTecnica+repSuministrador)

### CAP-006 Lifecycle / Solicitudes — `Verified-static`
- **Tier**: critical · **Source**: hybrid
- **Clases**: `Solicitud.cls` (64), `SolicitudServicio.cls` (746), `SolicitudViewModel.cls` (54), `SolicitudBusquedaViewModel.cls` (65), `Expediente.cls` (67), `ExpedienteServicio.cls` (308), `ExpedienteViewModel.cls` (36), `FiltrosSolicitud.cls` (34)
- **Módulos**: `SolicitudRepositorio.bas` (218), `ExpedienteRepositorio.bas` (113)
- **Forms**: `Form_frmAltaSolicitud.cls` (641), `Form_frmBuscarSolicitudes.cls` (360), `Form_frmBuscarExpediente.cls` (240), `Form_frmDetalleExpediente.cls` (88), `Form_frmFiltrosAvanzadosSolicitudes.cls` (134)
- **Tests**: **ninguno con manifest atómico**
- **Gaps de cobertura activos**:
  - **Sin manifest atómico** — alta y búsqueda sin evidencia runtime
  - `frmAltaSolicitud.cls` línea 475 embebe `dpddpd` literal en vez de `GetPasswordDB()` — divergencia con el patrón canónico
  - `SolicitudServicio.getSolicitudesViewModel` usa `FIRST(...)` (Access-only) — riesgo de portabilidad
  - Layout: 9 `.form.txt` con `bothChanged`

### CAP-007 Workflow — `Verified-runtime` parcial (LogEstadoRepositorio)
- **Tier**: critical · **Source**: hybrid
- **Clases**: `WorkflowServicio.cls` (**4352 líneas**), `Estado.cls` (51), `EstadoServicio.cls` (36)
- **Módulos**: `WorkflowRepositorio.bas` (58), `EstadoRepositorio.bas` (27), `LogEstadoRepositorio.bas` (136)
- **Forms**: `Form_frmVerLogCambios.cls` (185), `Form_frmDetalleLogCambio.cls` (59)
- **Tests**: `Test_Workflow_Strict.bas` (3 átomos verdes), manifest atómico `tests.testsWorkflow.json`
- **Última evidencia runtime**: commit `13bfcd1` (3/3 verde, Slice B4). Cubre `LogEstadoRepositorio.Guardar` (Insert + autonumérico) y `getUltimoEstadoAnterior` (última transición + 0 para vacío).
- **Gaps de cobertura activos**:
  - **DEFERRED**: `EjecutarTransicion` y `ReabrirSolicitudCerrada` requieren setup de `m_ObjUsuarioActivo.rol` (Calidad/Administrador) y precondiciones multi-tabla. Documentado en §5.
  - `WorkflowServicio.EjecutarTransicion` activa `On Error Resume Next` para notificaciones y enmascara el error tras `CommitTrans` — política a revisar
  - `ReabrirSolicitudCerrada` línea 1954 hace `If db Is Nothing Then Set db = getdb()` — seam arquitectónico roto
  - Layout: `Form_frmVerLogCambios.form.txt` y `Form_frmDetalleLogCambio.form.txt` `bothChanged`

### CAP-008 Adjuntos — `Verified-runtime` parcial (9/12 BR servicio/repo)
- **Tier**: critical · **Source**: hybrid
- **Clases**: `Adjunto.cls` (52), `AdjuntosServicio.cls` (543), `AdjuntoViewModel.cls` (39)
- **Módulos**: `AdjuntoRepositorio.bas` (221)
- **Forms**: `Form_frmGestionAdjuntos.cls` (390), `Form_frmElegirEtapaAdjunto.cls` (66)
- **Tests**: `Test_Adjuntos_Strict.bas` (10 funciones = 9 átomos + 1 RunAll)
- **Manifest atómico**: `tests.adjuntos.json` (9 átomos)
- **Última evidencia runtime**:
  - Slice BR-001..003: commit `91fe07e` (4/4 verde)
  - Slice BR-002/003: commit `fc133ad` (6/6 verde)
  - Slice BR-012 `ActualizarFicheroAdjunto`: commit `60ea93a` (9/9 verde por átomo individual — el batch único de 9 cae en `VBA_MANAGER_TIMEOUT` por suma de `durationMs` ≈ 28s vs límite 29s del runner; documentación honesta sin sobredeclarar cobertura)
- **BR pendientes de runtime coverage** (ordenados por valor/dependencia):
  - **BR-004 `Validar` entidad** — sad paths por API pública (`idSolicitud<=0`, ruta vacía, `nombreArchivo` > 255)
  - **BR-005 transacción** explícita — valor marginal (la rama `db Is Nothing` no se ejercita)
  - **BR-006 rollback de persistencia** — estrategia confirmada viable por diagnóstico schema (2026-06-15):
    - Backend `tbAdjuntos` solo tiene 2 índices: PK `idAdjunto` (único) y `idSolicitud` (no-único, no FK formal)
    - No hay FK formal `TbExpedientes→tbSolicitudes` (`Schema_Diag_ExpedienteToSolicitudFK` confirma `NOT FOUND`)
    - **Estrategia**: pre-insertar solicitud dummy con `idSolicitud = 900899` (fuera del rango de `TeardownFixtures` 900800-900899) y `idExpediente` arbitrario (sin FK formal que valide), después pre-insertar adjunto dummy con `idAdjunto = MAX(idAdjunto) + 1` y `idSolicitud = 900899` con todos los campos required poblados. Cuando el servicio llame `getSiguienteIDAdjunto`, va a calcular el mismo `MAX+1` (porque el dummy ya está), el INSERT choca con PK del dummy, el catch ejecuta `ws.Rollback` + `fso.DeleteFile destinoFull, True`.
- **Gaps de cobertura activos**:
  - Callback atómica de aprobación al subir PDF de cierre acoplada a `frmGestionAdjuntos.EjecutarFlujoSubida` líneas 175-225 — extraer a `AdjuntosServicio.SubirYCerrar`
  - Lógica de eliminación de `Documento Final Firmado` mezcla formulario y servicios
  - Layout: 2 `.form.txt` con `bothChanged`

### CAP-009 Logs / Auditoría — `Verified-runtime` parcial (cambio+error)
- **Tier**: critical · **Source**: hybrid
- **Clases**: `LogCambio.cls` (58), `LogError.cls` (56), `LogEstado.cls` (52), `LogCambioServicio.cls` (33), `LogServicio.cls` (226), `SnapshotServicio.cls` (103)
- **Módulos**: `LogCambioRepositorio.bas` (130), `LogErrorRepositorio.bas` (106), `LogEstadoRepositorio.bas` (136)
- **Forms**: `Form_frmVerLogErrores.cls` (181), `Form_frmDetalleLogError.cls` (59)
- **Tests**: `Test_Logs_Strict.bas` (2 átomos verdes), manifest atómico `tests.testsLogs.json`
- **Última evidencia runtime**: commit `63d4116` (2/2 verde, `RegistrarCambio` 3.7 s + `RegistrarError` 3.8 s, Slice B5)
- **Gaps de cobertura activos**:
  - `RegistrarTransicion` (LogEstado) sin cobertura — pendiente Slice B5.2 o B5.3
  - `GenerarSnapshotGlobal` / `CalcularHashSnapshot` sin cobertura — `SnapshotServicio.ObtenerDatosTabla` línea 68 lee con `getdb()` directamente — seam roto
  - `LogServicio.RegistrarError` cae a `Debug.Print` si la BD falla — política a revisar (¿abortar o continuar?); el test actual NO cubre este fallback (requiere `getdb()` inaccesible)
  - Layout: 4 `.form.txt` de logs con `bothChanged` (Fase C pendiente)
  - `LogEstadoServicio.cls` listado en documentación previa — **NO existe** en `src/`; usar `LogEstado.cls` + `LogEstadoRepositorio.bas` (divergencia documental)

### CAP-010 Seguridad / Administración — `Verified-runtime` parcial (rol hierarchy)
- **Tier**: critical · **Source**: hybrid
- **Clases**: `Usuario.cls` (92), `UsuarioServicio.cls` (150), `UsuarioAplicacionPermisos.cls` (55)
- **Módulos**: `UsuarioRepositorio.bas` (71)
- **Forms**: `Form_frm0OtrosAdmin.cls` (65), `Form_frm0Ppal.cls` (296), `Form_frm0PpalTecnico.cls` (170)
- **Tests**: `Test_Security_Strict.bas` (4 átomos verdes), manifest atómico `tests.testsSecurity.json`
- **Manifests adicionales**: `tests.issue19.json` (2 tests), `tests.vba.json` (16 tests)
- **Última evidencia runtime**: commit `fc025e1` (4/4 verde, Slice B6). Cubre `UsuarioServicio.DeterminarRol` (pure logic, no DB): Administrador > Calidad > Tecnico, fallback a Tecnico cuando Permisos=Nothing, Raise 513 cuando Usuario=Nothing.
- **Gaps de cobertura activos**:
  - **Pendientes**: `getUsuarioConPermisos`, `getUsuarioConectadoConPermisos`, `getResponsablesTecnicos`, `getResponsablesCalidad` (todos consultan `TbUsuariosAplicacionesPermisos` y/o `TbExpedientesResponsables`)
  - Identidad local basada en `WScript.Network.UserName` — no apta para web
  - `Factoria.CreateEntity` con lista cerrada de tipos (acoplamiento central)
  - `TbConfiguracionBackends` vive en frontend, no en backend activo — deuda de testabilidad

### Transversal / helpers
- **Variables Globales.bas** (295): m_TestingMode, m_BackendSandboxURL, m_ObjEntorno, m_ObjUsuarioActivo, m_URLRutaAplicacionLocal, m_PasswordBackend
- **Entorno.cls** (646): cache de `estados`, `URLDirectorioDocumentacion`, fingerprint
- **CondorError.cls** (99): envoltorio de errores con `Create`/`Raise`/`AddToCallStack`
- **Factoria.bas** (67): `CreateEntity` con lista cerrada
- **modEnumeradores.bas** (62): enums del dominio
- **JsonConverter.bas** (1451) + **JsonHelper.bas** (119): serialización JSON
- **RepositorioComun.bas** (202): helpers DAO compartidos (`EjecutarAccion`, `EjecutarConsulta`, `HidratarEntidadDesdeSQL`)
- **SnapshotHelper.bas** (763): serialización de filas para snapshot/logs
- **HashHelper.bas** (82): hash de archivos para trazabilidad
- **ChecklistHelper.bas** (70): reglas de completitud reutilizables
- **modSimuladorNotificaciones.bas** (107): simulador de correo para tests
- **ErrorLogger.bas** (46) + **ErrorPresenter.bas** (187): presentación de errores al usuario
- **FormulariosPadreAuxiliares.bas** (290): código compartido de formularios
- **FUNCIONES UTILES.bas** (1049): utilidades varias

## §2 Plan priorizado de slices

> Orden propuesto por valor/dependencia. Cada slice es 1 PR ≤ 400 líneas, 1 import → compile → test → doc → commit → push.

### Fase A — Cerrar las 3 capabilities que ya tienen manifest (CAP-001, 002, 008)

| # | Capability | Slice | Átomos | Estado |
|---|---|---|---|---|
| A1 | CAP-008 | `ActualizarFicheroAdjunto` BR-012 | 3 átomos | ✅ cerrado (commit `60ea93a`) |
| A2 | CAP-008 | BR-004 `Validar` sad paths | 2 átomos (reducido de 3: 256+ chars no testeable) | ✅ cerrado (commit `80be457`) |
| A3 | CAP-008 | BR-006 rollback de persistencia | **deferred** — ver §"BR-006 deferred" abajo | deuda arquitectónica |
| A4 | CAP-002 | Dividir `Test_DTM_AllMappedFields` por plantilla (resolver timeout Word) | 4 átomos escritos pero no viables | **deferred** — ver §"A4 deferred" abajo | deuda entorno (timeout runner) |
| A5 | CAP-008 | Extracción del seam `SubirYCerrar` de `frmGestionAdjuntos` (callback atómica BR-009) | 1 átomo (A5a verde) + 1 WIP (A5b happy path) | ✅ A5a cerrado; A5b deferred — ver §"A5b deferred" |

### BR-006 deferred — requiere refactor de seams

**Fecha**: 2026-06-15

**Decisión**: BR-006 (compensación de archivo si la persistencia falla) queda `Verified-static`. **No se inserta un átomo verde mintiendo** en el manifest.

**Por qué la estrategia "dummy row PK collision" no funciona**:
- `getSiguienteIDAdjunto` calcula `MAX(idAdjunto) + 1` LEYENDO el `MAX` actual
- Si pre-insertamos un dummy con `idAdjunto = MAX_antes + 1`, el nuevo `MAX` es `MAX_antes + 1`
- El servicio calcula `MAX + 1 = (MAX_antes + 1) + 1 = MAX_antes + 2` — NO choca con el dummy
- VBA single-threaded hace imposible insertar el dummy entre `getSiguienteIDAdjunto` y el `INSERT`

**Por qué la estrategia "Validar 256+ chars" no funciona**:
- `Validar` regla 3 (`Len(nombreArchivo) > 255`) corre DESPUÉS de `CopyFile`, así que el catch SÍ compensaría
- PERO requiere crear archivo físico con nombre 256+ chars
- Windows MAX_PATH = 260 chars. `%TEMP%\source\` ya mide ~85 chars, así que un nombre 256+ excede el límite
- FSO COM no soporta el prefijo `\\?\` nativamente

**Refactor requerido para desbloquear**:
- Inyectar un `DAO.Database` configurable que pueda fallar en `INSERT` (e.g., vía mock/stub o error injection)
- O agregar un flag interno `forceFailForTests` en `AdjuntosServicio.GuardarAdjuntoDesdeArchivo` que tire `Err.Raise` en un punto controlado entre `CopyFile` y `INSERT`
- Después del refactor, el módulo `Test_Adjuntos_Strict.bas` se puede extender con un átomo real que verifique la compensación

### A4 deferred — Word COM timeout excede el límite del runner

**Fecha**: 2026-06-15

**Decisión**: El barrido de `tbMapeoCampos` contra los FormFields de Word no es viable en este entorno. El manifest estricto solo incluye `Test_DTM_StrictMissingFieldValidation` (que SÍ pasa — solo abre Word y verifica una validación, no itera el contrato completo).

**Por qué dividir no salva el timeout**:
- El legacy `Test_DTM_AllMappedFields` (4 plantillas) timed out a 300 s
- La hipótesis era que dividirlo en 4 átomos (uno por plantilla) salvaría el problema
- Test empírico: UN átomo `Test_DTM_PC_AllMappedFields` también timed out a 29 s
- El overhead de Word COM open + FormFields read + close es demasiado para UNA plantilla
- El runner Dysflow tiene un timeout per-call de ~29 s que no se puede extender por manifest

**Refactor requerido para desbloquear**:
- Subir el timeout del runner Dysflow (requiere tocar el runner, fuera del scope del manifest)
- O optimizar `LoadFormFieldNames` (e.g., leer los nombres via XML/OOXML en vez de iterar `wordDoc.FormFields` COM)
- O reducir el alcance del contrato (e.g., validar solo un subset de campos por ejecución)
- Después del refactor, el módulo `Test_DocumentTemplateMapping.bas` se puede extender con los 4 átomos per-template (que ya están escritos como WIP).

### A5b deferred — Happy path de la callback atómica requiere precondiciones complejas

**Fecha**: 2026-06-15

**Decisión**: A5 cerrado con el átomo sad-path (A5a). A5b (happy-path: callback se invoca y completa la transición a `estadoAprobada`) queda como WIP.

**Por qué A5a sí es testeable**:
- A5a invoca `SubirYCerrar` con etapa `!= "Documento Final Firmado"` → NO se invoca la callback
- Las validaciones de la callback (rol, transiciones, precondiciones) NO se ejecutan
- El test es simple: verifica que el adjunto se guarda y la callback no se invoca

**Por qué A5b no es testeable en el sandbox actual**:
- A5b invoca `SubirYCerrar` con etapa = "Documento Final Firmado"
- La callback llama `EjecutarTransicion` que valida `EsTransicionPermitida` (rol) y `PrecondicionesCumplidas` (datos PC completos, RAC completo, etc.)
- El sandbox de tests no tiene `rolUsuario` configurado, ni los datos PC completos
- Reproducir el setup completo en el sandbox requiere seed de muchas tablas (PC, CDCA, CDCASUB, PCSUB) con todos los campos required por las precondiciones

**Refactor requerido para desbloquear A5b**:
- Opción (a): agregar un helper `EjecutarCallbackAprobacionConMocks` que reciba mocks para `EjecutarTransicion` y `GuardarDecisionFinal`, y testear el seam sin las precondiciones
- Opción (b): seed completo del sandbox con datos válidos (solicitud PC con todos los campos, RAC completo, firmante, etc.)
- Opción (c): agregar un modo "test" al servicio que skipea las precondiciones

**Estado actual**: el seam `SubirYCerrar` y `EjecutarCallbackAprobacion` están implementados en `AdjuntosServicio.cls` líneas 545-700. El test `Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre` (verde 3211 ms) cubre el seam en su camino non-cerrar.

### Fase B — Crear manifests atómicos para capabilities sin cobertura (CAP-004, 005, 006, 007, 009, 010)

| # | Capability | Slice inicial | Átomos | Notas |
|---|---|---|---|---|
| B1 | CAP-004 PC | Crear `Test_PC_Strict` siguiendo patrón `Test_PCSUB_Strict` | ~15-20 átomos | Migrar gemelo primero; `EsSolicitudEnValidacion` requiere extracción de seam |
| B2 | CAP-005 CDCASUB | Crear `Test_CDCASUB_Strict` | ~15-20 átomos | RAC delegado pendiente de producto |
| B3 | CAP-006 Lifecycle | Crear `Test_Lifecycle_Strict` (alta + búsqueda) | ~10-15 átomos | Requiere fix de password embebido en `frmAltaSolicitud.cls:475` antes |
| B4 | CAP-007 Workflow | Crear `Test_Workflow_Strict` (`LogEstadoRepositorio` repo-level; `EjecutarTransicion`/`ReabrirSolicitudCerrada` deferred) | ~10-15 átomos | `EjecutarTransicion`/`ReabrirSolicitudCerrada` requieren setup de `m_ObjUsuarioActivo.rol` y precondiciones multi-tabla. Cubierto el repo seam. |
| B5 | CAP-009 Logs | Crear `Test_Logs_Strict` | ~10 átomos | `SnapshotServicio.ObtenerDatosTabla` requiere refactor para inyectar `db` |
| B6 | CAP-010 Seguridad | Crear `Test_Security_Strict` (cubre `DeterminarRol` pure logic; `getUsuarioConPermisos` y `getResponsables*` requieren setup de `TbConfiguracionBackends`) | ~8-10 átomos | `TbConfiguracionBackends` vive en frontend; para tests aislados hace falta mover a backend o mockear. |

### Fase C — Cerrar la deuda UI/forms transversal

| # | Slice | Notas |
|---|---|---|
| C1 | Extraer callback atómica de `frmGestionAdjuntos.EjecutarFlujoSubida` a `AdjuntosServicio.SubirYCerrar` | Resuelve deuda UI de CAP-008 BR-009/010 |
| C2 | Reconciliar layout de los 6 subformularios PC (`Form_subfrmDatosPC_*.form.txt`) | Bloquea C1 en su mayoría; requiere export de Access y diff manual |
| C3 | Reconciliar layout de los 6 subformularios PCSUB | Ídem |
| C4 | Reconciliar layout de los 6 subformularios CDCA | Ídem |
| C5 | Reconciliar layout de los 6 subformularios CDCASUB | Ídem |
| C6 | Reconciliar layout de los 9 formularios de lifecycle/búsqueda | Ídem |
| C7 | Reconciliar layout de los 2 formularios de workflow log | Ídem |
| C8 | Reconciliar layout de los 2 formularios de adjuntos | Ídem |
| C9 | Reconciliar layout de los 4 formularios de logs | Ídem |

### Fase D — Cerrar gaps estructurales

| # | Slice | Notas |
|---|---|---|
| D1 | Migrar CDCA de `Nothing`-as-db a `DAO.Database` explícito | Refactor + re-migrar a v2.4.2 |
| D2 | Quitar password embebido de `frmAltaSolicitud.cls:475` | Usar `GetPasswordDB()` |
| D3 | Refactor `WorkflowServicio.ReabrirSolicitudCerrada` para inyectar `db` | Quitar `If db Is Nothing Then Set db = getdb()` |
| D4 | Refactor `SnapshotServicio.ObtenerDatosTabla` para inyectar `db` | Quitar `getdb()` directo |
| D5 | Confirmar con producto: identidad local → IdP corporativo (Kerberos/SAML/OAuth) | Bloqueante para migración web |
| D6 | Sustituir `FIRST(...)` por `MIN`/`MAX` + `TOP 1` ordenado | Riesgo de portabilidad |

## §3 Métricas de progreso

- **Funciones `Public Function Test_*` totales**: 126
- **Funciones en manifests atómicos strict v2.4.2**: 38 (29 PCSUB + 9 Adjuntos) = 30%
- **Capabilities con `Verified-runtime` parcial o total**: 3/10 = 30%
- **Manifests strict v2.4.2**: 2 (`tests.pcsub.json`, `tests.adjuntos.json`); legacy: 8
- **Archivos VBA productivos**: 158 (60 .cls + 53 .bas + 45 forms)
- **PRs cerrados a `origin/staging`** con strict TDD en esta épica: 4 (`f27f693`, `91fe07e`, `fc133ad`, `60ea93a`)

## §4 Decisiones pendientes del usuario

1. **Prioridad**: ¿Fase A (cerrar las 3 con manifest) antes de Fase B (crear manifests para las 6 sin cobertura)? Recomendación: sí, porque cada slice de Fase A entrega valor incremental pequeño y se puede pular en PRs de ≤ 400 líneas sin chained.
2. **Slice A2 vs A3**: BR-004 `Validar` (3 átomos) parece más predecible que BR-006 rollback (ahora con plan blindado por el diag schema). Mi recomendación: **A2 primero** (riesgo bajo, entrega rápido), después A3 con la estrategia PK confirmada por el diag.
3. **B1 (CAP-004 PC)**: requiere refactor de `EsSolicitudEnValidacion` para que el test pueda inyectar `db`. ¿Aceptás ese refactor antes del manifest?
4. **D2 (password)**: ¿querés que lo cierre como parte de B3, o como hotfix previo?
5. **C2..C9 (reconciliación layout)**: ¿bloquea acceptance, o se documenta como `Divergent / blocker UI` mientras se cubren los seams testeables?
6. **Estado del lock del ACCDB (resuelto 2026-06-15)**: 3 PIDs `MSACCESS.EXE -Embedding` (10524, 30852, 33796) cerrados por diagnóstico+autorización. Lock file `.laccdb` removido manualmente. `SchemaInspector.bas` importado y compilado. Diag schema ejecutado. ACCDB operativo.

## §5 Trazabilidad de la épica

| Hito | Commit | Slice | Átomos verdes | Notas |
|---|---|---|---|---|
| Baseline PCSUB strict | `f27f693` | CAP-001 Fase A (legacy → v2.4.2) | 29/29 | Migración de harness + cardinalidad + DB explícito |
| Adjuntos strict BR-001 | `91fe07e` | CAP-008 Fase A.1 | 4/4 | Primer manifest atómico de adjuntos |
| Adjuntos strict BR-002/003 | `fc133ad` | CAP-008 Fase A.2 | 6/6 | PDF-only + dedup nombre destino |
| Adjuntos strict BR-012 | `60ea93a` | CAP-008 Fase A.3 | 9/9 (individual) | `ActualizarFicheroAdjunto` + sad paths |
| Épica + diag schema | `eebf826` | Roadmap operacional + WIP SchemaInspector | n/a | Doc de épica + diag blindando BR-006 |
| SchemaInspector.bas importado + diag corrido | `20f1dc9` | Pre-check de BR-006 | n/a | Confirmó estrategia PK colisión viable; cerró lock del ACCDB |
| Adjuntos strict BR-004 sad paths | `80be457` | CAP-008 Fase A.4 | 11/11 (individual) | `GuardarAdjuntoDesdeArchivo` rechaza idSolicitud<=0 + ruta vacía |
| BR-006 deferred | `04d3fe0` | CAP-008 Fase A.3 (cancelada) | n/a | Documentada limitación: requiere refactor de seams para forzar falla de persistencia |
| SubirYCerrar seam + sad path | `1093cb6` | CAP-008 Fase A.5a | 1/1 | Refactor extrae `SubirYCerrar` + `EjecutarCallbackAprobacion` de `Form_frmGestionAdjuntos` |
| DTM AllMappedFields deferred | `665f8e5` | CAP-002 Fase A.4 (cancelada) | n/a | Subprocess desde VBA embebido en runner Dysflow cuelga; `scripts/run-dtm-tests.ps1` documenta flujo manual |
| Lifecycle getSolicitudesViewModel | `8fa7337` | CAP-006 Fase B.3 | 2/2 | `getSolicitudesViewModel_RetornaLasCreadas` + filtro por palabra clave |
| PC first strict atoms | `54be6aa` | CAP-004 Fase B.1 | 3/3 | `EsDetalleCompleto` (2 átomos) + `GuardarAprobacionSuministrador` UPDATE; bug encontrado: `DatosPCRepositorio` destruye conexión post-Update |
| Logs strict atoms | `63d4116` | CAP-009 Fase B.5 | 2/2 | `LogServicio.RegistrarCambio` (db inyectado) + `RegistrarError` (firma real, cardinalidad global) |
| CDCASUB strict atoms | `854f32b` | CAP-005 Fase B.2 | 3/3 | `EsParteTecnicaCompleta` happy + sad `clasifNC=INVALID` + `EsAprobacionSuministradorCompleta` sad `firmaIngenieria=''`. Valida divergencias vs PC. |
| Workflow LogEstado atoms | `13bfcd1` | CAP-007 Fase B.4 | 3/3 | `LogEstadoRepositorio.Guardar` (Insert + autonumérico) + `getUltimoEstadoAnterior` (última + 0 vacío). `EjecutarTransicion`/`ReabrirSolicitudCerrada` deferred. |
| Security DeterminarRol atoms | `fc025e1` | CAP-010 Fase B.6 | 4/4 | `UsuarioServicio.DeterminarRol` pure logic: Administrador + Calidad + Tecnico fallback + Nothing guard. Sin DB. |
| Manifests expansión (B1.1, B2.1, B3.x, B4.1, B5.1) | `946474c` `3942150` `309a6b6` `25822d7` | B+ slice expansions | +11 átomos | PC +3 (EsDatosGeneralesCompleta, EsParteTecnicaCompleta, EsMotivosCompleto); CDCASUB +3 (EsDictamenRACCompleta, EsDecisionFinalCompleta, EsDatosGeneralesCompleta); Workflow +2 (getHistorial, EliminarPorIdSolicitud); Logs +1 (LogCambioRepositorio.EliminarPorSolicitud); Lifecycle +2 (getSolicitudPorID happy + sad); Lifecycle +1 (getExpedientePorID sad path only — happy path deferred por `getdbExpedientes()`). |

### Deferred (documentados, sin commit verde)

| Slice | Capability | Motivo | Workaround aplicado |
|---|---|---|---|
| A.3 BR-006 rollback | CAP-008 | `getSiguienteIDAdjunto` calcula `MAX+1` DESPUÉS del seed; `Validar` regla 3 (256+ chars) choca con Windows MAX_PATH | Documentada limitación; el test happy path verifica cardinalidad normal |
| A.4 DTM AllMappedFields | CAP-002 | Subprocess PowerShell/Shell.Application desde VBA embebido cuelga en runner Dysflow | `scripts/run-dtm-tests.ps1` documenta flujo manual con `timeoutMs` per-call |
| A.5b SubirYCerrar happy path | CAP-008 | Requiere `CumplePasoAAprobada` (adjunto con `etapaWF="Cierre"`) + `rolUsuario` configurado en sandbox | El sad path (etapa != "Documento Final Firmado") sí está cubierto |
| B4 `EjecutarTransicion` / `ReabrirSolicitudCerrada` | CAP-007 | Requieren `m_ObjUsuarioActivo.rol` en {Calidad, Administrador} + precondiciones multi-tabla. LogEstadoRepositorio SÍ está cubierto (commit `13bfcd1`) | Se cubre el seam `LogEstadoRepositorio` (3 átomos); los servicios quedan pendientes hasta que se refactorice el setup del rol. |

## §6 Referencias

- [`access-vba-tdd` v2.4.2 skill](../../Users/adm1/.config/opencode/skills/access-vba-tdd/SKILL.md)
- [`access-vba-capability-docs` skill](../../Users/adm1/.config/opencode/skills/access-vba-capability-docs/SKILL.md)
- [Índice de capabilities](../capabilities/capabilities-index.md)
- [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md)
- [`docs/ERD/condor_datos.md`](../ERD/condor_datos.md) — schema del backend
- [`docs/DISCOVERY_MAP.md`](../DISCOVERY_MAP.md) — puntero al OpenSpec externo
