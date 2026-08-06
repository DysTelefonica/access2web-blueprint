# Capacidad: Workflow técnico, precondiciones, transiciones, rechazo y formalización

## §0 Identidad

- **ID de capacidad**: CAP-007
- **Tier**: critical
- **Estado**: active con deuda de seams transaccionales, rechazo y reconciliación UI/logs
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary` sobre `WorkflowServicio`, `WorkflowRepositorio`, `Estado*`, `Rechazo*`, `ValidacionRevision*`, `modEnumeradores` y formularios de log. Motor/repositorios/code-behind `.cls` están `matched`; `Form_frmVerLogCambios.form.txt` y `Form_frmDetalleLogCambio.form.txt` tienen diferencias accionables `bothChanged`. No existe módulo separado `PrecondicionesWorkflowServicio`; las precondiciones viven en `WorkflowServicio`.
- **Confianza global**: `Verified-runtime` parcial. Las precondiciones y transiciones están implementadas y el código está sincronizado con el binario, pero la falta de manifest workflow strict, la mezcla UI/transacciones, `db Is Nothing -> getdb()` y la deriva UI de formularios de log bloquean la promoción a `Verified-runtime`.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). Workflow es la capacidad con mayor deuda por la cantidad de ramas y la mezcla UI/servicio.

**Contrato TDD vigente**: las pruebas de workflow deben migrarse a `access-vba-tdd` v2.4.2 — `Public Function` con retorno JSON canónico, fixture propio por cada transición, schema-first, `DAO.Database` inyectado, cardinalidad `countBefore`/`countAfter` sobre `tbLogEstados`/`tbSolicitudes`, manifests atómicos y cero mutación de `TbConfiguracionBackends`.

**Justificación del nivel**: crítico, porque el workflow es el motor de cambio de estado de TODAS las solicitudes. Una regresión aquí invalida Pruebas de aceptación enteras y compromete SLAs.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: gobernar el ciclo de vida de cualquier solicitud (PC, PCSUB, CD/CA, CDCASUB) mediante transiciones entre estados (`Preregistro`, `Registro`, `Desarrollo Técnico`, `Modificación`, `Validación`, `Revisión`, `Formalización`, `Aprobada`, `Rechazada`) y precondiciones por estado.
- **Usuarios / perfiles**: técnicos (Desarrollo Técnico, Modificación, Validación), calidad (aceptar/rechazar, validación, formalización), RAC (revisión), administrador (reapertura), y todos los roles para borrado y consulta.
- **Problema que resuelve**: centraliza las reglas de transición, los permisos por rol, las precondiciones por tipo de solicitud y el registro de auditoría, para que ningún formulario pueda mutar el estado por SQL directo.
- **Valor de negocio**: orden del proceso, trazabilidad de cambios de estado, separación por roles, soporte de re-envíos tras rechazo y conservación de historial completo.
- **No-objetivos**: este documento no cubre la captura de datos técnicos (CAP-003/CAP-004/CAP-005/CAP-001), los adjuntos (CAP-008) ni los documentos (CAP-002). Sí los vincula.
- **Origen de la intención**: código actual, SDD `pcsub-staging-recovery` (reconciliación transaccional), SDD `pcsub-guardar-phase-advancement`, y el patrón de "firma del técnico" y "rechazo de calidad" definidos en PRD.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** una solicitud en cualquier estado **CUANDO** se pide `WorkflowServicio.getTransicionesValidas(idSolicitud, rolUsuarioActual)` **ENTONCES** se devuelven los estados destino cuyo `rolRequerido` en `tbTransiciones` sea satisfecho por el rol y que pasen `PrecondicionesCumplidas`. **Estado**: `Verified-static`; la lógica combina `WorkflowRepositorio.getPosiblesDestinos` con `PrecondicionesCumplidas` y `PermisoSuficiente`.
- **DADO** una solicitud en `estadoPreregistro` **CUANDO** se guarda Datos Generales completos **ENTONCES** `DatosXServicio.GuardarDatosGenerales` invoca `EjecutarTransicion(sol, estadoRegistro, m_ObjUsuarioActivo, db)` dentro de la misma transacción. **Estado**: `Verified-static`.
- **DADO** una solicitud en `estadoDesarrolloTecnico` con Detalle y Motivos completos **CUANDO** el usuario confirma el envío a Calidad **ENTONCES** se transiciona a `estadoModificacion` y se pregunta por rechazo. **Estado**: `Verified-static`; la navegación final es responsabilidad del formulario (CAP-004).
- **DADO** una solicitud en `estadoModificacion` **CUANDO** Calidad acepta **ENTONCES** se transiciona a `estadoValidacion` y `ValidacionRevisionRepositorio.CrearNuevoCicloValidacion` abre un ciclo de validación nuevo. **Estado**: `Verified-static`.
- **DADO** una solicitud en `estadoValidacion` con RAC que devuelve `APROBADO` **ENTONCES** `CumplePasoAFormalizacion` permite `estadoFormalizacion`. **Estado**: `Verified-static`.
- **DADO** una solicitud en `estadoFormalizacion` con PDF de cierre subido **ENTONCES** `CumplePasoAAprobada` permite `estadoAprobada`. **Estado**: `Verified-static`.
- **DADO** una solicitud en cualquier estado editable **CUANDO** se ejecuta `EjecutarTransicion` con precondición fallida **ENTONCES** se lanza `Err.Raise 513` con el mensaje de validación. **Estado**: `Verified-static`.
- **DADO** una solicitud con rechazo activo de Calidad **CUANDO** el usuario guarda Impacto/Propuesta **ENTONCES** se registra el delta JSON por `JsonHelper.RegistrarCambio` y `CumplePasoAModificacion` exige `VerificarCambiosTrasRechazo` para permitir la transición. **Estado**: `Verified-static`.
- **DADO** una solicitud en `estadoAprobada` **CUANDO** se ejecuta `ReabrirSolicitudCerrada` por Administrador o Calidad **ENTONCES** se borra la decisión final del tipo correspondiente, se calcula el estado anterior con `LogEstadoRepositorio.getUltimoEstadoAnterior` y se llama a `RevertirAFaseAnterior`. **Estado**: `Verified-static`.
- **DADO** una transición exitosa **ENTONCES** se persiste `tbLogEstados` (idEstadoAnterior, idEstadoNuevo, fecha, usuario) y `LogCambioRepositorio` registra un cambio de tipo `TRANSICIÓN_WF`. **Estado**: `Verified-static`.
- **DADO** una transición exitosa **ENTONCES** `NotificacionServicio` envía un correo no bloqueante; si falla, el error se limpia y no se propaga (la transición ya está confirmada). **Estado**: `Verified-static`; revisión humana pendiente sobre la semántica de "no bloqueante".

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | `EjecutarTransicion` exige precondiciones (`PrecondicionesCumplidas_PC/...`) y rol suficiente (`PermisoSuficiente`). Si no, lanza `Err.Raise 513`. | Código | Sí: `WorkflowServicio.EjecutarTransicion` líneas 1317-1323. | Pendiente. | Verified-static |
| BR-002 | Para llegar a `estadoModificacion` la parte técnica del tipo debe estar completa y, si hay rechazo, debe haber cambios verificables. | Código | Sí: `CumplePasoAModificacion` y `VerificarCambiosTrasRechazo`. | Pendiente. | Verified-static |
| BR-003 | Para llegar a `estadoValidacion` se exige parte técnica completa + RAC completo + Aprobación Suministrador completa. | Código | Sí: `CumplePasoAValidacion` líneas 1117-1173. | Pendiente. | Verified-static |
| BR-004 | Para llegar a `estadoRevision` debe existir adjunto en la etapa de `estadoValidacion`. | Código | Sí: `CumplePasoARevision` líneas 1175-1197. | Pendiente. | Verified-static |
| BR-005 | Para llegar a `estadoFormalizacion` el último `ValidacionRevision.resultado` debe ser `APROBADO`. | Código | Sí: `CumplePasoAFormalizacion` líneas 1199-1229. | Pendiente. | Verified-static |
| BR-006 | Para llegar a `estadoAprobada` debe existir adjunto de cierre (etapa `Cierre` o equivalente). | Código | Sí: `CumplePasoAAprobada` líneas 1261-1283. | Pendiente. | Verified-static |
| BR-007 | Las transiciones Modificación → Validación desactivan rechazos previos (`RechazoRepositorio.DesactivarRechazosPrevios`) para que la subsanación quede registrada. | Código | Sí: `EjecutarTransicion` líneas 1340-1342. | Pendiente. | Verified-static |
| BR-008 | Al llegar a `estadoValidacion` se crea un ciclo de validación nuevo (`CrearNuevoCicloValidacion`). | Código | Sí: `EjecutarTransicion` líneas 1345-1349. | Pendiente. | Verified-static |
| BR-009 | Al transicionar a `estadoDesarrolloTecnico` se reinicia `revisionCalidadEstado` a `PENDIENTE`. | Código | Sí: `EjecutarTransicion` líneas 1351-1353. | Pendiente. | Verified-static |
| BR-010 | La jerarquía de roles es: `Admin` ⊇ `Calidad` ⊇ `Tecnico` (más otros roles `Ingenieria`/`Economia`/`Secretaria`/`SinAcceso`/`CalidadAvisos`). | Código | Sí: `PermisoSuficiente` líneas 1483-1491. | Pendiente. | Verified-static |
| BR-011 | `ReabrirSolicitudCerrada` solo lo permite `Administrador` o `Calidad`, y solo desde `estadoAprobada`. Limpia la decisión final del tipo y revierte al estado anterior registrado. | Código | Sí: `ReabrirSolicitudCerrada` líneas 1948-1996. | Pendiente. | Verified-static |
| BR-012 | `getPaginaActivaPC/...` (y CDCA/CDCASUB/PCSUB) delegan en `getPaginaActivaGenerica` y mapean estado a pestaña (`tabGeneral`, `tabPropuesta`, `tabAprobacionSuministrador`, `tabDictamenRAC`, `tabDecisionFinal`). | Código | Sí: `WorkflowServicio.getPaginaActivaGenerica` líneas 2002-2064. | Pendiente. | Verified-static |
| BR-013 | El visualizador de estado (`GenerarHTML_VisualizadorDeEstado`) genera HTML por estado, con detalle forense, RAC, rechazos y adjuntos; sirve cacheado 5 minutos. | Código (Spec-118) | Sí: `WorkflowServicio.GenerarHTML_VisualizadorDeEstado` y `m_cacheTimeline`. | Pendiente. | Verified-static |
| BR-014 | El rechazo desde formalización borra Decisión Final, deja `racDecision = "RECHAZADO"`, persiste `racRechazoMotivos` y transiciona a `estadoRechazada`. | Código (PC, PCSUB, CDCA, CDCASUB) | Sí: `XServicio.RegistrarRechazoDesdeFormalizacion`. | Pendiente. | Verified-static |
| BR-015 | `PermiteEdicion_CDCA`, `PermiteEdicion_CDCASUB`, `PermiteEdicion_PC`, `PermiteEdicion_PCSUB` aplican reglas por bloque y estado; comparten estructura gemela. | Código | Sí: líneas 1557-1825. | Pendiente. | Verified-static |

### Validaciones observadas

- `EsTransicionPermitida`: estado origen/destino y rol suficiente.
- `PrecondicionesCumplidas_PC/...`: por tipo y por estado destino.
- En CAP-004/CAP-005 ya se listan las validaciones por bloque (Datos Generales, Propuesta, Impacto, Aprobación, RAC, Decisión Final).

### Transiciones de estado canónicas

- `estadoPreregistro` → `estadoRegistro` (Datos Generales completos).
- `estadoRegistro` → `estadoDesarrolloTecnico` (decisión del usuario o por Calidad tras revisión).
- `estadoDesarrolloTecnico` → `estadoModificacion` (parte técnica completa y, si hay rechazo, cambios verificables).
- `estadoModificacion` → `estadoValidacion` (aceptación de Calidad).
- `estadoModificacion` → `estadoDesarrolloTecnico` (rechazo de Calidad; permite subsanación con `revisionCalidadEstado = "RECHAZADO"`).
- `estadoValidacion` → `estadoRevision` (adjunto de borrador Word + envío a RAC).
- `estadoValidacion` → `estadoModificacion` (rechazo de RAC; trata como `REVERTIR_TECNICA`).
- `estadoRevision` → `estadoFormalizacion` (RAC `APROBADO`).
- `estadoRevision` → `estadoRechazada` (cerrar como rechazada).
- `estadoFormalizacion` → `estadoAprobada` (PDF de cierre adjunto + decisión final del tipo).
- `estadoAprobada` → estado anterior (reapertura por Calidad/Admin).

### Casos límite y hallazgos

- En `EjecutarTransicion` el control de errores activa `On Error Resume Next` antes de invocar `notifServ.EnviarNotificacion`. Si el envío de correo falla, el `Err.Clear` posterior enmascara el error para que no se propague; revisar la conveniencia de esta política con producto.
- `RechazoRepositorio.DesactivarRechazosPrevios` se llama solo en la transición `estadoModificacion → estadoValidacion`. Esto es coherente con el modelo de subsanación, pero cualquier intento de subsanar saltando pasos intermedios queda fuera del contrato actual.
- `GenerarHTML_VisualizadorDeEstado` genera HTML enorme (>500 líneas) inline; el tamaño de la cadena puede ser un problema de performance en formularios de solo lectura. Considerar un motor de plantillas externo o generar incremental.
- `VerificarCambiosTrasRechazo` se apoya en `JsonHelper` y puede fallar si el JSON de rechazo no existe o está corrupto; conviene blindar el helper.

### Señales de aceptación / presencia

- Existen `WorkflowServicio.cls` (motor central), `Estado.cls`/`EstadoServicio.cls`/`EstadoRepositorio.bas`, `WorkflowRepositorio.bas` (lectura de `tbTransiciones`).
- Existen `Rechazo.cls`/`RechazoServicio.cls`/`RechazoRepositorio.bas` para la gestión de rechazos y su historial.
- Existen `ValidacionRevision.cls`/`ValidacionRevisionServicio.cls`/`ValidacionRevisionRepositorio.bas` para el ciclo de validación/revisión.
- `modEnumeradores.bas` define los enums: `enumEstados`, `enumBloqueFormulario`, `enumTipoSolicitud`, `rol`.
- `m_ObjEntorno.estados` carga los IDs canónicos de los estados y permite mapear por nombre.
- No existe un módulo separado `PrecondicionesWorkflowServicio`; las precondiciones por tipo/estado residen en `WorkflowServicio`.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frmVerLogCambios` y `Form_frmDetalleLogCambio` para auditoría de transiciones (CAP-009).
  - `frmGestionSolicitud` consume `getTransicionesValidas` indirectamente vía el visualizador HTML.
  - El visualizador web (`frmWebVisor`) consume `GenerarHTML_VisualizadorDeEstado` con cache y expone `colaComandos` (`APROBAR_RAC`, `REVERTIR_VALIDACION`, `EXPORTAR_BORRADOR`, `AVANZAR_ESTADO`, etc.).
  - En CAP-004/CAP-005 ya se listan los entry points por tipo (`GuardarDesdeSubform` que delega en `EjecutarTransicion`).
- **Puntos de entrada de código**:
  - `WorkflowServicio.EjecutarTransicion` (atómico: valida, persiste, registra log, notifica no bloqueante).
  - `WorkflowServicio.getTransicionesValidas`, `EsTransicionPermitida`, `UsuarioPuedeTransicionarDesde`.
  - `WorkflowServicio.PuedeEditarBloque` (usado por cada formulario de datos para habilitar/ deshabilitar UI).
  - `WorkflowServicio.PermiteEdicion_CDCA`, `PermiteEdicion_CDCASUB`, `PermiteEdicion_PC`, `PermiteEdicion_PCSUB`.
  - `WorkflowServicio.getPaginaActivaPC/CDCA/CDCASUB/PCSUB`, `getPaginaActivaGenerica`, `GetProximaVersionBorrador`, `ObtenerHayCambiosValidacion`.
  - `WorkflowServicio.GenerarYAdjuntarDocumentoBorrador`, `GenerarHTML_VisualizadorDeEstado`, `m_cacheTimeline`, `InvalidarCacheTimeline`, `LimpiarCacheTimeline`.
  - `WorkflowServicio.EjecutarCierreFormalizacion`, `ReabrirSolicitudCerrada`, `RevertirAFaseAnterior`, `RegistrarBorradorRAC`, `TieneAdjuntoEnCicloActual`.
  - `RechazoServicio.GetUltimoRechazoActivo`, `DesactivarRechazosPrevios`, `LimpiarHistorialRechazos`.
  - `ValidacionRevisionServicio.RegistrarEnvioRAC`, `RegistrarRespuestaRAC`, `RegistrarBorrador`, `LimpiarPorSolicitud`, `GetUltimoOrdinal`, `GetUltimoPendienteId`.
  - `EstadoServicio`/`EstadoRepositorio`/`WorkflowRepositorio` para lectura.
- **Datos afectados**:
  - `tbEstados`/`tbTransiciones`: lectura para precondiciones; escritura solo si se modifica la configuración.
  - `tbLogEstados`: escritura en cada transición (`idEstadoAnterior`, `idEstadoNuevo`, `fechaTransicion`, `usuarioTransicion`).
  - `tbLogCambios`: escritura de tipo `TRANSICIÓN_WF` por transición.
  - `tbRechazos`/`tbHistorialRechazos`: lectura/escritura para la trazabilidad del rechazo.
  - `tbValidacionRevision`: lectura/escritura para el ciclo de validación (envío a RAC, respuesta, ordinal, hash, idAdjunto).
  - `tbSolicitudes`: estado, fechaModificacion, usuarioModificacion.
- **Dependencias**:
  - `DatosPC/PCSUB/CDCA/CDCASUBServicio` para precondiciones por tipo.
  - `AdjuntosServicio` para `CumplePasoARevision`/`CumplePasoAAprobada` y para borrar adjuntos de etapas.
  - `DocumentoServicio` para generar el borrador RAC.
  - `NotificacionServicio` para envío no bloqueante.
  - `m_ObjEntorno.estados`, `m_ObjUsuarioActivo`, `m_ObjUsuarioReal`, `rolUsuario`.
- **Sincronización fuente↔binario**: si se modifican `WorkflowServicio`, `WorkflowRepositorio`, `RechazoRepositorio`, `ValidacionRevisionRepositorio` o `EstadoRepositorio`, basta `dysflow.import_code` (sin `import-form`). Los formularios de logs (`frmVerLogCambios`, `frmDetalleLogCambio`) consumen la salida pero no la generan.
- **Valoración de diseño (tal-como-está vs ideal)**: la separación `WorkflowServicio` ↔ `WorkflowRepositorio` ↔ `DatosXServicio` es correcta. La forma `getTransicionesValidas(idSolicitud, rol)` que filtra por precondiciones es la dirección correcta, pero su uso depende de la rama UI. La deuda principal está en (a) `EjecutarTransicion` con `On Error Resume Next` para notificaciones, (b) `ReabrirSolicitudCerrada` con `If db Is Nothing Then Set db = getdb()` (rompe inyección de `db`), (c) `VerificarCambiosTrasRechazo` con dependencia de JSON, (d) ausencia de seam testeable para precondiciones y transiciones. La pieza está bien hecha para producción; no se recomienda `Verified-runtime` sin migrar la suite.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `WorkflowServicio.cls`, `WorkflowRepositorio.bas`, `Estado.cls`, `EstadoServicio.cls`, `EstadoRepositorio.bas`.
2. Restaurar `Rechazo.cls`, `RechazoServicio.cls`, `RechazoRepositorio.bas` y `ValidacionRevision.cls`, `ValidacionRevisionServicio.cls`, `ValidacionRevisionRepositorio.bas`.
3. Confirmar `modEnumeradores.bas` con `enumEstados`, `enumBloqueFormulario`, `enumTipoSolicitud` y `rol`.
4. Confirmar `tbEstados`, `tbTransiciones`, `tbRechazos`, `tbHistorialRechazos`, `tbValidacionRevision` con su esquema.
5. Restaurar `m_ObjEntorno.estados` con los IDs canónicos (`estadoPreregistro`, `estadoRegistro`, `estadoDesarrolloTecnico`, `estadoModificacion`, `estadoValidacion`, `estadoRevision`, `estadoFormalizacion`, `estadoAprobada`, `estadoRechazada`).
6. Importar con `dysflow.import_modules` y compilar con `dysflow.compile_vba`. Verificar binario con `dysflow.verify_binary`; actualmente el código del motor está sincronizado, pero los `.form.txt` de `frmVerLogCambios` y `frmDetalleLogCambio` están `bothChanged`.
7. Demostrar los escenarios de §2 con un manifest atómico `tests/tests.workflow.json` que cubra cada transición y cada precondición. Mientras no exista, esta capacidad queda en `Verified-static`.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/WorkflowServicio.cls` (2183+ líneas).
  - `src/classes/Estado.cls`, `EstadoServicio.cls`, `WorkflowRepositorio.bas`, `EstadoRepositorio.bas`.
  - `src/classes/Rechazo.cls`, `RechazoServicio.cls`, `RechazoRepositorio.bas`.
  - `src/classes/ValidacionRevision.cls`, `ValidacionRevisionServicio.cls`, `ValidacionRevisionRepositorio.bas`.
  - `src/modules/modEnumeradores.bas` (enums).
  - `src/forms/Form_frmVerLogCambios.cls`, `Form_frmDetalleLogCambio.cls`.
  - `docs/ERD/condor_datos.md` → `tbEstados`, `tbTransiciones`, `tbRechazos`, `tbHistorialRechazos`, `tbValidacionRevision`.
- **Evidencia Dysflow incorporada**:
  - `dysflow.verify_binary`: `WorkflowServicio`, `WorkflowRepositorio`, `EstadoServicio`, `EstadoRepositorio`, `RechazoServicio`, `RechazoRepositorio`, `ValidacionRevisionServicio`, `ValidacionRevisionRepositorio` y `modEnumeradores` matched.
  - `Form_frmVerLogCambios.cls` y `Form_frmDetalleLogCambio.cls` matched; sus `.form.txt` están `bothChanged` y requieren reconciliación UI.
  - Primer intento con `PrecondicionesWorkflowServicio` falló con `VBA_MODULE_NOT_FOUND`, confirmando que no existe como módulo independiente.
- **Tests existentes**: manifest atómico `tests/testsWorkflow.json` (commit `13bfcd1`, 2026-06-15, Slice B4). Cubre `LogEstadoRepositorio.Guardar` (Insert + autonumérico `idLogEstado`) y `getUltimoEstadoAnterior` (última transición + 0 para vacío). 3/3 átomos verdes. `EjecutarTransicion` y `ReabrirSolicitudCerrada` quedan deferred — requieren setup de `m_ObjUsuarioActivo.rol` (Calidad/Administrador) y precondiciones multi-tabla.
- **Evidencia runtime Dysflow**:
  - `Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields`: **VERDE** 2.7 s, sandbox `condor_datos.accdb` local, `idSolicitud=900911`, transición 2→3 (`estadoRegistro`→`estadoDesarrolloTecnico`), `usuarioTransicion='workflow-test-user'`.
  - `Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado`: **VERDE** 2.7 s, `idLogEstado` antes de Guardar=0, después de Guardar Access asigna autonumérico (id=2).
  - `Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition`: **VERDE** 2.7 s, secuencia 2→3→4→5 con fechas crecientes; `getUltimoEstadoAnterior=4` (anterior de la última 4→5); `getUltimoEstadoAnterior(999999)=0` para idSolicitud sin historial.
- **SDD/intención consultada**:
  - SDD `pcsub-staging-recovery`.
  - SDD `pcsub-guardar-phase-advancement`.
  - Specs internos: Spec-003 (Regla estricta de Modificación para Calidad), Spec-007 (Subsanación), Spec-118 (cache de timeline), Spec-132/137 (mensajes en Desarrollo Técnico), Spec-151 (checklist tras Aprobación Suministrador).

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| Una transición válida no aparece en el visualizador | `getTransicionesValidas` filtra por precondición que no se cumple o por rol insuficiente. | prueba focal con `getTransicionesValidas` y un manifest que cubra cada rama. | §2 BR-001 |
| La transición Modificación → Validación no desactiva rechazos | `RechazoRepositorio.DesactivarRechazosPrevios` no se invoca por path distinto. | prueba con `idEstadoAnterior=4` y `idEstadoDestino=5`. | §2 BR-007 |
| La transición a Validación no crea ciclo de validación | `CrearNuevoCicloValidacion` no se invoca. | prueba focal con cardinalidad en `tbValidacionRevision`. | §2 BR-008 |
| La reapertura desde Aprobada no funciona | `ReabrirSolicitudCerrada` no encuentra la decisión final o el estado anterior. | prueba con `idEstadoInterno=8` y un idSolicitud con Decisión Final. | §2 BR-011 |
| El visualizador no se actualiza tras transición | `InvalidarCacheTimeline` no se invoca o el cache se sirve stale. | prueba con TTL y transición. | §2 BR-013 |
| Notificación no se envía | `notifServ.EnviarNotificacion` lanza error o destinatario vacío. | prueba con stub de `notifServ`. | §2 BR-001 |
| Permisos UI deshabilitados en un estado que debería permitir edición | `PermiteEdicion_*` no se ha actualizado al nuevo estado del workflow. | prueba con matriz estado × bloque × rol. | §2 BR-015 |

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Workflow y precondiciones | Pendiente | Pendiente de confirmación | pending | Pendiente | Pendiente | Motor/repositorios sincronizados; pendiente manifest atómico y reconciliación de layout en formularios de log. |

## §6 Notas de migración web

- **Conservar**: precondiciones por tipo y estado, jerarquía de roles, registro en `tbLogEstados` y `tbLogCambios`, formalización con PDF de cierre, RAC como ciclo ordinal, cache de visor con invalidación por transición, política de no bloqueante para notificaciones (con revisión).
- **Transformar**: `On Error Resume Next` para notificaciones a un job de notificaciones con reintentos y observabilidad; `DAO.Workspace` local + `MS Access;PWD=…` a transacciones en servidor con ORM; HTML inline a un motor de plantillas server-side; reglas de precondición a un motor de reglas declarativo.
- **NO copiar**: dependencia de `m_ObjEntorno`/`m_ObjUsuarioActivo` globales, `getdb()` como singleton, mezcla de UI y transacciones, política de "notificación no bloqueante" sin reintento explícito.
- **Preguntas abiertas**: ¿La transición Modificación → Validación debe desactivar el rechazo o basta con registrarlo como histórico? (responsable de producto). ¿La reapertura debe borrar también la traza del rechazo? (responsable de calidad). ¿Las precondiciones deben vivir en una capa de reglas declarativa accesible al visualizador? (equipo técnico).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| `WorkflowServicio.EjecutarTransicion` valida rol, precondiciones, persiste, registra log y emite notificación no bloqueante. | Verified-static | `WorkflowServicio.cls` líneas 1284-1481. | 2026-06-15 |
| `getTransicionesValidas` filtra por rol y precondiciones. | Verified-static | `WorkflowServicio.cls` líneas 820-873. | 2026-06-15 |
| `CumplePasoAModificacion` exige parte técnica completa y, si hay rechazo, cambios verificables. | Verified-static | `WorkflowServicio.cls` líneas 1058-1115. | 2026-06-15 |
| `CumplePasoAValidacion` exige parte técnica + RAC + Aprobación Suministrador. | Verified-static | `WorkflowServicio.cls` líneas 1117-1173. | 2026-06-15 |
| `CumplePasoARevision` exige adjunto de borrador Word. | Verified-static | `WorkflowServicio.cls` líneas 1175-1197. | 2026-06-15 |
| `CumplePasoAFormalizacion` exige último `ValidacionRevision.resultado = "APROBADO"`. | Verified-static | `WorkflowServicio.cls` líneas 1199-1229. | 2026-06-15 |
| `CumplePasoAAprobada` exige adjunto de cierre. | Verified-static | `WorkflowServicio.cls` líneas 1261-1283. | 2026-06-15 |
| `ReabrirSolicitudCerrada` solo lo permite `Admin`/`Calidad` y revierte al estado anterior registrado. | Verified-static | `WorkflowServicio.cls` líneas 1948-1996. | 2026-06-15 |
| `EjecutarTransicion` activa `On Error Resume Next` para notificaciones y enmascara el fallo tras COMMIT. | Verified-static / política a revisar | `WorkflowServicio.cls` líneas 1384-1458. | 2026-06-15 |
| `PermiteEdicion_*` aplica reglas por bloque, estado y rol. | Verified-static | `WorkflowServicio.cls` líneas 1557-1825. | 2026-06-15 |
| `getPaginaActivaGenerica` mapea estado a pestaña activa. | Verified-static | `WorkflowServicio.cls` líneas 2002-2048. | 2026-06-15 |
| El visualizador web cachea HTML 5 minutos y se invalida al transicionar. | Verified-static | `WorkflowServicio.cls` líneas 15-71 y `InvalidarCacheTimeline`. | 2026-06-15 |
| El rechazo desde formalización queda registrado en `tbRechazos` con `racDecision = "RECHAZADO"`. | Verified-static | `DatosXServicio.RegistrarRechazoDesdeFormalizacion` (PC/PCSUB/CDCA/CDCASUB). | 2026-06-15 |
| `ReabrirSolicitudCerrada` hace `If db Is Nothing Then Set db = getdb()` lo que rompe la inyección de `db`. | Verified-static / deuda | `WorkflowServicio.cls` línea 1954. | 2026-06-15 |
| Existe un manifest atómico de pruebas de workflow que cumpla `access-vba-tdd` v2.4.2. | Verified-runtime (parcial) / pendiente `EjecutarTransicion`+`ReabrirSolicitudCerrada` | `tests/testsWorkflow.json` con tres átomos verdes (commit `13bfcd1`): `LogEstadoRepositorio.Guardar` 2.7 s + autonumérico 2.7 s + `getUltimoEstadoAnterior` 2.7 s. Pendientes: `EjecutarTransicion` (requiere `m_ObjUsuarioActivo.rol` setup), `ReabrirSolicitudCerrada` (idem), `GenerarYAdjuntarDocumentoBorrador`. || 2026-06-15 |
| Motor y repositorios workflow están sincronizados fuente↔binario. | Verified-static / binary-synced | `verify_binary`: módulos workflow/estado/rechazo/validación/enums matched. | 2026-06-15 |
| Formularios de log workflow están reconciliados fuente↔binario. | Divergent / blocker UI | `verify_binary`: `Form_frmVerLogCambios.form.txt` y `Form_frmDetalleLogCambio.form.txt` con `bothChanged`. | 2026-06-15 |
| Phase 0 vincula los casos UAT de navegación y cambio de estado a átomos de workflow verdes antes de aceptación. | Verified-static / deuda Phase 0 | `audit-e2e-thin-forms-phase-0.md` §7-§8; fila `Test_Workflow_EjecutarTransicion_PreconditionFailureDoesNotMutateState`. | 2026-06-26 |
| HTML UAT de transiciones/precondiciones queda bloqueado hasta tener átomo verde de `EjecutarTransicion`, manifest focal y `ref` firmado. | Divergent / pendiente | El manifest actual solo cubre repositorio de log; no cubre precondiciones ni mutación de `tbSolicitudes`. | 2026-06-26 |
| PCSUB `Bloque_Generales` puede ejecutar la transición `estadoRegistro -> estadoDesarrolloTecnico` mediante `WorkflowServicio` al guardar desde subformulario con aceptación del usuario, y no planifica prompt/workflow si la validación/persistencia falla. | Verified-runtime limitado a Slice 2 / pendiente cobertura workflow general | `tests/tests.pcsub.form.json` → 2/2 verde: `Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow` y `Test_PCSUB_GuardarDesdeSubform_Generales_InvalidDoesNotPlanPromptOrWorkflow`; evidencia de cardinalidad `tbDatosPCSUB: 0 -> 1` en happy path, estado/log de workflow registrado, y sad path sin persistencia ni workflow. No cubre todas las precondiciones ni el motor workflow completo. | 2026-06-27 |
| Slice 3.3 refactor: `WorkflowServicio.EjecutarTransicion` delega las decisiones de precondiciones en 3 helpers puros compartidos — `PrecondicionesWorkflowHelper` (4 funciones cubriendo BR-002..BR-006 de los `CumplePasoA*` con `p_RazonFallo` por callsite), `JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde` (BR-010, tabla hash Admin ⊇ Calidad ⊇ Tecnico, sin uso desde PermisoSuficiente todavía), `NotificacionHelper_NotificarSiPosible` (5 args reales: `vm, asunto, destinatarioEmail, copiaEmail, [nombreEstadoNotif]`, swallow de `On Error Resume Next` interno con `Err.Clear`). Wording user-facing preservado en cada callsite (decisión D3/B3); `On Error Resume Next` inline del notif block eliminado en `WorkflowServicio.cls ~1395-1478`. Helper de tildes `PrecondicionesWorkflowHelper_NormalizarTildes` (Chr$(193/201/...)) tolera `GetNombreEstadoSafe` devolviendo `Validación` con tilde. Manifest focal `tests/tests.workflow.strict.json` con **25 átomos** (12 Precondiciones + 8 JerarquiaRoles + 5 Notificacion). Commits `8452b8b` (helpers + tests + manifest), `7f0e06d` (B1 fix notif 5-args), `4c6b062` (WorkflowServicio consume helpers). `PermisoSuficiente` queda con su tabla inline por scope (requiere contexto origen/destino que 4 callsites no aportan); delega a slice posterior. `_RACCompleto` no se delega en `CumplePasoAValidacion` por dependencia gemelo-specific en `EsDictamenRACCompleto`; otra decisión para slice posterior. | Verified-runtime / pendiente `PermisoSuficiente`+`_RACCompleto`+UAT | Dysflow 2.9.0 import OK para los 8 paths (WorkflowServicio + MockNotifServ + 3 helpers + 3 tests); `dysflow.run_vba` workaround por `AutoExec`-disable bug de `dysflow.test_vba`/`verify_code`; pendiente confirmación de 25/25 átomos verde post-VBE-compile. Spec en `changes/e2e-methodology-exhaustive-rollout/specs/workflow-precondicion-transicion/spec.md` (329 líneas, 4 Requirements, 12 scenarios). | 2026-07-13 |

**Divergencias pendientes de revisión humana**:

- BR-001: la política de notificación no bloqueante enmascara el error de envío. Confirmar con producto si la transición debe fallar si el correo falla, o si la política actual es aceptable.
- BR-011: `ReabrirSolicitudCerrada` hace `If db Is Nothing Then Set db = getdb()`. Migrar a `db` siempre inyectado por el llamador para que las pruebas puedan inyectar sandbox.
- BR-002: `CumplePasoAModificacion` exige cambios tras rechazo pero `VerificarCambiosTrasRechazo` se apoya en JSON; conviene blindar la lectura.
- BR-014: el rechazo desde formalización es polimórfico en código pero cada tipo tiene su método; considerar un helper `XServicio.RegistrarRechazoFormalizacion` factoría para evitar divergencia futura.
