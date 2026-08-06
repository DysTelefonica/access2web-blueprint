# Capacidad: Logs, auditoría y snapshots

## §0 Identidad

- **ID de capacidad**: CAP-009
- **Tier**: critical
- **Estado**: active con deuda de pruebas focal, cache de visor y reconciliación UI/layout
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary` sobre logs, snapshots, error handling, migraciones y formularios de auditoría. Código principal y code-behind `.cls` están `matched`; `SnapshotServicio`/`SnapshotHelper` solo `caseOnly`; los cuatro `.form.txt` de logs tienen diferencias accionables `bothChanged`.
- **Confianza global**: mayoritariamente `Verified-static`. La estructura está bien hecha y la transacción atómica es correcta, pero no hay manifest focal `access-vba-tdd` v2.4.2, los consultores UI son formularios legacy y sus layouts están divergentes. La promoción a `Verified-runtime` requiere suite focal y reconciliación UI.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). Logs y auditoría necesitan `BeginTestSession`/`EndTestSession`, `DAO.Database` explícito, cardinalidad de mutaciones y manifests atómicos.

**Contrato TDD vigente**: las pruebas de logs deben migrarse a `access-vba-tdd` v2.4.2 — `Public Function` con retorno JSON canónico, fixture propio, schema-first, `DAO.Database` inyectado, cardinalidad `countBefore`/`countAfter` sobre `tbLogCambios`/`tbLogErrores`/`tbLogEstados`, manifests atómicos.

**Justificación del nivel**: crítico, porque logs y snapshots son el ancla de auditoría legal y de subsanación tras rechazo. Sin logs no hay trazabilidad de cambio.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: registrar todos los cambios técnicos (`tbLogCambios`), los errores en tiempo de ejecución (`tbLogErrores`) y las transiciones de estado (`tbLogEstados`) de una solicitud, además de producir snapshots JSON para detectar cambios tras rechazo y soportar el ciclo de validación del RAC.
- **Usuarios / perfiles**: auditores, calidad, técnicos, administradores y todos los roles con acceso a la solicitud.
- **Problema que resuelve**: centraliza la trazabilidad y permite reconstruir el estado de una solicitud en cualquier punto, detectar cambios del técnico tras un rechazo y mostrar el historial al usuario final.
- **Valor de negocio**: auditoría legal, detección de cambios, base del ciclo de validación del RAC (Spec-105), soporte del visualizador web.
- **No-objetivos**: este documento no cubre la captura de datos técnicos (CAP-003/CAP-004/CAP-005/CAP-001), la gestión de adjuntos (CAP-008) ni la búsqueda (CAP-006). Sí los vincula.
- **Origen de la intención**: código actual y SDD `pcsub-staging-recovery` (consolidación de logs).
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** cualquier operación de servicio que muta una solicitud **CUANDO** se ejecuta **ENTONCES** `LogServicio.RegistrarCambio(tabla, registroID, campo, valorAnterior, valorNuevo, tipoOperacion, db)` persiste en `tbLogCambios` con `fechaHora = Now`, `usuario = m_ObjUsuarioActivo.nombre`, `suplantadoPor` (si impersonación) y valida los cinco campos. **Estado**: `Verified-static`.
- **DADO** un error en tiempo de ejecución **CUANDO** el servicio lo captura con `CondorError` **ENTONCES** `LogServicio.RegistrarError` lo persiste en `tbLogErrores` con `modulo`, `procedimiento`, `numeroError`, `descripcionError`, `contexto`. Si el registro del propio error falla, se imprime en `Debug.Print` para evitar bucle. **Estado**: `Verified-static`.
- **DADO** una transición de estado exitosa **CUANDO** se ejecuta `WorkflowServicio.EjecutarTransicion` **ENTONCES** se persiste en `tbLogEstados` con `idEstadoAnterior`, `idEstadoNuevo`, `fechaTransicion`, `usuarioTransicion` y se registra adicionalmente un `LogCambio` de tipo `TRANSICIÓN_WF`. **Estado**: `Verified-static`.
- **DADO** un rechazo activo y un guardado de Impacto/Propuesta en `estadoValidacion` **CUANDO** el usuario guarda **ENTONCES** se persiste un delta JSON por campo bajo `JsonHelper.RegistrarCambio` para que `VerificarCambiosTrasRechazo` pueda compararlo. **Estado**: `Verified-static`.
- **DADO** una solicitud en `estadoModificacion` con `revisionCalidadEstado = "RECHAZADO"` **CUANDO** se necesita detectar cambios del técnico **ENTONCES** `SnapshotServicio.CalcularHashSnapshot` produce un SHA-256 del JSON de `tbSolicitudes` + `tbDatosPC` + `tbDatosCDCA` + `tbDatosCDCASUB` excluyendo campos de auditoría. **Estado**: `Verified-static`.
- **DADO** un usuario consultando logs **CUANDO** abre `frmVerLogCambios` o `frmVerLogErrores` con filtros **ENTONCES** el servicio `getLogs`/`getErrores` devuelve los registros ordenados por `fechaHora DESC`. **Estado**: `Verified-static`; el formulario es legacy.
- **DADO** una entrada de log con `idLogCambio` o `idLogError` **CUANDO** el usuario abre el detalle **ENTONCES** `frmDetalleLogCambio`/`frmDetalleLogError` muestra valor anterior y nuevo. **Estado**: `Verified-static`.
- **DADO** un ciclo de validación nuevo **CUANDO** se llega a `estadoValidacion` **ENTONCES** `ValidacionRevisionRepositorio.CrearNuevoCicloValidacion` registra el ordinal y el hash inicial (vacío) en `tbValidacionRevision`. **Estado**: `Verified-static`.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | `LogServicio.RegistrarCambio` exige `usuario`, `tabla`, `campo` no vacíos y `registro > 0`. | Código | Sí: `LogServicio.ValidarCambio`. | Pendiente. | Verified-static |
| BR-002 | `LogServicio.RegistrarError` exige `usuario`, `modulo`, `procedimiento` no vacíos. Si el registro del error falla, `Debug.Print`. | Código | Sí: `LogServicio.ValidarError` y bloque `Errores` (líneas 83-89). | Pendiente. | Verified-static |
| BR-003 | `LogEstado` se persiste en `tbLogEstados` con la fecha de la solicitud, no `Now()`, para mantener consistencia temporal. | Código | Sí: `WorkflowServicio.EjecutarTransicion` línea 1361. | Pendiente. | Verified-static |
| BR-004 | `SnapshotServicio` excluye campos `fechacreacion`, `fechamodificacion`, `usuariocreacion`, `usuariomodificacion`, `idusuariocreacion`, `idusuariomodificacion` del hash. | Código (Spec-105) | Sí: `SnapshotServicio.EsCampoExcluido` líneas 91-101. | Pendiente. | Verified-static |
| BR-005 | El hash de snapshot se calcula con `HashHelper.CalcularSHA256(json)`; la `tbValidacionRevision` guarda `hashDatos` y `HashDatos` se usa para detectar cambios. | Código (Spec-105) | Sí: `SnapshotServicio.CalcularHashSnapshot` línea 55. | Pendiente. | Verified-static |
| BR-006 | `LogCambioServicio.LimpiarHistorialCambios` borra todos los cambios de una solicitud (usado en `SolicitudServicio.EliminarSolicitudCompleta`). | Código | Sí: `LogCambioServicio.cls` líneas 19-32. | Pendiente. | Verified-static |
| BR-007 | `LogEstadoRepositorio.getUltimoEstadoAnterior(idSolicitud, db)` se usa en `ReabrirSolicitudCerrada` para calcular el estado anterior. | Código | Sí: `WorkflowServicio.ReabrirSolicitudCerrada` línea 1984. | Pendiente. | Verified-static |
| BR-008 | `ErrorLogger.bas` y `ErrorPresenter.bas` centralizan la captura y presentación de errores; `CondorError.cls` es el modelo canónico con `AddToCallStack` para preservar la traza. | Código | Sí: `ErrorLogger.bas` y `CondorError.cls`. | Pendiente. | Verified-static |
| BR-009 | `modActualizaciones.bas` ofrece un punto de migración de esquemas; las actualizaciones se aplican en código. | Código | Sí: `modActualizaciones.bas`. | Pendiente. | Verified-static |
| BR-010 | `getLogs` y `getErrores` filtran por `usuario`, `tabla`/`modulo` y `registro`/`procedimiento` con `LIKE` parametrizado. | Código | Sí: `LogServicio.cls` líneas 165-191. | Pendiente. | Verified-static |
| BR-011 | `tbValidacionRevision` registra `hashDatos` para que el visualizador pueda determinar si hay cambios entre la versión exportada y la actual. | Código (Spec-105) | Sí: `SnapshotHelper.CalcularHashFaseValidacion` y `WorkflowServicio.ObtenerHayCambiosValidacion`. | Pendiente. | Verified-static |

### Validaciones observadas

- `LogCambio` con `registro <= 0` produce error 513.
- `LogError` con `procedimiento` vacío produce error 513.
- `p_RegistroID` no numérico se ignora (no se añade a `params`).

### Transiciones de estado y navegación

- `LogCambio` y `LogEstado` se escriben como efecto secundario de las transiciones, sin ser transiciones ellos mismos.
- El visualizador web (`GenerarHTML_VisualizadorDeEstado`) consume `LogEstado` para pintar el timeline y `LogCambio` para mostrar el histórico.

### Casos límite y hallazgos

- `LogServicio.RegistrarError` no aborta la operación en curso si el registro falla (se imprime `Debug.Print`). Es una política defensiva correcta, pero deja la app sin log persistente del error; conviene un canal secundario.
- `SnapshotServicio` lee directamente con `getdb()` (`db.OpenRecordset(sql, dbOpenSnapshot)`) sin inyección de `db`. Es un seam roto: la prueba focal no puede inyectar sandbox.
- `LogEstadoServicio.cls` estaba listado en una versión previa del documento, pero no existe ni en `src/` ni en el proyecto VBA; la capacidad opera con `LogEstado.cls` y `LogEstadoRepositorio.bas`.
- `modActualizaciones.bas` es código de migración ad-hoc, no un sistema de migraciones con versionado. Conviene formalizarlo.

### Señales de aceptación / presencia

- Existen `LogCambio.cls`, `LogCambioServicio.cls`, `LogCambioRepositorio.bas`, `LogError.cls`, `LogServicio.cls`, `LogErrorRepositorio.bas`, `LogEstado.cls`, `LogEstadoRepositorio.bas`. No existe `LogEstadoServicio.cls`.
- Existen `SnapshotServicio.cls`, `SnapshotHelper.bas`, `ErrorLogger.bas`, `ErrorPresenter.bas`, `CondorError.cls`, `modActualizaciones.bas`.
- Existen `Form_frmVerLogCambios.cls`, `Form_frmVerLogErrores.cls`, `Form_frmDetalleLogCambio.cls`, `Form_frmDetalleLogError.cls`; su code-behind está sincronizado, pero sus `.form.txt` tienen deriva accionable `bothChanged`.
- `tbLogCambios` se persiste por `idLogCambio`; `tbLogErrores` por `idLogError`; `tbLogEstados` por `idLogEstado`; `tbValidacionRevision` por `Id`.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frmVerLogCambios` y `Form_frmDetalleLogCambio` para auditoría de cambios técnicos.
  - `Form_frmVerLogErrores` y `Form_frmDetalleLogError` para auditoría de errores.
  - El visualizador web `frmWebVisor` consume indirectamente `LogEstado` y `LogCambio` vía `WorkflowServicio.GenerarHTML_VisualizadorDeEstado`.
- **Puntos de entrada de código**:
  - `LogServicio.RegistrarCambio`, `RegistrarError`, `getLogs`, `getErrores`, `getLogCambioPorID`, `getLogErrorPorID`, `ValidarCambio`, `ValidarError`, `getLogsSQL`, `getErroresSQL`.
  - `LogCambioServicio.LimpiarHistorialCambios`.
  - `LogEstadoRepositorio` (consulta + persistencia de estados; no existe una clase `LogEstadoServicio`).
  - `LogEstadoRepositorio.getHistorialPorIdSolicitud`, `getUltimoEstadoAnterior`, `EliminarPorIdSolicitud`, `Guardar`.
  - `LogCambioRepositorio.GuardarLog`, `getLogCambioPorID`, `EliminarPorSolicitud`.
  - `LogErrorRepositorio.GuardarError`, `getLogErrorPorID`.
  - `SnapshotServicio.GenerarSnapshotGlobal`, `CalcularHashSnapshot`, `ObtenerDatosTabla`, `EsCampoExcluido`.
  - `SnapshotHelper.CalcularHashFaseValidacion`.
  - `ErrorLogger`, `ErrorPresenter`, `CondorError` (con `AddToCallStack`).
- **Datos afectados**:
  - `tbLogCambios`: `idLogCambio`, `fechaHora`, `usuario`, `suplantadoPor`, `tabla`, `registro`, `campo`, `valorAnterior`, `valorNuevo`, `tipoOperacion`.
  - `tbLogErrores`: `idLogError`, `fechaHora`, `usuario`, `suplantadoPor`, `modulo`, `procedimiento`, `numeroError`, `descripcionError`, `contexto`.
  - `tbLogEstados`: `idLogEstado`, `idSolicitud`, `idEstadoAnterior`, `idEstadoNuevo`, `fechaTransicion`, `usuarioTransicion`.
  - `tbValidacionRevision`: `Id`, `idSolicitud`, `ordinal`, `idAdjunto`, `fechaEnvio`, `fechaRecepcion`, `Resultado`, `Comentarios`, `Usuario`, `HashDatos`.
- **Dependencias**:
  - `m_ObjUsuarioActivo`, `m_ObjUsuarioReal`, `g_blnImpersonando` (globals).
  - `JsonHelper` para `JsonHelper.ObtenerJsonRechazo` y `JsonHelper.GuardarJsonRechazo`.
  - `HashHelper.CalcularSHA256` para snapshot.
  - `RepositorioComun.HidratarColeccionDesdeSQL` para `getLogs`/`getErrores`.
- **Sincronización fuente↔binario**: si se modifican `Log*` o `Snapshot*`, basta `dysflow.import_code`. Los formularios son consultores de solo lectura.
- **Valoración de diseño (tal-como-está vs ideal)**: el modelo es razonable y la transacción es atómica. La deuda principal está en (a) `SnapshotServicio` con `getdb()` directo (seam roto), (b) `LogServicio.RegistrarError` con `Debug.Print` como fallback (sin canal secundario), (c) falta de suite focal de logs que cumpla v2.4.2, (d) `modActualizaciones` sin sistema de migraciones. La pieza está bien hecha para el día a día; no se recomienda `Verified-runtime` sin seam.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `LogCambio.cls`, `LogCambioServicio.cls`, `LogCambioRepositorio.bas`, `LogError.cls`, `LogServicio.cls`, `LogErrorRepositorio.bas`, `LogEstado.cls`, `LogEstadoServicio.cls`, `LogEstadoRepositorio.bas`.
2. Restaurar `SnapshotServicio.cls` y `SnapshotHelper.bas`.
3. Restaurar `ErrorLogger.bas`, `ErrorPresenter.bas`, `CondorError.cls`, `modActualizaciones.bas`.
4. Confirmar esquemas `tbLogCambios` (10 columnas), `tbLogErrores` (9 columnas), `tbLogEstados` (6 columnas), `tbValidacionRevision` (10 columnas).
5. Restaurar `Form_frmVerLogCambios.cls`, `Form_frmVerLogErrores.cls`, `Form_frmDetalleLogCambio.cls`, `Form_frmDetalleLogError.cls`.
6. Importar con `dysflow.import_modules` y compilar con `dysflow.compile_vba`. Verificar binario con `dysflow.verify_binary`; actualmente los cuatro formularios de logs tienen `.form.txt` `bothChanged`.
7. Demostrar los escenarios de §2 con un manifest atómico `tests/testsLogs.json` que cubra: cambio, error, transición, hash de snapshot, `HayCambiosValidacion`. **Slice B5 (2026-06-15, commit `63d4116`)**: manifest focal con dos átomos strict TDD v2.4.2 para `LogServicio.RegistrarCambio` (con `db` inyectado, cardinalidad `countBefore=0` → `countAfter=1`, verificación de `campo` y `valorNuevo` persistidos) y `LogServicio.RegistrarError` (firma real sin `db`, cardinalidad global, verificación de `modulo`/`procedimiento`/`numeroError`). Pendientes: transición (`tbLogEstados`), hash de snapshot y `HayCambiosValidacion`.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/LogCambio.cls`, `LogCambioServicio.cls` (33 líneas), `LogError.cls`, `LogServicio.cls` (226 líneas), `LogEstado.cls`.
  - `src/modules/LogCambioRepositorio.bas`, `LogErrorRepositorio.bas`, `LogEstadoRepositorio.bas`.
  - `src/classes/SnapshotServicio.cls` (103 líneas), `src/modules/SnapshotHelper.bas`.
  - `src/modules/ErrorLogger.bas`, `ErrorPresenter.bas`, `src/classes/CondorError.cls`, `src/modules/modActualizaciones.bas`.
  - `src/forms/Form_frmVerLogCambios.cls`, `Form_frmVerLogErrores.cls`, `Form_frmDetalleLogCambio.cls`, `Form_frmDetalleLogError.cls`.
  - `docs/ERD/condor_datos.md` → `tbLogCambios`, `tbLogErrores`, `tbLogEstados`, `tbValidacionRevision`.
- **Evidencia Dysflow incorporada**:
  - Primer `dysflow.verify_binary` falló porque se pidió el módulo inexistente `LogEstadoServicio`.
  - Reintento con módulos reales: `LogCambio*`, `LogError*`, `LogEstado`, `LogEstadoRepositorio`, `LogServicio`, `CondorError`, `ErrorLogger`, `ErrorPresenter`, `modActualizaciones` y code-behind de formularios `matched`.
  - `SnapshotServicio` y `SnapshotHelper` solo diferencias no accionables `caseOnly`.
  - `Form_frmVerLogCambios.form.txt`, `Form_frmVerLogErrores.form.txt`, `Form_frmDetalleLogCambio.form.txt` y `Form_frmDetalleLogError.form.txt` están `bothChanged` y requieren reconciliación UI.
- **Tests existentes**: manifest atómico `tests/testsLogs.json` (commit `63d4116`, 2026-06-15, Slice B5). Cubre `LogServicio.RegistrarCambio` (db inyectado, cardinalidad, persistencia de `campo` y `valorNuevo`) y `LogServicio.RegistrarError` (firma real del servicio sin `db` parametrizable, cardinalidad global, persistencia de `modulo`/`procedimiento`/`numeroError`).
- **Evidencia runtime Dysflow**:
  - `Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields`: **VERDE** 3.7 s, sandbox `condor_datos.accdb` local, `idSolicitud=900631`, `idEstadoInterno` 2→8, `tipoOperacion=UPDATE`.
  - `Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields`: **VERDE** 3.8 s, sandbox `condor_datos.accdb` local, `idLogError` calculado por `LogErrorRepositorio.getSiguienteIDLogErrores(MAX+1)`, `numeroError=513`, `modulo=Test_Logs_Strict`, `procedimiento=RegistrarError_InsertsRow`.
- **SDD/intención consultada**: Spec-105 (hash de snapshot), código actual.

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| `LogServicio.RegistrarCambio` no persiste | `ValidarCambio` falla o `LogCambioRepositorio.GuardarLog` rompe. | prueba focal con cardinalidad. | §2 BR-001 |
| `RegistrarError` no aparece en `tbLogErrores` | Fallo en `LogErrorRepositorio.GuardarError` o `Debug.Print` consumido. | prueba focal. | §2 BR-002 |
| `SnapshotServicio.CalcularHashSnapshot` no detecta cambios tras rechazo | Campos excluidos demasiado amplios o error de serialización JSON. | prueba con solicitud modificada. | §2 BR-004/BR-005 |
| `WorkflowServicio.EjecutarTransicion` no registra `tbLogEstados` | `LogEstadoRepositorio.Guardar` no se invoca. | prueba con cardinalidad. | §2 BR-003 |
| `HayCambiosValidacion` siempre devuelve `NO` o `V1` | `tbValidacionRevision.HashDatos` no se actualiza. | prueba con transición Validación → modificación → Validación. | §2 BR-011 |
| `ReabrirSolicitudCerrada` no encuentra estado anterior | `LogEstadoRepositorio.getUltimoEstadoAnterior` devuelve 0. | prueba con historial. | §2 BR-007 |

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Logs, auditoría y snapshots | Pendiente | Pendiente de confirmación | pending | Pendiente | Pendiente | Manifest atómico parcial (cambio+error) commit `63d4116` Slice B5; pendiente transición, hash de snapshot, `HayCambiosValidacion` y reconciliación de layout `.form.txt` de formularios de logs. |

## §6 Notas de migración web

- **Conservar**: registro de `tbLogCambios`/`tbLogErrores`/`tbLogEstados` como efecto secundario, hash de snapshot para detectar cambios tras rechazo, ciclo de validación en `tbValidacionRevision`, modelo `CondorError` con `AddToCallStack`.
- **Transformar**: `Debug.Print` fallback a un canal de observabilidad (Sentry, Application Insights); `getdb()` singleton a un ORM/contexto; `SnapshotServicio.GenerarSnapshotGlobal` con `getdb()` a un repositorio inyectable.
- **NO copiar**: `modActualizaciones` ad-hoc, dependencia de `m_ObjUsuarioActivo` global, `SnapshotServicio` con `getdb()` directo, falta de canal secundario para `LogServicio.RegistrarError` cuando la BD falla.
- **Preguntas abiertas**: ¿El hash de snapshot debe ser SHA-256 o un hash más barato (e.g. xxHash) para minimizar coste en cada transición? (equipo técnico). ¿La política de "no abortar si falla el log" debe revisarse para errores críticos? (responsable de calidad).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| `LogServicio.RegistrarCambio` valida y persiste `tbLogCambios` con los cinco campos mínimos. | Verified-static | `LogServicio.cls` líneas 18-48. | 2026-06-15 |
| `LogServicio.RegistrarError` cae a `Debug.Print` si el registro del error falla. | Verified-static | `LogServicio.cls` líneas 83-89. | 2026-06-15 |
| `WorkflowServicio.EjecutarTransicion` persiste `tbLogEstados` con `fechaTransicion = sol.fechaModificacion`. | Verified-static | `WorkflowServicio.cls` línea 1361. | 2026-06-15 |
| `SnapshotServicio` excluye los seis campos de auditoría del hash. | Verified-static | `SnapshotServicio.cls` líneas 91-101. | 2026-06-15 |
| `SnapshotServicio.GenerarSnapshotGlobal` lee con `getdb()` directamente (seam roto). | Verified-static / deuda | `SnapshotServicio.cls` línea 68. | 2026-06-15 |
| `LogCambioRepositorio.GuardarLog` persiste por `LogCambioRepositorio` (módulo DAO). | Verified-static | `LogCambioRepositorio.bas`. | 2026-06-15 |
| `tbValidacionRevision.HashDatos` se usa para detectar cambios entre versión exportada y actual. | Verified-static | `WorkflowServicio.ObtenerHayCambiosValidacion` y `SnapshotHelper.CalcularHashFaseValidacion`. | 2026-06-15 |
| `ReabrirSolicitudCerrada` usa `LogEstadoRepositorio.getUltimoEstadoAnterior`. | Verified-static | `WorkflowServicio.cls` línea 1984. | 2026-06-15 |
| `modActualizaciones.bas` aplica migraciones ad-hoc. | Verified-static / deuda | `modActualizaciones.bas`. | 2026-06-15 |
| `LogEstadoServicio.cls` existe como parte de la capacidad. | Divergent / documentación corregida | `dysflow.verify_binary`: `VBA_MODULE_NOT_FOUND`; `src/` solo contiene `LogEstado.cls` y `LogEstadoRepositorio.bas`. | 2026-06-15 |
| Formularios de logs están reconciliados fuente↔binario. | Divergent / blocker UI | `verify_binary`: cuatro `.form.txt` de logs con `bothChanged`; code-behind `.cls` matched. | 2026-06-15 |
| Existe un manifest atómico de pruebas de logs que cumpla `access-vba-tdd` v2.4.2. | Verified-runtime (parcial) / pendiente transición+snapshot | `tests/testsLogs.json` con dos átomos verdes (`RegistrarCambio` 3.7 s, `RegistrarError` 3.8 s, commit `63d4116`). Pendientes: `tbLogEstados`, hash de snapshot, `HayCambiosValidacion`. | 2026-06-15 |

**Divergencias pendientes de revisión humana**:

- BR-002: la política de "no abortar si falla el log" deja la app sin traza. Revisar con producto si esto es aceptable o si conviene un canal secundario.
- BR-004/BR-005: `SnapshotServicio` con `getdb()` directo rompe la inyección de `db`. Mover a un seam testeable.
- `modActualizaciones`: considerar migrar a un sistema de migraciones con versionado explícito.
- Los cuatro formularios de logs tienen deriva UI/layout en `.form.txt`; no importar/exportar layouts hasta revisar cuál lado es fuente de verdad.
