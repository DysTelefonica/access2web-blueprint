<!--
Documento de capacidad CAP-002 — Edición de expediente.
Linaje: PRUEBA-002 spec + dysflow_get_schema + lectura ExpedienteOperaciones.cls (líneas 19, 566, 811, 855) + Form_FormExpediente.cls (línea 133)
Versión corregida tras critical review 2026-06-15. La spec PRUEBA-002 estaba errada: el form de edición NO llama `Registrar`; solo llama `RegistrarExpEntidades` (UPDATE de entidades-derivadas, no de cabecera).
Idioma: castellano de España. Identificadores de código/test se mantienen tal cual.
-->

# Capacidad: Edición de expediente

## §0 Identidad
- **ID de capacidad**: CAP-002
- **Tier**: critical
- **Estado**: active
- **Source**: hybrid (SDD + código actual)
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: 2026-06-15 mediante lectura de `src/classes/ExpedienteOperaciones.cls` (líneas 19, 566, 811, 855-1122) y `src/forms/Form_FormExpediente.cls` (línea 133)
- **Confianza global**: mixta — ver §7 (entry point verificado por lectura; sin tests verdes específicos de edición; algunas divergencias con la spec PRUEBA-002 §3.2)

## §1 Intención de negocio (≈ proposal SDD) — POR QUÉ
- **Propósito**: actualizar los datos derivados de un expediente existente en `TbExpedientesConEntidades` (cadenas concatenadas de Comerciales, CPVs, Lugares, PECAL, RACs, Responsables, etc.) preservando la cabecera en `TbExpedientes`. La edición NO toca la fila principal — solo la fila de "entidades-derivadas" — porque los datos de cabecera son trazabilidad/audit, no edición cotidiana.
- **Usuarios / perfiles**: gestor de área (edición cotidiana de entidades vinculadas); administrador (corrección de cadenas, vinculaciones rotas).
- **Problema que resuelve**: la cabecera del expediente es trazabilidad pura (quién/qué/cuándo se dio de alta); los datos derivados se reescriben muchas veces durante la vida del expediente (nuevo comercial, nuevo RAC, reasignación de responsable). Sin esta separación, cada edición reescribiría la cabecera y rompería el audit-trail.
- **Valor de negocio**: edición sin tocar audit-trail, recálculo controlado de cadenas (`CadenaComerciales`, `CadenaJPs`, `CadenaRACs`, etc.) en una sola fila, validación exhaustiva de unicidad solo cuando el campo cambió.
- **No-objetivos**: edición de la cabecera `TbExpedientes` (CAP-001 alta crea la cabecera; cambios estructurales requieren re-alta con `RegistrarCambioTipo`, no edición); edición masiva; cambio de tipo (`ExpedienteOperaciones.RegistrarCambioTipo`, línea 2266, CAP-004); edición de anexos (CAP-018).
- **Origen de la intención**: spec PRUEBA-002 §3.2 (BR-02-01..02) + lectura del código.
- **Referencia de tracker de origen**: PRUEBA-002 PR-C, PRUEBA-003 REFAC-3a (pendiente).

## §2 Contrato de comportamento (≈ spec SDD) — QUÉ ⟵ ANCLA DE REGRESIÓN

### Escenarios (Dado / Cuando / Entonces)
- **DADO** un DTO con cambios en entidades vinculadas (e.g., nuevo Comercial) y `p_ExpedienteAlInicio` con los valores previos **CUANDO** se invoca `ExpedienteOperaciones.RegistrarExpEntidades(p_Error, p_db, p_Ambito, p_ForceCacheFailureForTest)` (línea 811) **ENTONCES** se abre transacción DAO, se delega a `ExpedienteEntidadOperaciones.Registrar` con el `p_Ambito` (default `Todo` = recalcular todas las cadenas), UPDATE de `TbExpedientesConEntidades` (1 fila, FK = `IDExpediente`), INSERT en `TbUltimoCambio` con `FechaCambio = Now()` y `IDUsuarioCambio = g_UsuarioConectado`, COMMIT, retorna el `IDExpediente`.
- **DADO** un DTO con `Titulo = ""` **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** la cadena interna de validación (`MotivoNoOK`) rechaza con `p_Error = "El título del expediente es obligatorio"`. (Validación 1 de `MotivoNoOK`, línea 874).
- **DADO** un DTO con `EsAM = "Sí"` y `EsLote = "Sí"` **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** `p_Error = "No puede ser Acuerdo Marco y Lote a la vez"`. (Validación 3, línea 914).
- **DADO** un DTO con `AGEDYSAplica` cambiando de "Sí" a "No" cuando ya hay DPDs vinculados (`CadenaDPDs <> ""`) **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** `p_Error = "No se puede quitar AGEDYS: ya tiene DPDs vinculados."`. (Validación 5, línea 976).
- **DADO** un DTO con `Titulo` cambiado a un valor que ya existe en otro expediente **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** `p_Error = "El Título ya existe en el Expediente ID: <ID>"`. (Validación de unicidad condicional, línea 987-996 — solo consulta BD si el campo CAMBIÓ respecto a `p_ObjExpedienteAlInicio`).
- **DADO** un DTO con `TienePECAL = Sí` y `Adjudicado = "Sí"` y `ResponsableCalidad Is Nothing` **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** `p_Error = "Expediente adjudicado con PECAL requiere un Responsable de Calidad"`. (Validación 8, línea 1075).
- **DADO** un DTO con `AGEDYSAplica = "Sí"`, `AGEDYSGenerico <> "Sí"`, y `ColResponsables` sin Jefe de Proyecto **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** `p_Error = "Para Agedys se requiere al menos un Jefe de Proyecto."`. (Validación 10, línea 1092 — solo cuando `IDExpediente = ""` = solo alta).
- **DADO** un DTO con `ColAnexos` que tiene una ruta no existente como archivo **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** `p_Error = "El anexo no es accesible: <ruta>"`. (Validación 11, línea 1107 — `fso.FileExists(m_ID)`).
- **DADO** un DTO con una FK violation **CUANDO** se invoca `RegistrarExpEntidades` **ENTONCES** se hace rollback y `p_Error` contiene `Err.Description`.

### Reglas de negocio

| ID regla | Enunciado (leído de `MotivoNoOK`, líneas 855-1122) | Autoridad | ¿Aplicada en código? | Prueba | Confianza |
|---|---|---|---|---|---|
| BR-02-01 | `Titulo` es obligatorio y, si cambió, único en `TbExpedientes.Titulo` | spec + código | Sí — `MotivoNoOK` líneas 874 + 987-996 (consulta condicional) | `Test_Helper_ExpedienteEdicion_ValidarEdicion_TituloVacio_DevuelveFalse` + `Test_Helper_ExpedienteEdicion_ValidarEdicion_TituloDuplicado_DevuelveFalse` (target PR-C) | Verified-static |
| BR-02-02 | `EsAM`/`EsLote`/`EsExpediente`/`EsBasado` son excluyentes; al menos uno "Sí" | spec + código | Sí — líneas 878-929 | `Test_Helper_ExpedienteEdicion_ValidarEdicion_TipologiaInvalida_DevuelveFalse` (target PR-C) | Verified-static |
| BR-02-03 | `APLICAESTADO` ∈ {"Sí","No"}; si "Sí" entonces `FechasIrregulares` debe ser "No\|..." | spec + código | Sí — líneas 894-906 | `Test_Helper_ExpedienteEdicion_ValidarEdicion_FechasIrregulares_DevuelveFalse` (target PR-C) | Verified-static |
| BR-02-04 | `AGEDYSAplica` + `AGEDYSGenerico` (si AGEDYS aplica) | spec + código | Sí — líneas 958-967 | idem CAP-001 BR-01-03 | Verified-static |
| BR-02-05 | `AGEDYSAplica` no puede cambiar de "Sí" a "No" si hay DPDs vinculados | spec + código | Sí — líneas 970-980 | `Test_Helper_ExpedienteEdicion_ValidarEdicion_QuitarAGEDYSConDPDs_DevuelveFalse` (target PR-C) | Verified-static |
| BR-02-06 | `TIpo` (campo `Tipo`) obligatorio; autocalculado si vacío | spec + código | Sí — líneas 932-943 | `Test_Helper_ExpedienteEdicion_ValidarEdicion_SinTipo_DevuelveFalse` (target PR-C) | Verified-static |
| BR-02-07 | `CodExp` obligatorio salvo Lote; si cambió, único | spec + código | Sí — líneas 945-948 + 999-1010 | idem CAP-001 BR-01-05 | Verified-static |
| BR-02-08 | `CodExpLargo`, `Nemotecnico`, `NPedido`, `CodProyecto` (cuando presentes) deben ser únicos **si cambiaron** | spec + código | Sí — líneas 1012-1066 (consulta condicional) | `Test_Helper_ExpedienteEdicion_ValidarEdicion_UnicidadCondicional_*` (target PR-C) | Verified-static |
| BR-02-09 | `Ambito` ∈ {Defensa, Fuera, HPS} | spec + código | Sí — líneas 1068-1071 | idem CAP-001 BR-01-01 | Verified-static |
| BR-02-10 | PECAL con Adjudicado="Sí" requiere `ResponsableCalidad` no nulo | spec + código | Sí — líneas 1073-1078 | `Test_Helper_ExpedienteEdicion_ValidarEdicion_PECALSinResponsable_DevuelveFalse` (target PR-C) | Verified-static |
| BR-02-11 | Anexos: cada ruta (no ID numérico) debe existir como archivo (`fso.FileExists`) | spec + código | Sí — líneas 1102-1113 | `Test_Helper_ExpedienteEdicion_ValidarEdicion_AnexoNoExiste_DevuelveFalse` (target PR-C) | Verified-static |
| BR-02-12 | `RegistrarExpEntidades` con DTO válido hace UPDATE de `TbExpedientesConEntidades` (1 fila) + INSERT en `TbUltimoCambio` dentro de transacción DAO | spec + código | Sí — `ExpedienteEntidadOperaciones.Registrar` (delegada) | `Test_Helper_ExpedienteEdicion_Actualizar_DTOValido_ActualizaYDevuelveTrue` (target PR-C) | Verified-static |
| BR-02-13 | `RegistrarExpEntidades` hace rollback ante FK violation | spec + código | Sí — transacción DAO | `Test_Helper_ExpedienteEdicion_Actualizar_FKViolation_Rollback` (target PR-C) | Verified-static |
| BR-02-14 (subregla) | Dirty detection implícito: la unicidad se valida solo si el campo cambió (línea 987, 1000, etc.) | spec implícito | Sí — `MotivoNoOK` consulta condicional | `Test_Helper_ExpedienteEdicion_ValidarEdicion_NoConsultaSiNoCambio_*` (target PR-C) | Verified-static |
| BR-02-15 (subregla) | AGEDYS + responsables + JefeProyecto/CorreoSiempre: solo se valida en alta (`IDExpediente = ""`) | spec implícito | Sí — línea 1081 (guard `If .IDExpediente = ""`) | cubierto por CAP-001 BR-01-08 | Verified-static |
| BR-02-16 (subregla) | Edición registra `UltimoCambio` con `FechaCambio = Now()` y `IDUsuarioCambio = g_UsuarioConectado` | spec + código | Sí — `ExpedienteEntidadOperaciones.Registrar` (transaccional) | `Test_Helper_UltimoCambio_Registrar_InsertaUsuarioYFecha` (target PR-F) | Verified-static |

### Validaciones
- 1. Campos obligatorios: `Titulo`, `EsAM`, `EsLote`, `EsExpediente`, `EsBasado`, `APLICAESTADO`, `TIpo`, `Ambito`, `AGEDYSAplica`, `AGEDYSGenerico` (si AGEDYS), `IDOrganoContratacion`
- 2. Coherencia de fechas: `FechasIrregulares` no contiene "Sí|..." cuando `APLICAESTADO = "Sí"`
- 3. Exclusividad de tipo: AM, Lote, Expediente, Basado son excluyentes; al menos uno
- 4. Tipología: `IDExpedientePadre = ""` para AM; permite Lote con padre
- 5. Coherencia AGEDYS: no quitar si hay DPDs
- 6. Unicidad condicional: Titulo, CodExp, CodExpLargo, NPedido, Nemotecnico, CodProyecto — solo se valida si CAMBIÓ respecto al `p_ObjExpedienteAlInicio`
- 7. PECAL adjudicado: requiere ResponsableCalidad
- 8. AGEDYS sin genérico: requiere JefeProyecto + CorreoSiempre (solo alta)
- 9. Anexos: cada ruta debe existir como archivo

### Transiciones de estado
- No aplica cambio de estado explícito; el ESTADO se recalcula por `Expediente.ESTADOCalculadoTexto` (línea 68 de `Registrar`) a partir de las fechas y `APLICAESTADO`.

### Casos límite y de error
- Concurrencia: dos ediciones simultáneas pueden pasar la validación de unicidad pre-UPDATE pero colisionar en el UPDATE mismo → `Err.Description` + rollback
- `Ordinal` no numérico → `p_Error = "El ordinal ha de ser un valor numérico"` (línea 952)
- `FechasIrregulares` con formato inesperado → propagado por `Expediente.Error` (línea 934)

### Señales de aceptación / presencia ⟵ cómo saber que la funcionalidad EXISTE y funciona
- `ExpedienteOperaciones.RegistrarExpEntidades` existe en el binario
- `ExpedienteEntidadOperaciones.Registrar` existe en el binario
- `Form_FormExpediente.ComandoActualizarCompleto_Click` invoca `m_ExpOp.RegistrarExpEntidades p_Error:=m_Error` (línea 148)
- `dysflow.test_vba` con filtro `expediente-edicion` retorna `ok: true` con 5+/5+ PASS (target post-PRUEBA-002 PR-C)
- `dysflow.count_rows` antes/después de UPDATE: `TbExpedientesConEntidades` no cambia en count (UPDATE de 1 fila, no INSERT), `TbUltimoCambio` aumenta en 1

## §3 Mapa de implementación (≈ design SDD) — CÓMO
- **Puntos de entrada de UI**:
  - `Form_FormExpediente.cls` → `ComandoActualizarCompleto_Click` (línea 133) → `m_ExpOp.RegistrarExpEntidades p_Error:=m_Error` (línea 148)
  - El form hace `With m_ExpOp: Set .Expediente = m_ObjExpedienteDTOActivo.Expediente: .RegistrarExpEntidades p_Error:=m_Error` (líneas 146-152)
- **Puntos de entrada de código** (clase `ExpedienteOperaciones`):
  - `Public Function Registrar(p_DTO As ExpedienteDTO, p_ExpedienteAlInicio As Expediente, Optional ByRef p_Error As String) As String` (línea 19) — entry comprensivo; `p_ExpedienteAlInicio Is Nothing` → alta, otherwise → edición
  - `Public Function RegistrarAlta(p_DTO As ExpedienteDTO, Optional ByRef p_Error As String) As String` (línea 566) — wrapper de alta
  - `Public Function RegistrarExpEntidades(Optional ByRef p_Error As String, Optional p_db As DAO.Database = Nothing, Optional p_Ambito As EnumAmbitoActualizacion = EnumAmbitoActualizacion.Todo, Optional p_ForceCacheFailureForTest As Boolean = False) As String` (línea 811) — entry de edición (UPDATE entidades-derivadas)
  - `Private Function MotivoNoOK(p_DTO, p_ObjExpedienteAlInicio, p_Error) As String` (línea 855) — 11 grupos de validaciones para edición+alta
  - `Private Function MotivoNoOKAlta(...)` (línea 1123) — validaciones adicionales para alta
  - `ExpedienteEntidadOperaciones.Registrar(...)` (delegada) — UPDATE real de `TbExpedientesConEntidades`
  - `Public Function RegistrarCambioTipo(...)` (línea 2266) — CAP-004
- **Datos afectados** (escritura, dentro de transacción DAO.Workspace):
  - `TbExpedientesConEntidades` (UPDATE 1 fila, FK = `IDExpediente`) — re-cálculo de `CadenaComerciales`/`CadenaJPs`/`CadenaRACs`/etc. según `p_Ambito`
  - `TbExpedientesComerciales`/`CPVs`/`LugaresEjecucion`/`PECAL`/`RACS`/`Responsables`/`CodigoCompras`/`Anexos`/`Hitos`/`Modificados` (UPDATE/INSERT/DELETE según `p_Ambito`)
  - `TbExpedientesSuministradores` (árbol)
  - `TbExpedientesAnualidades` (UPDATE/INSERT/DELETE)
  - `TbUltimoCambio` (INSERT 1 fila con `FechaCambio = Now()`, `IDUsuarioCambio = g_UsuarioConectado`)
  - **NO** se toca `TbExpedientes` (cabecera) — la edición preserva audit-trail
- **Salidas**: `String` (el `IDExpediente`); eventos UI los maneja el form
- **Dependencias e integraciones**:
  - `ExpedienteDTO` (clase DTO)
  - `ExpedienteEntidadOperaciones` (clase delegada para UPDATE de entidades-derivadas)
  - `EnumAmbitoActualizacion` (enum: `Todo`/`Comerciales`/`CPVs`/etc.) para control granular de recálculo
  - `DAO.Workspace.Workspaces(0)` para transacción
  - `getdb()` (gateado por `m_TestingMode`)
  - `g_UsuarioConectado` para `IDUsuarioCambio` en `TbUltimoCambio`
- **Sincronización fuente↔binario**: `dysflow.import_modules({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones", "ExpedienteEntidadOperaciones", "Form_FormExpediente", "Expediente", "ExpedienteDTO"] })`; `dysflow.verify_binary({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones", "ExpedienteEntidadOperaciones"], diff: true })` debe retornar `actionableOk: true`
- **Valoración de diseño (tal-como-está vs ideal)**: **MAYORMENTE BIEN HECHO** para el estado actual. La separación cabecera ↔ entidades-derivadas es intencional y correcta. Las validaciones exhaustivas con dirty detection implícito son una decisión de diseño sólida. Pendiente: extracción a `Helper_ExpedienteEdicion` (PRUEBA-003 REFAC-3a, todavía no ejecutado), tests atómicos que prueben la cadena `RegistrarExpEntidades → UPDATE entidades-derivadas + INSERT TbUltimoCambio + COMMIT`.

## §4 Receta de reconstrucción (≈ tasks SDD) — REPRODUCIBILIDAD
> Pasos ordenados para reconstruir esta capacidad desde cero. Todas las operaciones fuente↔binario pasan por el MCP de Dysflow.
1. Verificar `src/classes/ExpedienteOperaciones.cls` con las firmas de §3. Ya existe con 3034 líneas; las firmas son los anchors de PRUEBA-003 REFAC-3a.
2. Verificar `src/classes/ExpedienteEntidadOperaciones.cls` (la clase delegada para UPDATE de entidades-derivadas).
3. Verificar `src/forms/Form_FormExpediente.cls` — el handler `ComandoActualizarCompleto_Click` (línea 133) llama a `m_ExpOp.RegistrarExpEntidades p_Error:=m_Error` (línea 148).
4. **Importar** cambios → `dysflow.import_modules({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones", "ExpedienteEntidadOperaciones", "Form_FormExpediente", "Expediente", "ExpedienteDTO"] })`.
5. **Compilar** → `dysflow.compile_vba({ projectId: "expedientes" })` debe retornar `ok: true`.
6. **Verificar binario** → `dysflow.verify_binary({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones", "ExpedienteEntidadOperaciones"], diff: true })` debe retornar `actionableOk: true`.
7. **Demostrar el comportamiento** con tests atómicos: `dysflow.test_vba({ projectId: "expedientes", proceduresJson: "[{\"procedure\":\"Test_Helper_ExpedienteEdicion_Actualizar_DTOValido_ActualizaYDevuelveTrue\"}, ...]" })` (5+ tests target post-PR-C).

## §5 Evidencia y trazabilidad (≈ verify SDD)
- **Tests**: `src/modules/Test_HelperExpedienteEdicion.bas` (5+ tests atómicos target post-PRUEBA-002 PR-C). Pre-PRUEBA-002: 0 tests específicos de edición.
- **Trazabilidad de release** (solo-añadir):

| Elemento | Ref. tracker | Versión de staging | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Edición con validación exhaustiva (11 grupos) y transacción DAO | `ExpedienteOperaciones.MotivoNoOK` línea 855 + `RegistrarExpEntidades` línea 811 | Pendiente | pending | Pendiente | — | comportamiento actual del código; sin tests verdes específicos |
| Dirty detection implícito (línea 987, 1000, etc.) | `MotivoNoOK` | Pendiente | pending | Pendiente | — | unicidad condicional, no es UPDATE parcial explícito |
| Extracción a `Helper_ExpedienteEdicion` (PRUEBA-003 REFAC-3a) | PRUEBA-003 REFAC-3a | Pendiente | pending | Pendiente | — | pendiente implementación |

- **Tabla de diagnóstico de regresión**:

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla del documento |
|---|---|---|---|
| Edición acepta `Titulo` duplicado | `MotivoNoOK` no detecta el cambio | `Test_Helper_ExpedienteEdicion_ValidarEdicion_TituloDuplicado_DevuelveFalse` | §2 BR-02-01 |
| Edición quita AGEDYS con DPDs vinculados | guard de línea 976 falta | `Test_Helper_ExpedienteEdicion_ValidarEdicion_QuitarAGEDYSConDPDs_DevuelveFalse` | §2 BR-02-05 |
| Edición acepta AM+Lote simultáneos | guard de línea 914 falta | `Test_Helper_ExpedienteEdicion_ValidarEdicion_TipologiaInvalida_DevuelveFalse` | §2 BR-02-02 |
| Edición no registra `TbUltimoCambio` | servicio de audit no invocado | `count_rows` antes/después de `TbUltimoCambio` | §3 |

## §6 Notas de migración web
- **Conservar**: la separación cabecera ↔ entidades-derivadas; las 11 validaciones de `MotivoNoOK`; el dirty detection implícito; el audit-trail en `TbUltimoCambio`.
- **Transformar**: el `DAO.Workspace` → `BEGIN TRANSACTION`/`COMMIT`/`ROLLBACK` del motor SQL destino; el `p_Ambito` (granularidad de recálculo) se traduce a endpoints PATCH parciales por recurso; el `EnumAmbitoActualizacion` se traduce a un campo `fields[]` en el request body.
- **NO copiar (legado)**: el `fso.FileExists` para validar rutas de anexos — portar a validación server-side con storage abstracto; las validaciones de unicidad con `getExpedientePorCampo` pre-UPDATE — portar a UNIQUE constraints y capturar la violación.
- **Preguntas abiertas**:
  - ¿`MotivoNoOK` debería también validar unicidad de `IDExpediente` (PK) en alguna rama? Hoy no lo hace.
  - El gap de concurrencia跨 sesiones es idéntico a CAP-001 — ¿UNIQUE constraints en DB o `SELECT FOR UPDATE`?
  - ¿La extracción a `Helper_ExpedienteEdicion` debería cubrir solo `RegistrarExpEntidades` (lo que usa el form) o también `Registrar` (entry comprensivo, no usado por el form hoy)?

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| `ExpedienteOperaciones.Registrar` existe y es entry comprensivo (alta + edición) | Verified-runtime | `compile_vba` OK; lectura directa (línea 19) | 2026-06-15 |
| `RegistrarExpEntidades` es el entry que usa el form de edición | Verified-static | lectura directa `Form_FormExpediente.cls` línea 148 | 2026-06-15 |
| `MotivoNoOK` valida 11 grupos para edición (campos obligatorios, fechas, tipología, códigos, AGEDYS, unicidad condicional, ambito, PECAL, Anexos) | Verified-static | lectura directa (líneas 855-1122) | 2026-06-15 |
| Tabla real es `TbExpedientes` con 61 columnas; PK = `IDExpediente` | Verified-runtime | `dysflow_get_schema` | 2026-06-15 |
| Tabla real para entidades derivadas es `TbExpedientesConEntidades` (22 columnas) | Verified-runtime | `dysflow_get_schema` | 2026-06-15 |
| Audit trail en `TbUltimoCambio` (PK = `ID`, FK = `IDExpediente`) | Verified-runtime | `dysflow_get_schema` | 2026-06-15 |
| Form llama `m_ExpOp.RegistrarExpEntidades p_Error:=m_Error` (línea 148) | Verified-static | lectura directa del form | 2026-06-15 |
| Tests atómicos `Test_Helper_ExpedienteEdicion_*` aún no implementados | Verified-static | `src/modules/Test_HelperExpedienteEdicion.bas` no existe | 2026-06-15 |

**⚠️ Divergencias detectadas** (intención SDD vs realidad del código):
- La spec PRUEBA-002 §3.2 BR-02-01 nombra "actualiza solo los campos dirty" — el código NO hace UPDATE parcial en `TbExpedientes`; de hecho **no toca `TbExpedientes`**. La edición preserva la cabecera (intencional). El "dirty detection" existe solo para la consulta de unicidad (línea 987, 1000, etc.) — `MotivoNoOK` consulta la BD solo si el campo cambió, optimizando así.
- La spec PRUEBA-002 §3.2 BR-02-02 menciona "`EstadoBloqueo = True`" como validación — el código **no implementa ese flag**. La edición se hace siempre y se confía en la traza de `TbUltimoCambio`. **Deuda**: confirmar con responsable si se quiere un flag de bloqueo explícito.
- La spec habla de "`TbExpedientesAuditTrail`" — **no existe**; es `TbUltimoCambio`.
- La spec describe "`Helper_ExpedienteEdicion`" — **esa clase no existe**. La entry real es `ExpedienteOperaciones.RegistrarExpEntidades` (que el form usa) y `Registrar` (entry comprensivo, no usado por el form de edición).
- La spec dice "`ActualizarExpediente`" como método separado — el código **no tiene ese nombre**. La entry es `RegistrarExpEntidades` (línea 811) que delega a `ExpedienteEntidadOperaciones.Registrar`.
- La spec PRUEBA-002 lista "16 sub-tasks" para PR-REFAC-1a edición; muchas son aspiracionales y deben re-derivarse de las 11 validaciones reales de `MotivoNoOK` cuando PRUEBA-003 REFAC-3a se ejecute.
