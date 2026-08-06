<!--
Documento de capacidad CAP-001 — Alta de expediente.
Linaje: PRUEBA-002 spec + dysflow_get_schema de tablas reales + ExpedienteOperaciones.cls + Form_FormExpedienteAlta.cls
Versión corregida tras critical review 2026-06-15 (schema y entry point verificados contra código real).
Idioma: castellano de España. Identificadores de código/test se mantienen tal cual.
-->

# Capacidad: Alta de expediente

## §0 Identidad
- **ID de capacidad**: CAP-001
- **Tier**: critical
- **Estado**: active
- **Source**: hybrid (SDD + código actual)
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: 2026-06-15 mediante `dysflow_get_schema` (TbExpedientes, TbExpedientesConEntidades, TbUltimoCambio) + lectura de `src/classes/ExpedienteOperaciones.cls` (líneas 19, 566, 855, 1123) y `src/forms/Form_FormExpedienteAlta.cls` línea 52
- **Confianza global**: mixta — ver §7 (entry point y validación son `Verified-static` por lectura; sin tests verdes específicos de alta)

## §1 Intención de negocio (≈ proposal SDD) — POR QUÉ
- **Propósito**: registrar un nuevo expediente en el sistema con su cabecera en `TbExpedientes`, sus datos derivados en `TbExpedientesConEntidades` (cadenas concatenadas para JOIN rápido en listados), y sus vinculaciones a entidades (Comercial, CPV, Lugar, PECAL, RAC, Responsable, Anexos, Hitos, Modificados, etc.). El alta es la primera escritura del expediente y debe ser atómica: o se inserta todo o no se inserta nada.
- **Usuarios / perfiles**: gestor de área (alta regular, lote, basado, expediente individual); administrador (corrección, casos especiales); el alta HPS vive en otro flujo (`ExpedienteOperaciones.RegistrarSoloHPS`).
- **Problema que resuelve**: el alta está embebida en `Form_FormExpedienteAlta.cls` (675 líneas) con 11 handlers y 50+ validaciones cruzadas contra tablas. La lógica de INSERT vive en `ExpedienteOperaciones.RegistrarAlta` (que llama a `Registrar` con `p_ExpedienteAlInicio = Nothing`). Es la **única ruta de INSERT en cabecera** — la comparte con edición vía el flag `esAlta = (p_ExpedienteAlInicio Is Nothing)`.
- **Valor de negocio**: alta transaccional, validaciones de unicidad pre-INSERT, generación de ID derivada, cálculo de `Estado`/`FechaFinGarantia` a partir de reglas internas, registro de `UltimoCambio` post-INSERT.
- **No-objetivos**: alta HPS (`RegistrarSoloHPS`, otra rama); carga masiva; validación de UI (la UI solo bloquea; el backend es la autoridad).
- **Origen de la intención**: spec PRUEBA-002 §3.2 (BR-01-01..05) + design PRUEBA-003 §6.3 (refactor planeado a `Helper_ExpedienteAlta` aún no ejecutado).
- **Referencia de tracker de origen**: PRUEBA-002 PR-C, PRUEBA-003 REFAC-3a (pendiente).

## §2 Contrato de comportamiento (≈ spec SDD) — QUÉ ⟵ ANCLA DE REGRESIÓN

### Escenarios (Dado / Cuando / Entonces)
- **DADO** un `ExpedienteDTO` válido (Ámbito no vacío, APLICAESTADO "Sí"/"No", AGEDYSAplica no vacío, Titulo no vacío y único, CodExp no vacío y único salvo Lote, IDOrganoContratacion > 0, etc.) **CUANDO** se invoca `ExpedienteOperaciones.RegistrarAlta(p_DTO, p_Error)` (línea 566) **ENTONCES** la helper llama internamente a `Registrar(p_DTO, Nothing, p_Error)` que abre transacción DAO.Workspace, INSERT en `TbExpedientes` (1 fila), INSERT en `TbExpedientesConEntidades` (1 fila con cadenas y entidades pre-concatenadas), INSERT en las tablas de vinculación que correspondan, INSERT en `TbUltimoCambio` (1 fila con `FechaCambio = Now()`, `IDUsuarioCambio = g_UsuarioConectado`), COMMIT.
- **DADO** un DTO con `Ambito = ""` **CUANDO** se invoca `RegistrarAlta` **ENTONCES** `p_Error = "Se ha de indicar un ámbito (Defensa, Fuera o HPS)"` y NO se hace INSERT. (Validación 1 de `MotivoNoOKAlta`, línea 1136).
- **DADO** un DTO con `CodExp` ya existente en `TbExpedientes` **CUANDO** se invoca `RegistrarAlta` **ENTONCES** `p_Error = "El Código corto ya existe en el Expediente ID: <ID>"` y NO se hace INSERT. (Validación de unicidad, línea 1158-1164).
- **DADO** un DTO con `Titulo = ""` **CUANDO** se invoca `RegistrarAlta` **ENTONCES** `p_Error = "El título del expediente es obligatorio"`. (Línea 1174).
- **DADO** un DTO con `AGEDYSAplica = "Sí"` y `AGEDYSGenerico = ""` **CUANDO** se invoca `RegistrarAlta` **ENTONCES** `p_Error = "Si aplica Agedys, indicar si es genérico"`. (Línea 1148).
- **DADO** un DTO con cualquier validación fallando **CUANDO** se invoca `RegistrarAlta` **ENTONCES** `p_Error` contiene el mensaje, no se hace INSERT, NO se modifica la BD.
- **DADO** una FK violation en `TbExpedientesConEntidades` (e.g., `IdGradoClasificacion` no existe en `TbGradosClasificacion`) **CUANDO** se invoca `RegistrarAlta` **ENTONCES** `p_Error` contiene `Err.Description` y se hace ROLLBACK de las inserciones previas en la transacción.

### Reglas de negocio

| ID regla | Enunciado (leído de `MotivoNoOKAlta`, líneas 1136+) | Autoridad | ¿Aplicada en código? | Prueba | Confianza |
|---|---|---|---|---|---|
| BR-01-01 | `Ambito` es obligatorio (Defensa, Fuera, HPS) | spec + código | Sí — `MotivoNoOKAlta` línea 1136 | `Test_Helper_ExpedienteAlta_ValidarAlta_AmbitoVacio_DevuelveFalse` (target PR-C) | Verified-static |
| BR-01-02 | `APLICAESTADO` debe ser "Sí" o "No" | spec + código | Sí — línea 1140 | `Test_Helper_ExpedienteAlta_ValidarAlta_AplicaEstadoInvalido_DevuelveFalse` (target PR-C) | Verified-static |
| BR-01-03 | `AGEDYSAplica` es obligatorio; si "Sí" entonces `AGEDYSGenerico` también | spec + código | Sí — líneas 1144-1152 | `Test_Helper_ExpedienteAlta_ValidarAlta_AGEDYSIncompleto_DevuelveFalse` (target PR-C) | Verified-static |
| BR-01-04 | `Titulo` es obligatorio y único en `TbExpedientes.Titulo` | spec + código | Sí — líneas 1174-1182 | `Test_Helper_ExpedienteAlta_RegistrarAlta_TituloDuplicado_NoInserta` (target PR-C) | Verified-static |
| BR-01-05 | `CodExp` es obligatorio salvo `EsLote = "Sí"`; si presente, único | spec + código | Sí — líneas 1154-1164 | `Test_Helper_ExpedienteAlta_RegistrarAlta_CodExpDuplicado_NoInserta` + `Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinCodExp_Permite` (target PR-C) | Verified-static |
| BR-01-06 | `IDOrganoContratacion` es obligatorio y > 0 | spec + código | Sí — línea 1192 | `Test_Helper_ExpedienteAlta_ValidarAlta_SinOrgano_DevuelveFalse` (target PR-C) | Verified-static |
| BR-01-07 | `CodExpLargo`, `Nemotecnico`, `NPedido`, `CodProyecto` (cuando presentes) deben ser únicos | spec + código | Sí — líneas 1166-1210 | tests dedicados por unicidad (target PR-C) | Verified-static |
| BR-01-08 | `RegistrarAlta` con DTO válido inserta en `TbExpedientes` + `TbExpedientesConEntidades` + `TbUltimoCambio` + vinculaciones en transacción DAO | spec + código | Sí — `Registrar` línea 47 (BeginTrans) + commits | `Test_Helper_ExpedienteAlta_RegistrarAlta_DTOValido_InsertaYDevuelveID` (target PR-C) | Verified-static |
| BR-01-09 | `RegistrarAlta` hace rollback completo ante FK violation | spec + código | Sí — `Registrar` usa `DAO.Workspace.Workspaces(0).BeginTrans` con `CommitTrans`/`Rollback` | `Test_Helper_ExpedienteAlta_RegistrarAlta_FKViolation_Rollback` (target PR-C) | Verified-static |
| BR-01-10 (subregla) | `OrdinalE2E` se setea a `Null` por defecto al alta | spec CAP-047 | Sí | `Test_Expediente_OrdinalE2E_DefaultIsNull` (VERIFIED-runtime ya en `tests/tests.vba.json`) | Verified-runtime |

### Validaciones
- `Ambito` ∈ `{"Defensa", "Fuera", "HPS"}`
- `APLICAESTADO` ∈ `{"Sí", "No"}`; si vacío, se rechaza
- `AGEDYSAplica` ∈ `{"Sí", "No"}`; si "Sí" entonces `AGEDYSGenerico` ∈ `{"Sí", "No"}`
- `Titulo` no vacío, único
- `CodExp` no vacío (salvo Lote), único
- `CodExpLargo` opcional, único
- `Nemotecnico` opcional, único
- `NPedido` opcional, único
- `CodProyecto` opcional, único
- `IDOrganoContratacion` numérico, > 0
- `TIpo` (campo `Tipo` con typo intencional VBE-canonicalizado) no vacío
- `IDEstado` no validado explícitamente en `MotivoNoOKAlta` (gap, ver §6)

### Transiciones de estado
- No aplica cambio de estado explícito; el expediente arranca con el `ESTADO` calculado por `Me.Expediente.ESTADOCalculadoTexto` (línea 68) que combina `FechaInicioContrato`/`FechaFinContrato`/`APLICAESTADO` y reglas internas.

### Casos límite y de error
- DTO con `IDExpedientePadre` para Lote → se valida implícitamente al asignar `.IDExpedienteCalculado`
- Concurrencia: la transacción DAO.Workspace garantiza atomicidad multi-tabla dentro de la sesión; no hay locking explícito跨 sesiones (potencial race condition en unicidad de `CodExp`/`Titulo` — gap)
- `Nemotecnico` autogenerado si vacío (calculado por `Expediente.IDExpedienteCalculado` y derivados)

### Señales de aceptación / presencia ⟵ cómo saber que la funcionalidad EXISTE y funciona
- `ExpedienteOperaciones.cls` existe en el binario (`dysflow.exists` retorna `true`)
- `Form_FormExpedienteAlta.ComandoRegistrar_Click` invoca `m_ExpOp.RegistrarAlta m_ObjExpedienteDTOActivo, p_Error:=m_Error` (línea 52 del form)
- `dysflow.test_vba` con filtro `expediente-alta` retorna `ok: true` con 5+/5+ PASS (target post-PRUEBA-002 PR-C)
- `dysflow.count_rows` antes/después del alta: la BD tiene 1 fila más en `TbExpedientes` + 1 en `TbExpedientesConEntidades` + 1 en `TbUltimoCambio`
- Si al registrar un DTO válido en la UI el sistema no crea las 3 filas y muestra error → regresión

## §3 Mapa de implementación (≈ design SDD) — CÓMO
- **Puntos de entrada de UI**:
  - `Form_FormExpedienteAlta.cls` → `ComandoRegistrar_Click` (línea 34) → `m_ExpOp.RegistrarAlta m_ObjExpedienteDTOActivo, p_Error:=m_Error`
  - `Form_FormExpedienteAltaParaHPS.cls` → `ComandoRegistrar_Click` (HPS, línea 402) → `m_ExpOp.Registrar m_ObjExpedienteDTOActivo, Nothing, m_Error` (pasa `Nothing` como `p_ExpedienteAlInicio` = alta)
  - `Form_FormExpedienteAltaTipo.cls` → `ComandoElegir_*` y `ComandoLimpiar_*` (UI routing, no INSERT)
- **Puntos de entrada de código** (clase `ExpedienteOperaciones`):
  - `Public Function Registrar(p_DTO As ExpedienteDTO, p_ExpedienteAlInicio As Expediente, Optional ByRef p_Error As String) As String` (línea 19) — entry principal transaccional; `p_ExpedienteAlInicio Is Nothing` → alta
  - `Public Function RegistrarAlta(p_DTO As ExpedienteDTO, Optional ByRef p_Error As String) As String` (línea 566) — wrapper que llama a `Registrar` con `p_ExpedienteAlInicio = Nothing`
  - `Private Function MotivoNoOK(p_DTO, p_ExpedienteAlInicio, p_Error) As String` (línea 855) — para edición: validaciones + cálculo de cambios
  - `Private Function MotivoNoOKAlta(p_DTO, p_Error) As String` (línea 1123) — para alta: 16+ validaciones
  - `Private Function MotivoNoOKSoloHPS(...)` (línea 1291) — para HPS
  - `Public Function RegistrarSoloHPS(...)` (línea 1369) — INSERT HPS
  - `Public Function RegistrarExpEntidades(...)` (línea 811) — UPDATE `TbExpedientesConEntidades`
  - `Public Function RegistrarAnualidad(...)` (línea 1398), `RegistrarComercial(...)` (línea 1444), `RegistrarCodigoCompra(...)` (1478), `RegistrarCPV(...)` (1511), `RegistrarLugarEjecucion(...)` (1544), `RegistrarPECAL(...)` (1604), `RegistrarRAC(...)` (1666), `RegistrarAnexo(...)` (1785), `RegistrarHito(...)` (2033), `RegistrarModificado(...)` (2165), `RegistrarCambioTipo(...)` (2266) — vinculaciones específicas
- **Datos afectados** (escritura, dentro de transacción DAO.Workspace):
  - `TbExpedientes` (1 fila INSERT, PK = `IDExpediente`)
  - `TbExpedientesConEntidades` (1 fila INSERT o UPDATE; PK = `IDExpediente`)
  - `TbExpedientesComerciales`, `TbExpedientesCPVs`, `TbExpedientesLugaresEjecucion`, `TbExpedientesPECAL`, `TbExpedientesRACS`, `TbExpedientesResponsables`, `TbExpedientesCodigoCompras` (N filas INSERT, una por entidad vinculada)
  - `TbExpedientesAnexos`, `TbExpedientesHitos`, `TbExpedientesModificados` (N filas si aplica)
  - `TbExpedientesSuministradores` (árbol)
  - `TbExpedientesAnualidades` (1+ filas)
  - `TbUltimoCambio` (1 fila INSERT, PK = `ID`, FK = `IDExpediente`)
- **Salidas**: `String` (el nuevo `IDExpediente` retornado por `Registrar`/`RegistrarAlta`); eventos UI (`MsgBox` de éxito) los maneja el form, no la helper
- **Dependencias e integraciones**:
  - `ExpedienteDTO` (clase DTO) — contiene `Public Expediente As Expediente`
  - `Expediente` (clase) — modelo de dominio con propiedades + `IDExpedienteCalculado`/`ESTADOCalculadoTexto`/`FechaFinGarantiaCalculada`
  - `DAO.Workspace.Workspaces(0)` para transacción
  - `getdb()` (gateado por `m_TestingMode` para sandbox)
  - `g_UsuarioConectado` (singleton) para `IDUsuarioCambio` en `TbUltimoCambio`
- **Sincronización fuente↔binario**: `dysflow.import_modules({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones", "Form_FormExpedienteAlta", "ExpedienteDTO", "Expediente"] })`; `dysflow.verify_binary({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones"], diff: true })` debe retornar `actionableOk: true`
- **Valoración de diseño (tal-como-está vs ideal)**: **MAYORMENTE BIEN HECHO** para el estado actual. La transacción DAO.Workspace ya está implementada y la lógica de validación es exhaustiva. Pendiente: extracción a `Helper_ExpedienteAlta` (PRUEBA-003 REFAC-3a, todavía no ejecutado), validaciones de unicidad跨 sesiones (gap, ver §6), y un test atómico que pruebe la cadena `RegistrarAlta → INSERT 3 tablas + COMMIT`.

## §4 Receta de reconstrucción (≈ tasks SDD) — REPRODUCIBILIDAD
> Pasos ordenados para reconstruir esta capacidad desde cero. Todas las operaciones fuente↔binario pasan por el MCP de Dysflow.
1. Crear/verificar `src/classes/ExpedienteOperaciones.cls` con las firmas de §3. La clase ya existe con 3034 líneas; las firmas son los anchors de PRUEBA-003 REFAC-3a.
2. Verificar `src/classes/Expediente.cls` (modelo de dominio) y `src/classes/ExpedienteDTO.cls` (wrapper DTO con `Public Expediente As Expediente`).
3. Verificar `src/forms/Form_FormExpedienteAlta.cls` — el handler `ComandoRegistrar_Click` (línea 34) llama a `SetDTOFromGeneral Me, m_ObjExpedienteDTOActivo.Expediente, m_Error` (línea 49) y luego a `m_ExpOp.RegistrarAlta m_ObjExpedienteDTOActivo, p_Error:=m_Error` (línea 52).
4. **Importar** cambios → `dysflow.import_modules({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones", "Form_FormExpedienteAlta", "ExpedienteDTO", "Expediente"] })`.
5. **Compilar** → `dysflow.compile_vba({ projectId: "expedientes" })` debe retornar `ok: true`.
6. **Verificar binario** → `dysflow.verify_binary({ projectId: "expedientes", moduleNames: ["ExpedienteOperaciones"], diff: true })` debe retornar `actionableOk: true`.
7. **Demostrar el comportamiento** con tests atómicos: `dysflow.test_vba({ projectId: "expedientes", proceduresJson: "[{\"procedure\":\"Test_Helper_ExpedienteAlta_RegistrarAlta_DTOValido_InsertaYDevuelveID\"}, ...]" })` (5+ tests target post-PR-C).

## §5 Evidencia y trazabilidad (≈ verify SDD)
- **Tests**: `src/modules/Test_HelperExpedienteAlta.bas` (5+ tests atómicos target post-PRUEBA-002 PR-C). Pre-PRUEBA-002: 0 tests específicos de alta.
- **Trazabilidad de release** (solo-añadir):

| Elemento | Ref. tracker | Versión de staging | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Alta con transacción DAO.Workspace + 16 validaciones | `ExpedienteOperaciones.Registrar` línea 19 | Pendiente | pending | Pendiente | — | comportamiento actual del código; sin tests verdes específicos |
| Extracción a `Helper_ExpedienteAlta` (PRUEBA-003 REFAC-3a) | PRUEBA-003 REFAC-3a | Pendiente | pending | Pendiente | — | pendiente implementación |

- **Tabla de diagnóstico de regresión**:

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla del documento |
|---|---|---|---|
| Alta deja filas huérfanas en `TbExpedientesConEntidades` | transacción DAO no envuelve toda la operación | `Test_Helper_ExpedienteAlta_RegistrarAlta_FKViolation_Rollback` | §2 BR-01-09 / §3 |
| Alta acepta `Titulo` duplicado | unicidad no validada | `Test_Helper_ExpedienteAlta_RegistrarAlta_TituloDuplicado_NoInserta` | §2 BR-01-04 |
| Alta no genera `TbUltimoCambio` | `RegistrarAlta` no llama al servicio de audit | `count_rows` antes/después de `TbUltimoCambio` | §3 |
| Alta acepta `CodExp` duplicado en Lote | validación de unicidad rota | `Test_Helper_ExpedienteAlta_RegistrarAlta_CodExpDuplicado_NoInserta` | §2 BR-01-05 |

## §6 Notas de migración web
- **Conservar**: la transacción atómica multi-tabla, las 16 validaciones de `MotivoNoOKAlta`, la generación derivada de `IDExpediente` y `ESTADOCalculadoTexto`.
- **Transformar**: el `DAO.Workspace` se traduce a `BEGIN TRANSACTION`/`COMMIT`/`ROLLBACK` del motor SQL destino; el `ExpedienteDTO` se transforma en un request body validado server-side.
- **NO copiar (legado)**: las validaciones de unicidad de `CodExp`/`Titulo` con `getExpedientePorCampo` pre-INSERT — portar a UNIQUE constraints en la DB y capturar la violación como error legible.
- **Preguntas abiertas**:
  - ¿La validación de `IDEstado` ∈ estados válidos debería moverse a la DB (FK constraint) o quedarse en `MotivoNoOKAlta`?
  - El gap de concurrencia跨 sesiones (dos altas simultáneas con el mismo `CodExp`) — ¿CHECK constraints en DB o `SELECT FOR UPDATE`?
  - ¿La extracción a `Helper_ExpedienteAlta` (PRUEBA-003 REFAC-3a) mantiene `ExpedienteOperaciones` como thin wrapper o lo reemplaza?

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| `ExpedienteOperaciones.Registrar` existe y es transaccional | Verified-runtime | `compile_vba` OK; lectura directa del código (línea 19, 47, ws.BeginTrans) | 2026-06-15 |
| `RegistrarAlta` es wrapper que llama a `Registrar(p_DTO, Nothing, p_Error)` | Verified-static | lectura directa del código (línea 566) | 2026-06-15 |
| `MotivoNoOKAlta` valida 16 reglas (Ámbito, APLICAESTADO, AGEDYS, Titulo, CodExp, CodExpLargo, Nemotecnico, IDOrganoContratacion, NPedido, CodProyecto, TIpo, ...) | Verified-static | lectura directa del código (líneas 1123-1290) | 2026-06-15 |
| Tabla real es `TbExpedientes` con 61 columnas; PK = `IDExpediente` (Long, required) | Verified-runtime | `dysflow_get_schema` | 2026-06-15 |
| Tabla real para entidades derivadas es `TbExpedientesConEntidades` (22 columnas, PK = `IDExpediente`) | Verified-runtime | `dysflow_get_schema` | 2026-06-15 |
| Audit trail en `TbUltimoCambio` (PK = `ID`, FK = `IDExpediente`, campos `FechaCambio`/`IDUsuarioCambio`) | Verified-runtime | `dysflow_get_schema` | 2026-06-15 |
| Form llama `m_ExpOp.RegistrarAlta m_ObjExpedienteDTOActivo, p_Error:=m_Error` (línea 52) | Verified-static | lectura directa del form | 2026-06-15 |
| Tests atómicos `Test_Helper_ExpedienteAlta_*` aún no implementados | Verified-static | `src/modules/Test_HelperExpedienteAlta.bas` no existe en el repo | 2026-06-15 |
| `OrdinalE2E` se setea a `Null` por defecto al alta | Verified-runtime | `Test_Expediente_OrdinalE2E_DefaultIsNull` (ya en `tests/tests.vba.json` y verde) | 2026-06-15 |

**⚠️ Divergencias detectadas** (intención SDD vs realidad del código):
- La spec PRUEBA-002 §3.2 BR-01-04 nombra "`Importe < 0`" como validación; el código **no la implementa**. `ImporteLicitacion` e `ImporteContratacion` existen en `TbExpedientes` pero `MotivoNoOKAlta` no los revisa. **Deuda**: confirmar con responsable si la validación es necesaria.
- La spec PRUEBA-002 §3.2 BR-01-03 nombra "Lote sin `IDExpedientePadre`" como validación; el código **no la implementa explícitamente** (la lógica vive en `.IDExpedienteCalculado`). **Deuda**: confirmar con responsable.
- La spec habla de "`TbExpedienteEntidades`" como tabla de N entidades; **esa tabla no existe**. La realidad es `TbExpedientesConEntidades` (1 fila con cadenas concatenadas) + 7 tablas de vinculación específicas (`TbExpedientesComerciales`/`CPVs`/etc.). **La doc PRUEBA-002 estaba mal; la corrijo aquí.**
- La spec habla de "`TbExpedientesAuditTrail`"; **esa tabla no existe**. La realidad es `TbUltimoCambio`.
- La spec describe "`Helper_ExpedienteAlta`"; **esa clase no existe**. La entry real es `ExpedienteOperaciones.Registrar`/`RegistrarAlta`.
