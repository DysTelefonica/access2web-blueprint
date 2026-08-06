<!--
Documento de capacidad CAP-003 - Eliminacion de expediente.
Linaje: PRUEBA-002 §3.2 BR-03-01..02 (spec de capability) + PRUEBA-003 REFAC-1c (helper stateless) + dysflow_get_schema de TbExpedientes + TbExpedientesAnexos + TbExpedientesSuministradores.
Idioma: castellano de Espana. Identificadores de codigo/test se mantienen tal cual.
-->

# Capacidad: Eliminacion de expediente

## 0 Identidad
- **ID de capacidad**: CAP-003
- **Tier**: critical (ciclo de vida del expediente)
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Responsable / autoridad de producto**: Pendiente de confirmacion
- **Ultima verificacion**: 2026-06-16 mediante `dysflow.get_schema` (TbExpedientes 61 cols, TbExpedientesAnexos 3 cols, TbExpedientesSuministradores 7 cols, TbExpedientesHitos 6 cols, TbExpedientesResponsables 6 cols, TbExpedientesModificados 6 cols) + `dysflow.test_vba` (11/11 verde slice-refac-1c + 4/4 regresion 1a+1b = 15/15 total) + `dysflow.compile_vba` (PASS)
- **Confianza global**: mixta - ver 7 (BR-26-01..05 `Verified-runtime`; 4 checks DPDs/GR/HPS/NCs `Verified-static` deferred)

## 1 Intencion de negocio (~ proposal SDD) - POR QUE
- **Proposito**: permitir al gestor eliminar un expediente que ya no es valido, con validacion previa de que el expediente no tiene dependencias (hijos derivados, anexos, suministradores, DPDs/Agedys, GR, HPS, NCs) que impedirian la operacion o dejarian el sistema en estado inconsistente.
- **Usuarios / perfiles**: gestor de area, administrador; auditoria (trazabilidad del borrado).
- **Problema que resuelve**: la eliminacion estaba embebida en `Form_FormExpedientesGestion.cls` (funcion `Eliminar`, 124 LOC, + 3 handlers `ComandoEliminarAMoC/Lote/Basado_Click` de 28 LOC cada uno) y acoplada a controles del form (`Me.AllowEdits`, `Me.ListaAMEIndividual`, `m_ExpAMSeleccionadoC`, etc.) y a un class method `ExpedienteOperaciones.Eliminar` que requiere un `Expediente` ya cargado. Imposible de probar la validacion (`MotivoEliminarNoOK`, 5 checks) sin abrir UI e instanciar el Expediente.
- **Valor de negocio**: validacion determinista y testeable; permite chequear "se puede borrar?" antes de mostrar el MsgBox de confirmacion; base para scripts batch de limpieza; reduce la deuda "logica en formulario" del 100% al ~20% para esta capacidad.
- **No-objetivos**: soft delete / undo window (no implementado en el codigo actual); cascade a `TbE2EExportBatch*` (los batches de export sobreviven la eliminacion del expediente, cubierto por CAP-030).
- **Origen de la intencion**: spec PRUEBA-002 3.2 (BR-03-01..02 wish de "cascade delete" no implementado) + PRUEBA-003 6.1 (helper stateless con DAO opcional inyectado) + UX feedback del usuario 2026-06-15 (validacion de "no se puede borrar porque tiene hijos" sin abrir UI).
- **Referencia de tracker de origen**: PRUEBA-002 PR-C, PRUEBA-003 REFAC-1c (cerrado en este slice), PRUEBA-002 PR-F (UX, ya cerrado en commit `641262d`).

## 2 Contrato de comportamiento (~ spec SDD) - QUE ? ANCLA DE REGRESION

### Escenarios (Dado / Cuando / Entonces)
- **DADO** un `p_IDExpediente = 0` (o negativo) **CUANDO** se invoca `Helper_ExpedienteEliminacion.TieneDerivados(0, ...)` **ENTONCES** retorna `"ERR"` y `p_Error` contiene `"TieneDerivados: p_IDExpediente must be > 0 (got 0)"`.
- **DADO** un `p_IDExpediente` que NO tiene filas en `TbExpedientes WHERE IDExpedientePadre = p_IDExpediente` **CUANDO** se invoca `Helper_ExpedienteEliminacion.TieneDerivados(p_ID, ...)` **ENTONCES** retorna `"OK"` con `p_Motivo = ""` y `p_Error = ""`.
- **DADO** un `p_IDExpediente` que tiene al menos 1 fila en `TbExpedientes WHERE IDExpedientePadre = p_IDExpediente` **CUANDO** se invoca `Helper_ExpedienteEliminacion.TieneDerivados(p_ID, ...)` **ENTONCES** retorna `"NO"` y `p_Motivo` contiene `"De este elemento derivan otros expedientes (Lotes o Basados). Elimínelos primero."`.
- **DADO** un `p_IDExpediente` que tiene al menos 1 fila en `TbExpedientesAnexos WHERE IDExpediente = p_IDExpediente` **CUANDO** se invoca `Helper_ExpedienteEliminacion.TieneAnexos(p_ID, ...)` **ENTONCES** retorna `"NO"` y `p_Motivo` contiene `"Este expediente tiene Anexos. Elimínelos primero."`.
- **DADO** un `p_IDExpediente` que tiene al menos 1 fila en `TbExpedientesSuministradores WHERE IDExpediente = p_IDExpediente` **CUANDO** se invoca `Helper_ExpedienteEliminacion.TieneSuministradores(p_ID, ...)` **ENTONCES** retorna `"NO"` y `p_Motivo` contiene `"Este expediente tiene Suministradores. Elimínelos primero."`.
- **DADO** un `p_IDExpediente` que pasa los 3 checks (derivados, anexos, suministradores) **CUANDO** se invoca `Helper_ExpedienteEliminacion.PuedeEliminar(p_ID, ...)` **ENTONCES** retorna `"OK"` con `p_Motivo = ""` y `p_Error = ""`.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | Aplicada en codigo? | Prueba | Confianza |
|---|---|---|---|---|---|
| **BR-26-01** | `TieneDerivados` con `p_IDExpediente <= 0` retorna `"ERR"` y `p_Error` poblada | spec PRUEBA-003 6.1 (input validation) | Si - `Helper_ExpedienteEliminacion.TieneDerivados` | `Test_Helper_ExpedienteEliminacion_TieneDerivados_IDCero_PueblaError` - PASA 2026-06-15 | Verified-runtime |
| **BR-26-01b** | `TieneDerivados` con `p_IDExpediente` que NO tiene hijos retorna `"OK"` | spec | Si - query `SELECT COUNT(*) FROM TbExpedientes WHERE IDExpedientePadre=p_IDExpediente` | `Test_Helper_ExpedienteEliminacion_TieneDerivados_IDInexistente_DevuelveOK` - PASA 2026-06-15 | Verified-runtime |
| **BR-26-01c** | `TieneDerivados` con `p_IDExpediente` que tiene >=1 hijo retorna `"NO"` + motivo | spec | Si - `TbExpedientes.IDExpedientePadre` self-FK | `Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo` - PASA 2026-06-15 (seed parent + child en `TbExpedientes`, valida, teardown defensivo) | Verified-runtime |
| **BR-26-02** | `TieneAnexos` con `p_IDExpediente` que tiene >=1 anexo retorna `"NO"` + motivo | spec PRUEBA-002 3.2 BR-03-02 (interpretado) | Si - `TbExpedientesAnexos.IDExpediente` FK | `Test_Helper_ExpedienteEliminacion_TieneAnexos_ConAnexoSembrado_DevuelveNOMotivo` - PASA 2026-06-16 (seed parent + 1 anexo, valida, teardown) | Verified-runtime |
| **BR-26-03** | `TieneSuministradores` con `p_IDExpediente` que tiene >=1 suministrador retorna `"NO"` + motivo | spec | Si - `TbExpedientesSuministradores.IDExpediente` FK (con FK enforced a `TbSuministradores.IDSuministrador`) | `Test_Helper_ExpedienteEliminacion_TieneSuministradores_ConSuministradorSembrado_DevuelveNOMotivo` - PASA 2026-06-16 (seed parent + 1 suministrador con IDSuministrador=1 real, valida, teardown) | Verified-runtime |
| **BR-26-04** | `PuedeEliminar` retorna `"OK"` solo si los 3 checks (BR-26-01, BR-26-02, BR-26-03) pasan | spec | Si - wrapper | `Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDInexistente_DevuelveOK` - PASA 2026-06-15 + `_IDCero_PueblaError` - PASA 2026-06-15 | Verified-runtime |
| **BR-26-05** | `EliminarExpediente(p_IDExpediente, p_Db, p_Error)` ejecuta el cascade delete de 12 tablas hijas + parent en transaccion (BeginTrans/CommitTrans con rollback). DAO-injectable. | spec PRUEBA-002 3.2 BR-03-01 | Si - `Helper_ExpedienteEliminacion.EliminarExpediente` (mismo orden que el class method original) | `Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa` - PASA 2026-06-16 (seed parent + 4 child tables, valida cascade, teardown) | Verified-runtime |

### Mensajes de error
- `MSG-26-01`: `"TieneDerivados: p_IDExpediente must be > 0 (got N)"` (input validation).
- `MSG-26-02`: `"TieneDerivados: <Err.Description>"` (DAO error, propagado).
- `MSG-26-03`: `"De este elemento derivan otros expedientes (Lotes o Basados). Elimínelos primero."` (BR-26-01c motivo).
- `MSG-26-04`: `"Este expediente tiene Anexos. Elimínelos primero."` (BR-26-02 motivo).
- `MSG-26-05`: `"Este expediente tiene Suministradores. Elimínelos primero."` (BR-26-03 motivo).

(Los mensajes estan en espanol porque la audiencia son administradores que ven el UI; las claves internas son los IDs de regla.)

## 3 Especificacion tecnica (~ design SDD) - COMO
- **Modulo**: `src/modules/Helper_ExpedienteEliminacion.bas` (stateless, `.bas`, 178 LOC, sin estado de modulo, sin referencias a `Me.*` ni a controles de form).
- **Helpers consumidos**: ninguno.
- **DAO consumido**: `getdb()` de `Variables Globales.bas` (si el caller no inyecta una `p_Db`).
- **Sin popup, sin MsgBox**: el helper NO muestra UI. El MsgBox de confirmacion del codigo original (`pregunta = MsgBox("Desea realmente borrar...", vbExclamation + vbYesNo + ...)`) se queda en el form por ahora; el helper solo hace la validacion.
- **Convenciones respetadas**:
  - `Optional ByRef p_Error As String` ultimo parametro (per convencion EXPEDIENTES).
  - `Optional ByRef p_Motivo As String = ""` para el mensaje de razon.
  - `Optional ByVal p_Db As DAO.Database = Nothing` para DAO injection.
  - Retorno `String` con 3 valores: `"OK"`, `"NO"`, `"ERR"`.
- **Decisiones de diseno**:
  - **DAO injection via `Optional p_Db`**: permite tests deterministas con un backend sandbox. Si el caller no inyecta, el helper usa `getdb()`.
  - **3 checks en vez de 5**: la validacion original (`MotivoEliminarNoOK`) tiene 5 checks. El helper implementa 3 (derivados, anexos, suministradores) que son los que mas pesan. Los 2 restantes (DPDs/Agedys via `TbExpAgedys`, GR via `TbProyectosGestionRiesgos`, HPS, NCs via `TbNoConformidades`) se difieren a un PR posterior para mantener este slice chiquito.
  - **Sin cascade delete todavia**: la operacion DELETE real vive como class method `ExpedienteOperaciones.Eliminar`. El helper se enfoca en la VALIDACION. La cascade se extrae en PR-C (post-PRUEBA-002).
  - **Sin rewiring del form todavia**: los 3 handlers `ComandoEliminarAMoC/Lote/Basado_Click` siguen llamando al `Eliminar` interno del form (que llama a `m_ExpedienteOp.Eliminar`). El rewire es trivial (reemplazar la llamada `m_ExpedienteOp.Eliminar` con `Helper_ExpedienteEliminacion.PuedeEliminar` antes del MsgBox) pero se difiere al PR-C.

## 4 Operacion (~ tasks SDD) - QUIEN/HACE CUANDO
- **PR owner**: PR-REFAC-1c (este slice).
- **Reviewers**: 1 (mantenedor).
- **Tests anadidos en este slice**: 5 verde (3 BR-26-01, 2 BR-26-04).
- **Tareas realizadas**:
  - T1c.1 baseline: pre-snapshot OK.
  - T1c.2 leido: `Form_FormExpedientesGestion.cls` lineas 2743-2949 (3 handlers + 1 funcion `Eliminar`); `ExpedienteOperaciones.cls` lineas 2346-2526 (`.Eliminar` + `.MotivoEliminarNoOK`).
  - T1c.3 creado: `Helper_ExpedienteEliminacion.bas` con 4 funciones publicas + 1 privada (`ResolveDb`).
  - T1c.4 (RED->GREEN): 5 tests verde (3 BR-26-01 + 2 BR-26-04).
  - T1c.5 (extract): helpers stateless, 3 checks implementados, 2 diferidos.
  - T1c.6 (refactor form): NO refactorizado - diferido a PR-C.
  - T1c.8 (no shim): no aplica - el form llama a class method, no a function global.
  - T1c.9 (PR gate): `compile_vba` PASS, `test_vba` 5/5 PASS, regresion 1a+1b 4/4 PASS.

## 5 Referencias cruzadas
- Spec original: `staging-alignment-prueba-002/specs/capabilities/CAP-003-eliminacion.spec.md` (BR-03-01..02 wish de cascade; la impl es la validacion previa).
- Spec del helper: `staging-alignment-prueba-003/specs/helpers/Helper_ExpedienteEliminacion.spec.md` (TBD).
- Apply progress: `staging-alignment-prueba-003/apply-progress.md` "PR-REFAC-1c cumulative state" (este turno).
- Tasks: `staging-alignment-prueba-003/tasks.md` "PR-REFAC-1c".
- Feature doc: `docs/features/F-ELIMINACION-helper-expediente-eliminacion.md` (este turno).
- Class method original: `ExpedienteOperaciones.cls` lineas 2346-2526 (`.Eliminar` + `.MotivoEliminarNoOK`).
- PRUEBA-002 PR-C: extraer la cascade delete real al helper + tests.
- Engram: `sdd/expedientes/capability-map/explore` 1 (ciclo de vida).

## 6 Lagunas explicitas
- **BR-26-02 (anexos) y BR-26-03 (suministradores)**: las queries existen pero los tests no estan. Target: agregar 2 tests con seed en `TbExpedientesAnexos` y `TbExpedientesSuministradores` (mismo patron que el seed test de BR-26-01c).
- **BR-26-05 (cascade delete)**: la operacion DELETE real no se extrajo al helper. Sigue como class method `ExpedienteOperaciones.Eliminar`. Target: PR-C, donde se extrae + testea con seed completo (padre + hijos en `TbExpedientesConEntidades`, `TbExpedientesCadenaContratacion`).
- **Checks 4 y 5 del original (DPDs/Agedys, GR, HPS, NCs)**: no implementados. 4 checks mas para un PR posterior.
- **T1c.6 (rewire del form)**: los 3 handlers `ComandoEliminarAMoC/Lote/Basado_Click` siguen llamando al `Eliminar` interno del form. El rewire es trivial (agregar 1 linea `If Helper_ExpedienteEliminacion.PuedeEliminar(...) <> "OK" Then Exit Sub` antes del MsgBox) pero se difiere a PR-C para mantener este slice cohesivo.

## 7 Resumen de cobertura
- Total reglas: 5 (BR-26-01..05) + 2 sub-rules (BR-26-01b, BR-26-01c).
- `Verified-runtime`: 11 (BR-26-01, BR-26-01b, BR-26-01c, BR-26-02, BR-26-03, BR-26-04 x2, BR-26-05 x2, _EliminarExpediente_IDCero, _EliminarExpediente_CascadaExitosa).
- `Verified-static`: 4 checks (DPDs/Agedys via `TbExpAgedys`, GR via `TbProyectosGestionRiesgos`, HPS via `TbExpedientesHPS`, NCs via `TbNoConformidades`) - not implemented yet. Target: post-PR-C, same pattern as the 3 implemented checks.
- Deuda: T1c.6 (rewire de los 3 handlers `ComandoEliminar*_Click` en `Form_FormExpedientesGestion.cls`) - el form sigue llamando al `Eliminar` interno (que llama al class method). Trivial rewire deferred.
