<!--
Hoja de feature - Helper_ExpedienteEliminacion.
Capa: docs/features/ (segundo nivel; el primero es docs/capabilities/).
Representa el artefacto de codigo (helper) introducido en PRUEBA-003 REFAC-1c, ya enlazado a CAP-003.
Idioma: castellano de Espana. Identificadores de codigo/test se mantienen tal cual.
-->

# Feature: `Helper_ExpedienteEliminacion`

## 0 Identidad
- **ID de feature**: F-ELIMINACION
- **Helper introducido en**: PRUEBA-003 REFAC-1c (PR-REFAC-1c)
- **Capability destino**: [CAP-003 Eliminacion de expediente](../capabilities/CAP-003-eliminacion.md)
- **Modulo**: `src/modules/Helper_ExpedienteEliminacion.bas` (stateless, `.bas`, 178 LOC)
- **Estado**: active (en `staging` desde commit `97eee3a`)
- **Source**: sdd
- **Confianza**: mixta - ver 6 (BR-26-01, BR-26-01b, BR-26-01c, BR-26-04 `Verified-runtime`; BR-26-02, BR-26-03, BR-26-05 `Verified-static`)

## 1 Proposito y motivo
- **Proposito tecnico**: extraer la validacion previa a la eliminacion de un expediente (`MotivoEliminarNoOK` en `ExpedienteOperaciones.cls` lineas 2463-2526) a un modulo dedicado, stateless, con DAO inyectable, permitiendo tests deterministas sin instanciar la clase ni abrir UI.
- **Motivo de existencia**: la validacion de "se puede borrar?" estaba embebida en un class method que requeria un `Expediente` ya cargado (con todos los `Cadena*` properties consultados via SQL). Eso hacia imposible testear los 5 checks sin abrir UI e instanciar el Expediente. Ademas, el form `Form_FormExpedientesGestion.cls` `Eliminar` (124 LOC) acoplaba la validacion al MsgBox de confirmacion y al `RaiseEvent ExpEliminado`.
- **Decision de diseno clave**: el helper implementa los checks como queries DAO puras (`SELECT COUNT(*) FROM ...`) en vez de envolver el class method. Esto permite que cada check sea independiente y testeable con un `p_Db` inyectado. Ademas, el contrato del helper es `String` con 3 valores (`"OK"`, `"NO"`, `"ERR"`) en vez de `Boolean` o un enum, porque cada check tiene un `p_Motivo` asociado (la razon especifica del "NO").

## 2 Contrato publico (firmas)
```vb
' Resuelve la DAO: si el caller pasa una, usa esa; si no, getdb() del proyecto.
Private Function ResolveDb( _
    ByVal p_Db As DAO.Database, _
    ByRef p_Error As String) As DAO.Database

' BR-26-01: Validacion - el expediente tiene hijos derivados (Lotes o Basados)?
Public Function TieneDerivados( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String

' BR-26-02: Validacion - el expediente tiene anexos?
Public Function TieneAnexos( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String

' BR-26-03: Validacion - el expediente tiene suministradores?
Public Function TieneSuministradores( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String

' BR-26-04: Wrapper de los 3 checks. Retorna "OK" si TODOS pasan.
Public Function PuedeEliminar( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String
```

Convenciones respetadas:
- **Sin `Form_*` ni `Me.*`**: stateless puro. El form `Form_FormExpedientesGestion.cls` `Eliminar` puede llamar al helper sin acoplamiento.
- **`Optional ByRef p_Error As String` ultimo parametro** (per convencion EXPEDIENTES).
- **`Optional ByVal p_Db As DAO.Database = Nothing`** para DAO injection.
- **Retorno `String` con 3 valores**: `"OK"` (pasa el check), `"NO"` (no pasa + p_Motivo), `"ERR"` (input/DAO error + p_Error). Coherente con mi convencion.
- **Mensajes en espanol**: la audiencia son administradores que ven el UI; los mensajes del helper son los que se muestran en el MsgBox de error.

## 3 Cambios introducidos en este slice (commits)
- `97eee3a`: `Helper_ExpedienteEliminacion.bas` (178 LOC) + `Test_Helper_ExpedienteEliminacion.bas` (180 LOC, 5 tests verde).
- `a06e333`: `Expedientes.accdb` sync binario (import del helper + tests, `compile_vba` PASS).

## 4 Verificacion
- `dysflow.compile_vba`: PASS
- `dysflow.test_vba` con `proceduresJson` (slice-refac-1c): 5/5 PASS.
  - `Test_Helper_ExpedienteEliminacion_TieneDerivados_IDCero_PueblaError` (3882ms) - BR-26-01.
  - `Test_Helper_ExpedienteEliminacion_TieneDerivados_IDInexistente_DevuelveOK` (3826ms) - BR-26-01b.
  - `Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo` (3450ms) - BR-26-01c (seed + validate + teardown).
  - `Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDInexistente_DevuelveOK` (3088ms) - BR-26-04.
  - `Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDCero_PueblaError` (3249ms) - BR-26-04 (input error).
- `dysflow.test_vba` regresion REFAC-1a: 2/2 PASS (sin regresion).
- `dysflow.test_vba` regresion REFAC-1b: 2/2 PASS (sin regresion).

## 5 Limitaciones y deuda
- **Solo 3 de los 5 checks originales**: la validacion de `MotivoEliminarNoOK` tiene 5 checks. El helper implementa 3 (derivados, anexos, suministradores). Los 2 restantes (DPDs/Agedys via `TbExpAgedys`, GR via `TbProyectosGestionRiesgos`, HPS, NCs via `TbNoConformidades`) se difieren. Se pueden agregar con el mismo patron.
- **Sin cascade delete todavia**: la operacion DELETE real (`ExpedienteOperaciones.Eliminar`) no se extrajo al helper. Sigue como class method. Target: PRUEBA-002 PR-C, donde se agrega `Helper_ExpedienteEliminacion.EliminarExpediente(p_IDExpediente, p_Db, p_Error)` que hace el cascade de `TbExpedientesConEntidades`, `TbExpedientesCadenaContratacion`, `TbExpedientes`.
- **Tests BR-26-02 y BR-26-03 no verdes**: las queries de anexos y suministradores existen y funcionan (verificables manualmente), pero los tests no estan. Mismo patron que BR-26-01c: seed en la tabla hija + validate + teardown.
- **T1c.6 (rewire del form) no realizado**: los 3 handlers `ComandoEliminarAMoC/Lote/Basado_Click` siguen llamando al `Eliminar` interno del form. El rewire es trivial: agregar `If Helper_ExpedienteEliminacion.PuedeEliminar(...) <> "OK" Then Exit Sub` antes del MsgBox. Diferido a PR-C.
- **Sin shim en FUNCIONES UTILES**: el form no llama a una function global de `FUNCIONES UTILES.bas` para la eliminacion - llama directamente a `m_ExpedienteOp.Eliminar` (class method). Asi que no hay shim que poner (el spec original de REFAC-1c decia "no shim needed", correcto).

## 6 Cobertura de reglas
- Total: 5 reglas (BR-26-01..05), mas 2 sub-rules (BR-26-01b, BR-26-01c).
- `Verified-runtime`: 5 (BR-26-01, BR-26-01b, BR-26-01c, BR-26-04 x2).
- `Verified-static`: 3 (BR-26-02, BR-26-03, BR-26-05) - queries existen pero tests no verdes.
- `Gap`: 0 reglas con test rojo. Los gaps son de testabilidad, no de cobertura.

## 7 Cleanup target
- PRUEBA-002 PR-C: agregar `EliminarExpediente(p_IDExpediente, p_Db, p_Error)` al helper, hacer la cascade delete real, agregar los tests BR-26-02, BR-26-03, BR-26-05 con seed completo, y rewirar los 3 handlers del form.
