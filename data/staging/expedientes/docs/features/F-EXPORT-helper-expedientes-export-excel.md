<!--
Hoja de feature - Helper_ExpedientesExportExcel.
Capa: docs/features/ (segundo nivel; el primero es docs/capabilities/).
Representa el artefacto de código (helper) introducido en PRUEBA-003 REFAC-1b, ya enlazado a CAP-025.
Idioma: castellano de España. Identificadores de código/test se mantienen tal cual.
-->

# Feature: `Helper_ExpedientesExportExcel`

## §0 Identidad
- **ID de feature**: F-EXPORT
- **Helper introducido en**: PRUEBA-003 REFAC-1b (PR-REFAC-1b)
- **Capability destino**: [CAP-025 Exportación Excel](../capabilities/CAP-025-excel.md)
- **Módulo**: `src/modules/Helper_ExpedientesExportExcel.bas` (stateless, `.bas`, 327 LOC)
- **Estado**: active (en `staging` desde commit `dbafcc3`)
- **Source**: sdd
- **Confianza**: mixta - ver §6 (BR-25-01..02 `Verified-runtime`; BR-25-03..08 `Verified-static`)

## §1 Propósito y motivo
- **Propósito técnico**: extraer las funciones de exportación Excel `GenerarConsultaExpedientes`, `RellenarLinea`, `ConvertirATabla`, `AjustarCeldas` (originalmente en `FUNCIONES UTILES.bas`, ~180 LOC) a un módulo dedicado, stateless, con cancelación cooperativa vía `g_OperationCancelled` para responder al bug UX recurrente de popup que no se puede cerrar.
- **Motivo de existencia**: PRUEBA-002 §3.2 BR-25-01..02 requiere la capacidad de exportación; la lógica monolítica en `FUNCIONES UTILES.bas` no es testeable. La cancelación es un feedback explícito del usuario 2026-06-15: "si el popup no se puede cerrar, ¿qué pasa si Excel COM se cuelga?".
- **Decisión de diseño clave**: el chequeo de `g_OperationCancelled` está en 5 puntos del helper (pre-popup, post-popup, y en cada uno de los 3 loops). Si está set, el helper cierra el workbook sin guardar, sale de Excel, cierra el popup, y retorna `""`.

## §2 Contrato público (firmas)
```vb
' Genera el archivo Excel con AM/Lotes/Basados/Tecnica.
' Retorna la URL del .xlsx generado, o "" si no hay datos / se canceló.
Public Function GenerarConsultaExpedientes( _
    ByVal p_ColCampos As Scripting.Dictionary, _
    ByVal p_ColAM As Scripting.Dictionary, _
    ByVal p_ColLotes As Scripting.Dictionary, _
    ByVal p_ColBasados As Scripting.Dictionary, _
    ByVal p_ColTecnica As Scripting.Dictionary, _
    ByVal p_IncluirDerivados As Boolean, _
    Optional ByRef p_Error As String) As String

' Puebla la fila intFila de wbHoja con los valores del ExpedienteCompleto
' en el orden de p_ColCampos.Keys. Recursivo para Derivados si p_IncluirDerivados.
Public Function RellenarLinea( _
    ByVal p_ColCampos As Scripting.Dictionary, _
    ByRef m_ColExpUsados As Scripting.Dictionary, _
    ByRef wbHoja As Object, _
    ByRef intFila As Long, _
    ByVal p_ExpC As ExpedienteCompleto, _
    ByVal p_IncluirDerivados As Boolean, _
    Optional ByRef p_Error As String) As String

' Convierte el rango con datos de wbHoja en un ListObject (tabla estructurada).
Public Function ConvertirATabla( _
    ByVal p_Hoja As Object, _
    Optional ByRef p_Error As String) As String

' Aplica el formato por defecto a las celdas de p_Hoja.
Public Function AjustarCeldas( _
    ByVal p_Hoja As Object, _
    Optional ByRef p_Error As String) As String
```

Convenciones respetadas:
- **Sin `Form_*` ni `Me.*`**: stateless puro. El form `Form_FormExpedientesGestion.cls` `ComandoExportarExcel*_Click` handlers llaman a `GenerarConsultaExpedientes` (ahora shim) sin saber que el cuerpo está en otro módulo.
- **`Optional ByRef p_Error As String` último parámetro** (per convención EXPEDIENTES).
- **Globales consumidos**: `m_ObjEntorno` (provee `appExcel`, `URLDirectorioLocal`); `g_OperationCancelled` (cancelación).
- **Popup helpers consumidos** (no inyectables hoy): bloquean el COM test runner.

## §3 Cambios introducidos en este slice (commits)
- `dbafcc3` (size:exception 415 LOC): `Helper_ExpedientesExportExcel.bas` (327 LOC) + `Test_HelperExpedientesExportExcel.bas` (88 LOC, 2 tests verde + 1 slot @Deprecated).
- `823d938`: `FUNCIONES UTILES.bas` shim de 4 funciones (1 línea cada una) + `Public g_OperationCancelled As Boolean` global + reset del flag en `CerrarPopupProgreso`.
- `f90ca8b`: `Expedientes.accdb` sync binario (import del helper + shim + tests, `compile_vba` PASS).

## §4 Verificación
- `dysflow.compile_vba`: PASS.
- `dysflow.test_vba` con `proceduresJson` (slice-refac-1b): 2/2 PASS.
  - `Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_ColVacia_NoGeneraExcel` (5009ms) — BR-25-01.
  - `Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_pColCamposNothing_NoGeneraExcel` (3733ms) — BR-25-02.
- `dysflow.test_vba` regresión slice-refac-1a: 4/4 PASS (ConstruirWhereBusqueda 2 + GenerarFilasAM 1 + ContarExpedientes 1).
- `dysflow.verify_binary` scoped al helper: `actionableOk=true` con residual `caseOnly`/`encodingOnly` (no funcional).

## §5 Limitaciones y deuda
- **Excel COM bloquea tests**: la spec original de PRUEBA-002 §3.2 BR-25-01..02 mencionaba un `Exportar(Recordset, Columnas, Destino, Error)` CSV-style con UTF-8 BOM. Esa función NUNCA fue implementada en el código legacy. La implementación real es Excel COM (no CSV), por lo que el helper tiene 4 funciones (`GenerarConsultaExpedientes`, `RellenarLinea`, `ConvertirATabla`, `AjustarCeldas`) y NO un `Exportar`. La spec se reformula en este doc para alinearse con el código real.
- **Popup helpers bloquean el COM test runner**: las 3 funciones que invocan `MostrarPopupProgreso` (que llama `DoCmd.OpenForm "frmBusy"`) cuelgan el test runner. Por eso solo BR-25-01..02 (early-exit) son `Verified-runtime` y BR-25-03..08 son `Verified-static`.
- **UX gap** ✅ **CERRADO en commit `641262d`**: el form `frmBusy` ahora tiene la X estándar del sistema habilitada (`ControlBox = 0` + `CloseButton = 0` en `.form.txt`) y un handler `Form_Unload` que setea `g_OperationCancelled = True`. El usuario cierra el popup con la X; el helper aborta Excel limpiamente. Cancel mechanism completo y operativo. PRUEBA-002 PR-F se cierra por la vía del approach alternativo (X del sistema en lugar de botón custom).

## §6 Cobertura de reglas
- Total: 8 reglas (BR-25-01..08).
- `Verified-runtime`: 2 (BR-25-01, BR-25-02) — ramas testeables sin popup.
- `Verified-static`: 6 (BR-25-03..08) — ramas que tocan popup/Excel COM.
- `Gap`: 0 reglas con test rojo. El gap es de testabilidad, no de cobertura.

## §7 Cleanup target
- PRUEBA-002 PR-E: borrar los 4 shims de `FUNCIONES UTILES.bas` (post-confirm de cobertura del helper). El call site `Form_FormExpedientesGestion.cls` debe ser actualizado en ese PR para llamar al helper directamente, no al shim.
