Attribute VB_Name = "Test_ExpedienteHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data TDD atoms for modExpedienteHelper.bas
' (Form_FormExpediente, see docs/audit/form-expediente-pure-data.md).
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY.

' === Module-level constants (all at top per vba-access §10.1) ====================

Private Const TEST_BASE_ID As Long = 900840


' === Local helpers (all at top per vba-access §10.1) ============================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubExpediente ---------------------------------------------------------
Private Function BuildStubExpediente( _
    ByVal p_IDExpediente As String, _
    ByVal p_TituloFormulario As String, _
    ByVal p_UltimaModificacionTexto As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDExpediente") = p_IDExpediente
    d("TituloFormulario") = p_TituloFormulario
    d("UltimaModificacionTexto") = p_UltimaModificacionTexto
    Set BuildStubExpediente = d
End Function

' --- BuildStubDTO ----------------------------------------------------------------
Private Function BuildStubDTO(ByVal p_Expediente As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("Expediente") = p_Expediente
    Set BuildStubDTO = d
End Function


' === Public atoms ===============================================================

' ---------------------------------------------------------------------------
' 1. Form_Load_Init — happy admin editable
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_Form_Load_Init_HappyAdminEditable() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_Form_Load_Init( _
        "", True, "TEST", True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("abiertoParaEditar")) <> True Then _
        Err.Raise 1002, , "expected abiertoParaEditar=true"
    If CBool(parsed("payload")("allowEdits")) <> True Then _
        Err.Raise 1002, , "expected allowEdits=true"
    If CBool(parsed("payload")("comandoRegistrarEnabled")) <> True Then _
        Err.Raise 1002, , "expected comandoRegistrarEnabled=true"

    Test_ExpedienteHelper_Form_Load_Init_HappyAdminEditable = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_Form_Load_Init_HappyAdminEditable = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Form_Load_Init — happy non-admin no edit
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_Form_Load_Init_HappyNonAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_Form_Load_Init( _
        "", False, "TEST", True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("allowEdits")) <> False Then _
        Err.Raise 1002, , "expected allowEdits=false for non-admin"

    Test_ExpedienteHelper_Form_Load_Init_HappyNonAdmin = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_Form_Load_Init_HappyNonAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Form_Load_Init — Sololectura
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_Form_Load_Init_Sololectura() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_Form_Load_Init( _
        "Sololectura", True, "TEST", True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("abiertoParaEditar")) <> False Then _
        Err.Raise 1002, , "expected abiertoParaEditar=false for Sololectura"
    If CBool(parsed("payload")("allowEdits")) <> False Then _
        Err.Raise 1002, , "expected allowEdits=false for Sololectura"

    Test_ExpedienteHelper_Form_Load_Init_Sololectura = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_Form_Load_Init_Sololectura = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. Form_Load_Init — edge (empty backend -> caption empty brackets)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_Form_Load_Init_EdgeEmptyBackend() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_Form_Load_Init( _
        "", True, "", True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If InStr(CStr(parsed("payload")("caption")), " []") = 0 Then _
        Err.Raise 1002, , "expected caption with empty brackets"

    Test_ExpedienteHelper_Form_Load_Init_EdgeEmptyBackend = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_Form_Load_Init_EdgeEmptyBackend = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. EstablecerDatos — happy existente (id set)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_EstablecerDatos_HappyExistente() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("EXP-1", "Mi Título", "2026-06-26 10:00"))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_EstablecerDatos(stub, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("lblTitulo")) <> "Mi Título" Then _
        Err.Raise 1002, , "expected lblTitulo='Mi Título'"
    If CBool(parsed("payload")("lblUltimaModificacionVisible")) <> True Then _
        Err.Raise 1002, , "expected lblUltimaModificacionVisible=true"
    If CBool(parsed("payload")("comandoActualizarCompletoVisible")) <> True Then _
        Err.Raise 1002, , "expected comandoActualizarCompletoVisible=true"
    If CBool(parsed("payload")("esNuevo")) <> False Then _
        Err.Raise 1002, , "expected esNuevo=false"

    Test_ExpedienteHelper_EstablecerDatos_HappyExistente = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_EstablecerDatos_HappyExistente = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. EstablecerDatos — happy nuevo (id empty)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_EstablecerDatos_HappyNuevo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_EstablecerDatos(stub, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("lblUltimaModificacionVisible")) <> False Then _
        Err.Raise 1002, , "expected lblUltimaModificacionVisible=false for nuevo"
    If CBool(parsed("payload")("comandoActualizarCompletoVisible")) <> False Then _
        Err.Raise 1002, , "expected comandoActualizarCompletoVisible=false for nuevo"
    If CBool(parsed("payload")("esNuevo")) <> True Then _
        Err.Raise 1002, , "expected esNuevo=true"

    Test_ExpedienteHelper_EstablecerDatos_HappyNuevo = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_EstablecerDatos_HappyNuevo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. EstablecerDatos — sad (DTO Nothing)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_EstablecerDatos_SadNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_EstablecerDatos(Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteHelper_EstablecerDatos_SadNothing = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_EstablecerDatos_SadNothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. EstablecerDatos — sad (DTO without Expediente)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_EstablecerDatos_SadNoExpediente() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_EstablecerDatos(stub, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteHelper_EstablecerDatos_SadNoExpediente = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_EstablecerDatos_SadNoExpediente = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. DebeGuardarPestana — happy General editable con cambios
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_DebeGuardarPestana_HappyEditableConCambios() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_DebeGuardarPestana( _
        "FormExpedienteGeneral", True, True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("debeGuardar")) <> True Then _
        Err.Raise 1002, , "expected debeGuardar=true"

    Test_ExpedienteHelper_DebeGuardarPestana_HappyEditableConCambios = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_DebeGuardarPestana_HappyEditableConCambios = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. DebeGuardarPestana — happy General editable sin cambios -> false
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_DebeGuardarPestana_EditableSinCambios() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_DebeGuardarPestana( _
        "FormExpedienteGeneral", True, True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("debeGuardar")) <> False Then _
        Err.Raise 1002, , "expected debeGuardar=false (no changes)"

    Test_ExpedienteHelper_DebeGuardarPestana_EditableSinCambios = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_DebeGuardarPestana_EditableSinCambios = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. DebeGuardarPestana — non-General tab -> false
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_DebeGuardarPestana_NonGeneralTab() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_DebeGuardarPestana( _
        "FormExpedienteEntidades", True, True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("debeGuardar")) <> False Then _
        Err.Raise 1002, , "expected debeGuardar=false for non-General tab"

    Test_ExpedienteHelper_DebeGuardarPestana_NonGeneralTab = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_DebeGuardarPestana_NonGeneralTab = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. DebeGuardarPestana — edge (allowEdits=false -> false)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_DebeGuardarPestana_EdgeAllowEditsFalse() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_DebeGuardarPestana( _
        "FormExpedienteFechas", True, False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("debeGuardar")) <> False Then _
        Err.Raise 1002, , "expected debeGuardar=false when allowEdits=false"

    Test_ExpedienteHelper_DebeGuardarPestana_EdgeAllowEditsFalse = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_DebeGuardarPestana_EdgeAllowEditsFalse = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. Tab_Seleccionar_Guardar — editable -> intentarGuardar=true
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_Tab_Seleccionar_Guardar_Editable() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_Tab_Seleccionar_Guardar(True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("intentarGuardar")) <> True Then _
        Err.Raise 1002, , "expected intentarGuardar=true"

    Test_ExpedienteHelper_Tab_Seleccionar_Guardar_Editable = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_Tab_Seleccionar_Guardar_Editable = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. Tab_Seleccionar_Guardar — no editable -> intentarGuardar=false
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_Tab_Seleccionar_Guardar_NoEditable() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteHelper.Expediente_Tab_Seleccionar_Guardar(False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("intentarGuardar")) <> False Then _
        Err.Raise 1002, , "expected intentarGuardar=false when not editable"

    Test_ExpedienteHelper_Tab_Seleccionar_Guardar_NoEditable = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteHelper_Tab_Seleccionar_Guardar_NoEditable = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_ExpedienteHelper_Form_Load_Init_HappyAdminEditable", _
        "Test_ExpedienteHelper_Form_Load_Init_HappyNonAdmin", _
        "Test_ExpedienteHelper_Form_Load_Init_Sololectura", _
        "Test_ExpedienteHelper_Form_Load_Init_EdgeEmptyBackend", _
        "Test_ExpedienteHelper_EstablecerDatos_HappyExistente", _
        "Test_ExpedienteHelper_EstablecerDatos_HappyNuevo", _
        "Test_ExpedienteHelper_EstablecerDatos_SadNothing", _
        "Test_ExpedienteHelper_EstablecerDatos_SadNoExpediente", _
        "Test_ExpedienteHelper_DebeGuardarPestana_HappyEditableConCambios", _
        "Test_ExpedienteHelper_DebeGuardarPestana_EditableSinCambios", _
        "Test_ExpedienteHelper_DebeGuardarPestana_NonGeneralTab", _
        "Test_ExpedienteHelper_DebeGuardarPestana_EdgeAllowEditsFalse", _
        "Test_ExpedienteHelper_Tab_Seleccionar_Guardar_Editable", _
        "Test_ExpedienteHelper_Tab_Seleccionar_Guardar_NoEditable" _
    )

    Dim passed As Long
    Dim failed As Long
    Dim firstFailure As String
    Dim i As Long
    For i = LBound(atoms) To UBound(atoms)
        Dim result As String
        result = Application.Run(CStr(atoms(i)))
        If InStr(result, """ok"":true") > 0 Then
            passed = passed + 1
        Else
            failed = failed + 1
            If Len(firstFailure) = 0 Then firstFailure = CStr(atoms(i)) & " -> " & result
        End If
    Next i

    logs(0) = "RunAll: passed=" & passed & " failed=" & failed
    If failed > 0 Then
        logs(1) = "RunAll: firstFailure=" & firstFailure
        Test_ExpedienteHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_ExpedienteHelper_RunAll = BuildOk("all-passed", logs)
    End If
    Exit Function
EH:
    Dim p_Error As String
    p_Error = "Test_ExpedienteHelper_RunAll EH: " & Err.Description
    Test_ExpedienteHelper_RunAll = BuildFail(p_Error, logs)
End Function