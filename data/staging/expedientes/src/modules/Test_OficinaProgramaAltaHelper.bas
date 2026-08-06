Attribute VB_Name = "Test_OficinaProgramaAltaHelper"
Option Compare Database
Option Explicit

' Test_OficinaProgramaAltaHelper — REWORK (2026-06-26)
' Pure-data TDD atoms for modOficinaProgramaAltaHelper.bas
' (Form_FormOficinaPrograma Alta/Edición, see docs/audit/oficina-programa-alta-thin.md).
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - Test_OficinaProgramaAltaHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
'     DoCmd.OpenForm TEST_FORM_NAME and return Forms(TEST_FORM_NAME) as p_Form.
'     This caused VBE UI interruption in headless runs (user-reported 2026-06-26).
'
' New design (per access-vba-e2e-methodology rule #1):
'   - ZERO DoCmd.OpenForm calls in tests.
'   - ZERO Forms(...) references in tests.
'   - Stubs are Scripting.Dictionary instances, NOT real form-bound objects.
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY.
'   - §4.2 no-humo: atoms assert concrete values.
'
' Convention:
'   - Atoms use ONLY Global public names (access-vba-tdd §1.1.1).
'   - Each atom returns JSON: {"ok":true|false,"value":...,"payload":null,"error":...,"logs":[...]}.

' === Module-level constants (all at top per vba-access §10.1) =======================

Private Const TEST_BASE_ID As Long = 900730

Private Const TEST_MODO_ALTA As String = "alta"
Private Const TEST_MODO_EDICION As String = "edicion"


' === Local helpers (all at top per vba-access §10.1) ================================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubValores -----------------------------------------------------------
' Builds a Scripting.Dictionary with both OficinaPrograma fields.
Private Function BuildStubValores( _
    ByVal p_OficinaPrograma As String, _
    ByVal p_Descripcion As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("OficinaPrograma") = p_OficinaPrograma
    d("DESCRIPCION") = p_Descripcion
    Set BuildStubValores = d
End Function


' === Public atoms ====================================================================

' ---------------------------------------------------------------------------
' 1. Abrir_Inicializar — sad (invalid modo)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadInvalidModo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Entidad As Object
    Set p_Entidad = Nothing
    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Abrir_Inicializar( _
        "invalid", "1", p_Entidad, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for invalid modo, got " & json
    End If

    Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildOk("invalid-modo-rejected", logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy (alta mode, no ID)
'    Expected: titulo=ALTA..., hasEntidad=false, p_Entidad=empty Dictionary
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_Abrir_Inicializar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim p_Entidad As Object
    Set p_Entidad = Nothing
    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Abrir_Inicializar( _
        TEST_MODO_ALTA, "", p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1002, , "expected ok=true"

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1003, , "expected modo=alta"
    If CStr(parsed("payload")("titulo")) <> "ALTA DE OFICINA DE PROGRAMA" Then
        Err.Raise 1004, , "expected titulo=ALTA DE OFICINA DE PROGRAMA, got " & parsed("payload")("titulo")
    End If
    If parsed("payload")("hasEntidad") <> False Then Err.Raise 1005, , "expected hasEntidad=false"
    If p_Entidad Is Nothing Then Err.Raise 1006, , "expected p_Entidad populated as empty Dictionary"

    Test_OficinaProgramaAltaHelper_Abrir_Inicializar_HappyAlta = BuildOk("alta-initialized", logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_Abrir_Inicializar_HappyAlta = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Abrir_Inicializar — sad (edicion mode, missing ID)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadEdicionSinID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Entidad As Object
    Set p_Entidad = Nothing
    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Abrir_Inicializar( _
        TEST_MODO_EDICION, "", p_Entidad, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for edicion sin ID, got " & json
    End If
    If InStr(json, "p_IDEntidad is required") = 0 Then
        Err.Raise 1003, , "expected error to mention 'p_IDEntidad is required', got " & json
    End If

    Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildOk("edicion-sin-id-rejected", logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. VerificarCambios — happy (alta mode, one field non-empty)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_VerificarCambios_AltaOneField() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("NEW-OF", "")

    Dim originales As Object
    Set originales = Nothing

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then
        Err.Raise 1002, , "expected hayCambios=true (alta with one non-empty field)"
    End If
    If parsed("payload")("diffs")("OficinaPrograma") <> True Then
        Err.Raise 1003, , "expected diffs.OficinaPrograma=true"
    End If

    Test_OficinaProgramaAltaHelper_VerificarCambios_AltaOneField = BuildOk(True, logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_VerificarCambios_AltaOneField = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. VerificarCambios — happy (alta mode, all fields empty)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_VerificarCambios_AltaAllEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("", "")

    Dim originales As Object
    Set originales = Nothing

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then
        Err.Raise 1002, , "expected hayCambios=false (alta with all empty)"
    End If

    Test_OficinaProgramaAltaHelper_VerificarCambios_AltaAllEmpty = BuildOk(False, logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_VerificarCambios_AltaAllEmpty = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. VerificarCambios — happy (edicion mode, all fields equal -> hayCambios=False)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionAllEqual() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ALPHA-OF", "DESC")
    Dim originales As Object
    Set originales = BuildStubValores("ALPHA-OF", "DESC")

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then
        Err.Raise 1002, , "expected hayCambios=false (edicion, all equal)"
    End If

    Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionAllEqual = BuildOk(False, logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionAllEqual = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. VerificarCambios — happy (edicion mode, one field differs -> hayCambios=True)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionOneDiff() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ALPHA-OF-CHANGED", "DESC")
    Dim originales As Object
    Set originales = BuildStubValores("ALPHA-OF", "DESC")

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then
        Err.Raise 1002, , "expected hayCambios=true (OficinaPrograma differs)"
    End If
    If parsed("payload")("diffs")("OficinaPrograma") <> True Then
        Err.Raise 1003, , "expected diffs.OficinaPrograma=true"
    End If
    If parsed("payload")("diffs")("DESCRIPCION") <> False Then
        Err.Raise 1004, , "expected diffs.DESCRIPCION=false"
    End If

    Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionOneDiff = BuildOk(True, logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionOneDiff = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. VerificarCambios — sad (no actuales -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_VerificarCambios_SadNoActuales() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_VerificarCambios( _
        Nothing, Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for no actuales, got " & json
    End If

    Test_OficinaProgramaAltaHelper_VerificarCambios_SadNoActuales = BuildOk("no-actuales-rejected", logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_VerificarCambios_SadNoActuales = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. Registrar — sad (no valores -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_Registrar_SadNoValores() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Registrar(Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for no valores, got " & json
    End If

    Test_OficinaProgramaAltaHelper_Registrar_SadNoValores = BuildOk("no-valores-rejected", logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_Registrar_SadNoValores = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Cerrar — happy (m_ObjOficinaProgramaActiva cleared)
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_Cerrar_HappyCleared() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    ' Pre-set the global so we can verify the reset.
    Dim preObj As New OficinaPrograma
    preObj.IDOficinaPrograma = "999999"
    Set m_ObjOficinaProgramaActiva = preObj

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Cerrar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then
        Err.Raise 1002, , "expected ok=true, got " & json
    End If

    If Not m_ObjOficinaProgramaActiva Is Nothing Then
        Err.Raise 1003, , "expected m_ObjOficinaProgramaActiva cleared, still set"
    End If

    Test_OficinaProgramaAltaHelper_Cerrar_HappyCleared = BuildOk("cleared", logs)
    Exit Function
EH:
    Test_OficinaProgramaAltaHelper_Cerrar_HappyCleared = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_OficinaProgramaAltaHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadInvalidModo", _
        "Test_OficinaProgramaAltaHelper_Abrir_Inicializar_HappyAlta", _
        "Test_OficinaProgramaAltaHelper_Abrir_Inicializar_SadEdicionSinID", _
        "Test_OficinaProgramaAltaHelper_VerificarCambios_AltaOneField", _
        "Test_OficinaProgramaAltaHelper_VerificarCambios_AltaAllEmpty", _
        "Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionAllEqual", _
        "Test_OficinaProgramaAltaHelper_VerificarCambios_EdicionOneDiff", _
        "Test_OficinaProgramaAltaHelper_VerificarCambios_SadNoActuales", _
        "Test_OficinaProgramaAltaHelper_Registrar_SadNoValores", _
        "Test_OficinaProgramaAltaHelper_Cerrar_HappyCleared" _
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
        Test_OficinaProgramaAltaHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_OficinaProgramaAltaHelper_RunAll = BuildOk("all-passed", logs)
    End If
    Exit Function
EH:
    Dim p_Error As String
    p_Error = "Test_OficinaProgramaAltaHelper_RunAll EH: " & Err.Description
    Test_OficinaProgramaAltaHelper_RunAll = BuildFail(p_Error, logs)
End Function
