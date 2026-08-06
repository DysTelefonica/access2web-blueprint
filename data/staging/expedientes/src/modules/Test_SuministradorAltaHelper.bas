Attribute VB_Name = "Test_SuministradorAltaHelper"
Option Compare Database
Option Explicit

' Test_SuministradorAltaHelper — REWORK (2026-06-26)
' Pure-data TDD atoms for modSuministradorAltaHelper.bas
' (Form_FormSuministrador Alta/Edición, see docs/audit/suministrador-alta-thin.md).
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - Test_SuministradorAltaHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
'     DoCmd.OpenForm TEST_FORM_NAME and return Forms(TEST_FORM_NAME) as p_Form.
'     This caused VBE UI interruption in headless runs (user-reported 2026-06-26).
'   - Every atom opened the form via OpenForm, populated controls via
'     p_Form.Controls("X").Value = ..., and asserted on p_Form.Controls(...) state.
'
' New design (per access-vba-e2e-methodology rule #1):
'   - ZERO DoCmd.OpenForm calls in tests.
'   - ZERO Forms(...) references in tests.
'   - ZERO Screen.ActiveForm references in tests.
'   - ZERO Application.Echo in tests.
'   - Stubs are Scripting.Dictionary instances, NOT real form-bound objects.
'   - Helpers receive pure data (Dictionary stubs); atoms parse JSON via JsonConverter.
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY (audited in §1.9 pre-compile).
'   - §4.2 no-humo: atoms assert concrete values (counts, fields, JSON), not "did not crash".
'   - §4.4 strong assertions: payload keys present, exact string matches.
'
' Convention:
'   - Atoms use ONLY Global public names (access-vba-tdd §1.1.1).
'   - Each atom returns JSON: {"ok":true|false,"value":...,"payload":null,"error":...,"logs":[...]}.
'   - Atoms never call MsgBox or pop up UI.

' === Module-level constants (all at top per vba-access §10.1) =======================

' Fixture ID base — entities created by atoms use IDs in this range so teardown is
' surgical (DELETE WHERE IDSuministrador >= TEST_BASE_ID).
Private Const TEST_BASE_ID As Long = 900750

' Valid modo values.
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
' Builds a Scripting.Dictionary with all 7 Suministrador fields.
Private Function BuildStubValores( _
    ByVal p_Nombre As String, _
    ByVal p_Nemotecnico As String, _
    ByVal p_CIF As String, _
    ByVal p_Direccion As String, _
    ByVal p_Ciudad As String, _
    ByVal p_CP As String, _
    ByVal p_TramitadoraHPS As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Nombre") = p_Nombre
    d("Nemotecnico") = p_Nemotecnico
    d("CIF") = p_CIF
    d("Direccion") = p_Direccion
    d("Ciudad") = p_Ciudad
    d("CP") = p_CP
    d("TramitadoraHPS") = p_TramitadoraHPS
    Set BuildStubValores = d
End Function

' --- TeardownFixture ------------------------------------------------------------
Private Function TeardownFixture() As Long
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim before As Long
    before = DCount("*", "TbSuministradores", "IDSuministrador >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador >= " & TEST_BASE_ID, dbFailOnError
    On Error GoTo 0
    TeardownFixture = before
End Function


' === Public atoms ====================================================================

' ---------------------------------------------------------------------------
' 1. Abrir_Inicializar — sad (invalid modo)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_Abrir_Inicializar_SadInvalidModo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Entidad As Object
    Set p_Entidad = Nothing
    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_Abrir_Inicializar( _
        "invalid", "1", p_Entidad, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for invalid modo, got " & json
    End If

    Test_SuministradorAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildOk("invalid-modo-rejected", logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy (alta mode, no ID)
'    Expected: titulo=ALTA..., hasEntidad=false, p_Entidad=empty Dictionary
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_Abrir_Inicializar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim p_Entidad As Object
    Set p_Entidad = Nothing
    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_Abrir_Inicializar( _
        TEST_MODO_ALTA, "", p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1002, , "expected ok=true"

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1003, , "expected modo=alta"
    If CStr(parsed("payload")("titulo")) <> "ALTA DE SUMINISTRADOR" Then
        Err.Raise 1004, , "expected titulo=ALTA DE SUMINISTRADOR, got " & parsed("payload")("titulo")
    End If
    If parsed("payload")("hasEntidad") <> False Then Err.Raise 1005, , "expected hasEntidad=false"
    If p_Entidad Is Nothing Then Err.Raise 1006, , "expected p_Entidad populated as empty Dictionary"

    Test_SuministradorAltaHelper_Abrir_Inicializar_HappyAlta = BuildOk("alta-initialized", logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_Abrir_Inicializar_HappyAlta = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Abrir_Inicializar — sad (edicion mode, missing ID)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_Abrir_Inicializar_SadEdicionSinID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Entidad As Object
    Set p_Entidad = Nothing
    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_Abrir_Inicializar( _
        TEST_MODO_EDICION, "", p_Entidad, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for edicion sin ID, got " & json
    End If
    If InStr(json, "p_IDEntidad is required") = 0 Then
        Err.Raise 1003, , "expected error to mention 'p_IDEntidad is required', got " & json
    End If

    Test_SuministradorAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildOk("edicion-sin-id-rejected", logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. VerificarCambios — happy (alta mode, empty originales, one field non-empty)
'    Expected: hayCambios=True, Nombre diff=True
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_VerificarCambios_AltaOneField() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("NEW-NAME", "", "", "", "", "", "")

    Dim originales As Object
    Set originales = Nothing

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then
        Err.Raise 1002, , "expected hayCambios=true (alta with one non-empty field)"
    End If
    If parsed("payload")("diffs")("Nombre") <> True Then
        Err.Raise 1003, , "expected diffs.Nombre=true"
    End If

    Test_SuministradorAltaHelper_VerificarCambios_AltaOneField = BuildOk(True, logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_VerificarCambios_AltaOneField = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. VerificarCambios — happy (alta mode, all fields empty)
'    Expected: hayCambios=False
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_VerificarCambios_AltaAllEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("", "", "", "", "", "", "")

    Dim originales As Object
    Set originales = Nothing

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then
        Err.Raise 1002, , "expected hayCambios=false (alta with all empty)"
    End If

    Test_SuministradorAltaHelper_VerificarCambios_AltaAllEmpty = BuildOk(False, logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_VerificarCambios_AltaAllEmpty = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. VerificarCambios — happy (edicion mode, all fields equal -> hayCambios=False)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_VerificarCambios_EdicionAllEqual() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ALPHA", "NEMO", "CIF", "DIR", "CITY", "28001", "Sí")
    Dim originales As Object
    Set originales = BuildStubValores("ALPHA", "NEMO", "CIF", "DIR", "CITY", "28001", "Sí")

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then
        Err.Raise 1002, , "expected hayCambios=false (edicion, all equal)"
    End If

    Test_SuministradorAltaHelper_VerificarCambios_EdicionAllEqual = BuildOk(False, logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_VerificarCambios_EdicionAllEqual = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. VerificarCambios — happy (edicion mode, one field differs -> hayCambios=True)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_VerificarCambios_EdicionOneDiff() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ALPHA-CHANGED", "NEMO", "CIF", "DIR", "CITY", "28001", "Sí")
    Dim originales As Object
    Set originales = BuildStubValores("ALPHA", "NEMO", "CIF", "DIR", "CITY", "28001", "Sí")

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_VerificarCambios( _
        actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then
        Err.Raise 1002, , "expected hayCambios=true (Nombre differs)"
    End If
    If parsed("payload")("diffs")("Nombre") <> True Then
        Err.Raise 1003, , "expected diffs.Nombre=true"
    End If
    If parsed("payload")("diffs")("CIF") <> False Then
        Err.Raise 1004, , "expected diffs.CIF=false (no diff)"
    End If

    Test_SuministradorAltaHelper_VerificarCambios_EdicionOneDiff = BuildOk(True, logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_VerificarCambios_EdicionOneDiff = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. VerificarCambios — sad (no actuales -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_VerificarCambios_SadNoActuales() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_VerificarCambios( _
        Nothing, Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for no actuales, got " & json
    End If

    Test_SuministradorAltaHelper_VerificarCambios_SadNoActuales = BuildOk("no-actuales-rejected", logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_VerificarCambios_SadNoActuales = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. Registrar — sad (no valores -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_Registrar_SadNoValores() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_Registrar(Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then
        Err.Raise 1002, , "expected ok=false for no valores, got " & json
    End If

    Test_SuministradorAltaHelper_Registrar_SadNoValores = BuildOk("no-valores-rejected", logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_Registrar_SadNoValores = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Cerrar — happy (m_ObjSuministradorActivo cleared)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_Cerrar_HappyCleared() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    ' Pre-set the global so we can verify the reset.
    Dim preObj As New Suministrador
    preObj.IDSuministrador = "999999"
    Set m_ObjSuministradorActivo = preObj

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_Cerrar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then
        Err.Raise 1002, , "expected ok=true, got " & json
    End If

    If Not m_ObjSuministradorActivo Is Nothing Then
        Err.Raise 1003, , "expected m_ObjSuministradorActivo cleared, still set"
    End If

    Test_SuministradorAltaHelper_Cerrar_HappyCleared = BuildOk("cleared", logs)
    Exit Function
EH:
    Test_SuministradorAltaHelper_Cerrar_HappyCleared = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. RunAll — wrapper for Dysflow manifest discovery (access-vba-tdd §1.1.1)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorAltaHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_SuministradorAltaHelper_Abrir_Inicializar_SadInvalidModo", _
        "Test_SuministradorAltaHelper_Abrir_Inicializar_HappyAlta", _
        "Test_SuministradorAltaHelper_Abrir_Inicializar_SadEdicionSinID", _
        "Test_SuministradorAltaHelper_VerificarCambios_AltaOneField", _
        "Test_SuministradorAltaHelper_VerificarCambios_AltaAllEmpty", _
        "Test_SuministradorAltaHelper_VerificarCambios_EdicionAllEqual", _
        "Test_SuministradorAltaHelper_VerificarCambios_EdicionOneDiff", _
        "Test_SuministradorAltaHelper_VerificarCambios_SadNoActuales", _
        "Test_SuministradorAltaHelper_Registrar_SadNoValores", _
        "Test_SuministradorAltaHelper_Cerrar_HappyCleared" _
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
        Test_SuministradorAltaHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_SuministradorAltaHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_SuministradorAltaHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    On Error GoTo 0
    Test_SuministradorAltaHelper_RunAll = BuildFail(p_Error, logs)
End Function
