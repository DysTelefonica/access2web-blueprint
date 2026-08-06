Attribute VB_Name = "Test_DatosTipoSolicitudHelper_PCSUB_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_DatosTipoSolicitudHelper_PCSUB_Strict — v2.4.2 strict TDD atoms
'
' Skill:    access-vba-tdd v2.4.2
' SDD/CAP:  CAP-001 (PCSUB), delta alta-solicitud (Spec-002)
' Branch:   staging
' Project:  condor (Dysflow projectId)
'
' Helper under test (pure):
'   DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo(ByVal contratistaPrincipal As String) As String
'
' Contract (from design.md):
'   - Principal   -> contains PC and CD_CA, excludes PC_SUB and CD_CA_SUB
'   - Subcontratista -> contains PC, PC_SUB, CD_CA_SUB, excludes CD_CA
'   - Principal detection MUST trim and be case-insensitive for "Sí"
'   - Safe default for empty/unknown role: subcontratista list
'
' TDD skill v2.4.2 compliance
'   - Public Function Test_*() As String returning JSON {ok,value,payload,error,logs}
'   - Pure atoms: no DAO, no Screen, no MsgBox, no Debug.Print
'   - Helper is pure: no db needed, atoms run without BeginTestSession
' ============================================================================

' --- Fixture constants for RowSource exact-contract assertions ---
Private Const TOKEN_PC As String = "'PC';'PC - Propuesta de Cambio'"
Private Const TOKEN_CD_CA As String = "'CD_CA';'CD/CA - Concesión/Desviación'"
Private Const TOKEN_PC_SUB As String = "'PC_SUB';'PC-SUB - Subcontratista'"
Private Const TOKEN_CD_CA_SUB As String = "'CD_CA_SUB';'CD/CA-SUB - Subcontratista'"
Private Const EXPECTED_PRINCIPAL_ROWSOURCE As String = TOKEN_PC & ";" & TOKEN_CD_CA
Private Const EXPECTED_SUBCONTRATISTA_ROWSOURCE As String = TOKEN_PC & ";" & TOKEN_PC_SUB & ";" & TOKEN_CD_CA_SUB

' ============================================================================
' JSON wrappers (per skill §1.8 — must be local wrappers)
' ============================================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = TestHelper.BuildJsonOk(CStr(value), logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = TestHelper.BuildJsonFail(msg, logs)
End Function

' ============================================================================
' Pure helper assertion (no DAO, no UI)
' ============================================================================
Private Function RowSourceContains(ByVal rowSource As String, ByVal token As String) As Boolean
    RowSourceContains = (InStr(1, rowSource, token, vbBinaryCompare) > 0)
End Function

Private Function RowSourceEquals(ByVal rowSource As String, ByVal expectedRowSource As String) As Boolean
    RowSourceEquals = (StrComp(rowSource, expectedRowSource, vbBinaryCompare) = 0)
End Function

' ============================================================================
' ATOMS — Principal role
' ============================================================================
Public Function Test_DTS_Principal_Contains_PC_And_CD_CA() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: helper accepts principal role"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo(""Sí"")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("Sí")
    logs(2) = "3. Act: actual = """ & actual & """"

    If Not RowSourceEquals(actual, EXPECTED_PRINCIPAL_ROWSOURCE) Then
        errMsg = "Principal RowSource must equal exact contract; expected: " & EXPECTED_PRINCIPAL_ROWSOURCE & "; got: " & actual
        Test_DTS_Principal_Contains_PC_And_CD_CA = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Principal_Contains_PC_And_CD_CA = BuildOk("principal exact RowSource is PC and CD_CA", logs)
End Function

Public Function Test_DTS_Principal_Excludes_PC_SUB_And_CD_CA_SUB() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: principal must not expose subcontractor types"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo(""Sí"")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("Sí")
    logs(2) = "3. Act: actual = """ & actual & """"

    If RowSourceContains(actual, TOKEN_PC_SUB) Then
        errMsg = "Principal RowSource must NOT contain PC_SUB token; got: " & actual
        Test_DTS_Principal_Excludes_PC_SUB_And_CD_CA_SUB = BuildFail(errMsg, logs)
        Exit Function
    End If

    If RowSourceContains(actual, TOKEN_CD_CA_SUB) Then
        errMsg = "Principal RowSource must NOT contain CD_CA_SUB token; got: " & actual
        Test_DTS_Principal_Excludes_PC_SUB_And_CD_CA_SUB = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Principal_Excludes_PC_SUB_And_CD_CA_SUB = BuildOk("principal excludes PC_SUB and CD_CA_SUB", logs)
End Function

' ============================================================================
' ATOMS — Subcontratista role
' ============================================================================
Public Function Test_DTS_Subcontratista_Contains_PC_PC_SUB_CD_CA_SUB() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: subcontratista must see plain PC plus sub types"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo(""No"")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("No")
    logs(2) = "3. Act: actual = """ & actual & """"

    If Not RowSourceEquals(actual, EXPECTED_SUBCONTRATISTA_ROWSOURCE) Then
        errMsg = "Subcontratista RowSource must equal exact contract; expected: " & EXPECTED_SUBCONTRATISTA_ROWSOURCE & "; got: " & actual
        Test_DTS_Subcontratista_Contains_PC_PC_SUB_CD_CA_SUB = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Subcontratista_Contains_PC_PC_SUB_CD_CA_SUB = BuildOk("subcontratista exact RowSource is PC, PC_SUB, CD_CA_SUB", logs)
End Function

Public Function Test_DTS_Subcontratista_Excludes_CD_CA() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: subcontratista must not expose CD_CA"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo(""No"")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("No")
    logs(2) = "3. Act: actual = """ & actual & """"

    If RowSourceContains(actual, TOKEN_CD_CA) Then
        errMsg = "Subcontratista RowSource must NOT contain CD_CA token; got: " & actual
        Test_DTS_Subcontratista_Excludes_CD_CA = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Subcontratista_Excludes_CD_CA = BuildOk("subcontratista excludes CD_CA", logs)
End Function

' ============================================================================
' ATOMS — Safe defaults
' ============================================================================
Public Function Test_DTS_Empty_Role_Defaults_To_Subcontratista() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: empty role is safe default to subcontratista"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo("""")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("")
    logs(2) = "3. Act: actual = """ & actual & """"

    If Not RowSourceEquals(actual, EXPECTED_SUBCONTRATISTA_ROWSOURCE) Then
        errMsg = "Empty role must default to exact subcontratista RowSource; expected: " & EXPECTED_SUBCONTRATISTA_ROWSOURCE & "; got: " & actual
        Test_DTS_Empty_Role_Defaults_To_Subcontratista = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Empty_Role_Defaults_To_Subcontratista = BuildOk("empty role defaults to exact subcontratista RowSource", logs)
End Function

Public Function Test_DTS_Unknown_Role_Defaults_To_Subcontratista() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: unknown role is safe default to subcontratista"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo(""UnknownRole"")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("UnknownRole")
    logs(2) = "3. Act: actual = """ & actual & """"

    If Not RowSourceEquals(actual, EXPECTED_SUBCONTRATISTA_ROWSOURCE) Then
        errMsg = "Unknown role must default to exact subcontratista RowSource; expected: " & EXPECTED_SUBCONTRATISTA_ROWSOURCE & "; got: " & actual
        Test_DTS_Unknown_Role_Defaults_To_Subcontratista = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Unknown_Role_Defaults_To_Subcontratista = BuildOk("unknown role defaults to exact subcontratista RowSource", logs)
End Function

' ============================================================================
' ATOMS — Trim and case-insensitive "Sí"
' ============================================================================
Public Function Test_DTS_Principal_SI_CaseInsensitive_Lowercase() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: lowercase 'sí' must be recognized as principal"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo(""sí"")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("sí")
    logs(2) = "3. Act: actual = """ & actual & """"

    If Not RowSourceEquals(actual, EXPECTED_PRINCIPAL_ROWSOURCE) Then
        errMsg = "Lowercase 'sí' must map to exact principal RowSource; expected: " & EXPECTED_PRINCIPAL_ROWSOURCE & "; got: " & actual
        Test_DTS_Principal_SI_CaseInsensitive_Lowercase = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Principal_SI_CaseInsensitive_Lowercase = BuildOk("lowercase sí recognized as exact principal RowSource", logs)
End Function

Public Function Test_DTS_Principal_SI_Trimmed_Whitespace() As String
    Dim logs(0 To 2) As String
    Dim actual As String
    Dim errMsg As String

    logs(0) = "1. Arrange: '  Sí  ' (with surrounding whitespace) must be principal"
    logs(1) = "2. Act: ConstruirRowSourceComboTipo(""  Sí  "")"
    actual = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo("  Sí  ")
    logs(2) = "3. Act: actual = """ & actual & """"

    If Not RowSourceEquals(actual, EXPECTED_PRINCIPAL_ROWSOURCE) Then
        errMsg = "Whitespace-padded 'Sí' must map to exact principal RowSource; expected: " & EXPECTED_PRINCIPAL_ROWSOURCE & "; got: " & actual
        Test_DTS_Principal_SI_Trimmed_Whitespace = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_DTS_Principal_SI_Trimmed_Whitespace = BuildOk("whitespace-padded Sí recognized as exact principal RowSource", logs)
End Function
