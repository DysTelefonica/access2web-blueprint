Attribute VB_Name = "Test_TestingModeSandbox"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_TestingModeSandbox_ForceLocalBackend_ActivatesTestingMode() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim errMsg As String
    logs(0) = "1. Arrange: reset test session"
    Test_Helper.ResetTestSession errMsg

    logs(1) = "2. Act: ForceLocalBackend validates BackendSandbox"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_TestingModeSandbox_ForceLocalBackend_ActivatesTestingMode = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Assert: m_TestingMode=True"
    If Not m_TestingMode Then
        Test_TestingModeSandbox_ForceLocalBackend_ActivatesTestingMode = BuildFail("m_TestingMode no fue activado", logs)
        Exit Function
    End If

    logs(3) = "4. Assert: m_BackendSandboxURL configurado"
    If m_BackendSandboxURL = "" Then
        Test_TestingModeSandbox_ForceLocalBackend_ActivatesTestingMode = BuildFail("m_BackendSandboxURL quedo vacio", logs)
        Test_Helper.ResetTestSession errMsg
        Exit Function
    End If

    Test_TestingModeSandbox_ForceLocalBackend_ActivatesTestingMode = BuildOk("testing_mode_active", logs)
    Test_Helper.ResetTestSession errMsg
    Exit Function

EH:
    Test_TestingModeSandbox_ForceLocalBackend_ActivatesTestingMode = BuildFail(Err.Description, logs)
    Test_Helper.ResetTestSession errMsg
End Function

Public Function Test_TestingModeSandbox_ResetTestSession_DisablesTestingMode() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim errMsg As String
    logs(0) = "1. Arrange: ForceLocalBackend activa modo testing"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_TestingModeSandbox_ResetTestSession_DisablesTestingMode = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Act: ResetTestSession"
    Test_Helper.ResetTestSession errMsg
    If errMsg <> "" Then
        Test_TestingModeSandbox_ResetTestSession_DisablesTestingMode = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Assert: m_TestingMode=False"
    If m_TestingMode Then
        Test_TestingModeSandbox_ResetTestSession_DisablesTestingMode = BuildFail("m_TestingMode sigue activo", logs)
        Exit Function
    End If

    logs(3) = "4. Assert: cache global marcado cerrado"
    If m_DBOpen Then
        Test_TestingModeSandbox_ResetTestSession_DisablesTestingMode = BuildFail("m_DBOpen sigue activo", logs)
        Exit Function
    End If

    Test_TestingModeSandbox_ResetTestSession_DisablesTestingMode = BuildOk("testing_mode_disabled", logs)
    Exit Function

EH:
    Test_TestingModeSandbox_ResetTestSession_DisablesTestingMode = BuildFail(Err.Description, logs)
End Function

Public Function Test_TestingModeSandbox_GetDbFailsWhenSandboxMissing() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim errMsg As String
    Dim originalTestingMode As Boolean
    Dim originalSandbox As String
    Dim originalTempVar As String

    logs(0) = "1. Arrange: capture global backend state"
    originalTestingMode = m_TestingMode
    originalSandbox = m_BackendSandboxURL
    originalTempVar = Nz(Application.TempVars("BackendPathConfigurado"), "")

    logs(1) = "2. Arrange: m_TestingMode=True con sandbox inexistente"
    m_TestingMode = True
    m_BackendSandboxURL = CurrentProject.path & "\__missing_testing_sandbox__.accdb"
    Application.TempVars("BackendPathConfigurado") = originalTempVar

    logs(2) = "3. Act: getdb"
    Dim dbSandbox As DAO.Database
    Set dbSandbox = getdb(errMsg)

    logs(3) = "4. Assert: getdb devuelve Nothing y error"
    If Not dbSandbox Is Nothing Then
        Test_TestingModeSandbox_GetDbFailsWhenSandboxMissing = BuildFail("getdb abrio una conexion con sandbox inexistente", logs)
        GoTo Cleanup
    End If
    If errMsg = "" Then
        Test_TestingModeSandbox_GetDbFailsWhenSandboxMissing = BuildFail("getdb no devolvio p_Error", logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: no se uso fallback a backend activo"
    Test_TestingModeSandbox_GetDbFailsWhenSandboxMissing = BuildOk("sandbox_missing_guarded", logs)

Cleanup:
    m_TestingMode = originalTestingMode
    m_BackendSandboxURL = originalSandbox
    If Len(originalTempVar) > 0 Then
        Application.TempVars("BackendPathConfigurado") = originalTempVar
    Else
        On Error Resume Next
        Application.TempVars.Remove "BackendPathConfigurado"
        On Error GoTo EH
    End If
    Set dbSandbox = Nothing
    Exit Function

EH:
    Test_TestingModeSandbox_GetDbFailsWhenSandboxMissing = BuildFail(Err.Description, logs)
    Resume Cleanup
End Function
