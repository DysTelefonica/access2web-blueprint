Attribute VB_Name = "Test_TestingHarness"
Option Compare Database
Option Explicit

Public Function Test_TestingHarness_ForceLocalBackend_ActivaModoTest() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim errMsg As String

    logs(0) = "1. Arrange: reset test session"
    Test_Helper.ResetTestSession errMsg
    If errMsg <> "" Then
        Test_TestingHarness_ForceLocalBackend_ActivaModoTest = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Act: ForceLocalBackend"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_TestingHarness_ForceLocalBackend_ActivaModoTest = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Assert: m_TestingMode=True"
    If Not m_TestingMode Then
        Test_TestingHarness_ForceLocalBackend_ActivaModoTest = Test_Helper.BuildJsonFail("m_TestingMode no fue activado", logs)
        GoTo Cleanup
    End If

    logs(3) = "4. Assert: backend de test configurado"
    If Trim$(GetTestingBackendURL()) = "" Then
        Test_TestingHarness_ForceLocalBackend_ActivaModoTest = Test_Helper.BuildJsonFail("GetTestingBackendURL devolvió vacío", logs)
        GoTo Cleanup
    End If

    Test_TestingHarness_ForceLocalBackend_ActivaModoTest = Test_Helper.BuildJsonOk("testing_mode_enabled", logs)

Cleanup:
    Test_Helper.ResetTestSession errMsg
    Exit Function
EH:
    Test_TestingHarness_ForceLocalBackend_ActivaModoTest = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_TestingHarness_ResetTestSession_DesactivaModoTest() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim errMsg As String

    logs(0) = "1. Arrange: ForceLocalBackend activa el modo test"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_TestingHarness_ResetTestSession_DesactivaModoTest = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Act: ResetTestSession"
    If Not Test_Helper.ResetTestSession(errMsg) Then
        Test_TestingHarness_ResetTestSession_DesactivaModoTest = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Assert: m_TestingMode=False"
    If m_TestingMode Then
        Test_TestingHarness_ResetTestSession_DesactivaModoTest = Test_Helper.BuildJsonFail("m_TestingMode sigue activo", logs)
        Exit Function
    End If

    logs(3) = "4. Assert: backend de test limpiado"
    If GetTestingBackendURL() <> "" Then
        Test_TestingHarness_ResetTestSession_DesactivaModoTest = Test_Helper.BuildJsonFail("GetTestingBackendURL no se limpió", logs)
        Exit Function
    End If

    Test_TestingHarness_ResetTestSession_DesactivaModoTest = Test_Helper.BuildJsonOk("testing_mode_disabled", logs)
    Exit Function
EH:
    Test_TestingHarness_ResetTestSession_DesactivaModoTest = Test_Helper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_TestingHarness_Getdb_UsaSandboxEnModoTest() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim errMsg As String
    Dim dbSandbox As DAO.Database
    Dim expectedPath As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_TestingHarness_Getdb_UsaSandboxEnModoTest = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    expectedPath = GetTestingBackendURL()

    logs(1) = "2. Act: getdb() en modo test"
    Set dbSandbox = getdb(, errMsg)
    If errMsg <> "" Then
        Test_TestingHarness_Getdb_UsaSandboxEnModoTest = Test_Helper.BuildJsonFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(2) = "3. Assert: conexión abierta"
    If dbSandbox Is Nothing Then
        Test_TestingHarness_Getdb_UsaSandboxEnModoTest = Test_Helper.BuildJsonFail("getdb devolvió Nothing", logs)
        GoTo Cleanup
    End If

    logs(3) = "4. Assert: la conexión apunta al sandbox"
    If StrComp(dbSandbox.Name, expectedPath, vbTextCompare) <> 0 Then
        Test_TestingHarness_Getdb_UsaSandboxEnModoTest = Test_Helper.BuildJsonFail("getdb abrió " & dbSandbox.Name & " en lugar de " & expectedPath, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: m_TestingMode sigue activo"
    If Not m_TestingMode Then
        Test_TestingHarness_Getdb_UsaSandboxEnModoTest = Test_Helper.BuildJsonFail("m_TestingMode se desactivó durante getdb", logs)
        GoTo Cleanup
    End If

    Test_TestingHarness_Getdb_UsaSandboxEnModoTest = Test_Helper.BuildJsonOk("getdb_uses_sandbox", logs)

Cleanup:
    Test_Helper.ResetTestSession errMsg
    Set dbSandbox = Nothing
    Exit Function
EH:
    Test_TestingHarness_Getdb_UsaSandboxEnModoTest = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function
