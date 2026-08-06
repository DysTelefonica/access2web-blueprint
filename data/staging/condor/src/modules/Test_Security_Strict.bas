Attribute VB_Name = "Test_Security_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Security_Strict — CAP-010 Seguridad strict TDD v2.4.2 atoms (Fase B6, 2026-06-15)
'
' Touched objects: in-memory Usuario + UsuarioAplicacionPermisos.
' DeterminarRol is pure logic over the Usuario object — no DB access.
' Test the rol hierarchy: Administrador > Calidad > Tecnico, with Permisos=Nothing
' falling back to Tecnico.
'
' v2.4.2 migration: SetupSecuritySandbox (private) replaced with the canonical
'   TestHelper.BeginTestSession(logs, errMsg) + ApplyProdEnvForTest. The
'   test does not open a DAO.Database because DeterminarRol is pure
'   logic; the BeginTestSession call is kept for lifecycle consistency
'   (skill §3) and to validate the sandbox URL via ForceLocalBackend.
' ============================================================================

Private Const ID_USR_ADMIN As Long = 901011
Private Const ID_USR_CALIDAD As Long = 901012
Private Const ID_USR_TECNICO As Long = 901013

' ----------------------------------------------------------------------------
' Fixture builders (pure in-memory; no DB access)
' ----------------------------------------------------------------------------

Private Sub BuildUsuarioAdministrador(ByRef p_Usuario As Usuario)
    Set p_Usuario = New Usuario
    p_Usuario.nombre = "Admin Test"
    p_Usuario.EsAdministrador = "Sí"
    ' Permisos left as Nothing: Administrador check fires first.
    Set p_Usuario.Permisos = Nothing
End Sub

Private Sub BuildUsuarioCalidad(ByRef p_Usuario As Usuario)
    Set p_Usuario = New Usuario
    p_Usuario.nombre = "Calidad Test"
    p_Usuario.EsAdministrador = "No"
    Dim perms As New UsuarioAplicacionPermisos
    perms.EsUsuarioCalidad = "Sí"
    perms.EsUsuarioTecnico = "No"
    Set p_Usuario.Permisos = perms
End Sub

Private Sub BuildUsuarioTecnico(ByRef p_Usuario As Usuario)
    Set p_Usuario = New Usuario
    p_Usuario.nombre = "Tecnico Test"
    p_Usuario.EsAdministrador = "No"
    ' Permisos left as Nothing: defaults to Tecnico.
    Set p_Usuario.Permisos = Nothing
End Sub

' ----------------------------------------------------------------------------
' ApplyProdEnvForTest + EnsureFolder + CleanupTempRoot migrados a
' TestHelper.SetupProdGlobalsForTest / TestHelper.CleanupProdTempRoot
' (v2.4.3). Cada test ahora llama directamente al helper canónico.
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' Atoms
' ----------------------------------------------------------------------------

Public Function Test_Security_Strict_DeterminarRol_ReturnsAdministrador_WhenEsAdministradorSi() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim tempRoot As String
    Dim svc As New UsuarioServicio
    Dim usr As Usuario
    Dim result As rol

    logs = TestHelper.NewLogsArray(4)

    ' 1. BeginTestSession + apply prod env (no db: pure logic)
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_Security_Strict_DeterminarRol_ReturnsAdministrador_WhenEsAdministradorSi = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap010_sec", "CAP010 Test User")
    logs(0) = "1. BeginTestSession OK; prod env applied"

    ' 2. Act
    Call BuildUsuarioAdministrador(usr)
    logs(1) = "2. Built Usuario with EsAdministrador='Sí' and Permisos=Nothing"

    result = svc.DeterminarRol(usr)
    logs(2) = "3. DeterminarRol returned: " & result

    ' 3. Assert
    If result <> rol.Administrador Then
        logs(3) = "4. FAILED: expected rol.Administrador=" & rol.Administrador & ", got " & result
        Test_Security_Strict_DeterminarRol_ReturnsAdministrador_WhenEsAdministradorSi = _
            TestHelper.BuildJsonFail("expected Administrador", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. result=Administrador as expected"

    Test_Security_Strict_DeterminarRol_ReturnsAdministrador_WhenEsAdministradorSi = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    Set usr = Nothing
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Security_Strict_DeterminarRol_ReturnsAdministrador_WhenEsAdministradorSi = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Security_Strict_DeterminarRol_ReturnsCalidad_WhenEsUsuarioCalidadSi() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim tempRoot As String
    Dim svc As New UsuarioServicio
    Dim usr As Usuario
    Dim result As rol

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_Security_Strict_DeterminarRol_ReturnsCalidad_WhenEsUsuarioCalidadSi = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap010_sec", "CAP010 Test User")
    logs(0) = "1. BeginTestSession OK; prod env applied"

    Call BuildUsuarioCalidad(usr)
    logs(1) = "2. Built Usuario with EsAdministrador='No', Permisos.EsUsuarioCalidad='Sí'"

    result = svc.DeterminarRol(usr)
    logs(2) = "3. DeterminarRol returned: " & result

    If result <> rol.Calidad Then
        logs(3) = "4. FAILED: expected rol.Calidad=" & rol.Calidad & ", got " & result
        Test_Security_Strict_DeterminarRol_ReturnsCalidad_WhenEsUsuarioCalidadSi = _
            TestHelper.BuildJsonFail("expected Calidad", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. result=Calidad as expected"

    Test_Security_Strict_DeterminarRol_ReturnsCalidad_WhenEsUsuarioCalidadSi = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    Set usr = Nothing
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Security_Strict_DeterminarRol_ReturnsCalidad_WhenEsUsuarioCalidadSi = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Security_Strict_DeterminarRol_ReturnsTecnico_ByDefault() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim tempRoot As String
    Dim svc As New UsuarioServicio
    Dim usr As Usuario
    Dim result As rol

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_Security_Strict_DeterminarRol_ReturnsTecnico_ByDefault = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap010_sec", "CAP010 Test User")
    logs(0) = "1. BeginTestSession OK; prod env applied"

    Call BuildUsuarioTecnico(usr)
    logs(1) = "2. Built Usuario with EsAdministrador='No' and Permisos=Nothing (default fallback)"

    result = svc.DeterminarRol(usr)
    logs(2) = "3. DeterminarRol returned: " & result

    If result <> rol.Tecnico Then
        logs(3) = "4. FAILED: expected rol.Tecnico=" & rol.Tecnico & ", got " & result
        Test_Security_Strict_DeterminarRol_ReturnsTecnico_ByDefault = _
            TestHelper.BuildJsonFail("expected Tecnico", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. result=Tecnico as expected (fallback when Permisos=Nothing)"

    Test_Security_Strict_DeterminarRol_ReturnsTecnico_ByDefault = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    Set usr = Nothing
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Security_Strict_DeterminarRol_ReturnsTecnico_ByDefault = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim tempRoot As String
    Dim svc As New UsuarioServicio
    Dim result As rol
    Dim errorNumber As Long
    Dim errorDescription As String

    logs = TestHelper.NewLogsArray(5)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap010_sec", "CAP010 Test User")
    logs(0) = "1. BeginTestSession OK; prod env applied"

    On Error Resume Next
    result = svc.DeterminarRol(Nothing)
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo EH
    logs(1) = "2. Called DeterminarRol(Nothing)"

    If errorNumber = 0 Then
        logs(2) = "3. FAILED: expected CondorError Raise (513), but no error fired (result=" & result & ")"
        Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing = _
            TestHelper.BuildJsonFail("expected Raise, got " & result, logs)
        GoTo CleanExit
    End If
    logs(2) = "3. Caught error " & errorNumber & ": " & errorDescription

    If errorNumber <> 513 Then
        logs(3) = "4. FAILED: expected errorNumber=513, got " & errorNumber
        Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing = _
            TestHelper.BuildJsonFail("expected 513, got " & errorNumber, logs)
        GoTo CleanExit
    End If
    logs(3) = "4. errorNumber=513 as expected (custom application error)"

    Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ----------------------------------------------------------------------------
' Smoke aggregator — globally unique Public Function per §1.1.1
' ----------------------------------------------------------------------------
Public Function Test_Security_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String
    Dim result As String

    logs = TestHelper.NewLogsArray(5)

    result = Test_Security_Strict_DeterminarRol_ReturnsAdministrador_WhenEsAdministradorSi()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(0) = "1. DeterminarRol returns Administrador when EsAdministrador='Sí'"

    result = Test_Security_Strict_DeterminarRol_ReturnsCalidad_WhenEsUsuarioCalidadSi()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(1) = "2. DeterminarRol returns Calidad when Permisos.EsUsuarioCalidad='Sí'"

    result = Test_Security_Strict_DeterminarRol_ReturnsTecnico_ByDefault()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(2) = "3. DeterminarRol returns Tecnico as fallback when Permisos=Nothing"

    result = Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(3) = "4. DeterminarRol raises CondorError when Usuario is Nothing"

    logs(4) = "5. PASS"
    Test_Security_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

Failed:
    Test_Security_Strict_RunAll = TestHelper.BuildJsonFail("one CAP-010 Security atom failed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Security_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_Security_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
