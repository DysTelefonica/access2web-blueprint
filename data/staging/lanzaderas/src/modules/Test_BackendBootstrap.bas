Attribute VB_Name = "Test_BackendBootstrap"
Option Compare Database
Option Explicit

Private Const FIXTURE_BACKEND_PROD As String = "\\fixture\backend\prod.accdb"
Private Const FIXTURE_BACKEND_LOCAL As String = "C:\fixture\backend\local.accdb"
Private Const FIXTURE_APP_PROD As String = "\\fixture\app\prod"
Private Const FIXTURE_APP_LOCAL As String = "C:\fixture\app\local"
Private Const FIXTURE_APP_ID As String = "12"

Private Function SeedRuntimeConfig(ByVal p_BackendActivo As String, _
                                   ByVal p_EnPruebas As String, _
                                   Optional ByRef p_Error As String = "") As Boolean
    SeedRuntimeConfig = Test_ConfigFixtures.SeedSingleConfigRow( _
                        p_BackendActivo:=p_BackendActivo, _
                        p_BackendProduccion:=FIXTURE_BACKEND_PROD, _
                        p_BackendSandbox:=FIXTURE_BACKEND_LOCAL, _
                        p_PasswordBackend:="", _
                        p_IDAplicacion:=FIXTURE_APP_ID, _
                        p_RutaDirectorioAplicacion_PROD:=FIXTURE_APP_PROD, _
                        p_RutaDirectorioAplicacion_LOCAL:=FIXTURE_APP_LOCAL, _
                        p_EnPruebas:=p_EnPruebas, _
                        p_Error:=p_Error)
End Function

Public Function Test_BackendBootstrap_LeeConfiguracionLocal_RespetaBackendConfigurado() As String
    On Error GoTo EH

    Dim logs(0 To 6) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultValue As String
    Dim resultJson As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendBootstrap_LeeConfiguracionLocal_RespetaBackendConfigurado = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed fila válida BackendActivo=PROD y EnPruebas=No"
    If Not SeedRuntimeConfig("PROD", "No", errMsg) Then
        Test_BackendBootstrap_LeeConfiguracionLocal_RespetaBackendConfigurado = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ResetGlobals + LeeConfiguracionLocal()"
    Application.TempVars.RemoveAll
    ResetGlobals errMsg
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg
    resultValue = LeeConfiguracionLocal(errMsg)
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg

    logs(3) = "4. Assert: LeeConfiguracionLocal devuelve OK"
    If resultValue <> "OK" Then
        resultJson = Test_Helper.BuildJsonFail("LeeConfiguracionLocal devolvió: " & resultValue, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: BackendActivo runtime=PROD"
    If UCase$(Trim$(m_BackendActivo)) <> "PROD" Then
        resultJson = Test_Helper.BuildJsonFail("m_BackendActivo esperado=PROD; actual=" & m_BackendActivo, logs)
        GoTo Cleanup
    End If

    logs(5) = "6. Assert: TempVars BackendActivo y ruta activa reflejan la semilla"
    If UCase$(Trim$(Test_Helper.SafeTempVarValue("BackendActivo"))) <> "PROD" Then
        resultJson = Test_Helper.BuildJsonFail("TempVar BackendActivo esperado=PROD; actual=" & Test_Helper.SafeTempVarValue("BackendActivo"), logs)
        GoTo Cleanup
    End If
    If Test_Helper.SafeTempVarValue("RutaDirectorioAplicacion") <> NormalizarRuta(FIXTURE_APP_PROD) Then
        resultJson = Test_Helper.BuildJsonFail("RutaDirectorioAplicacion no respetó la ruta PROD sembrada", logs)
        GoTo Cleanup
    End If

    logs(6) = "7. Assert: cardinalidad de config permanece en 1"
    If Test_ConfigFixtures.CountConfigRows(errMsg) <> 1 Then
        resultJson = Test_Helper.BuildJsonFail("TbConfiguracionBackends debe quedar con una única fila válida", logs)
        GoTo Cleanup
    End If

    resultJson = Test_Helper.BuildJsonOk("backend_config_respected", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendBootstrap_LeeConfiguracionLocal_RespetaBackendConfigurado = resultJson
    Exit Function

EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_BackendBootstrap_LeeConfiguracionLocal_ForceLocalSoloConParametro() As String
    On Error GoTo EH

    Dim logs(0 To 6) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultValue As String
    Dim resultJson As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendBootstrap_LeeConfiguracionLocal_ForceLocalSoloConParametro = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed fila válida BackendActivo=PROD y EnPruebas=No"
    If Not SeedRuntimeConfig("PROD", "No", errMsg) Then
        Test_BackendBootstrap_LeeConfiguracionLocal_ForceLocalSoloConParametro = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ResetGlobals + LeeConfiguracionLocal(..., LOCAL)"
    Application.TempVars.RemoveAll
    ResetGlobals errMsg
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg
    resultValue = LeeConfiguracionLocal(errMsg, "LOCAL")
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg

    logs(3) = "4. Assert: LeeConfiguracionLocal devuelve OK"
    If resultValue <> "OK" Then
        resultJson = Test_Helper.BuildJsonFail("LeeConfiguracionLocal devolvió: " & resultValue, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: BackendActivo runtime=LOCAL"
    If UCase$(Trim$(m_BackendActivo)) <> "LOCAL" Then
        resultJson = Test_Helper.BuildJsonFail("m_BackendActivo esperado=LOCAL; actual=" & m_BackendActivo, logs)
        GoTo Cleanup
    End If

    logs(5) = "6. Assert: TempVars BackendActivo y ruta activa reflejan override"
    If UCase$(Trim$(Test_Helper.SafeTempVarValue("BackendActivo"))) <> "LOCAL" Then
        resultJson = Test_Helper.BuildJsonFail("TempVar BackendActivo esperado=LOCAL; actual=" & Test_Helper.SafeTempVarValue("BackendActivo"), logs)
        GoTo Cleanup
    End If
    If Test_Helper.SafeTempVarValue("RutaDirectorioAplicacion") <> NormalizarRuta(FIXTURE_APP_LOCAL) Then
        resultJson = Test_Helper.BuildJsonFail("RutaDirectorioAplicacion no respetó la ruta LOCAL sembrada", logs)
        GoTo Cleanup
    End If

    logs(6) = "7. Assert: cardinalidad de config permanece en 1"
    If Test_ConfigFixtures.CountConfigRows(errMsg) <> 1 Then
        resultJson = Test_Helper.BuildJsonFail("TbConfiguracionBackends debe quedar con una única fila válida", logs)
        GoTo Cleanup
    End If

    resultJson = Test_Helper.BuildJsonOk("force_local_only_with_parameter", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendBootstrap_LeeConfiguracionLocal_ForceLocalSoloConParametro = resultJson
    Exit Function

EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_BackendBootstrap_LeeConfiguracionLocal_NoFuerzaLocalPorEnPruebas() As String
    On Error GoTo EH

    Dim logs(0 To 5) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultJson As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendBootstrap_LeeConfiguracionLocal_NoFuerzaLocalPorEnPruebas = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed fila válida BackendActivo=PROD y EnPruebas=Sí"
    If Not SeedRuntimeConfig("PROD", "Sí", errMsg) Then
        Test_BackendBootstrap_LeeConfiguracionLocal_NoFuerzaLocalPorEnPruebas = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ResetGlobals + LeeConfiguracionLocal()"
    Application.TempVars.RemoveAll
    ResetGlobals errMsg
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg
    Call LeeConfiguracionLocal(errMsg)
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg

    logs(3) = "4. Assert: BackendActivo runtime sigue en PROD"
    If UCase$(Trim$(m_BackendActivo)) <> "PROD" Then
        resultJson = Test_Helper.BuildJsonFail("m_BackendActivo esperado=PROD; actual=" & m_BackendActivo, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: ruta activa resuelta desde PROD sembrado"
    If ResolverRutaSegunConfiguracion("LOCAL_PATH", "PROD_PATH", errMsg) <> "PROD_PATH" Then
        resultJson = Test_Helper.BuildJsonFail("ResolverRutaSegunConfiguracion no respetó BackendActivo=PROD", logs)
        GoTo Cleanup
    End If

    logs(5) = "6. Assert: cardinalidad de config permanece en 1"
    If Test_ConfigFixtures.CountConfigRows(errMsg) <> 1 Then
        resultJson = Test_Helper.BuildJsonFail("TbConfiguracionBackends debe quedar con una única fila válida", logs)
        GoTo Cleanup
    End If

    resultJson = Test_Helper.BuildJsonOk("en_pruebas_no_force_local", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendBootstrap_LeeConfiguracionLocal_NoFuerzaLocalPorEnPruebas = resultJson
    Exit Function
EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_BackendBootstrap_NombreAplicacionSegunConfiguracion_UsaEnPruebas() As String
    On Error GoTo EH

    Dim logs(0 To 5) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultJson As String
    Dim applicationName As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendBootstrap_NombreAplicacionSegunConfiguracion_UsaEnPruebas = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed fila válida BackendActivo=PROD y EnPruebas=Sí"
    If Not SeedRuntimeConfig("PROD", "Sí", errMsg) Then
        Test_BackendBootstrap_NombreAplicacionSegunConfiguracion_UsaEnPruebas = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ResetGlobals + NombreAplicacionSegunConfiguracion()"
    Application.TempVars.RemoveAll
    ResetGlobals errMsg
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg
    applicationName = NombreAplicacionSegunConfiguracion(errMsg)
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendBootstrap", errMsg

    logs(3) = "4. Assert: usa nombre de pruebas"
    If applicationName <> "LANZADERA PRUEBAS" Then
        resultJson = Test_Helper.BuildJsonFail("NombreAplicacionSegunConfiguracion esperado=LANZADERA PRUEBAS; actual=" & applicationName, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: TempVar EnPruebas refleja configuración sembrada"
    If Test_Helper.SafeTempVarValue("EnPruebas") <> "Sí" Then
        resultJson = Test_Helper.BuildJsonFail("TempVar EnPruebas esperado=Sí; actual=" & Test_Helper.SafeTempVarValue("EnPruebas"), logs)
        GoTo Cleanup
    End If

    logs(5) = "6. Assert: cardinalidad de config permanece en 1"
    If Test_ConfigFixtures.CountConfigRows(errMsg) <> 1 Then
        resultJson = Test_Helper.BuildJsonFail("TbConfiguracionBackends debe quedar con una única fila válida", logs)
        GoTo Cleanup
    End If

    resultJson = Test_Helper.BuildJsonOk("nombre_aplicacion_depende_config", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendBootstrap_NombreAplicacionSegunConfiguracion_UsaEnPruebas = resultJson
    Exit Function
EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

