Attribute VB_Name = "Test_BackendConfigHelper"
Option Compare Database
Option Explicit

Private Const EXPECTED_BACKEND_LOCAL As String = "C:\00repos\datos\Lanzadera_Datos.accdb"
Private Const EXPECTED_BACKEND_REMOTO As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
Private Const EXPECTED_APP_REMOTA As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera"
Private Const EXPECTED_APP_LOCAL As String = "C:\Users\adm1\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\0Lanzadera"
Private Const FIXTURE_APP_ID As String = "12"
Private Const FIXTURE_APP_ID_UPDATE As String = "99"
Private Const FIXTURE_PASSWORD_BACKEND As String = "fixture-password"
Private Const FIXTURE_EN_PRUEBAS As String = "Sí"

Private Function AssertHelperRow(ByRef p_Logs() As String, _
                                 Optional ByVal p_ExpectedIDAplicacion As String = FIXTURE_APP_ID, _
                                 Optional ByVal p_ExpectedEnPruebas As String = "No", _
                                 Optional ByVal p_ExpectedPasswordBackend As String = "", _
                                 Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rowCount As Long

    rowCount = Test_ConfigFixtures.CountConfigRows(p_Error)
    If p_Error <> "" Then Exit Function
    If rowCount <> 1 Then
        p_Error = "TbConfiguracionBackends debe quedar con exactamente una fila; actual=" & CStr(rowCount)
        Exit Function
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM TbConfiguracionBackends", dbOpenSnapshot)
    If rs.EOF Then
        p_Error = "TbConfiguracionBackends quedó sin filas tras ejecutar el helper"
        GoTo Cleanup
    End If

    If UCase$(Trim$(Nz(rs.Fields("BackendActivo").value, ""))) <> "LOCAL" Then
        p_Error = "BackendActivo esperado=LOCAL; actual=" & Nz(rs.Fields("BackendActivo").value, "")
        GoTo Cleanup
    End If

    If CStr(Nz(rs.Fields("BackendProduccion").value, "")) <> EXPECTED_BACKEND_REMOTO Then
        p_Error = "BackendProduccion no coincide con la ruta remota esperada"
        GoTo Cleanup
    End If
    If CStr(Nz(rs.Fields("BackendSandbox").value, "")) <> EXPECTED_BACKEND_LOCAL Then
        p_Error = "BackendSandbox no coincide con la ruta local esperada"
        GoTo Cleanup
    End If
    If CStr(Nz(rs.Fields("RutaDirectorioAplicacion_PROD").value, "")) <> NormalizarRuta(EXPECTED_APP_REMOTA) Then
        p_Error = "RutaDirectorioAplicacion_PROD no coincide con la ruta remota esperada"
        GoTo Cleanup
    End If
    If CStr(Nz(rs.Fields("RutaDirectorioAplicacion_LOCAL").value, "")) <> NormalizarRuta(EXPECTED_APP_LOCAL) Then
        p_Error = "RutaDirectorioAplicacion_LOCAL no coincide con la ruta local esperada"
        GoTo Cleanup
    End If
    If CStr(Nz(rs.Fields("IDAplicacion").value, "")) <> p_ExpectedIDAplicacion Then
        p_Error = "IDAplicacion no coincide con el valor esperado"
        GoTo Cleanup
    End If
    If CStr(Nz(rs.Fields("EnPruebas").value, "")) <> p_ExpectedEnPruebas Then
        p_Error = "EnPruebas no coincide con el valor esperado"
        GoTo Cleanup
    End If
    If CStr(Nz(rs.Fields("PasswordBackend").value, "")) <> p_ExpectedPasswordBackend Then
        p_Error = "PasswordBackend no coincide con el valor esperado"
        GoTo Cleanup
    End If

    AssertHelperRow = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

EH:
    p_Error = "AssertHelperRow: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

Public Function Test_BackendConfigHelper_CreaFilaDesdeTablaVacia() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultJson As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendConfigHelper_CreaFilaDesdeTablaVacia = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: tabla vacía controlada"
    If Not Test_ConfigFixtures.ClearConfigRows(errMsg) Then
        Test_BackendConfigHelper_CreaFilaDesdeTablaVacia = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ejecutar helper local/remoto"
    If ConfigurarBackendLanzadera_LocalConReposLocalYRemoto(errMsg) <> "OK" Then
        If errMsg = "" Then errMsg = "ConfigurarBackendLanzadera_LocalConReposLocalYRemoto devolvió un resultado distinto de OK"
        resultJson = Test_Helper.BuildJsonFail(errMsg, logs)
        GoTo Cleanup
    End If
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendConfigHelper", errMsg

    logs(3) = "4. Assert: helper crea una única fila legal con valores esperados"
    If Not AssertHelperRow(logs, FIXTURE_APP_ID, "No", "", errMsg) Then
        resultJson = Test_Helper.BuildJsonFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: helper create path completado"
    resultJson = Test_Helper.BuildJsonOk("backend_config_helper_create_ok", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendConfigHelper_CreaFilaDesdeTablaVacia = resultJson
    Exit Function

EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_BackendConfigHelper_ActualizaFilaExistente() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultJson As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendConfigHelper_ActualizaFilaExistente = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed de una fila válida con valores distintos"
    If Not Test_ConfigFixtures.SeedSingleConfigRow("PROD", "\\seed\backend\prod-original.accdb", "C:\seed\backend\local-original.accdb", FIXTURE_PASSWORD_BACKEND, FIXTURE_APP_ID_UPDATE, "\\seed\app\prod-original", "C:\seed\app\local-original", FIXTURE_EN_PRUEBAS, errMsg) Then
        Test_BackendConfigHelper_ActualizaFilaExistente = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ejecutar helper local/remoto"
    If ConfigurarBackendLanzadera_LocalConReposLocalYRemoto(errMsg) <> "OK" Then
        If errMsg = "" Then errMsg = "ConfigurarBackendLanzadera_LocalConReposLocalYRemoto devolvió un resultado distinto de OK"
        resultJson = Test_Helper.BuildJsonFail(errMsg, logs)
        GoTo Cleanup
    End If
    If errMsg <> "" Then Err.Raise 1000, "Test_BackendConfigHelper", errMsg

    logs(3) = "4. Assert: helper mantiene cardinalidad=1, actualiza target y preserva campos no tocados"
    If Not AssertHelperRow(logs, FIXTURE_APP_ID_UPDATE, FIXTURE_EN_PRUEBAS, FIXTURE_PASSWORD_BACKEND, errMsg) Then
        resultJson = Test_Helper.BuildJsonFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: helper update path completado"
    resultJson = Test_Helper.BuildJsonOk("backend_config_helper_update_ok", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendConfigHelper_ActualizaFilaExistente = resultJson
    Exit Function

EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_BackendConfigHelper_BloqueaDuplicados() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultValue As String
    Dim resultJson As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendConfigHelper_BloqueaDuplicados = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed de dos filas inválidas para el contrato"
    If Not Test_ConfigFixtures.SeedTwoConfigRows(errMsg) Then
        Test_BackendConfigHelper_BloqueaDuplicados = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ejecutar helper local/remoto"
    resultValue = ConfigurarBackendLanzadera_LocalConReposLocalYRemoto(errMsg)

    logs(3) = "4. Assert: helper falla con error funcional de duplicados"
    If errMsg = "" Then
        resultJson = Test_Helper.BuildJsonFail("Se esperaba error por filas duplicadas y no ocurrió. Resultado=" & resultValue, logs)
        GoTo Cleanup
    End If
    If InStr(1, errMsg, "más de un registro", vbTextCompare) = 0 Then
        resultJson = Test_Helper.BuildJsonFail("El error no indica bloqueo por duplicados: " & errMsg, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: helper duplicate-row path completado"
    resultJson = Test_Helper.BuildJsonOk("backend_config_helper_duplicate_blocked", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendConfigHelper_BloqueaDuplicados = resultJson
    Exit Function

EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_BackendConfigHelper_RechazaBackendActivoInvalido() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim snapshot As TbConfiguracionSnapshot
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultValue As String
    Dim resultJson As String

    logs(0) = "1. Arrange: snapshot del entorno para cleanup"
    If Not Test_ConfigFixtures.SnapshotSingleConfigRow(snapshot, errMsg) Then
        Test_BackendConfigHelper_RechazaBackendActivoInvalido = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: tabla vacía controlada"
    If Not Test_ConfigFixtures.ClearConfigRows(errMsg) Then
        Test_BackendConfigHelper_RechazaBackendActivoInvalido = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ejecutar helper genérico con BackendActivo inválido"
    resultValue = ConfigurarTbConfiguracionBackends( _
                    p_BackendActivo:="INVALIDO", _
                    p_BackendRemoto:=EXPECTED_BACKEND_REMOTO, _
                    p_BackendLocal:=EXPECTED_BACKEND_LOCAL, _
                    p_RutaTrabajoRemota:=EXPECTED_APP_REMOTA, _
                    p_RutaTrabajoLocal:=EXPECTED_APP_LOCAL, _
                    p_Error:=errMsg)

    logs(3) = "4. Assert: helper falla con error funcional por BackendActivo inválido"
    If errMsg = "" Then
        resultJson = Test_Helper.BuildJsonFail("Se esperaba error por BackendActivo inválido y no ocurrió. Resultado=" & resultValue, logs)
        GoTo Cleanup
    End If
    If InStr(1, errMsg, "BackendActivo inválido", vbTextCompare) = 0 Then
        resultJson = Test_Helper.BuildJsonFail("El error no indica BackendActivo inválido: " & errMsg, logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: helper invalid-backend path completado"
    resultJson = Test_Helper.BuildJsonOk("backend_config_helper_invalid_backend_blocked", logs)

Cleanup:
    If Not Test_ConfigFixtures.RestoreSingleConfigRow(snapshot, cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If

    Test_BackendConfigHelper_RechazaBackendActivoInvalido = resultJson
    Exit Function

EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function



