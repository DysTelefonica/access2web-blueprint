Attribute VB_Name = "Test_Helper_ExpedienteEliminacion"
Option Compare Database
Option Explicit
' Tests atómicos para Helper_ExpedienteEliminacion (PRUEBA-003 REFAC-1c).
' Cobertura: BR-26-01..03 (Verified-runtime), BR-26-04..05 (Verified-static, deferred).
'
' Convenciones:
'   - 0-arg Public Function
'   - JSON return: BuildJsonOk / BuildJsonFail
'   - Fixture-first: cada test que toca datos siembra su propio escenario
'   - IDs deterministicos con prefijo TEST-REFAC-1C- para evitar colisiones
'   - Teardown defensivo en orden inverso de FK (hijo antes que padre)

' BR-26-01: p_IDExpediente <= 0 -> ERR + p_Error poblada
Public Function Test_Helper_ExpedienteEliminacion_TieneDerivados_IDCero_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=0"
    result = Helper_ExpedienteEliminacion.TieneDerivados(0, Nothing, motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "ERR" Then
        Test_Helper_ExpedienteEliminacion_TieneDerivados_IDCero_PueblaError = BuildJsonFail("expected ERR, got: " & result, logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_ExpedienteEliminacion_TieneDerivados_IDCero_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: ERR + errMsg populated"
    Test_Helper_ExpedienteEliminacion_TieneDerivados_IDCero_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-26-01: ID no existente -> OK (no tiene derivados)
Public Function Test_Helper_ExpedienteEliminacion_TieneDerivados_IDInexistente_DevuelveOK() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=99999999 (no parent row in DB)"
    result = Helper_ExpedienteEliminacion.TieneDerivados(99999999, Nothing, motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "OK" Then
        Test_Helper_ExpedienteEliminacion_TieneDerivados_IDInexistente_DevuelveOK = BuildJsonFail("expected OK, got: " & result & " motivo: " & motivo, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteEliminacion_TieneDerivados_IDInexistente_DevuelveOK = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    logs(2) = "Assert: OK + no error"
    Test_Helper_ExpedienteEliminacion_TieneDerivados_IDInexistente_DevuelveOK = BuildJsonOk("ok", logs)
End Function

' BR-26-01: ID con hijo sembrado -> NO + motivo
' Siembra padre + hijo en TbExpedientes (FK self-reference), valida, hace teardown defensivo.
' IDs deterministicos en rango 99990001..99990002 (TEST-REFAC-1C prefix conceptual).
Public Function Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo = BuildJsonFail("test did not complete", logs)

    Dim db As DAO.Database
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    Const TEST_PADRE As Long = 99990001
    Const TEST_HIJO As Long = 99990002
    Dim cleanupError As String

    On Error GoTo HandleError
    Set db = getdb()

    ' Teardown preventivo (por si quedo algo de un test fallido previo)
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_HIJO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    On Error GoTo HandleError

    ' Seed: padre + hijo (hijo referencia padre por IDExpedientePadre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (" & TEST_PADRE & ", 'TEST-REFAC-1C-PADRE')", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, IDExpedientePadre) VALUES (" & TEST_HIJO & ", 'TEST-REFAC-1C-HIJO', " & TEST_PADRE & ")", dbFailOnError
    logs(0) = "Arrange: seed parent (" & TEST_PADRE & ") + child (" & TEST_HIJO & ") in TbExpedientes"

    ' Act
    result = Helper_ExpedienteEliminacion.TieneDerivados(TEST_PADRE, Nothing, motivo, errMsg)
    logs(1) = "Act: TieneDerivados called on parent"

    ' Assert
    If result <> "NO" Then
        Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo = BuildJsonFail("expected NO, got: " & result & " motivo: " & motivo, logs)
        GoTo HandleError
    End If
    If InStr(motivo, "Lotes o Basados") = 0 Then
        Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo = BuildJsonFail("expected motivo about Lotes/Basados, got: " & motivo, logs)
        GoTo HandleError
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo = BuildJsonFail("expected no error, got: " & errMsg, logs)
        GoTo HandleError
    End If
    logs(2) = "Assert: NO + motivo about Lotes/Basados"

    Test_Helper_ExpedienteEliminacion_TieneDerivados_ConHijoSembrado_DevuelveNOMotivo = BuildJsonOk("ok", logs)
    GoTo cleanup_success

HandleError:
    ' Teardown defensivo en orden inverso de FK (hijo antes que padre)
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_HIJO
    If Err.Number <> 0 Then cleanupError = "cleanup hijo: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    If Err.Number <> 0 Then cleanupError = cleanupError & " / cleanup padre: " & Err.Description
    If cleanupError <> "" Then logs(2) = logs(2) & " | Cleanup error: " & cleanupError
    Exit Function

cleanup_success:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_HIJO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
End Function

' BR-26-04: PuedeEliminar con ID no existente -> OK (pasa los 3 checks)
Public Function Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDInexistente_DevuelveOK() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=99999999 (no relationships in DB)"
    result = Helper_ExpedienteEliminacion.PuedeEliminar(99999999, Nothing, motivo, errMsg)
    logs(1) = "Act: PuedeEliminar called"
    If result <> "OK" Then
        Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDInexistente_DevuelveOK = BuildJsonFail("expected OK, got: " & result & " motivo: " & motivo, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDInexistente_DevuelveOK = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    logs(2) = "Assert: OK + no error"
    Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDInexistente_DevuelveOK = BuildJsonOk("ok", logs)
End Function

' BR-26-04: PuedeEliminar con ID <= 0 -> ERR (short-circuit en el primer check)
Public Function Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDCero_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=0"
    result = Helper_ExpedienteEliminacion.PuedeEliminar(0, Nothing, motivo, errMsg)
    logs(1) = "Act: PuedeEliminar called"
    If result <> "ERR" Then
        Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDCero_PueblaError = BuildJsonFail("expected ERR, got: " & result, logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDCero_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: ERR + errMsg"
    Test_Helper_ExpedienteEliminacion_PuedeEliminar_IDCero_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-26-05: EliminarExpediente con ID <= 0 -> ERR (input validation)
Public Function Test_Helper_ExpedienteEliminacion_EliminarExpediente_IDCero_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=0"
    result = Helper_ExpedienteEliminacion.EliminarExpediente(0, Nothing, errMsg)
    logs(1) = "Act: EliminarExpediente called"
    If result <> "ERR" Then
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_IDCero_PueblaError = BuildJsonFail("expected ERR, got: " & result, logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_IDCero_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: ERR + errMsg"
    Test_Helper_ExpedienteEliminacion_EliminarExpediente_IDCero_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-26-09: TieneAgedys con ID <= 0 -> ERR
Public Function Test_Helper_ExpedienteEliminacion_TieneAgedys_IDCero_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=0"
    result = Helper_ExpedienteEliminacion.TieneAgedys(0, Nothing, motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "ERR" Then
        Test_Helper_ExpedienteEliminacion_TieneAgedys_IDCero_PueblaError = BuildJsonFail("expected ERR, got: " & result, logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_ExpedienteEliminacion_TieneAgedys_IDCero_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: ERR + errMsg"
    Test_Helper_ExpedienteEliminacion_TieneAgedys_IDCero_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-26-09: TieneAgedys con ID no existente -> OK (no DPDs vinculados)
Public Function Test_Helper_ExpedienteEliminacion_TieneAgedys_IDInexistente_DevuelveOK() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=99999999 (no DPDs en TbExpAgedys)"
    result = Helper_ExpedienteEliminacion.TieneAgedys(99999999, Nothing, motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "OK" Then
        Test_Helper_ExpedienteEliminacion_TieneAgedys_IDInexistente_DevuelveOK = BuildJsonFail("expected OK, got: " & result & " motivo: " & motivo, logs)
        Exit Function
    End If
    logs(2) = "Assert: OK"
    Test_Helper_ExpedienteEliminacion_TieneAgedys_IDInexistente_DevuelveOK = BuildJsonOk("ok", logs)
End Function

' BR-26-02: TieneAnexos con ID <= 0 -> ERR
Public Function Test_Helper_ExpedienteEliminacion_TieneAnexos_IDCero_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=0"
    result = Helper_ExpedienteEliminacion.TieneAnexos(0, Nothing, motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "ERR" Then
        Test_Helper_ExpedienteEliminacion_TieneAnexos_IDCero_PueblaError = BuildJsonFail("expected ERR, got: " & result, logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_ExpedienteEliminacion_TieneAnexos_IDCero_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: ERR + errMsg"
    Test_Helper_ExpedienteEliminacion_TieneAnexos_IDCero_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-26-02: TieneAnexos con anexo sembrado -> NO + motivo
' Siembra parent + 1 anexo en TbExpedientesAnexos (PK IDDocumento, FK IDExpediente opcional,
' NombreDocumento opcional). Llama al helper, valida "NO" + motivo sobre anexos. Teardown defensivo.
Public Function Test_Helper_ExpedienteEliminacion_TieneAnexos_ConAnexoSembrado_DevuelveNOMotivo() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteEliminacion_TieneAnexos_ConAnexoSembrado_DevuelveNOMotivo = BuildJsonFail("test did not complete", logs)

    Dim db As DAO.Database
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    Const TEST_PADRE As Long = 99990200
    Const TEST_ANEXO As Long = 99990201
    Dim cleanupError As String

    On Error GoTo HandleError
    Set db = getdb()

    ' Teardown preventivo
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    On Error GoTo HandleError

    ' Seed parent + 1 anexo
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (" & TEST_PADRE & ", 'TEST-REFAC-1C-BR02')", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesAnexos (IDDocumento, IDExpediente, NombreDocumento) VALUES (" & TEST_ANEXO & ", " & TEST_PADRE & ", 'TEST-REFAC-1C-ANEXO')", dbFailOnError
    logs(0) = "Arrange: seed parent + 1 anexo"

    ' Act
    result = Helper_ExpedienteEliminacion.TieneAnexos(TEST_PADRE, Nothing, motivo, errMsg)
    logs(1) = "Act: TieneAnexos called on parent"

    ' Assert
    If result <> "NO" Then
        Test_Helper_ExpedienteEliminacion_TieneAnexos_ConAnexoSembrado_DevuelveNOMotivo = BuildJsonFail("expected NO, got: " & result & " motivo: " & motivo, logs)
        GoTo HandleError
    End If
    If InStr(motivo, "Anexos") = 0 Then
        Test_Helper_ExpedienteEliminacion_TieneAnexos_ConAnexoSembrado_DevuelveNOMotivo = BuildJsonFail("expected motivo about Anexos, got: " & motivo, logs)
        GoTo HandleError
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteEliminacion_TieneAnexos_ConAnexoSembrado_DevuelveNOMotivo = BuildJsonFail("expected no error, got: " & errMsg, logs)
        GoTo HandleError
    End If
    logs(2) = "Assert: NO + motivo about Anexos"

    Test_Helper_ExpedienteEliminacion_TieneAnexos_ConAnexoSembrado_DevuelveNOMotivo = BuildJsonOk("ok", logs)
    GoTo cleanup_success

HandleError:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    If Err.Number <> 0 Then cleanupError = "anexo: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    If Err.Number <> 0 Then cleanupError = cleanupError & " padre: " & Err.Description
    If cleanupError <> "" Then logs(2) = logs(2) & " | Cleanup error: " & cleanupError
    Exit Function

cleanup_success:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
End Function

' BR-26-03: TieneSuministradores con ID <= 0 -> ERR
Public Function Test_Helper_ExpedienteEliminacion_TieneSuministradores_IDCero_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_IDExpediente=0"
    result = Helper_ExpedienteEliminacion.TieneSuministradores(0, Nothing, motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "ERR" Then
        Test_Helper_ExpedienteEliminacion_TieneSuministradores_IDCero_PueblaError = BuildJsonFail("expected ERR, got: " & result, logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_ExpedienteEliminacion_TieneSuministradores_IDCero_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: ERR + errMsg"
    Test_Helper_ExpedienteEliminacion_TieneSuministradores_IDCero_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-26-03: TieneSuministradores con suministrador sembrado -> NO + motivo
' Siembra parent + 1 suministrador en TbExpedientesSuministradores (PK IDExpedienteSuministrador,
' FK IDExpediente NOT NULL, FK IDSuministrador NOT NULL, IDPadre nullable, Descripcon/Contratista/
' SubContratista opcionales). Llama al helper, valida "NO" + motivo sobre suministradores. Teardown.
Public Function Test_Helper_ExpedienteEliminacion_TieneSuministradores_ConSuministradorSembrado_DevuelveNOMotivo() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteEliminacion_TieneSuministradores_ConSuministradorSembrado_DevuelveNOMotivo = BuildJsonFail("test did not complete", logs)

    Dim db As DAO.Database
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    Const TEST_PADRE As Long = 99990300
    Const TEST_SUMIN As Long = 99990301
    Const TEST_PROVEEDOR As Long = 1   ' IDSuministrador real (FK enforced a TbSuministradores)
    Dim cleanupError As String

    On Error GoTo HandleError
    Set db = getdb()

    ' Teardown preventivo
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador=" & TEST_SUMIN
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    On Error GoTo HandleError

    ' Seed parent + 1 suministrador (IDSuministrador=1 para no violar FK a TbSuministradores)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (" & TEST_PADRE & ", 'TEST-REFAC-1C-BR03')", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesSuministradores (IDExpedienteSuministrador, IDExpediente, IDSuministrador) VALUES (" & TEST_SUMIN & ", " & TEST_PADRE & ", " & TEST_PROVEEDOR & ")", dbFailOnError
    logs(0) = "Arrange: seed parent + 1 suministrador"

    ' Act
    result = Helper_ExpedienteEliminacion.TieneSuministradores(TEST_PADRE, Nothing, motivo, errMsg)
    logs(1) = "Act: TieneSuministradores called on parent"

    ' Assert
    If result <> "NO" Then
        Test_Helper_ExpedienteEliminacion_TieneSuministradores_ConSuministradorSembrado_DevuelveNOMotivo = BuildJsonFail("expected NO, got: " & result & " motivo: " & motivo, logs)
        GoTo HandleError
    End If
    If InStr(motivo, "Suministradores") = 0 Then
        Test_Helper_ExpedienteEliminacion_TieneSuministradores_ConSuministradorSembrado_DevuelveNOMotivo = BuildJsonFail("expected motivo about Suministradores, got: " & motivo, logs)
        GoTo HandleError
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteEliminacion_TieneSuministradores_ConSuministradorSembrado_DevuelveNOMotivo = BuildJsonFail("expected no error, got: " & errMsg, logs)
        GoTo HandleError
    End If
    logs(2) = "Assert: NO + motivo about Suministradores"

    Test_Helper_ExpedienteEliminacion_TieneSuministradores_ConSuministradorSembrado_DevuelveNOMotivo = BuildJsonOk("ok", logs)
    GoTo cleanup_success

HandleError:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador=" & TEST_SUMIN
    If Err.Number <> 0 Then cleanupError = "sumin: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    If Err.Number <> 0 Then cleanupError = cleanupError & " padre: " & Err.Description
    If cleanupError <> "" Then logs(2) = logs(2) & " | Cleanup error: " & cleanupError
    Exit Function

cleanup_success:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador=" & TEST_SUMIN
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
End Function

' BR-26-05: EliminarExpediente con cascade completo.
' Siembra parent + filas en 4 tablas hijas (TbExpedientesHitos, TbExpedientesResponsables,
' TbExpedientesModificados, TbExpedientesConEntidades), llama al helper, valida "OK" y
' verifica que las filas hijas estan todas borradas. Teardown defensivo en orden inverso de FK.
' IDs deterministicos con prefijo TEST-REFAC-1C-CASCADE-.
Public Function Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("test did not complete", logs)

    Dim db As DAO.Database
    Dim errMsg As String
    Dim result As String
    Dim rs As DAO.Recordset
    Const TEST_PADRE As Long = 99990100
    Const TEST_HITO As Long = 99990101
    Const TEST_RESPONSABLE As Long = 99990102
    Const TEST_USUARIO As Long = 99990103   ' requerido por TbExpedientesResponsables
    Const TEST_MODIFICADO As Long = 99990104
    Dim cleanupError As String

    On Error GoTo HandleError
    Set db = getdb()

    ' Teardown preventivo (por si quedo algo de un test fallido previo)
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesHitos WHERE IDHitoExpediente=" & TEST_HITO
    db.Execute "DELETE FROM TbExpedientesResponsables WHERE IDExpedienteResponsable=" & TEST_RESPONSABLE
    db.Execute "DELETE FROM TbExpedientesModificados WHERE IDExpedienteModificado=" & TEST_MODIFICADO
    db.Execute "DELETE FROM TbExpedientesConEntidades WHERE IDExpediente=" & TEST_PADRE
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    On Error GoTo HandleError

    ' Seed parent + 4 child rows (PK + required fields only)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (" & TEST_PADRE & ", 'TEST-REFAC-1C-CASCADE')", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesConEntidades (IDExpediente) VALUES (" & TEST_PADRE & ")", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesHitos (IDHitoExpediente, IDExpediente) VALUES (" & TEST_HITO & ", " & TEST_PADRE & ")", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesResponsables (IDExpedienteResponsable, IdExpediente, IdUsuario) VALUES (" & TEST_RESPONSABLE & ", " & TEST_PADRE & ", " & TEST_USUARIO & ")", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesModificados (IDExpedienteModificado, IDExpediente) VALUES (" & TEST_MODIFICADO & ", " & TEST_PADRE & ")", dbFailOnError
    logs(0) = "Arrange: seed parent + 4 child rows"

    ' Act
    result = Helper_ExpedienteEliminacion.EliminarExpediente(TEST_PADRE, Nothing, errMsg)
    logs(1) = "Act: EliminarExpediente called"

    ' Assert: result is OK
    If result <> "OK" Then
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("expected OK, got: " & result & " err: " & errMsg, logs)
        GoTo HandleError
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("expected no error, got: " & errMsg, logs)
        GoTo HandleError
    End If

    ' Assert: parent is gone
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS N FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE, dbReadOnly)
    If rs!n <> 0 Then
        rs.Close
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("expected parent gone, found " & rs!n, logs)
        GoTo HandleError
    End If
    rs.Close
    logs(2) = "Assert: parent row deleted"

    ' Assert: 4 child tables are empty for this parent
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS N FROM TbExpedientesConEntidades WHERE IDExpediente=" & TEST_PADRE, dbReadOnly)
    If rs!n <> 0 Then
        rs.Close
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("expected ConEntidades empty, found " & rs!n, logs)
        GoTo HandleError
    End If
    rs.Close

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS N FROM TbExpedientesHitos WHERE IDExpediente=" & TEST_PADRE, dbReadOnly)
    If rs!n <> 0 Then
        rs.Close
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("expected Hitos empty, found " & rs!n, logs)
        GoTo HandleError
    End If
    rs.Close

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS N FROM TbExpedientesResponsables WHERE IDExpediente=" & TEST_PADRE, dbReadOnly)
    If rs!n <> 0 Then
        rs.Close
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("expected Responsables empty, found " & rs!n, logs)
        GoTo HandleError
    End If
    rs.Close

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS N FROM TbExpedientesModificados WHERE IDExpediente=" & TEST_PADRE, dbReadOnly)
    If rs!n <> 0 Then
        rs.Close
        Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonFail("expected Modificados empty, found " & rs!n, logs)
        GoTo HandleError
    End If
    rs.Close
    logs(3) = "Assert: 4 child tables cascade deleted"

    Test_Helper_ExpedienteEliminacion_EliminarExpediente_CascadaExitosa = BuildJsonOk("ok", logs)
    GoTo cleanup_success

HandleError:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesHitos WHERE IDHitoExpediente=" & TEST_HITO
    If Err.Number <> 0 Then cleanupError = "hito: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientesResponsables WHERE IDExpedienteResponsable=" & TEST_RESPONSABLE
    If Err.Number <> 0 Then cleanupError = cleanupError & " resp: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientesModificados WHERE IDExpedienteModificado=" & TEST_MODIFICADO
    If Err.Number <> 0 Then cleanupError = cleanupError & " mod: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientesConEntidades WHERE IDExpediente=" & TEST_PADRE
    If Err.Number <> 0 Then cleanupError = cleanupError & " conEnt: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    If Err.Number <> 0 Then cleanupError = cleanupError & " padre: " & Err.Description
    If cleanupError <> "" Then logs(2) = logs(2) & " | Cleanup error: " & cleanupError
    Exit Function

cleanup_success:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesHitos WHERE IDHitoExpediente=" & TEST_HITO
    db.Execute "DELETE FROM TbExpedientesResponsables WHERE IDExpedienteResponsable=" & TEST_RESPONSABLE
    db.Execute "DELETE FROM TbExpedientesModificados WHERE IDExpedienteModificado=" & TEST_MODIFICADO
    db.Execute "DELETE FROM TbExpedientesConEntidades WHERE IDExpediente=" & TEST_PADRE
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
End Function

' BR-26-04: PuedeEliminar retorna "NO" con el motivo del primer check que falle
' cuando hay multiples relaciones. El wrapper corre los checks en orden:
' derivados -> anexos -> suministradores -> agedys. El primero que falla
' es el que se reporta.
Public Function Test_Helper_ExpedienteEliminacion_PuedeEliminar_ConMultiplesRelaciones_DevuelveNOMotivoPrioritario() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteEliminacion_PuedeEliminar_ConMultiplesRelaciones_DevuelveNOMotivoPrioritario = BuildJsonFail("test did not complete", logs)

    Dim db As DAO.Database
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    Const TEST_PADRE As Long = 99990500
    Const TEST_HIJO As Long = 99990501
    Const TEST_ANEXO As Long = 99990502
    Dim cleanupError As String

    On Error GoTo HandleError
    Set db = getdb()

    ' Teardown preventivo
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_HIJO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    On Error GoTo HandleError

    ' Seed: padre + hijo (derivados) + anexo
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (" & TEST_PADRE & ", 'TEST-REFAC-1C-MULTI-PADRE')", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, IDExpedientePadre) VALUES (" & TEST_HIJO & ", 'TEST-REFAC-1C-MULTI-HIJO', " & TEST_PADRE & ")", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesAnexos (IDDocumento, IDExpediente, NombreDocumento) VALUES (" & TEST_ANEXO & ", " & TEST_PADRE & ", 'TEST-REFAC-1C-MULTI-ANEXO')", dbFailOnError
    logs(0) = "Arrange: seed parent + 1 hijo (derivados) + 1 anexo"

    ' Act: PuedeEliminar deberia retornar NO con motivo de derivados (primer check que falla)
    result = Helper_ExpedienteEliminacion.PuedeEliminar(TEST_PADRE, Nothing, motivo, errMsg)
    logs(1) = "Act: PuedeEliminar called"

    ' Assert: NO + motivo de derivados
    If result <> "NO" Then
        Test_Helper_ExpedienteEliminacion_PuedeEliminar_ConMultiplesRelaciones_DevuelveNOMotivoPrioritario = BuildJsonFail("expected NO, got: " & result, logs)
        GoTo HandleError
    End If
    If InStr(motivo, "Lotes o Basados") = 0 Then
        Test_Helper_ExpedienteEliminacion_PuedeEliminar_ConMultiplesRelaciones_DevuelveNOMotivoPrioritario = BuildJsonFail("expected motivo about Lotes/Basados (primer check), got: " & motivo, logs)
        GoTo HandleError
    End If
    logs(2) = "Assert: NO + motivo de derivados (primer check que falla)"

    Test_Helper_ExpedienteEliminacion_PuedeEliminar_ConMultiplesRelaciones_DevuelveNOMotivoPrioritario = BuildJsonOk("ok", logs)
    GoTo cleanup_success

HandleError:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    If Err.Number <> 0 Then cleanupError = "anexo: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_HIJO
    If Err.Number <> 0 Then cleanupError = cleanupError & " hijo: " & Err.Description
    Err.Clear
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    If Err.Number <> 0 Then cleanupError = cleanupError & " padre: " & Err.Description
    If cleanupError <> "" Then logs(2) = logs(2) & " | Cleanup error: " & cleanupError
    Exit Function

cleanup_success:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_HIJO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
End Function
