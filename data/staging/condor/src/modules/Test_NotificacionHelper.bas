Attribute VB_Name = "Test_NotificacionHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' Test_NotificacionHelper - Slice 3.3 BR-001 non-blocking notification seam
'
' Each atom is a Public Function (global, unique) that returns the
' canonical JSON via TestHelper.BuildJsonOk / BuildJsonFail. Pure
' helper tests: NO DAO, NO controls, NO MsgBox, no fixtures required.
'
' The MockNotifServ class (src/classes/MockNotifServ.cls) provides the
' five-argument EnviarNotificacion(vm, asunto, destinatarioEmail,
' copiaEmail, [nombreEstadoNotif]) stub that mirrors the real
' NotificacionServicio signature, without requiring the real backend.
' ==========================================================================

Public Function Test_NotificacionHelper_EnviarNotificacionExitosa_OK() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim stub As New MockNotifServ
    Dim status As String

    logs(0) = "1. Arrange: MockNotifServ sin error"

    stub.ResetMock
    status = NotificacionHelper_NotificarSiPosible(stub, "", "X", "a@b", "")
    logs(1) = "2. Act: status='" & status & "'"

    If status <> "OK" Then _
        Err.Raise 513, , "Happy path debe devolver 'OK', obtuvo '" & status & "'"
    If Not stub.WasCalled Then _
        Err.Raise 513, , "El helper debio invocar EnviarNotificacion"
    If stub.Asunto <> "X" Or stub.Destinatario <> "a@b" Then _
        Err.Raise 513, , "Argumentos no propagados: asunto='" & stub.Asunto & "' destinatario='" & stub.Destinatario & "'"
    logs(2) = "3. PASS: status=OK + stub fue llamado"
    logs(3) = "4. PASS: argumentos propagados correctamente"

    Test_NotificacionHelper_EnviarNotificacionExitosa_OK = TestHelper.BuildJsonOk("OK", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_NotificacionHelper_EnviarNotificacionExitosa_OK = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_NotificacionHelper_EnviarNotificacionFailureSwallowed() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim stub As New MockNotifServ
    Dim status As String

    logs(0) = "1. Arrange: MockNotifServ configurado para RaiseError=True (error 424)"

    stub.ResetMock
    stub.RaiseError = True

    status = NotificacionHelper_NotificarSiPosible(stub, "", "APROBADO_RAC", "a@b", "")
    logs(1) = "2. Act: status='" & status & "' (debe iniciar con 'NOTIF_ERROR_')"

    If Len(status) = 0 Then Err.Raise 513, , "El helper debio devolver un status"
    If Left$(status, 12) <> "NOTIF_ERROR_" Then _
        Err.Raise 513, , "Status debe empezar con 'NOTIF_ERROR_', obtuvo '" & status & "'"
    If InStr(status, "424") = 0 Then _
        Err.Raise 513, , "Status debe incluir el codigo 424, obtuvo '" & status & "'"
    If Not stub.WasCalled Then _
        Err.Raise 513, , "El helper debio invocar EnviarNotificacion aun cuando fallo"
    logs(2) = "3. PASS: status NOTIF_ERROR_424_*"

    ' Confirmar que el helper NO propago el error al caller.
    ' Si hubiera propagado, la proxima linea elevaria 424.
    Dim sentinel As String
    sentinel = "still here"
    logs(3) = "4. PASS: caller sigue en control (sentinel='" & sentinel & "')"
    logs(4) = "5. PASS: BR-001 cumplido (silencio + motivo capturable)"

    Test_NotificacionHelper_EnviarNotificacionFailureSwallowed = TestHelper.BuildJsonOk("swallowed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_NotificacionHelper_EnviarNotificacionFailureSwallowed = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_NotificacionHelper_NullNotifServ_NoFalla() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)
    Dim status As String

    logs(0) = "1. Arrange: p_NotifServ=Nothing (no hay backend)"

    status = NotificacionHelper_NotificarSiPosible(Nothing, "", "X", "a@b", "")
    logs(1) = "2. Act: status='" & status & "'"

    ' Politica BR-001: invocar un metodo sobre Nothing dispara un error de
    ' runtime (VBA reporta 91; algunos hosts 438). El helper debe TRAGAR ese
    ' error y devolver un marcador NOTIF_ERROR_*, nunca explotar. Aseveramos
    ' el prefijo (robusto ante 91 vs 438), no el numero exacto.
    If Left$(status, 12) <> "NOTIF_ERROR_" Then _
        Err.Raise 513, , "Nothing service debe devolver 'NOTIF_ERROR_*' (error tragado), obtuvo '" & status & "'"
    logs(2) = "3. PASS: Nothing -> 'NOTIF_ERROR_*' sin propagar error (no crash)"

    Test_NotificacionHelper_NullNotifServ_NoFalla = TestHelper.BuildJsonOk(status, logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_NotificacionHelper_NullNotifServ_NoFalla = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_NotificacionHelper_TipoNotificacionVacio_OK() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)
    Dim stub As New MockNotifServ
    Dim status As String

    logs(0) = "1. Arrange: asunto='' y destinatario='' (inputs vacios)"

    stub.ResetMock
    status = NotificacionHelper_NotificarSiPosible(stub, "", "", "", "")
    logs(1) = "2. Act: status='" & status & "'"

    ' El helper debe aceptar inputs vacios y devolver OK si el stub no falla.
    If status <> "OK" Then _
        Err.Raise 513, , "Inputs vacios no deben hacer fallar al helper, obtuvo '" & status & "'"
    If Not stub.WasCalled Then _
        Err.Raise 513, , "El helper debio invocar el stub con inputs vacios"
    logs(2) = "3. PASS: inputs vacios -> OK (helper no valida semantica del contenido)"

    Test_NotificacionHelper_TipoNotificacionVacio_OK = TestHelper.BuildJsonOk("OK", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_NotificacionHelper_TipoNotificacionVacio_OK = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_NotificacionHelper_FirmaCincoArgsReenviaAlServicio() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim stub As New MockNotifServ
    Dim status As String

    logs(0) = "1. Arrange: MockNotifServ captura asunto/destinatario/copia"

    stub.ResetMock
    status = NotificacionHelper_NotificarSiPosible( _
        stub, "", "AsuntoTransicion", "qa@telefonica.com", "dev@telefonica.com", "VALIDADO_POR_CALIDAD")
    logs(1) = "2. Act: status='" & status & "'"

    If status <> "OK" Then _
        Err.Raise 513, , "Firma de 5 args debe devolver 'OK', obtuvo '" & status & "'"
    If stub.Asunto <> "AsuntoTransicion" Then _
        Err.Raise 513, , "asunto no reenviado: obtuvo '" & stub.Asunto & "'"
    If stub.Destinatario <> "qa@telefonica.com" Then _
        Err.Raise 513, , "destinatario no reenviado: obtuvo '" & stub.Destinatario & "'"
    If stub.Copia <> "dev@telefonica.com" Then _
        Err.Raise 513, , "copia no reenviada: obtuvo '" & stub.Copia & "'"
    logs(2) = "3. PASS: asunto/destinatario/copia reenviados al servicio"
    logs(3) = "4. PASS: firma de 5 argumentos alineada con NotificacionServicio"

    Test_NotificacionHelper_FirmaCincoArgsReenviaAlServicio = TestHelper.BuildJsonOk("OK", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_NotificacionHelper_FirmaCincoArgsReenviaAlServicio = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
