Attribute VB_Name = "Test_CorreoPruebas"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_CorreoPruebas_RedirigeDestinatariosAUsuarioConectado() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: correo en modo pruebas con destinatario real"
    logs(1) = "2. Act: AplicarRedireccionCorreoPruebas"
    logs(2) = "3. Assert: destinatario final es el usuario conectado"

    Dim correo As Correo
    Set correo = New Correo

    Dim vRedir() As Variant
    vRedir = correo.AplicarRedireccionCorreoPruebas("cliente@example.com", "", "", "tester@example.com", True)

    If CStr(vRedir(0)) <> "tester@example.com" Then
        Test_CorreoPruebas_RedirigeDestinatariosAUsuarioConectado = BuildFail("Debe redirigir Destinatarios al correo del usuario conectado", logs)
        Exit Function
    End If

    Test_CorreoPruebas_RedirigeDestinatariosAUsuarioConectado = BuildOk("destinatario_redirigido", logs)
    Exit Function
EH:
    Test_CorreoPruebas_RedirigeDestinatariosAUsuarioConectado = BuildFail("Test_CorreoPruebas_RedirigeDestinatariosAUsuarioConectado: " & Err.Description, logs)
End Function

Public Function Test_CorreoPruebas_LimpiaCopiasEnPruebas() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: correo en modo pruebas con CC y BCC reales"
    logs(1) = "2. Act: AplicarRedireccionCorreoPruebas"
    logs(2) = "3. Assert: CC y BCC quedan vacios"

    Dim correo As Correo
    Set correo = New Correo

    Dim vRedir() As Variant
    vRedir = correo.AplicarRedireccionCorreoPruebas("cliente@example.com", "cc@example.com", "bcc@example.com", "tester@example.com", True)

    If CStr(vRedir(1)) <> "" Then
        Test_CorreoPruebas_LimpiaCopiasEnPruebas = BuildFail("DestinatariosConCopia debe quedar vacio en pruebas", logs)
        Exit Function
    End If

    If CStr(vRedir(2)) <> "" Then
        Test_CorreoPruebas_LimpiaCopiasEnPruebas = BuildFail("DestinatariosConCopiaOculta debe quedar vacio en pruebas", logs)
        Exit Function
    End If

    Test_CorreoPruebas_LimpiaCopiasEnPruebas = BuildOk("copias_limpiadas", logs)
    Exit Function
EH:
    Test_CorreoPruebas_LimpiaCopiasEnPruebas = BuildFail("Test_CorreoPruebas_LimpiaCopiasEnPruebas: " & Err.Description, logs)
End Function

Public Function Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: correo en pruebas con To, CC y BCC originales"
    logs(1) = "2. Act: AplicarRedireccionCorreoPruebas"
    logs(2) = "3. Assert: banner visible lista originales escapados"

    Dim correo As Correo
    Set correo = New Correo

    Dim vRedir() As Variant
    vRedir = correo.AplicarRedireccionCorreoPruebas("real<to>&@example.com", "cc@example.com", "bcc@example.com", "tester@example.com", True)

    Dim banner As String
    banner = CStr(vRedir(3))

    If Left$(banner, 4) <> "<div" Then
        Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales = BuildFail("El banner debe ser visible y estar listo para ir al inicio del HTML", logs)
        Exit Function
    End If

    If InStr(1, banner, "AVISO DE PRUEBAS", vbTextCompare) = 0 Then
        Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales = BuildFail("El banner debe identificar que es un correo de pruebas", logs)
        Exit Function
    End If

    If InStr(1, banner, "real&lt;to&gt;&amp;@example.com", vbBinaryCompare) = 0 Then
        Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales = BuildFail("El banner debe incluir Destinatarios originales con HTML seguro", logs)
        Exit Function
    End If

    If InStr(1, banner, "Con copia:</strong> cc@example.com", vbTextCompare) = 0 Then
        Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales = BuildFail("El banner debe incluir DestinatariosConCopia originales", logs)
        Exit Function
    End If

    If InStr(1, banner, "Con copia oculta:</strong> bcc@example.com", vbTextCompare) = 0 Then
        Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales = BuildFail("El banner debe incluir DestinatariosConCopiaOculta originales", logs)
        Exit Function
    End If

    Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales = BuildOk("banner_originales", logs)
    Exit Function
EH:
    Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales = BuildFail("Test_CorreoPruebas_BannerIncluyeDestinatariosYCopiasOriginales: " & Err.Description, logs)
End Function

Public Function Test_CorreoPruebas_ProduccionNoMutaCorreo() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: correo fuera de modo pruebas"
    logs(1) = "2. Act: AplicarRedireccionCorreoPruebas"
    logs(2) = "3. Assert: destinatarios, copias y cuerpo no cambian"

    Dim correo As Correo
    Set correo = New Correo

    Dim cuerpoOriginal As String
    cuerpoOriginal = "<p>Cuerpo real</p>"

    Dim vRedir() As Variant
    vRedir = correo.AplicarRedireccionCorreoPruebas("cliente@example.com", "cc@example.com", "bcc@example.com", "tester@example.com", False)

    Dim cuerpoFinal As String
    cuerpoFinal = cuerpoOriginal
    If CStr(vRedir(3)) <> "" Then
        cuerpoFinal = CStr(vRedir(3)) & vbCrLf & cuerpoFinal
    End If

    If CStr(vRedir(0)) <> "cliente@example.com" Then
        Test_CorreoPruebas_ProduccionNoMutaCorreo = BuildFail("Destinatarios no debe cambiar fuera de pruebas", logs)
        Exit Function
    End If

    If CStr(vRedir(1)) <> "cc@example.com" Or CStr(vRedir(2)) <> "bcc@example.com" Then
        Test_CorreoPruebas_ProduccionNoMutaCorreo = BuildFail("Copias no deben cambiar fuera de pruebas", logs)
        Exit Function
    End If

    If cuerpoFinal <> cuerpoOriginal Then
        Test_CorreoPruebas_ProduccionNoMutaCorreo = BuildFail("Cuerpo no debe cambiar fuera de pruebas", logs)
        Exit Function
    End If

    Test_CorreoPruebas_ProduccionNoMutaCorreo = BuildOk("produccion_sin_mutacion", logs)
    Exit Function
EH:
    Test_CorreoPruebas_ProduccionNoMutaCorreo = BuildFail("Test_CorreoPruebas_ProduccionNoMutaCorreo: " & Err.Description, logs)
End Function
