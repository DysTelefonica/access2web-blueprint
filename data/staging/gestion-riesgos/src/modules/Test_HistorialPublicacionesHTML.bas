Attribute VB_Name = "Test_HistorialPublicacionesHTML"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function NormalizeSqlForAssert(ByVal sql As String) As String
    sql = Replace(sql, vbCrLf, " ")
    sql = Replace(sql, vbCr, " ")
    sql = Replace(sql, vbLf, " ")
    sql = Replace(sql, vbTab, " ")
    Do While InStr(1, sql, "  ", vbBinaryCompare) > 0
        sql = Replace(sql, "  ", " ")
    Loop
    NormalizeSqlForAssert = Trim$(sql)
End Function

Public Function Test_HistorialPublicacionesHTML_DocumentoTelefonicaUTF8() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: cuerpo de historial con texto acentuado"
    logs(1) = "2. Act: HTMLDocumentoHistorialTelefonica"
    logs(2) = "3. Assert: declara UTF-8"
    logs(3) = "4. Assert: incluye tokens visuales Telefonica"

    Dim html As String
    html = HTMLDocumentoHistorialTelefonica("Historial de publicaciones", "Proyecto Ñ · Edición 1", "<p>Publicación aprobada</p>")

    If InStr(1, html, "<meta charset='UTF-8'>", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_DocumentoTelefonicaUTF8 = BuildFail("El HTML debe declarar charset UTF-8", logs)
        Exit Function
    End If

    If InStr(1, html, "#0066ff", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_DocumentoTelefonicaUTF8 = BuildFail("El HTML debe incluir el token azul Telefonica", logs)
        Exit Function
    End If

    If InStr(1, html, "tf-card", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_DocumentoTelefonicaUTF8 = BuildFail("El HTML debe usar estructura visual de tarjeta", logs)
        Exit Function
    End If

    Test_HistorialPublicacionesHTML_DocumentoTelefonicaUTF8 = BuildOk("historial_telefonica_utf8", logs)
    Exit Function
EH:
    Test_HistorialPublicacionesHTML_DocumentoTelefonicaUTF8 = BuildFail("Test_HistorialPublicacionesHTML_DocumentoTelefonicaUTF8: " & Err.Description, logs)
End Function

Public Function Test_HistorialPublicacionesHTML_DocumentoConservaCuerpo() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: cuerpo con tabla de historial"
    logs(1) = "2. Act: HTMLDocumentoHistorialTelefonica"
    logs(2) = "3. Assert: el cuerpo no queda reducido al titulo"

    Dim body As String
    Dim html As String
    body = "<div class='tf-table-wrap'><table><tr><th>Fecha</th></tr><tr><td>19/05/2026</td></tr></table></div>"
    html = HTMLDocumentoHistorialTelefonica("Historial", "Subtitulo", body)

    If InStr(1, html, "19/05/2026", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_DocumentoConservaCuerpo = BuildFail("El HTML debe conservar los datos del historial", logs)
        Exit Function
    End If

    If InStr(1, html, "<table", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_DocumentoConservaCuerpo = BuildFail("El HTML debe conservar la tabla del historial", logs)
        Exit Function
    End If

    Test_HistorialPublicacionesHTML_DocumentoConservaCuerpo = BuildOk("historial_con_cuerpo", logs)
    Exit Function
EH:
    Test_HistorialPublicacionesHTML_DocumentoConservaCuerpo = BuildFail("Test_HistorialPublicacionesHTML_DocumentoConservaCuerpo: " & Err.Description, logs)
End Function

Public Function Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: IDProyecto conocido"
    logs(1) = "2. Act: getSQLLogPublicacionesProyecto"
    logs(2) = "3. Assert: une publicaciones con ediciones"
    logs(3) = "4. Assert: filtra por proyecto, no por edicion activa"

    Dim errMsg As String
    Dim sql As String
    sql = NormalizeSqlForAssert(Constructor.getSQLLogPublicacionesProyecto("123", errMsg))
    If errMsg <> "" Then
        Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto = BuildFail(errMsg, logs)
        Exit Function
    End If

    If InStr(1, sql, "FROM TbLogPublicaciones INNER JOIN TbProyectosEdiciones", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto = BuildFail("El historial debe relacionar publicaciones con ediciones del proyecto", logs)
        Exit Function
    End If

    If InStr(1, sql, "TbProyectosEdiciones.IDProyecto=123", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto = BuildFail("El historial debe filtrar por IDProyecto", logs)
        Exit Function
    End If

    If InStr(1, sql, "WHERE IDEdicion=", vbTextCompare) > 0 Then
        Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto = BuildFail("El historial de publicaciones no debe limitarse a la edicion activa", logs)
        Exit Function
    End If

    Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto = BuildOk("historial_publicaciones_proyecto", logs)
    Exit Function
EH:
    Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto = BuildFail("Test_HistorialPublicacionesHTML_SQLHistorialProyectoFiltraPorProyecto: " & Err.Description, logs)
End Function

Public Function Test_HistorialPublicacionesHTML_SQLHistorialProyectoOrdenPredecible() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: IDProyecto conocido"
    logs(1) = "2. Act: getSQLLogPublicacionesProyecto"
    logs(2) = "3. Assert: ordena por FechaRegistro e ID"

    Dim errMsg As String
    Dim sql As String
    sql = NormalizeSqlForAssert(Constructor.getSQLLogPublicacionesProyecto("123", errMsg))
    If errMsg <> "" Then
        Test_HistorialPublicacionesHTML_SQLHistorialProyectoOrdenPredecible = BuildFail(errMsg, logs)
        Exit Function
    End If

    If InStr(1, sql, "ORDER BY TbLogPublicaciones.FechaRegistro, TbLogPublicaciones.ID;", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLHistorialProyectoOrdenPredecible = BuildFail("El historial debe tener orden predecible por FechaRegistro e ID", logs)
        Exit Function
    End If

    Test_HistorialPublicacionesHTML_SQLHistorialProyectoOrdenPredecible = BuildOk("historial_publicaciones_orden_predecible", logs)
    Exit Function
EH:
    Test_HistorialPublicacionesHTML_SQLHistorialProyectoOrdenPredecible = BuildFail("Test_HistorialPublicacionesHTML_SQLHistorialProyectoOrdenPredecible: " & Err.Description, logs)
End Function

Public Function Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: IDProyecto conocido"
    logs(1) = "2. Act: getSQLLogCorreosResponsablesProyecto"
    logs(2) = "3. Assert: une correos responsables con ediciones"
    logs(3) = "4. Assert: filtra por proyecto, no por edicion activa"

    Dim errMsg As String
    Dim sql As String
    sql = NormalizeSqlForAssert(Constructor.getSQLLogCorreosResponsablesProyecto("123", errMsg))
    If errMsg <> "" Then
        Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto = BuildFail(errMsg, logs)
        Exit Function
    End If

    If InStr(1, sql, "FROM TbProyectoEdicionesCorreoRevision INNER JOIN TbProyectosEdiciones", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto = BuildFail("El historial debe relacionar correos responsables con ediciones del proyecto", logs)
        Exit Function
    End If

    If InStr(1, sql, "TbProyectosEdiciones.IDProyecto=123", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto = BuildFail("El historial de correos responsables debe filtrar por IDProyecto", logs)
        Exit Function
    End If

    If InStr(1, sql, "WHERE IDEdicion=", vbTextCompare) > 0 Then
        Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto = BuildFail("El historial de correos responsables no debe limitarse a la edicion activa", logs)
        Exit Function
    End If

    Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto = BuildOk("historial_correos_responsables_proyecto", logs)
    Exit Function
EH:
    Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto = BuildFail("Test_HistorialPublicacionesHTML_SQLCorreosResponsablesFiltraPorProyecto: " & Err.Description, logs)
End Function

Public Function Test_HistorialPublicacionesHTML_SQLCorreosResponsablesOrdenPredecible() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: IDProyecto conocido"
    logs(1) = "2. Act: getSQLLogCorreosResponsablesProyecto"
    logs(2) = "3. Assert: ordena por FechaCorreoRevision, IDEdicion e IDEnvioCorreoTecnico"

    Dim errMsg As String
    Dim sql As String
    sql = NormalizeSqlForAssert(Constructor.getSQLLogCorreosResponsablesProyecto("123", errMsg))
    If errMsg <> "" Then
        Test_HistorialPublicacionesHTML_SQLCorreosResponsablesOrdenPredecible = BuildFail(errMsg, logs)
        Exit Function
    End If

    If InStr(1, sql, "ORDER BY TbProyectoEdicionesCorreoRevision.FechaCorreoRevision, TbProyectoEdicionesCorreoRevision.IDEdicion, TbProyectoEdicionesCorreoRevision.IDEnvioCorreoTecnico;", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLCorreosResponsablesOrdenPredecible = BuildFail("El historial de correos responsables debe tener orden predecible", logs)
        Exit Function
    End If

    Test_HistorialPublicacionesHTML_SQLCorreosResponsablesOrdenPredecible = BuildOk("historial_correos_responsables_orden_predecible", logs)
    Exit Function
EH:
    Test_HistorialPublicacionesHTML_SQLCorreosResponsablesOrdenPredecible = BuildFail("Test_HistorialPublicacionesHTML_SQLCorreosResponsablesOrdenPredecible: " & Err.Description, logs)
End Function

Public Function Test_HistorialPublicacionesHTML_SQLRegistroCorreoEnviadoContrato() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: datos de correo ya enviado"
    logs(1) = "2. Act: getSQLRegistrarCorreoEnviado"
    logs(2) = "3. Assert: inserta solo el log, sin invocar envio"

    Dim errMsg As String
    Dim sql As String
    Dim logCorreo As EdicionCorreoRevision
    Set logCorreo = New EdicionCorreoRevision
    sql = NormalizeSqlForAssert(logCorreo.getSQLRegistrarCorreoEnviado("901", "902", "903", #5/19/2026 10:11:12 AM#, "calidad'o", errMsg))
    If errMsg <> "" Then
        Test_HistorialPublicacionesHTML_SQLRegistroCorreoEnviadoContrato = BuildFail(errMsg, logs)
        Exit Function
    End If

    If InStr(1, sql, "INSERT INTO TbProyectoEdicionesCorreoRevision (IDEnvioCorreoTecnico, IDEdicion, FechaCorreoRevision, UsuarioCalidad, IDCorreo)", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLRegistroCorreoEnviadoContrato = BuildFail("El log debe insertar en TbProyectoEdicionesCorreoRevision con los campos esperados", logs)
        Exit Function
    End If

    If InStr(1, sql, "VALUES (901, 902, #2026/05/19 10:11:12#, 'calidad''o', 903);", vbTextCompare) = 0 Then
        Test_HistorialPublicacionesHTML_SQLRegistroCorreoEnviadoContrato = BuildFail("El log debe preservar IDEdicion, IDCorreo, fecha y UsuarioCalidad sin reenviar correo", logs)
        Exit Function
    End If

    Test_HistorialPublicacionesHTML_SQLRegistroCorreoEnviadoContrato = BuildOk("registro_correo_enviado_contrato", logs)
    Exit Function
EH:
    Test_HistorialPublicacionesHTML_SQLRegistroCorreoEnviadoContrato = BuildFail("Test_HistorialPublicacionesHTML_SQLRegistroCorreoEnviadoContrato: " & Err.Description, logs)
End Function
