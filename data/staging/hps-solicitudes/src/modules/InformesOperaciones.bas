Attribute VB_Name = "InformesOperaciones"
Option Compare Database
Option Explicit

' ==========================================================================================
' GENERACIÓN DEL INFORME DE CONTROL (GEMELO DEL VBSCRIPT)
' ==========================================================================================

Public Function GenerarCuerpoInformeControl(Optional ByRef p_Error As String) As String
    Dim html As String
    Dim cardClass As String
    Dim textoEstado As String
    Dim icono As String
    
    On Error GoTo errores
    
    ' 1. Obtener estado de Configuración Global
    ' Asumimos que m_ObjEntorno está inicializado
    If m_ObjEntorno.Configuracion.CorreosAutomaticos = "Sí" Then
        cardClass = "card-active"
        textoEstado = "CORREOS AUTOMÁTICOS ACTIVADOS"
        icono = "&#10004;" ' Checkmark
    Else
        cardClass = "card-inactive"
        textoEstado = "CORREOS AUTOMÁTICOS DESACTIVADOS (OFF)"
        icono = "&#9888;" ' Warning symbol
    End If

    ' 2. Cabecera y CSS
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='ISO-8859-1'>"
    html = html & GetCSS_Telefonica()
    html = html & "</head><body>"
    
    html = html & "<h1>Informe de Seguimiento de Solicitudes HPS (Vista Previa)</h1>"
    
    ' 3. Card de Estado
    html = html & "<div class='card " & cardClass & "'>"
    html = html & icono & " &nbsp; " & textoEstado
    html = html & "</div>"

    html = html & "<p>Informe generado el: <b>" & Format(Now, "dd/mm/yyyy HH:mm") & "</b> por " & m_ObjUsuarioConectado.Nombre & "</p>"
    
    ' TABLA 1: SOLICITUDES EN TRÁMITE (Se mantiene igual)
    html = html & "<h2>1. Solicitudes Pendientes de Trámite</h2>"
    html = html & GenerarTablaSolicitudesEnTramite_VBA()
    
    ' TABLA 2: CORREOS ENVIADOS AYER (Se mantiene igual)
    html = html & "<h2>2. Correos Enviados (Día Anterior)</h2>"
    html = html & GenerarTablaCorreosEnviadosAyer_VBA()
    
    ' TABLA 3: CORREOS PENDIENTES (Adaptado a Almudena: 5 días + Gestor)
    html = html & "<h2>3. Correos Pendientes de Envío (Próximos 5 días)</h2>"
    html = html & GenerarTablaCorreosPendientes_VBA()
    
    ' TABLA 4: ALERTAS (Adaptado a VBS: Correos en limbo)
    html = html & "<h2>4. Alerta: Avisos Previstos NO Programados</h2>"
    html = html & GenerarTablaAlertas_VBA()
    
    html = html & "<br><br><hr style='border: 0; border-top: 1px solid #D1D5E4;'><p style='color:#58617A; font-size: 9pt;'>Generado por Aplicativo HPS (VBA Twin).</p>"
    html = html & "</body></html>"
    
    GenerarCuerpoInformeControl = html
    Exit Function

errores:
    p_Error = "Error generando cuerpo del informe: " & Err.Description
End Function

' ==========================================================================================
' ESTILOS CSS
' ==========================================================================================
Private Function GetCSS_Telefonica() As String
    Dim s As String
    s = "<style>"
    s = s & "body { font-family: 'Telefónica Sans', 'Segoe UI', Arial, sans-serif; font-size: 11pt; line-height: 1.45; color: #031A34; background: #FFFFFF; }"
    s = s & "h1 { color: #0066FF; font-size: 20pt; margin-bottom: 15pt; border-bottom: 3px solid #0066FF; padding-bottom: 5pt; }"
    s = s & "h2 { color: #2B3447; font-size: 16pt; margin-top: 20pt; margin-bottom: 10pt; border-bottom: 2px solid #D1D5E4; padding-bottom: 5pt; }"
    s = s & "table { width: 100%; border-collapse: collapse; margin: 10pt 0; font-size: 10pt; }"
    s = s & "th, td { border: 1px solid #D1D5E4; padding: 8pt; text-align: left; }"
    s = s & "th { background-color: #F2F4FF; color: #031A34; font-weight: 600; }"
    s = s & "tr:nth-child(even) { background-color: #FBF0FF; }"
    s = s & "tr:hover { background-color: #EBFFFF; }"
    
    s = s & ".tag { padding: 2px 6px; border-radius: 4px; font-size: 9pt; font-weight: bold; }"
    s = s & ".tag-blue { background-color: #E6F2FF; color: #0066FF; }"
    s = s & ".tag-grey { background-color: #F2F4FF; color: #58617A; }"
    s = s & ".tag-red { background-color: #F9EEED; color: #843C34; }"
    
    s = s & ".card { padding: 15px; border-radius: 4px; margin-bottom: 20px; border-left: 6px solid; font-size: 12pt; font-weight: bold; }"
    s = s & ".card-active { background-color: #E6F2FF; border-color: #0066FF; color: #0066FF; }"
    s = s & ".card-inactive { background-color: #F9EEED; border-color: #D6786B; color: #843C34; }"
    s = s & "</style>"
    GetCSS_Telefonica = s
End Function


' ------------------------------------------------------------------------------------------
' TABLA 3: CORREOS PENDIENTES (GEMELO VBS: 5 DÍAS + GESTOR)
' ------------------------------------------------------------------------------------------
Private Function GenerarTablaCorreosPendientes_VBA() As String
    Dim m_SQL As String
    Dim rcd As DAO.Recordset
    Dim html As String
    Dim fechaProg As Variant
    Dim Gestor As String
    Dim diasRestantes As Long
    Dim mostrarFila As Boolean
    Dim filasGeneradas As Integer
    
    ' SQL Adaptada: Incluye S.Gestor y Filtro de Nulos
    m_SQL = "SELECT C.IDCorreo, C.IDSolicitud, S.Nombre, S.Apellido1, S.TIPO, C.FechaOrdenEnvio, C.Accion, S.Gestor " & _
            "FROM TbCorreosEnviados C LEFT JOIN TbSolicitudes S ON C.IDSolicitud = S.IDSolicitud " & _
            "WHERE C.FechaEnvio Is Null AND C.FechaOrdenEnvio Is Not Null " & _
            "ORDER BY C.FechaOrdenEnvio ASC;"
            
    Set rcd = getdb().OpenRecordset(m_SQL) ' Usamos tu función getdb()
    
    If rcd.EOF Then
        GenerarTablaCorreosPendientes_VBA = "<p><i>No hay correos programados pendientes.</i></p>"
        rcd.Close
        Exit Function
    End If
    
    html = "<table><thead><tr><th>ID Solicitud</th><th>Solicitante</th><th>Gestor</th><th>Tipo</th><th>Acción/Asunto</th><th>Fecha Programada</th></tr></thead><tbody>"
    
    filasGeneradas = 0
    
    Do While Not rcd.EOF
        fechaProg = rcd!FechaOrdenEnvio
        mostrarFila = False
        
        ' LÓGICA ALMUDENA: Ventana de 5 días
        diasRestantes = DateDiff("d", Date, fechaProg)
        If diasRestantes <= 5 Then
             mostrarFila = True
        End If
        
        If mostrarFila Then
            Gestor = Nz(rcd!Gestor, "")
            If Gestor = "" Then Gestor = "Sin Asignar"
            
            html = html & "<tr>"
            html = html & "<td><strong>" & Nz(rcd!IDSolicitud, "-") & "</strong></td>"
            html = html & "<td>" & Nz(rcd!Nombre, "") & " " & Nz(rcd!Apellido1, "") & "</td>"
            ' Nueva Columna Gestor
            html = html & "<td>" & Gestor & "</td>"
            html = html & "<td>" & Nz(rcd!TIPO, "") & "</td>"
            html = html & "<td>" & Nz(rcd!Accion, "") & "</td>"
            html = html & "<td><span class='tag tag-grey'>" & fechaProg & "</span></td>"
            html = html & "</tr>"
            
            filasGeneradas = filasGeneradas + 1
        End If
        
        rcd.MoveNext
    Loop
    
    html = html & "</tbody></table>"
    
    ' Si el filtro ocultó todo
    If filasGeneradas = 0 Then
        html = "<p><i>No hay correos programados para los próximos 5 días.</i></p>"
    End If
    
    rcd.Close
    Set rcd = Nothing
    
    GenerarTablaCorreosPendientes_VBA = html
End Function

' ------------------------------------------------------------------------------------------
' TABLA 4: ALERTAS (GEMELO VBS: CORREOS EN LIMBO)
' ------------------------------------------------------------------------------------------
Private Function GenerarTablaAlertas_VBA() As String
    Dim m_SQL As String
    Dim rcd As DAO.Recordset
    Dim html As String
    Dim Gestor As String
    
    ' SQL IDÉNTICA AL VBSCRIPT: Busca correos en TbCorreosEnviados sin fecha envío y sin fecha programación
    m_SQL = "SELECT C.IDCorreo, C.IDSolicitud, S.Nombre, S.Apellido1, S.TIPO, C.Accion, S.Gestor " & _
            "FROM TbCorreosEnviados C LEFT JOIN TbSolicitudes S ON C.IDSolicitud = S.IDSolicitud " & _
            "WHERE C.FechaEnvio Is Null AND C.FechaOrdenEnvio Is Null " & _
            "ORDER BY C.IDCorreo ASC;"
            
    Set rcd = getdb().OpenRecordset(m_SQL)
    
    ' Caso Vacío (Lo normal)
    If rcd.EOF Then
        GenerarTablaAlertas_VBA = "<p><i>No se han detectado avisos fuera de programación.</i></p>"
        rcd.Close
        Exit Function
    End If
    
    ' Caso con Datos (Anomalía)
    html = "<table><thead><tr><th>ID Solicitud</th><th>Solicitante</th><th>Gestor</th><th>Tipo</th><th>Acción/Asunto</th><th>Estado</th></tr></thead><tbody>"
    
    Do While Not rcd.EOF
        Gestor = Nz(rcd!Gestor, "")
        If Gestor = "" Then Gestor = "Sin Asignar"
        
        html = html & "<tr>"
        html = html & "<td><strong>" & Nz(rcd!IDSolicitud, "-") & "</strong></td>"
        html = html & "<td>" & Nz(rcd!Nombre, "") & " " & Nz(rcd!Apellido1, "") & "</td>"
        html = html & "<td>" & Gestor & "</td>"
        html = html & "<td>" & Nz(rcd!TIPO, "") & "</td>"
        html = html & "<td>" & Nz(rcd!Accion, "") & "</td>"
        html = html & "<td><span class='tag tag-red'>SIN FECHA</span></td>"
        html = html & "</tr>"
        rcd.MoveNext
    Loop
    
    html = html & "</tbody></table>"
    rcd.Close
    Set rcd = Nothing
    
    GenerarTablaAlertas_VBA = html
End Function
' ------------------------------------------------------------------------------------------
' TABLA 1: SOLICITUDES EN TRÁMITE (GEMELO VBS)
' ------------------------------------------------------------------------------------------
Private Function GenerarTablaSolicitudesEnTramite_VBA() As String
    Dim m_SQL As String
    Dim rcd As DAO.Recordset
    Dim html As String
    Dim Gestor As String
    
    ' Variables para "traducir" el estado si tu aplicación lo requiere
    Dim sEstadoInterno As String
    Dim sEstadoTitulo As String
    Dim lEnumEstado As Long
    
    ' SQL: Filtrar no finalizadas (RegistroHPS, Desestimado, Cancelado IS NULL)
    m_SQL = "SELECT S.IDSolicitud, S.TIPO, S.Nombre, S.Apellido1, S.Apellido2, S.Estado, S.Gestor " & _
            "FROM TbSolicitudes S INNER JOIN TbSolicitudesFechas F ON S.IDSolicitud = F.IDSolicitud " & _
            "WHERE (F.FechaRegistroEnHPS Is Null) AND (F.FechaDesestimado Is Null) AND (F.FechaCancelado Is Null) " & _
            "ORDER BY S.IDSolicitud DESC;"
            
    Set rcd = getdb().OpenRecordset(m_SQL)
    
    If rcd.EOF Then
        GenerarTablaSolicitudesEnTramite_VBA = "<p><i>No hay solicitudes pendientes de trámite.</i></p>"
        rcd.Close
        Exit Function
    End If
    
    ' CABECERA: 5 Columnas (igual que VBS)
    html = "<table><thead><tr>"
    html = html & "<th>ID</th><th>Tipo</th><th>Nombre Completo</th><th>Estado</th><th>Gestor</th>"
    html = html & "</tr></thead><tbody>"
    
    Do While Not rcd.EOF
        ' 1. Formateo Gestor (Color Coral si está vacío)
        Gestor = Nz(rcd!Gestor, "")
        If Gestor = "" Then Gestor = "<span style='color:#D6786B;'>Sin Asignar</span>"
        
        ' 2. Formateo Estado (Intento de traducción o valor directo)
        sEstadoInterno = Nz(rcd!Estado, "")
        sEstadoTitulo = sEstadoInterno
        
        ' Si usas colección de estados en memoria (m_ObjEntorno), intentamos traducir:
        On Error Resume Next
        If Not m_ObjEntorno Is Nothing Then
            If m_ObjEntorno.ColTextoEstados.Exists(sEstadoInterno) Then
                lEnumEstado = m_ObjEntorno.ColTextoEstados(sEstadoInterno)
                If m_ObjEntorno.ColEstadosParaTitulo.Exists(CStr(lEnumEstado)) Then
                    sEstadoTitulo = m_ObjEntorno.ColEstadosParaTitulo(CStr(lEnumEstado))
                End If
            End If
        End If
        On Error GoTo 0
        
        ' 3. Filas
        html = html & "<tr>"
        html = html & "<td><strong>" & Nz(rcd!IDSolicitud, "") & "</strong></td>"
        html = html & "<td>" & Nz(rcd!TIPO, "") & "</td>"
        html = html & "<td>" & Nz(rcd!Nombre, "") & " " & Nz(rcd!Apellido1, "") & " " & Nz(rcd!Apellido2, "") & "</td>"
        html = html & "<td><span class='tag tag-blue'>" & sEstadoTitulo & "</span></td>"
        html = html & "<td>" & Gestor & "</td>"
        html = html & "</tr>"
        
        rcd.MoveNext
    Loop
    
    html = html & "</tbody></table>"
    rcd.Close
    Set rcd = Nothing
    
    GenerarTablaSolicitudesEnTramite_VBA = html
End Function

' ------------------------------------------------------------------------------------------
' TABLA 2: CORREOS ENVIADOS AYER (GEMELO VBS)
' ------------------------------------------------------------------------------------------
Private Function GenerarTablaCorreosEnviadosAyer_VBA() As String
    Dim m_SQL As String
    Dim rcd As DAO.Recordset
    Dim html As String
    Dim fechaAyer As Date
    
    fechaAyer = DateAdd("d", -1, Date)
    
    m_SQL = "SELECT IDCorreo, Asunto, Destinatarios, FechaEnvio " & _
            "FROM TbCorreosEnviados " & _
            "WHERE Year(FechaEnvio) = " & Year(fechaAyer) & " " & _
            "AND Month(FechaEnvio) = " & Month(fechaAyer) & " " & _
            "AND Day(FechaEnvio) = " & Day(fechaAyer) & " " & _
            "ORDER BY FechaEnvio DESC;"
            
    Set rcd = getdb().OpenRecordset(m_SQL)
    
    If rcd.EOF Then
        GenerarTablaCorreosEnviadosAyer_VBA = "<p><i>No se enviaron correos el día de ayer (" & Format(fechaAyer, "dd/mm/yyyy") & ").</i></p>"
        rcd.Close
        Exit Function
    End If
    
    html = "<table><thead><tr><th>ID</th><th>Asunto</th><th>Destinatarios</th><th>Hora Envío</th></tr></thead><tbody>"
    
    Do While Not rcd.EOF
        html = html & "<tr>"
        html = html & "<td>" & Nz(rcd!IDCORREO, "") & "</td>"
        html = html & "<td>" & Nz(rcd!Asunto, "") & "</td>"
        html = html & "<td style='font-size:9pt;'>" & Nz(rcd!DESTINATARIOS, "") & "</td>"
        ' Solo mostramos la hora para ahorrar espacio
        html = html & "<td>" & Format(rcd!FechaEnvio, "Short Time") & "</td>"
        html = html & "</tr>"
        rcd.MoveNext
    Loop
    
    html = html & "</tbody></table>"
    rcd.Close
    Set rcd = Nothing
    
    GenerarTablaCorreosEnviadosAyer_VBA = html
End Function
