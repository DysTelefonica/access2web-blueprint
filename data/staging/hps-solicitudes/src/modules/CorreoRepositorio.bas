Attribute VB_Name = "CorreoRepositorio"

Option Compare Database
Option Explicit

' Función principal de búsqueda con filtros dinámicos
Public Function BuscarCorreos( _
                            Optional p_EstadoFiltro As String, _
                            Optional p_IDSolicitud As String, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Where As String
    Dim m_Correo As Correo
    
    On Error GoTo errores
    
    ' 1. Construcción de la consulta base
    m_SQL = "SELECT * FROM TbCorreosEnviados WHERE 1=1 "
    
    ' 2. Filtro por Estado (Lógica de Negocio aplicada a SQL)
    Select Case p_EstadoFiltro
        Case "Enviado"
            m_Where = m_Where & " AND FechaEnvio Is Not Null"
            
        Case "Pendiente"
            ' Pendiente = No enviado Y (No es programado O es programado pero ya venció la fecha)
            m_Where = m_Where & " AND FechaEnvio Is Null AND (FechaOrdenEnvio <= Now() OR FechaOrdenEnvio Is Null)"
            
        Case "Programado"
            ' Programado = No enviado Y tiene fecha futura
            m_Where = m_Where & " AND FechaEnvio Is Null AND FechaOrdenEnvio > Now()"
            
    End Select
    
    ' 3. Filtro por Solicitud
    If p_IDSolicitud <> "" And p_IDSolicitud <> "0" Then
        m_Where = m_Where & " AND IDSolicitud = " & p_IDSolicitud
    End If
    
    ' 4. Orden
    m_SQL = m_SQL & m_Where & " ORDER BY FechaGrabacion DESC;"
    
    ' 5. Ejecución y Mapeo (Hidratación)
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
    ' Usamos el RepositorioComun o lógica directa para hidratar la colección
    ' Aquí lo hago directo para no depender de un módulo Factoria externo si no lo tienes creado aún
    Set BuscarCorreos = New Scripting.Dictionary
    
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Correo = New Correo
                ' Hidratación manual rápida (o usar CallByName si prefieres)
                m_Correo.IDCORREO = Nz(.Fields("IDCorreo"), "")
                m_Correo.Asunto = Nz(.Fields("Asunto"), "")
                m_Correo.DESTINATARIOS = Nz(.Fields("Destinatarios"), "")
                m_Correo.FechaEnvio = Nz(.Fields("FechaEnvio"), "")
                m_Correo.FechaOrdenEnvio = Nz(.Fields("FechaOrdenEnvio"), "")
                m_Correo.Accion = Nz(.Fields("Accion"), "")
                m_Correo.IDSolicitud = Nz(.Fields("IDSolicitud"), "")
                ' ... resto de campos necesarios
                
                If Not BuscarCorreos.Exists(m_Correo.IDCORREO) Then
                    BuscarCorreos.Add m_Correo.IDCORREO, m_Correo
                End If
                
                Set m_Correo = Nothing
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rcdDatos = Nothing
    
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en CorreoRepositorio.BuscarCorreos: " & Err.Description
    End If
End Function
' En CorreoRepositorio.bas

' Devuelve True si ya existe un correo de ese tipo pendiente de envío (o programado) para esa solicitud
Public Function ExisteCorreoPendiente( _
                                    p_IDSolicitud As String, _
                                    p_TipoCorreo As EnumTipoEnvioCorreo, _
                                    ByRef p_Error As String) As Boolean
    
    Dim m_SQL As String
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    
    m_SQL = "SELECT Count(*) as Total FROM TbCorreosEnviados " & _
            "WHERE IDSolicitud = " & p_IDSolicitud & " " & _
            "AND TipoCorreo = " & p_TipoCorreo & " " & _
            "AND FechaEnvio Is Null;" ' Solo nos importan los que no han salido
            
    Set rcd = getdb().OpenRecordset(m_SQL)
    If Not rcd.EOF Then
        If rcd("Total") > 0 Then ExisteCorreoPendiente = True
    End If
    rcd.Close
    Set rcd = Nothing
    
    Exit Function
errores:
    p_Error = "Error comprobando existencia correo: " & Err.Description
End Function
