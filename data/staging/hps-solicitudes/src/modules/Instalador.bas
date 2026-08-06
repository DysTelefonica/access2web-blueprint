Attribute VB_Name = "Instalador"
Option Compare Database
Option Explicit
Public Function getLogsGeneralesSinIDSolicitud( _
                                                Optional ByRef p_Error As String _
                                                 ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Log As LogGeneral
    
    
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbLogsGeneral " & _
            "WHERE IDSolicitud Is Null ;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Log = New LogGeneral
               For Each m_Campo In m_Log.ColCampos
                  m_Log.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getLogsGeneralesSinIDSolicitud Is Nothing Then
                   Set getLogsGeneralesSinIDSolicitud = New Scripting.Dictionary
                   getLogsGeneralesSinIDSolicitud.CompareMode = TextCompare
               End If
               If Not getLogsGeneralesSinIDSolicitud.Exists(m_Log.IDLog) Then
                   getLogsGeneralesSinIDSolicitud.Add m_Log.IDLog, m_Log
               End If
               Set m_Log = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getLogsGeneralesSinIDSolicitud ha devuelto el error: " & Err.Description
    End If
End Function


Public Function RegularizarIDSolicitudLog( _
                                            Optional ByRef p_Error As String _
                                            ) As String
   
    
    Dim m_Log As LogGeneral
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_SQL As String
    Dim m_IDSolicitud As String
    On Error GoTo errores
    
    m_SQL = "DELETE TbLogsGeneral.DNI " & _
            "FROM TbLogsGeneral " & _
            "WHERE (((TbLogsGeneral.DNI) Is Null));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbLogsGeneral.* " & _
            "FROM TbLogsGeneral LEFT JOIN TbSolicitudes ON TbLogsGeneral.DNI = TbSolicitudes.DNI " & _
            "WHERE (((TbSolicitudes.DNI) Is Null));"
    getdb().Execute m_SQL
    Set m_Col = getLogsGeneralesSinIDSolicitud(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Log = m_Col(m_ID)
        m_IDSolicitud = m_Log.IDSolicitudCalculado
        m_SQL = "UPDATE TbLogsGeneral SET idSolicitud = " & m_IDSolicitud & " " & _
                    "WHERE IDLog=" & m_Log.IDLog & ";"
        getdb().Execute m_SQL
        Set m_Log = Nothing
    Next
    
    RegularizarIDSolicitudLog = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegularizarIDSolicitudLog ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RegularizarEstados( _
                                            Optional ByRef p_Error As String _
                                            ) As String
   
    
   
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Solicitud As solicitud
    Dim m_EstadoInicial As String
    Dim m_EstadoFinal As String
    On Error GoTo errores
    
    Set m_Col = m_ObjEntorno.ColSolicitudes
    
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Solicitud = m_Col(m_ID)
        m_EstadoInicial = m_Solicitud.Estado
        m_Solicitud.EstadoGrabar p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_EstadoFinal = m_Solicitud.Estado
        VBA.DoEvents
        Debug.Print m_ID, m_EstadoInicial, m_EstadoFinal
        If m_EstadoFinal = "" Then
            Stop
        End If
        
        Set m_Solicitud = Nothing
    Next
    
    RegularizarEstados = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegularizarEstados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

