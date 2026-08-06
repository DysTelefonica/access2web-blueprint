Attribute VB_Name = "cacheSuministrador"
Option Compare Database
Option Explicit

Public Function cache_suministradores_regenerar( _
                                                Optional ByRef p_Error As String _
                                                ) As String
                                
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    Dim m_Suministrador As Suministrador
    Dim m_linea As String
    
    On Error GoTo errores
    EVE
    m_SQL = "TbSuministradores"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            m_linea = "Borrando tabla caché Local "
            Avance m_linea
            DoEvents
            Debug.Print m_linea
            DoEvents
            m_SQL = "DELETE * FROM TbSuministradoresLocal;"
            CurrentDb().Execute m_SQL
            
            .MoveFirst
            Do While Not .EOF
                m_ID = .Fields("IDSuministrador")
                m_linea = "ID......." & m_ID
                Avance m_linea
                DoEvents
                Debug.Print m_linea
                DoEvents
                Set m_Suministrador = Constructor.getSuministrador(p_ID:=m_ID, p_Error:=p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                cache_suministrador_regenerar p_Suministrador:=m_Suministrador, p_Error:=p_Error
               
                If p_Error <> "" Then Err.Raise 1000
                .MoveNext
            Loop
        End If
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método cache_suministradores_regenerar ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
    
End Function
Public Function cache_suministrador_regenerar( _
                                                Optional p_Suministrador As Suministrador, _
                                                Optional p_IDSuministrador As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
                                
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Suministrador As Suministrador
    Dim m_Campo As Variant
    Dim m_Valor As String
    Dim m_ID As String
    
    On Error GoTo errores
    
    p_Error = ""
    If p_Suministrador Is Nothing Then
        Set m_Suministrador = Constructor.getSuministrador(p_ID:=p_IDSuministrador, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    Else
        Set m_Suministrador = p_Suministrador
    End If
    If m_Suministrador Is Nothing Then
        Exit Function
    End If
    m_ID = m_Suministrador.IDSuministrador
    m_SQL = "SELECT * FROM TbSuministradoresLocal " & _
            "WHERE IDSuministrador=" & m_ID & ";"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                For Each m_Campo In m_Suministrador.ColCampos
                
                    'Debug.Print m_Campo
                    'If CStr(m_Campo) = "Motivo_HPS" Then Stop
                    m_Valor = m_Suministrador.getPropiedad(m_Campo, p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If m_Valor <> "" Then
                        rcdDatos.Fields(m_Campo).value = m_Valor
                    Else
                        rcdDatos.Fields(m_Campo).value = Null
                    End If
                Next
            .Update
        Else
            .Edit
                For Each m_Campo In m_Suministrador.ColCampos
                
                    'Debug.Print m_Campo
                    'If CStr(m_Campo) = "Motivo_HPS" Then Stop
                    m_Valor = m_Suministrador.getPropiedad(m_Campo, p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If m_Valor <> "" Then
                        rcdDatos.Fields(m_Campo).value = m_Valor
                    Else
                        rcdDatos.Fields(m_Campo).value = Null
                    End If
                Next
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método cache_suministrador_regenerar ha devuelto el error: " & Err.Description
    End If
   
    
End Function





Public Function cache_usuario_suministrador_borrar( _
                                                p_IDSuministrador As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
                            
    
    Dim m_SQL As String
    
   On Error GoTo errores
    
    m_SQL = "DELETE * FROM TbSuministradoresLocal WHERE ID=" & p_IDSuministrador & ";"
    CurrentDb().Execute m_SQL
   
            
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método cache_usuario_suministrador_borrar ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function


