Attribute VB_Name = "cacheUsuario"
Option Compare Database
Option Explicit

Public Function cache_usuarios_regenerar( _
                                            Optional ByRef p_Error As String _
                                            ) As String
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    Dim m_DatosLocal As DatosLocal
    Dim m_linea As String
    
    On Error GoTo errores
    EVE
    m_SQL = "TbUsuarios"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            m_linea = "Borrando tabla caché Local "
            Avance m_linea
            DoEvents
            Debug.Print m_linea
            DoEvents
            m_SQL = "DELETE * FROM TbDatosLocal"
            CurrentDb().Execute m_SQL
            m_SQL = "DELETE * FROM TbDatosLocalParaIndicadores"
            CurrentDb().Execute m_SQL
            m_SQL = "DELETE * FROM TbUsuariosEntidades"
            getdb().Execute m_SQL
            
            .MoveFirst
            Do While Not .EOF
                m_ID = .Fields("ID")
                m_linea = "ID......." & m_ID
                Avance m_linea
                DoEvents
                Debug.Print m_linea
                DoEvents
                Set m_DatosLocal = getDatosLocalDeUsuario(p_ID:=m_ID, p_Error:=p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                cache_usuario_regenerar p_DatosLocal:=m_DatosLocal, p_Error:=p_Error
               
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
        p_Error = "El método cache_usuarios_regenerar ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function cache_usuario_regenerar( _
                                        Optional p_ID As String, _
                                        Optional p_DatosLocal As DatosLocal, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Valor As String
    Dim m_DatosLocal As DatosLocal
    Dim m_ID As String
   
    
    
    On Error GoTo errores
    If p_DatosLocal Is Nothing Then
        Set m_DatosLocal = getDatosLocalDeUsuario(p_ID:=p_ID, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    Else
        Set m_DatosLocal = p_DatosLocal
    End If
    If m_DatosLocal Is Nothing Then
        Exit Function
    End If
    m_ID = m_DatosLocal.ID
    
    m_SQL = "SELECT * FROM TbUsuariosEntidades " & _
            "WHERE ID=" & m_ID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                For Each m_Campo In m_DatosLocal.ColCampos
                
                    'Debug.Print m_Campo
                    'If CStr(m_Campo) = "Motivo_HPS" Then Stop
                    m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
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
                For Each m_Campo In m_DatosLocal.ColCampos
                
                    'Debug.Print m_Campo
                    'If CStr(m_Campo) = "Motivo_HPS" Then Stop
                    m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
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
    m_SQL = "SELECT * FROM TbDatosLocal " & _
            "WHERE ID=" & m_ID & ";"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                For Each m_Campo In m_DatosLocal.ColCampos
                
                    'Debug.Print m_Campo
                    'If CStr(m_Campo) = "Motivo_HPS" Then Stop
                    m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
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
                For Each m_Campo In m_DatosLocal.ColCampos
                
                    'Debug.Print m_Campo
                   ' If CStr(m_Campo) = "HPS_UE_F_Concesion" Then Stop
                    m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
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
    m_SQL = "SELECT * FROM TbDatosLocalParaIndicadores " & _
            "WHERE ID=" & m_ID & ";"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                For Each m_Campo In m_DatosLocal.ColCampos
                
                    'Debug.Print m_Campo
                    'If CStr(m_Campo) = "Motivo_HPS" Then Stop
                    m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
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
                For Each m_Campo In m_DatosLocal.ColCampos
                
                    'Debug.Print m_Campo
                    'If CStr(m_Campo) = "Motivo_HPS" Then Stop
                    m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
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
        p_Error = "El método cache_usuario_regenerar ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function cache_usuario_borrar( _
                                        p_ID As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                    
    
    Dim m_SQL As String
    
   On Error GoTo errores
    
    m_SQL = "DELETE * FROM TbDatosLocal WHERE ID=" & p_ID & ";"
    CurrentDb().Execute m_SQL
    m_SQL = "DELETE * FROM TbDatosLocalParaIndicadores WHERE ID=" & p_ID & ";"
    CurrentDb().Execute m_SQL
    m_SQL = "DELETE * FROM TbUsuariosEntidades WHERE ID=" & p_ID & ";"
    getdb().Execute m_SQL
            
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método cache_usuario_borrar ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function





