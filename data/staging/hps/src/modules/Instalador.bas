Attribute VB_Name = "Instalador"
Option Compare Database
Option Explicit


Public Function ActualizaCadenaContratistasHPS( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_usuario As UsuarioHPS
    
    On Error GoTo errores
    
    Set m_Col = m_ObjEntorno.UsuariosHPS
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_usuario = m_Col(m_ID)
        m_usuario.CadenaContratistasGrabar
        If m_usuario.CadenaContratistas = "" Then Stop
        Debug.Print m_usuario.NombreCompleto, m_usuario.CadenaContratistas
        Set m_usuario = Nothing
    Next
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizaCadenaContratistasHPS ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function ActualizaCadenaContratistasHistoricos( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_usuario As UsuarioHistorico
    
    On Error GoTo errores
    
    Set m_Col = Constructor.getUsuariosHistoricos(p_Error)
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_usuario = m_Col(m_ID)
        m_usuario.CadenaContratistasGrabar
        If m_usuario.CadenaContratistas = "" Then
            RegistrarSuministradorAExpediente m_usuario.IDExpediente, "12"
            m_usuario.CadenaContratistasGrabar
            If m_usuario.CadenaContratistas = "" Then Stop
        End If
        Debug.Print m_usuario.NombreCompleto, m_usuario.CadenaContratistas
        Set m_usuario = Nothing
    Next
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizaCadenaContratistasHistoricos ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function RegistrarSuministradorAExpediente( _
                                                p_IDExpediente As String, _
                                                p_IDSuministrador As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    m_SQL = "SELECT * " & _
            "FROM TbExpedientesSuministradores " & _
            "WHERE IDExpediente=" & p_IDExpediente & " AND " & _
            "IDSuministrador=" & p_IDSuministrador & ";"
    Set rcdDatos = getdbExpedientes.OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        rcdDatos.AddNew
            rcdDatos.Fields("IDExpedienteSuministrador") = DameID("TbExpedientesSuministradores", _
                                                        "IDExpedienteSuministrador", getdbExpedientes())
            rcdDatos.Fields("IDExpediente") = p_IDExpediente
            rcdDatos.Fields("IDSuministrador") = p_IDSuministrador
            rcdDatos.Fields("ContratistaPrincipal") = "Sí"
            rcdDatos.Fields("SubContratista") = "No"
        rcdDatos.Update
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizaCadenaContratistasHPS ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function getFechas( _
                            p_ID As String, _
                            p_Tipo As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Resultado As String
    
    On Error GoTo errores
    m_SQL = "SELECT TbHPS.F_Concesion, TbHPS.F_Caducidad " & _
            "FROM TbHPS " & _
            "WHERE (((TbHPS.IDUsuario)=" & p_ID & ") AND ((TbHPS.TipoHPS)='" & p_Tipo & "'));"
            
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        getFechas = ""
        Set rcdDatos = Nothing
        Exit Function
    End If
    getFechas = Nz(rcdDatos.Fields("F_Concesion"), "") & "|" & Nz(rcdDatos.Fields("F_Caducidad"), "")
    
    Exit Function
    
errores:
    getFechas = "#ERR"
    
End Function

Public Function AlgunHPS( _
                            p_ID As String _
                            ) As Boolean
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    m_SQL = "SELECT * " & _
            "FROM TbHPS " & _
            "WHERE (((TbHPS.IDUsuario)=" & p_ID & "));"
            
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Not rcdDatos.EOF Then
        AlgunHPS = True
        
        Exit Function
    End If
    
    
    Exit Function
    
errores:
    AlgunHPS = False
    
End Function
Public Function ActualizarRechas(p_Tabla As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Resultado As String
    Dim m_FechaConcesionReal As String
    Dim m_FechaCaducidadReal  As String
    Dim m_FechaConcesion As String
    Dim m_FechaCaducidad  As String
    Dim m_Tipo As String
    Dim m_ID As String
    Dim dato
    Dim db As DAO.Database
    
    If p_Tabla = "TbUsuariosEntidades" Then
        Set db = getdb()
    Else
        Set db = CurrentDb()
    End If
    m_SQL = "SELECT * " & _
            "FROM " & p_Tabla & ";"
    Set rcdDatos = db.OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        ActualizarRechas = "0"
        Set rcdDatos = Nothing
        Exit Function
    End If
    With rcdDatos
        Do While Not .EOF
            m_ID = .Fields("ID")
            m_Resultado = ""
            m_FechaConcesion = ""
            m_FechaCaducidad = ""
            If Not AlgunHPS(m_ID) Then
                'Stop
                GoTo siguiente
            End If
            .Edit
                m_Tipo = "Nacional"
                m_Resultado = getFechas(m_ID, m_Tipo)
                If InStr(1, m_Resultado, "|") <> 0 Then
                    dato = Split(m_Resultado, "|")
                    m_FechaConcesion = Nz(dato(0), "")
                    m_FechaCaducidad = Nz(dato(1), "")
                    If IsDate(m_FechaConcesion) Then
                        .Fields("HPS_NAC_F_Concesion") = m_FechaConcesion
                    Else
                        .Fields("HPS_NAC_F_Concesion") = Null
                    End If
                    If IsDate(m_FechaCaducidad) Then
                        .Fields("HPS_NAC_F_Caducidad") = m_FechaCaducidad
                    Else
                        .Fields("HPS_NAC_F_Caducidad") = Null
                    End If
                End If
                

                m_Tipo = "OTAN"
                m_Resultado = getFechas(m_ID, m_Tipo)
                If InStr(1, m_Resultado, "|") <> 0 Then
                    dato = Split(m_Resultado, "|")
                    m_FechaConcesion = Nz(dato(0), "")
                    m_FechaCaducidad = Nz(dato(1), "")
                    If IsDate(m_FechaConcesion) Then
                        .Fields("HPS_OTAN_F_Concesion") = m_FechaConcesion
                    Else
                        .Fields("HPS_OTAN_F_Concesion") = Null
                    End If
                    If IsDate(m_FechaCaducidad) Then
                        .Fields("HPS_OTAN_F_Caducidad") = m_FechaCaducidad
                    Else
                        .Fields("HPS_OTAN_F_Caducidad") = Null
                    End If
                End If
                

                m_Tipo = "ESA"
                m_Resultado = getFechas(m_ID, m_Tipo)
                If InStr(1, m_Resultado, "|") <> 0 Then
                    dato = Split(m_Resultado, "|")
                    m_FechaConcesion = Nz(dato(0), "")
                    m_FechaCaducidad = Nz(dato(1), "")
                    If IsDate(m_FechaConcesion) Then
                        .Fields("HPS_ESA_F_Concesion") = m_FechaConcesion
                    Else
                        .Fields("HPS_ESA_F_Concesion") = Null
                    End If
                    If IsDate(m_FechaCaducidad) Then
                        .Fields("HPS_ESA_F_Caducidad") = m_FechaCaducidad
                    Else
                        .Fields("HPS_ESA_F_Caducidad") = Null
                    End If
                End If
                
                m_Tipo = "UE"
                m_Resultado = getFechas(m_ID, m_Tipo)
                If InStr(1, m_Resultado, "|") <> 0 Then
                    dato = Split(m_Resultado, "|")
                    m_FechaConcesion = Nz(dato(0), "")
                    m_FechaCaducidad = Nz(dato(1), "")
                    If IsDate(m_FechaConcesion) Then
                        .Fields("HPS_UE_F_Concesion") = m_FechaConcesion
                    Else
                        .Fields("HPS_UE_F_Concesion") = Null
                    End If
                    If IsDate(m_FechaCaducidad) Then
                        .Fields("HPS_UE_F_Caducidad") = m_FechaCaducidad
                    Else
                        .Fields("HPS_UE_F_Caducidad") = Null
                    End If
                End If
                
            .Update
            VBA.DoEvents
            
            Debug.Print m_ID
siguiente:
            .MoveNext
        Loop
    End With
    
    Exit Function
  
End Function

