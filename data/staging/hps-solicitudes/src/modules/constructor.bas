Attribute VB_Name = "constructor"
Option Compare Database
Option Explicit
Public Function getcadenaContratistas( _
                                         p_IDExp As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Nombre As String
    
    On Error GoTo errores
    m_SQL = "SELECT TbSuministradores.Nombre, TbSuministradores.Nemotecnico " & _
            "FROM TbSuministradores INNER JOIN TbExpedientesSuministradores " & _
            "ON TbSuministradores.IDSuministrador = TbExpedientesSuministradores.IDSuministrador " & _
            "WHERE (((TbExpedientesSuministradores.IDExpediente)=" & p_IDExp & "));"
    
   
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_Nombre = Nz(.Fields("Nemotecnico"), "")
            If m_Nombre = "" Then
                m_Nombre = .Fields("Nombre")
            End If
            
            If getcadenaContratistas = "" Then
                getcadenaContratistas = m_Nombre
            Else
                getcadenaContratistas = getcadenaContratistas & "|" & m_Nombre
            End If
            
            m_Nombre = ""
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getcadenaContratistas ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuarioConectadoPorMaquina( _
                                                Optional ByRef p_Error As String _
                                                ) As Usuario
    Dim objNetwork As Object
    On Error GoTo errores
    Set objNetwork = CreateObject("Wscript.Network")
    Set getUsuarioConectadoPorMaquina = constructor.getUsuario(, objNetwork.UserName, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set objNetwork = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioConectadoPorMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getUsuariosAdministradores( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Usuario As Usuario
    
    
    On Error GoTo errores
    
    
    
    m_SQL = "SELECT * " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE EsAdministrador='Sí' AND FechaBaja Is Null;"

    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Usuario = New Usuario
                For Each m_Campo In m_Usuario.ColCampos
                    m_Usuario.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                     If p_Error <> "" Then
                         Err.Raise 1000
                     End If
                 Next
                 If getUsuariosAdministradores Is Nothing Then
                    Set getUsuariosAdministradores = New Scripting.Dictionary
                    getUsuariosAdministradores.CompareMode = TextCompare
                 End If
                 If Not getUsuariosAdministradores.Exists(m_Usuario.id) Then
                    getUsuariosAdministradores.Add m_Usuario.id, m_Usuario
                 End If
                 
                 Set m_Usuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicacionesPermisos INNER JOIN TbUsuariosAplicaciones " & _
            "ON TbUsuariosAplicacionesPermisos.CorreoUsuario = TbUsuariosAplicaciones.CorreoUsuario " & _
            "WHERE (((TbUsuariosAplicacionesPermisos.EsUsuarioAdministrador)='Sí') " & _
            "AND ((TbUsuariosAplicaciones.FechaBaja) Is Null) " & _
            "AND ((TbUsuariosAplicacionesPermisos.IDAplicacion)=" & IDAplicacion & "));"

    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Usuario = New Usuario
                For Each m_Campo In m_Usuario.ColCampos
                    m_Usuario.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                     If p_Error <> "" Then
                         Err.Raise 1000
                     End If
                 Next
                 If getUsuariosAdministradores Is Nothing Then
                    Set getUsuariosAdministradores = New Scripting.Dictionary
                    getUsuariosAdministradores.CompareMode = TextCompare
                 End If
                 If Not getUsuariosAdministradores.Exists(m_Usuario.id) Then
                    getUsuariosAdministradores.Add m_Usuario.id, m_Usuario
                 End If
                 
                 Set m_Usuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosAdministradores ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosTecnicos( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Usuario As Usuario
    
    
    On Error GoTo errores
    
    
    
   
    m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicacionesPermisos INNER JOIN TbUsuariosAplicaciones " & _
            "ON TbUsuariosAplicacionesPermisos.CorreoUsuario = TbUsuariosAplicaciones.CorreoUsuario " & _
            "WHERE (((TbUsuariosAplicacionesPermisos.EsUsuarioTecnico)='Sí') " & _
            "AND ((TbUsuariosAplicaciones.FechaBaja) Is Null) " & _
            "AND ((TbUsuariosAplicacionesPermisos.IDAplicacion)=" & IDAplicacion & "));"

    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Usuario = New Usuario
                For Each m_Campo In m_Usuario.ColCampos
                    m_Usuario.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                     If p_Error <> "" Then
                         Err.Raise 1000
                     End If
                 Next
                 If getUsuariosTecnicos Is Nothing Then
                    Set getUsuariosTecnicos = New Scripting.Dictionary
                    getUsuariosTecnicos.CompareMode = TextCompare
                 End If
                 If Not getUsuariosTecnicos.Exists(m_Usuario.id) Then
                    getUsuariosTecnicos.Add m_Usuario.id, m_Usuario
                 End If
                 
                 Set m_Usuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosTecnicos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosTramitadores( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Usuario As Usuario
    
    
    On Error GoTo errores
    
    
    
   
    m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicacionesPermisos INNER JOIN TbUsuariosAplicaciones " & _
            "ON TbUsuariosAplicacionesPermisos.CorreoUsuario = TbUsuariosAplicaciones.CorreoUsuario " & _
            "WHERE (((TbUsuariosAplicacionesPermisos.EsUsuarioAdministrador)='Sí') " & _
            "AND ((TbUsuariosAplicaciones.FechaBaja) Is Null) " & _
            "AND ((TbUsuariosAplicacionesPermisos.IDAplicacion)=" & IDAplicacion & "));"

    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Usuario = New Usuario
                For Each m_Campo In m_Usuario.ColCampos
                    m_Usuario.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                     If p_Error <> "" Then
                         Err.Raise 1000
                     End If
                 Next
                 If getUsuariosTramitadores Is Nothing Then
                    Set getUsuariosTramitadores = New Scripting.Dictionary
                    getUsuariosTramitadores.CompareMode = TextCompare
                 End If
                 If Not getUsuariosTramitadores.Exists(m_Usuario.id) Then
                    getUsuariosTramitadores.Add m_Usuario.id, m_Usuario
                 End If
                 
                 Set m_Usuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosTramitadores ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getConfiguracion( _
                                Optional ByRef p_Error As String _
                                ) As Configuracion
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    
    On Error GoTo errores
    
    
    m_SQL = "TbConfiguracion"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getConfiguracion = New Configuracion
        For Each m_Campo In getConfiguracion.ColCampos
            'If CStr(m_Campo) = "TipoInforme" Then Stop
            getConfiguracion.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getConfiguracion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUsuario( _
                            Optional p_ID As String, _
                            Optional p_UsuarioRed As String, _
                            Optional p_Nombre As String, _
                            Optional p_Correo As String, _
                            Optional ByRef p_Error As String _
                            ) As Usuario

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_NombreCampoID As String
    Dim m_EsNumeroID As Boolean
    Dim m_Where As String
    Dim m_ValorID As String
    Dim m_SQLInicial As String
    
    On Error GoTo errores
    If p_ID = "" And p_UsuarioRed = "" And p_Nombre = "" And p_Correo = "" Then
        Exit Function
    End If
    
    m_SQLInicial = "SELECT TbUsuariosAplicaciones.* " & _
                    "FROM TbUsuariosAplicaciones "
    
    If p_ID <> "" Then
        m_NombreCampoID = "ID"
        m_ValorID = p_ID
        m_EsNumeroID = True
    ElseIf p_UsuarioRed <> "" Then
        m_NombreCampoID = "UsuarioRed"
        m_ValorID = p_UsuarioRed
        m_EsNumeroID = False
    ElseIf p_Nombre <> "" Then
        m_NombreCampoID = "Nombre"
        m_ValorID = p_Nombre
        m_EsNumeroID = False
    ElseIf p_Correo <> "" Then
        m_NombreCampoID = "CorreoUsuario"
        m_ValorID = p_Correo
        m_EsNumeroID = False
    End If
    
    If m_EsNumeroID Then
        m_Where = m_NombreCampoID & "=" & m_ValorID & ";"
    Else
        m_Where = m_NombreCampoID & "='" & m_ValorID & "';"
    End If
    m_SQL = m_SQLInicial & "WHERE " & m_Where
    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuario = New Usuario
        For Each m_Campo In getUsuario.ColCampos
            getUsuario.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuario ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAplicacionesPermisos( _
                                            p_CorreoUsuario As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
            
    On Error GoTo errores
    If p_CorreoUsuario = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & p_CorreoUsuario & "';"
    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjUsuarioAplicacionPermisos = New UsuarioAplicacionPermisos
            For Each m_Campo In m_ObjUsuarioAplicacionPermisos.ColCampos
               m_ObjUsuarioAplicacionPermisos.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getAplicacionesPermisos Is Nothing Then
                Set getAplicacionesPermisos = New Scripting.Dictionary
                getAplicacionesPermisos.CompareMode = TextCompare
            End If
            If Not getAplicacionesPermisos.Exists(CStr(m_ObjUsuarioAplicacionPermisos.IDAplicacion)) Then
                getAplicacionesPermisos.Add m_ObjUsuarioAplicacionPermisos.IDAplicacion, m_ObjUsuarioAplicacionPermisos
            End If
            Set m_ObjUsuarioAplicacionPermisos = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAplicacionesPermisos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudFechas( _
                                        p_IDSolicitud As String, _
                                        Optional ByRef p_Error As String _
                                        ) As SolicitudFechas
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_IDSolicitud = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudesFechas " & _
            "WHERE IDSolicitud=" & p_IDSolicitud & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getSolicitudFechas = New SolicitudFechas
        For Each m_Campo In getSolicitudFechas.ColCampos
            'Debug.Print m_Campo
            'If CStr(m_Campo) = "FechaEnvioTraspasoONS" Then Stop
            getSolicitudFechas.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudFechas ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitud( _
                                p_IDSolicitud As String, _
                                Optional ByRef p_Error As String _
                                ) As solicitud
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_IDSolicitud = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudes " & _
            "WHERE IDSolicitud=" & p_IDSolicitud & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getSolicitud = New solicitud
        For Each m_Campo In getSolicitud.ColCampos
            getSolicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitud ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesPorDNI( _
                                        p_DNI As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    On Error GoTo errores
    If p_DNI = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudes " & _
            "WHERE DNI='" & p_DNI & "';"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               'If InStr(1, m_Solicitud.Nombre, "José Antonio") <> 0 Then Stop
               If getSolicitudesPorDNI Is Nothing Then
                   Set getSolicitudesPorDNI = New Scripting.Dictionary
                   getSolicitudesPorDNI.CompareMode = TextCompare
               End If
               If Not getSolicitudesPorDNI.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPorDNI.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPorDNI ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudPrimeraDeLog( _
                                        p_DNI As String, _
                                        Optional ByRef p_Error As String _
                                        ) As solicitud
    Dim m_ID As Variant
    Dim m_Solicitud As solicitud
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    If p_DNI = "" Then
        Exit Function
    End If
    Set m_Col = constructor.getSolicitudesPorDNI(p_DNI:=p_DNI, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set getSolicitudPrimeraDeLog = m_Col(m_ID)
        Exit Function
    Next
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudPrimeraDeLog ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getPendientesAltaSolicitud( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    Dim m_Fecha As String
    
    
    On Error GoTo errores
    
    m_Fecha = Format(DateAdd("m", -9, CDate(Date)), "mm/dd/yyyy")
    m_SQL = "SELECT DISTINCT TbUsuarios.* " & _
            "FROM TbUsuarios INNER JOIN TbHPS ON TbUsuarios.ID = TbHPS.IDUsuario " & _
            "WHERE " & _
            "F_Solicitud Is Null AND " & _
            "F_Caducidad<=#" & m_Fecha & "# " & _
            "AND F_Baja Is Null;"
    
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_UsuarioHPS = New UsuarioHPS
               For Each m_Campo In m_UsuarioHPS.ColCampos
                  m_UsuarioHPS.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getPendientesAltaSolicitud Is Nothing Then
                   Set getPendientesAltaSolicitud = New Scripting.Dictionary
                   getPendientesAltaSolicitud.CompareMode = TextCompare
               End If
               If Not getPendientesAltaSolicitud.Exists(m_UsuarioHPS.id) Then
                   getPendientesAltaSolicitud.Add m_UsuarioHPS.id, m_UsuarioHPS
               End If
               Set m_UsuarioHPS = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing

    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getPendientesAltaSolicitud ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getSolicitudesPendientesEnvioExcel( _
                                                    Optional p_Tramitador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    
    
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes LEFT JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "TbSolicitudesFechas.IDSolicitud Is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes LEFT JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "TbSolicitudesFechas.IDSolicitud Is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               
               For Each m_Campo In m_Solicitud.ColCampos
                'Debug.Print m_Campo
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesEnvioExcel Is Nothing Then
                   Set getSolicitudesPendientesEnvioExcel = New Scripting.Dictionary
                   getSolicitudesPendientesEnvioExcel.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesEnvioExcel.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesEnvioExcel.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "FechaEnvioExcel Is Null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "FechaEnvioExcel Is Null  AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
   
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesEnvioExcel Is Nothing Then
                   Set getSolicitudesPendientesEnvioExcel = New Scripting.Dictionary
                   getSolicitudesPendientesEnvioExcel.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesEnvioExcel.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesEnvioExcel.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesEnvioExcel ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesPendientesRecordatorioExcel1( _
                                                            Optional p_Tramitador As String, _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Not FechaEnvioExcel Is Null AND " & _
                "FechaRecepcionExcel is null AND " & _
                "FechaPrevistaCorreoRecordatorioExcel1 >=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioExcel1 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Not FechaEnvioExcel Is Null AND " & _
                "FechaRecepcionExcel is null AND " & _
                "FechaPrevistaCorreoRecordatorioExcel1 >=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioExcel1 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesRecordatorioExcel1 Is Nothing Then
                   Set getSolicitudesPendientesRecordatorioExcel1 = New Scripting.Dictionary
                   getSolicitudesPendientesRecordatorioExcel1.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesRecordatorioExcel1.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesRecordatorioExcel1.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesRecordatorioExcel1 ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesPendientesRecordatorioExcel2( _
                                                            Optional p_Tramitador As String, _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Not FechaEnvioExcel Is Null AND " & _
                "FechaRecepcionExcel is null AND " & _
                "FechaPrevistaCorreoRecordatorioExcel2 >=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioExcel2 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Not FechaEnvioExcel Is Null AND " & _
                "FechaRecepcionExcel is null AND " & _
                "FechaPrevistaCorreoRecordatorioExcel2 >=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioExcel2 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesRecordatorioExcel2 Is Nothing Then
                   Set getSolicitudesPendientesRecordatorioExcel2 = New Scripting.Dictionary
                   getSolicitudesPendientesRecordatorioExcel2.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesRecordatorioExcel2.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesRecordatorioExcel2.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesRecordatorioExcel2 ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesPendientesRellenoExcel( _
                                                    Optional p_Tramitador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    
    
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Not FechaEnvioExcel Is Null AND " & _
                "FechaRecepcionExcel Is Null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Not FechaEnvioExcel Is Null AND " & _
                "FechaRecepcionExcel Is Null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
    
    
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesRellenoExcel Is Nothing Then
                   Set getSolicitudesPendientesRellenoExcel = New Scripting.Dictionary
                   getSolicitudesPendientesRellenoExcel.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesRellenoExcel.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesRellenoExcel.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesRellenoExcel ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesPendientesAltaMarga( _
                                                    Optional p_Tramitador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Tipo<>'Traspaso' AND " & _
                "FechaTramitacionAltaMarga Is Null AND " & _
                "Not FechaRecepcionExcel is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Tipo<>'Traspaso' AND " & _
                "FechaTramitacionAltaMarga Is Null AND " & _
                "Not FechaRecepcionExcel is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                'If CStr(m_Campo) = "Nemotecnico" Then Stop
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesAltaMarga Is Nothing Then
                   Set getSolicitudesPendientesAltaMarga = New Scripting.Dictionary
                   getSolicitudesPendientesAltaMarga.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesAltaMarga.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesAltaMarga.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesAltaMarga ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesPendientesEnvioDPS( _
                                                    Optional p_Tramitador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Tipo<>'Traspaso' AND " & _
                "FechaEnvioDPS Is Null AND " & _
                "Not FechaTramitacionAltaMarga is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Tipo<>'Traspaso' AND " & _
                "FechaEnvioDPS Is Null AND " & _
                "Not FechaTramitacionAltaMarga is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesEnvioDPS Is Nothing Then
                   Set getSolicitudesPendientesEnvioDPS = New Scripting.Dictionary
                   getSolicitudesPendientesEnvioDPS.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesEnvioDPS.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesEnvioDPS.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesEnvioDPS ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesPendientesRecordatorioEnvioDPS1( _
                                                                    Optional p_Tramitador As String, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Tipo<>'Traspaso' AND " & _
                "NOT FechaTramitacionAltaMarga Is Null AND " & _
                "FechaEnvioDPS is null AND " & _
                "FechaPrevistaCorreoRecordatorioRellenoMarga1 <=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioRellenoMarga1 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Tipo<>'Traspaso' AND " & _
                "NOT FechaTramitacionAltaMarga Is Null AND " & _
                "FechaEnvioDPS is null AND " & _
                "FechaPrevistaCorreoRecordatorioRellenoMarga1 <=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioRellenoMarga1 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesRecordatorioEnvioDPS1 Is Nothing Then
                   Set getSolicitudesPendientesRecordatorioEnvioDPS1 = New Scripting.Dictionary
                   getSolicitudesPendientesRecordatorioEnvioDPS1.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesRecordatorioEnvioDPS1.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesRecordatorioEnvioDPS1.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesRecordatorioEnvioDPS1 ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesPendientesRecordatorioEnvioDPS2( _
                                                                    Optional p_Tramitador As String, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Tipo<>'Traspaso' AND " & _
                "NOT FechaTramitacionAltaMarga Is Null AND " & _
                "FechaEnvioDPS is null AND " & _
                "FechaPrevistaCorreoRecordatorioRellenoMarga2 <=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioRellenoMarga2 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Tipo<>'Traspaso' AND " & _
                "NOT FechaTramitacionAltaMarga Is Null AND " & _
                "FechaEnvioDPS is null AND " & _
                "FechaPrevistaCorreoRecordatorioRellenoMarga2 <=#" & Format(Date, "mm/dd/yyyy") & "# AND " & _
                "FechaCorreoRecordatorioRellenoMarga2 is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesRecordatorioEnvioDPS2 Is Nothing Then
                   Set getSolicitudesPendientesRecordatorioEnvioDPS2 = New Scripting.Dictionary
                   getSolicitudesPendientesRecordatorioEnvioDPS2.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesRecordatorioEnvioDPS2.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesRecordatorioEnvioDPS2.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesRecordatorioEnvioDPS2 ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesPendientesRegistroONS( _
                                                    Optional p_Tramitador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "FechaEnvioONS Is Null AND " & _
                "Not FechaEnvioDPS is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                 "FechaEnvioONS Is Null AND " & _
                "Not FechaEnvioDPS is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesRegistroONS Is Nothing Then
                   Set getSolicitudesPendientesRegistroONS = New Scripting.Dictionary
                   getSolicitudesPendientesRegistroONS.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesRegistroONS.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesRegistroONS.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesRegistroONS ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesPendientesRegistroEnHPS( _
                                                    Optional p_Tramitador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "FechaRegistroEnHPS Is Null AND " & _
                "Not FechaEnvioONS is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                 "FechaRegistroEnHPS Is Null AND " & _
                "Not FechaEnvioONS is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesRegistroEnHPS Is Nothing Then
                   Set getSolicitudesPendientesRegistroEnHPS = New Scripting.Dictionary
                   getSolicitudesPendientesRegistroEnHPS.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesRegistroEnHPS.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesRegistroEnHPS.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesRegistroEnHPS ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesPendientesEnvioTraspasoONS( _
                                                            Optional p_Tramitador As String, _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    On Error GoTo errores
    
   
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Tipo='Traspaso' AND " & _
                "FechaEnvioONS Is Null AND " & _
                "Not FechaRecepcionExcel is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Tipo='Traspaso' AND " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "FechaEnvioONS Is Null AND " & _
                "Not FechaRecepcionExcel is null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesEnvioTraspasoONS Is Nothing Then
                   Set getSolicitudesPendientesEnvioTraspasoONS = New Scripting.Dictionary
                   getSolicitudesPendientesEnvioTraspasoONS.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesEnvioTraspasoONS.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesEnvioTraspasoONS.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesEnvioTraspasoONS ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesPendientes( _
                                        p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                        Optional p_Tramitador As String = "Todos", _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    Dim m_Usuario As Usuario
    Dim m_Where As String
    Dim m_SQLInicio As String
    On Error GoTo errores
    If m_ObjEntorno.CorreosAutomaticos <> "Sí" Then
        Exit Function
    End If
    If p_Tramitador <> "Todos" Then
        Set m_Usuario = constructor.getUsuario(p_Nombre:=p_Tramitador, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Usuario Is Nothing Then
            p_Error = "Tramitador no reconocido"
            Err.Raise 1000
        End If
        p_Tramitador = m_Usuario.UsuarioRed
    Else
        p_Tramitador = ""
    End If
    m_SQLInicio = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud "
                
    m_Where = getWhereSolicitudesPendientes(p_EnumTipoEnvioCorreo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = m_SQLInicio & " " & m_Where
    
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "CorreosAutomaticos='Sí' " & _
                "AND FechaCorreoRecordatorioExcel1 Is Null " & _
                "AND Not FechaPrevistaCorreoRecordatorioExcel1 Is Null " & _
                "AND FechaDesestimado Is Null " & _
                "AND FechaCancelado Is Null " & _
                "AND FechaTramitacionAltaMarga Is Null " & _
                ";"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "CorreosAutomaticos='Sí' " & _
                "AND FechaCorreoRecordatorioExcel1 Is Null " & _
                "AND Not FechaPrevistaCorreoRecordatorioExcel1 Is Null " & _
                "AND FechaDesestimado Is Null " & _
                "AND FechaCancelado Is Null " & _
                "AND FechaTramitacionAltaMarga Is Null " & _
                "AND Gestor='" & p_Tramitador & "' " & _
                ";"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientes Is Nothing Then
                   Set getSolicitudesPendientes = New Scripting.Dictionary
                   getSolicitudesPendientes.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientes.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientes.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientes ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getWhereSolicitudesPendientes( _
                                                p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                                Optional p_Tramitador As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    
    On Error GoTo errores
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel1 Then
        getWhereSolicitudesPendientes = "WHERE " & _
                                        "CorreosAutomaticos='Sí' " & _
                                        "AND FechaCorreoRecordatorioExcel1 Is Null " & _
                                        "AND Not FechaPrevistaCorreoRecordatorioExcel1 Is Null " & _
                                        "AND FechaDesestimado Is Null " & _
                                        "AND FechaCancelado Is Null " & _
                                        "AND FechaTramitacionAltaMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionPreMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionMarga Is Null " & _
                                        ";"
    ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel2 Then
        getWhereSolicitudesPendientes = "WHERE " & _
                                        "CorreosAutomaticos='Sí' " & _
                                        "AND FechaCorreoRecordatorioExcel2 Is Null " & _
                                        "AND Not FechaPrevistaCorreoRecordatorioExcel2 Is Null " & _
                                        "AND FechaDesestimado Is Null " & _
                                        "AND FechaCancelado Is Null " & _
                                        "AND FechaTramitacionAltaMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionPreMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionMarga Is Null " & _
                                        ";"
    ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionPreMarga Then
        getWhereSolicitudesPendientes = "WHERE " & _
                                        "CorreosAutomaticos='Sí' " & _
                                        "AND FechaDesestimado Is Null " & _
                                        "AND FechaCancelado Is Null " & _
                                        "AND Not FechaPrevistaAutocancelacionPreMarga Is Null " & _
                                        ";"
    ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga1 Then
        getWhereSolicitudesPendientes = "WHERE " & _
                                        "CorreosAutomaticos='Sí' " & _
                                        "AND FechaCorreoRecordatorioRellenoMarga1 Is Null " & _
                                        "AND Not FechaPrevistaCorreoRecordatorioRellenoMarga1 Is Null " & _
                                        "AND FechaDesestimado Is Null " & _
                                        "AND FechaCancelado Is Null " & _
                                        "AND not FechaTramitacionAltaMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionPreMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionMarga Is Null " & _
                                        ";"
    ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga2 Then
        getWhereSolicitudesPendientes = "WHERE " & _
                                        "CorreosAutomaticos='Sí' " & _
                                        "AND FechaCorreoRecordatorioRellenoMarga2 Is Null " & _
                                        "AND Not FechaPrevistaCorreoRecordatorioRellenoMarga2 Is Null " & _
                                        "AND FechaDesestimado Is Null " & _
                                        "AND FechaCancelado Is Null " & _
                                        "AND not FechaTramitacionAltaMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionPreMarga Is Null " & _
                                        "AND FechaPrevistaAutocancelacionMarga Is Null " & _
                                        ";"
    ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionMarga Then
        getWhereSolicitudesPendientes = "WHERE " & _
                                        "CorreosAutomaticos='Sí' " & _
                                        "AND FechaDesestimado Is Null " & _
                                        "AND FechaCancelado Is Null " & _
                                        "AND Not FechaPrevistaAutocancelacionMarga Is Null " & _
                                        ";"
    Else
        p_Error = "Tipo de Correo no codificado"
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getWhereSolicitudesPendientes ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesPendientesEnvioONS( _
                                                    Optional p_Tramitador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Not FechaRecepcionExcel Is Null AND " & _
                "FechaEnvioONS Is Null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Not FechaRecepcionExcel Is Null AND " & _
                "FechaEnvioONS Is Null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesPendientesEnvioONS Is Nothing Then
                   Set getSolicitudesPendientesEnvioONS = New Scripting.Dictionary
                   getSolicitudesPendientesEnvioONS.CompareMode = TextCompare
               End If
               If Not getSolicitudesPendientesEnvioONS.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesPendientesEnvioONS.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesPendientesEnvioONS ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesTramitadas( _
                                            Optional p_Tramitador As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Not FechaEnvioONS Is Null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' AND " & _
                "Not FechaEnvioONS Is Null AND " & _
                "FechaDesestimado is null AND " & _
                "FechaCancelado is Null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesTramitadas Is Nothing Then
                   Set getSolicitudesTramitadas = New Scripting.Dictionary
                   getSolicitudesTramitadas.CompareMode = TextCompare
               End If
               If Not getSolicitudesTramitadas.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesTramitadas.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesTramitadas ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesEnTramite( _
                                            Optional p_Tramitador As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE (((TbSolicitudesFechas.[FechaRegistroEnHPS]) Is Null) " & _
                "And ((TbSolicitudesFechas.[FechaDesestimado]) Is Null) " & _
                "And ((TbSolicitudesFechas.[FechaCancelado]) Is Null)) " & _
                "Or (((TbSolicitudesFechas.[FechaEnvioONS]) Is Null) " & _
                "And ((TbSolicitudesFechas.[FechaDesestimado]) Is Null) " & _
                "And ((TbSolicitudesFechas.[FechaCancelado]) Is Null)) " & _
                "ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
                "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
                "WHERE (((TbSolicitudesFechas.[FechaRegistroEnHPS]) Is Null) " & _
                "AND ((TbSolicitudesFechas.[FechaDesestimado]) Is Null) " & _
                "AND ((TbSolicitudesFechas.[FechaCancelado]) Is Null) " & _
                "AND ((TbSolicitudes.Gestor)='" & p_Tramitador & "')) OR (((TbSolicitudesFechas.[FechaEnvioONS]) Is Null) " & _
                "AND ((TbSolicitudesFechas.[FechaDesestimado]) Is Null) " & _
                "AND ((TbSolicitudesFechas.[FechaCancelado]) Is Null) AND ((TbSolicitudes.Gestor)='" & p_Tramitador & "')) " & _
                "ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesEnTramite Is Nothing Then
                   Set getSolicitudesEnTramite = New Scripting.Dictionary
                   getSolicitudesEnTramite.CompareMode = TextCompare
               End If
               If Not getSolicitudesEnTramite.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesEnTramite.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesEnTramite ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesDesestimadas( _
                                            Optional p_Tramitador As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
            "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
            "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
            "WHERE " & _
            "Not FechaDesestimado is null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
       m_SQL = "SELECT TbSolicitudes.* " & _
            "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
            "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
            "WHERE " & _
            "Not FechaDesestimado is null  AND " & _
            "Gestor='" & p_Tramitador & "' ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesDesestimadas Is Nothing Then
                   Set getSolicitudesDesestimadas = New Scripting.Dictionary
                   getSolicitudesDesestimadas.CompareMode = TextCompare
               End If
               If Not getSolicitudesDesestimadas.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesDesestimadas.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesDesestimadas ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesCanceladas( _
                                            Optional p_Tramitador As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
         m_SQL = "SELECT TbSolicitudes.* " & _
            "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
            "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
            "WHERE " & _
            "Not FechaCancelado is null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
         m_SQL = "SELECT TbSolicitudes.* " & _
            "FROM TbSolicitudes INNER JOIN TbSolicitudesFechas " & _
            "ON TbSolicitudes.IDSolicitud = TbSolicitudesFechas.IDSolicitud " & _
            "WHERE " & _
            "Gestor='" & p_Tramitador & "' AND " & _
            "Not FechaCancelado is null ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If
   
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesCanceladas Is Nothing Then
                   Set getSolicitudesCanceladas = New Scripting.Dictionary
                   getSolicitudesCanceladas.CompareMode = TextCompare
               End If
               If Not getSolicitudesCanceladas.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesCanceladas.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesCanceladas ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudes( _
                                    Optional p_Tramitador As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    
    If p_Tramitador = "" Or p_Tramitador = "Todos" Then
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes ORDER BY TbSolicitudes.IDSolicitud DESC;"
    Else
        m_SQL = "SELECT TbSolicitudes.* " & _
                "FROM TbSolicitudes " & _
                "WHERE " & _
                "Gestor='" & p_Tramitador & "' ORDER BY TbSolicitudes.IDSolicitud DESC;"
    End If

    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               'If InStr(1, m_Solicitud.Nombre, "José Antonio") <> 0 Then Stop
               If getSolicitudes Is Nothing Then
                   Set getSolicitudes = New Scripting.Dictionary
                   getSolicitudes.CompareMode = TextCompare
               End If
               If Not getSolicitudes.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudes.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudes ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getMotivos( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Motivo As MotivoHPS
    
    
    On Error GoTo errores
    m_SQL = "SELECT * " & _
                "FROM TbMotivoHPS;"
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
              Set m_Motivo = New MotivoHPS
               For Each m_Campo In m_Motivo.ColCampos
                  m_Motivo.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getMotivos Is Nothing Then
                   Set getMotivos = New Scripting.Dictionary
                   getMotivos.CompareMode = TextCompare
               End If
               If Not getMotivos.Exists(m_Motivo.MotivoHPS) Then
                   getMotivos.Add m_Motivo.MotivoHPS, m_Motivo
               End If
               Set m_Motivo = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMotivos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getMotivo( _
                            p_Motivo As String, _
                            Optional ByRef p_Error As String _
                            ) As MotivoHPS
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                "FROM TbMotivoHPS " & _
                "WHERE MotivoHPS='" & p_Motivo & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getMotivo = New MotivoHPS
        For Each m_Campo In getMotivo.ColCampos
            getMotivo.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMotivo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosConMotivo( _
                                    p_Motivo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuarios " & _
            "WHERE Motivo_HPS='" & p_Motivo & "';"
    
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_UsuarioHPS = New UsuarioHPS
               For Each m_Campo In m_UsuarioHPS.ColCampos
                  m_UsuarioHPS.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               'If InStr(1, m_Solicitud.Nombre, "José Antonio") <> 0 Then Stop
               If getUsuariosConMotivo Is Nothing Then
                   Set getUsuariosConMotivo = New Scripting.Dictionary
                   getUsuariosConMotivo.CompareMode = TextCompare
               End If
               If Not getUsuariosConMotivo.Exists(m_UsuarioHPS.id) Then
                   getUsuariosConMotivo.Add m_UsuarioHPS.id, m_UsuarioHPS
               End If
               Set m_UsuarioHPS = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosConMotivo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSolicitudesConMotivo( _
                                        p_Motivo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Solicitud As solicitud
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudes " & _
            "WHERE Motivo_HPS='" & p_Motivo & "';"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               'If InStr(1, m_Solicitud.Nombre, "José Antonio") <> 0 Then Stop
               If getSolicitudesConMotivo Is Nothing Then
                   Set getSolicitudesConMotivo = New Scripting.Dictionary
                   getSolicitudesConMotivo.CompareMode = TextCompare
               End If
               If Not getSolicitudesConMotivo.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesConMotivo.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesConMotivo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getExpediente( _
                                    p_IDExpediente As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Expediente
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_IDExpediente = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbExpedientes " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getExpediente = New Expediente
        For Each m_Campo In getExpediente.ColCampos
            getExpediente.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpediente ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getSuministrador( _
                                Optional p_ID As String, _
                                Optional p_Nombre As String, _
                                Optional ByRef p_Error As String _
                                ) As Suministrador

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_ID = "" And p_Nombre = "" Then
        Exit Function
    End If
    If p_ID <> "" Then
        m_SQL = "SELECT * FROM TbSuministradores " & _
                "WHERE IDSuministrador=" & p_ID & ";"
    Else
        m_SQL = "SELECT * FROM TbSuministradores " & _
                "WHERE Nombre Like'*" & p_Nombre & "*';"
    End If
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getSuministrador = New Suministrador
        For Each m_Campo In getSuministrador.ColCampos
            getSuministrador.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSuministrador ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getSuministradores( _
                                Optional p_SoloTramitadorasHPS As EnumSiNo = EnumSiNo.no, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Suministrador As Suministrador
    
    On Error GoTo errores
    If p_SoloTramitadorasHPS = Empty Then
        p_SoloTramitadorasHPS = EnumSiNo.no
    End If
    
    If p_SoloTramitadorasHPS = EnumSiNo.no Then
        m_SQL = "SELECT * FROM TbSuministradores ORDER BY Nombre;"
    Else
        m_SQL = "SELECT * FROM TbSuministradores " & _
                "WHERE TramitadoraHPS='Sí' ORDER BY Nombre;"
    End If
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Suministrador = New Suministrador
            For Each m_Campo In m_Suministrador.ColCampos
                'Debug.Print m_Campo
                'If CStr(m_Campo) = "Nemotecnico" Then Stop
                m_Suministrador.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getSuministradores Is Nothing Then
                Set getSuministradores = New Scripting.Dictionary
                getSuministradores.CompareMode = TextCompare
            End If
            If Not getSuministradores.Exists(m_Suministrador.IDSuministrador) Then
                getSuministradores.Add m_Suministrador.IDSuministrador, m_Suministrador
            End If
            Set m_Suministrador = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSuministradores ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getCorreosPorEstados( _
                                    Optional p_Estado As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Correo As Correo
    
    On Error GoTo errores
    
    If p_Estado = "" Or p_Estado = "Todos" Then
        m_SQL = "TbCorreosEnviados"
    ElseIf p_Estado = "Enviados" Then
        m_SQL = "SELECT * " & _
                "FROM TbCorreosEnviados " & _
                "WHERE not FechaEnvio Is Null;"
    ElseIf p_Estado = "Por Enviar" Then
         m_SQL = "SELECT * " & _
                "FROM TbCorreosEnviados " & _
                "WHERE FechaEnvio Is Null;"
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Correo = New Correo
               For Each m_Campo In m_Correo.ColCampos
                  m_Correo.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getCorreosPorEstados Is Nothing Then
                   Set getCorreosPorEstados = New Scripting.Dictionary
                   getCorreosPorEstados.CompareMode = TextCompare
               End If
               If Not getCorreosPorEstados.Exists(m_Correo.IDCORREO) Then
                   getCorreosPorEstados.Add m_Correo.IDCORREO, m_Correo
               End If
               Set m_Correo = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreosPorEstados ha devuelto el error: " & Err.Description
    End If
End Function



Public Function getExpedientes( _
                                Optional p_Nombre As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Expediente As Expediente
    
    On Error GoTo errores
    
    
    If p_Nombre = "" Then
        m_SQL = "SELECT * " & _
                "FROM TbExpedientes;"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbExpedientes " & _
                "WHERE Nemotecnico Like '*" & p_Nombre & "*' " & _
                "OR CodExp Like '*" & p_Nombre & "*' " & _
                "OR Titulo Like '*" & p_Nombre & "*';"
       
    End If
    
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Expediente = New Expediente
               For Each m_Campo In m_Expediente.ColCampos
                  m_Expediente.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getExpedientes Is Nothing Then
                   Set getExpedientes = New Scripting.Dictionary
                   getExpedientes.CompareMode = TextCompare
               End If
               If Not getExpedientes.Exists(m_Expediente.IDExpediente) Then
                   getExpedientes.Add m_Expediente.IDExpediente, m_Expediente
               End If
               Set m_Expediente = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedientes ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getResponsables( _
                                Optional p_Nombre As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Responsables As Responsable
    
    On Error GoTo errores
    If p_Nombre = "" Then
        m_SQL = "TbResponsables"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbResponsables " & _
                "WHERE " & _
                "Nombre Like '*" & p_Nombre & "*';"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Responsables = New Responsable
               For Each m_Campo In m_Responsables.ColCampos
                  m_Responsables.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getResponsables Is Nothing Then
                   Set getResponsables = New Scripting.Dictionary
                   getResponsables.CompareMode = TextCompare
               End If
               If Not getResponsables.Exists(m_Responsables.IDResponsable) Then
                   getResponsables.Add m_Responsables.IDResponsable, m_Responsables
               End If
               Set m_Responsables = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getResponsables ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesConResponsable( _
                                            p_Correo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    On Error GoTo errores
    
    If p_Correo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbSolicitudes.* " & _
            "FROM TbSolicitudes INNER JOIN TbResponsables ON TbSolicitudes.emailResponsable = TbResponsables.Correo " & _
            "WHERE Correo='" & p_Correo & "';"
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesConResponsable Is Nothing Then
                   Set getSolicitudesConResponsable = New Scripting.Dictionary
                   getSolicitudesConResponsable.CompareMode = TextCompare
               End If
               If Not getSolicitudesConResponsable.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesConResponsable.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesConResponsable ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getLogGeneral( _
                                p_IDLog As String, _
                                Optional ByRef p_Error As String _
                                ) As LogGeneral
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_IDLog = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbLogsGeneral " & _
            "WHERE IDLogGeneral=" & p_IDLog & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getLogGeneral = New LogGeneral
        For Each m_Campo In getLogGeneral.ColCampos
            getLogGeneral.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getLogGeneral ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getLogsGenerales( _
                                    Optional p_IDSolicitud As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Log As LogGeneral
    
    
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbLogsGeneral " & _
            "WHERE IDSolicitud=" & p_IDSolicitud & ";"
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
               If getLogsGenerales Is Nothing Then
                   Set getLogsGenerales = New Scripting.Dictionary
                   getLogsGenerales.CompareMode = TextCompare
               End If
               If Not getLogsGenerales.Exists(m_Log.IDLog) Then
                   getLogsGenerales.Add m_Log.IDLog, m_Log
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
        p_Error = "El método getLogsGenerales ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getCorreo( _
                        p_IDCorreo As String, _
                        Optional ByRef p_Error As String _
                        ) As Correo
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_IDCorreo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbCorreosEnviados " & _
            "WHERE IDCorreo=" & p_IDCorreo & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getCorreo = New Correo
        For Each m_Campo In getCorreo.ColCampos
            getCorreo.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getCorreos( _
                            p_IDSolicitud As String, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Correo As Correo
    
    On Error GoTo errores
    If p_IDSolicitud = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbCorreosEnviados " & _
            "WHERE IDSolicitud=" & p_IDSolicitud & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Correo = New Correo
               For Each m_Campo In m_Correo.ColCampos
                  m_Correo.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getCorreos Is Nothing Then
                   Set getCorreos = New Scripting.Dictionary
                   getCorreos.CompareMode = TextCompare
               End If
               If Not getCorreos.Exists(m_Correo.IDCORREO) Then
                   getCorreos.Add m_Correo.IDCORREO, m_Correo
               End If
               Set m_Correo = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getListaColeccionesGradosHPS( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_TipoHPS As String
    Dim m_Grado As String
    Dim m_ColGradosNAC As Scripting.Dictionary
    Dim m_ColGradosOTAN As Scripting.Dictionary
    Dim m_ColGradosUE As Scripting.Dictionary
    Dim m_ColGradosESA As Scripting.Dictionary
    
    On Error GoTo errores
    p_Error = ""
    
    
    m_SQL = "SELECT * " & _
            "FROM TbHPSGrado;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_TipoHPS = Nz(.Fields("TipoHPS"), "")
                m_Grado = Nz(.Fields("Grado"), "")
                If m_TipoHPS = "Nacional" Then
                    If m_ColGradosNAC Is Nothing Then
                        Set m_ColGradosNAC = New Scripting.Dictionary
                        m_ColGradosNAC.CompareMode = TextCompare
                    End If
                    m_ColGradosNAC.Add m_Grado, m_Grado
                ElseIf m_TipoHPS = "OTAN" Then
                    If m_ColGradosOTAN Is Nothing Then
                        Set m_ColGradosOTAN = New Scripting.Dictionary
                        m_ColGradosOTAN.CompareMode = TextCompare
                    End If
                    If Not m_ColGradosOTAN.Exists(m_Grado) Then
                        m_ColGradosOTAN.Add m_Grado, m_Grado
                    End If
                    
                ElseIf m_TipoHPS = "UE" Then
                    If m_ColGradosUE Is Nothing Then
                        Set m_ColGradosUE = New Scripting.Dictionary
                        m_ColGradosUE.CompareMode = TextCompare
                    End If
                    If Not m_ColGradosUE.Exists(m_Grado) Then
                        m_ColGradosUE.Add m_Grado, m_Grado
                    End If
                ElseIf m_TipoHPS = "ESA" Then
                    If m_ColGradosESA Is Nothing Then
                        Set m_ColGradosESA = New Scripting.Dictionary
                        m_ColGradosESA.CompareMode = TextCompare
                    End If

                   If Not m_ColGradosESA.Exists(m_Grado) Then
                        m_ColGradosESA.Add m_Grado, m_Grado
                    End If
                Else
                    p_Error = "Tipo HPS desconocido"
                    Err.Raise 1000
                End If

                .MoveNext
            Loop

        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not m_ColGradosNAC Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("Nacional") Then
            getListaColeccionesGradosHPS.Add "Nacional", m_ColGradosNAC
        End If
        
    End If
    If Not m_ColGradosOTAN Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("OTAN") Then
            getListaColeccionesGradosHPS.Add "OTAN", m_ColGradosOTAN
        End If
        
    End If
    If Not m_ColGradosUE Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("UE") Then
            getListaColeccionesGradosHPS.Add "UE", m_ColGradosUE
        End If
        
    End If
    If Not m_ColGradosESA Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("ESA") Then
            getListaColeccionesGradosHPS.Add "ESA", m_ColGradosESA
        End If
        
    End If
    

    Exit Function
errores:

    If Err.Number <> 1000 Then
        p_Error = "El método getListaColeccionesGradosHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Function EsDNIValido(ByVal Documento As String) As Boolean
    Dim LetrasDNI As String
    Dim LetrasNIE As String
    Dim Numeros As String
    Dim Letra As String
    Dim Resto As Integer
    Dim LetraCalculada As String
    Dim PrimerCaracter As String
    
    On Error GoTo errores
    ' Inicializamos
    LetrasDNI = "TRWAGMYFPDXBNJZSQVHLCKE"
    LetrasNIE = "XYZ"
    Documento = UCase(Trim(Documento))
    
    ' Validamos longitud mínima
    If Len(Documento) <> 9 Then
        EsDNIValido = False
        Exit Function
    End If
    
    ' Extraemos letra y números
    Letra = Right(Documento, 1)
    Numeros = Left(Documento, Len(Documento) - 1)
    PrimerCaracter = Left(Numeros, 1)
    
    ' Comprobamos si es NIE
    If InStr(LetrasNIE, PrimerCaracter) > 0 Then
        ' Sustituimos X, Y, Z por su valor numérico
        Select Case PrimerCaracter
            Case "X"
                Numeros = "0" & Mid(Numeros, 2)
            Case "Y"
                Numeros = "1" & Mid(Numeros, 2)
            Case "Z"
                Numeros = "2" & Mid(Numeros, 2)
        End Select
    End If
    
    ' Validamos que los números sean realmente numéricos
    If Not IsNumeric(Numeros) Then
        EsDNIValido = False
        Exit Function
    End If
    
    ' Calculamos la letra correspondiente
    Resto = CLng(Numeros) Mod 23
    LetraCalculada = Mid(LetrasDNI, Resto + 1, 1)
    
    ' Comparamos con la letra proporcionada
    If Letra = LetraCalculada Then
        EsDNIValido = True
    Else
        EsDNIValido = False
    End If
    Exit Function
errores:
    Documento = False
End Function

Function EsCIFValido(CIF As String) As Boolean
    Dim LetrasIniciales As String
    Dim LetrasFinales As String
    Dim SumaPar As Integer
    Dim SumaImpar As Integer
    Dim total As Integer
    Dim Resto As Integer
    Dim DigitoControl As String
    Dim i As Integer
    EsCIFValido = True
    Exit Function
    ' Definir las letras iniciales y finales válidas
    LetrasIniciales = "ABCDEFGHJKLMNPQRSUVW"
    LetrasFinales = "JABCDEFGHI"
    
    ' Validar longitud y formato inicial
    If Len(CIF) <> 9 Then
        EsCIFValido = False
        Exit Function
    End If
    
    If InStr(1, LetrasIniciales, Left(CIF, 1)) = 0 Then
        EsCIFValido = False
        Exit Function
    End If
    
    ' Calcular suma de dígitos pares e impares
    SumaPar = 0
    SumaImpar = 0
    For i = 2 To 8 Step 2
        SumaPar = SumaPar + CInt(Mid(CIF, i, 1))
    Next i
    
    For i = 1 To 7 Step 2
        Dim Impar As Integer
        Impar = CInt(Mid(CIF, i + 1, 1)) * 2
        SumaImpar = SumaImpar + (Impar \ 10) + (Impar Mod 10)
    Next i
    
    ' Calcular el dígito de control
    total = SumaPar + SumaImpar
    Resto = total Mod 10
    If Resto <> 0 Then
        Resto = 10 - Resto
    End If
    
    ' Comparar dígito de control
    If IsNumeric(Right(CIF, 1)) Then
        DigitoControl = CStr(Resto)
    Else
        DigitoControl = Mid(LetrasFinales, Resto + 1, 1)
    End If
    
    If Right(CIF, 1) = DigitoControl Then
        EsCIFValido = True
    Else
        EsCIFValido = False
    End If
End Function



Public Function getHPSParaEmpresa( _
                                        p_IDEmpresa As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_JSON As Object
    Dim m_SQL As String
    Dim m_Col As Collection
    Dim i As Long
    Dim m_Valor As String
    Dim m_Nombre As String
    Dim m_DNI As String
    
    Dim m_Cadena As String
    
    On Error GoTo errores
    If p_IDEmpresa = "" Then
        Exit Function
    End If
    Set m_JSON = getJSonDeTabla("TbUsuarios", "IDEmpresaUsuario", p_IDEmpresa, getdbHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_JSON Is Nothing Then
        Set m_Col = m_JSON
        For i = 1 To m_Col.Count
            m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido_2")
            m_Nombre = TextoParsedoParaTxt(m_Valor)
            m_Valor = m_JSON(i)("DNI")
            m_DNI = TextoParsedoParaTxt(m_Valor)
            If m_Cadena = "" Then
                m_Cadena = m_Valor
            Else
                m_Cadena = m_Cadena & vbNewLine & m_Valor
            End If
        Next
    End If
    Set m_JSON = getJSonDeTabla("TbUsuarios", "IDEmpresaHPS", p_IDEmpresa, getdbHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_JSON Is Nothing Then
        Set m_Col = m_JSON
        For i = 1 To m_Col.Count
            m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido_2")
            m_Nombre = TextoParsedoParaTxt(m_Valor)
            m_Valor = m_JSON(i)("DNI")
            m_DNI = TextoParsedoParaTxt(m_Valor)
            If m_Cadena = "" Then
                m_Cadena = m_Valor
            Else
                m_Cadena = m_Cadena & vbNewLine & m_Valor
            End If
        Next
    End If
    Set m_JSON = getJSonDeTabla("TbUsuariosHistoricos", "IDEmpresaUsuario", p_IDEmpresa, getdbHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_JSON Is Nothing Then
        Set m_Col = m_JSON
        For i = 1 To m_Col.Count
            m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido_2")
            m_Nombre = TextoParsedoParaTxt(m_Valor)
            m_Valor = m_JSON(i)("DNI")
            m_DNI = TextoParsedoParaTxt(m_Valor)
            If m_Cadena = "" Then
                m_Cadena = m_Valor
            Else
                m_Cadena = m_Cadena & vbNewLine & m_Valor
            End If
        Next
    End If
    Set m_JSON = getJSonDeTabla("TbUsuariosHistoricos", "IDEmpresaHPS", p_IDEmpresa, getdbHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_JSON Is Nothing Then
        Set m_Col = m_JSON
        For i = 1 To m_Col.Count
            m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido_2")
            m_Nombre = TextoParsedoParaTxt(m_Valor)
            m_Valor = m_JSON(i)("DNI")
            m_DNI = TextoParsedoParaTxt(m_Valor)
            If m_Cadena = "" Then
                m_Cadena = m_Valor
            Else
                m_Cadena = m_Cadena & vbNewLine & m_Valor
            End If
        Next
    End If
    
    
    getHPSParaEmpresa = m_Cadena
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHPSParaEmpresa ha devuelto" & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function getHPSParaExpediente( _
                                        p_IDExpediente As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_JSON As Object
    Dim m_SQL As String
    Dim m_Col As Collection
    Dim i As Long
    Dim m_Valor As String
    Dim m_Nombre As String
    Dim m_DNI As String
    
    Dim m_Cadena As String
    
    On Error GoTo errores
    If p_IDExpediente = "" Then
        Exit Function
    End If
    Set m_JSON = getJSonDeTabla("TbUsuarios", "IDExpediente", p_IDExpediente, getdbHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_JSON Is Nothing Then
        Set m_Col = m_JSON
        For i = 1 To m_Col.Count
            m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido_2")
            m_Nombre = TextoParsedoParaTxt(m_Valor)
            m_Valor = m_JSON(i)("DNI")
            m_DNI = TextoParsedoParaTxt(m_Valor)
            If m_Cadena = "" Then
                m_Cadena = m_Valor
            Else
                m_Cadena = m_Cadena & vbNewLine & m_Valor
            End If
        Next
    End If
    
    Set m_JSON = getJSonDeTabla("TbUsuariosHistoricos", "IDExpediente", p_IDExpediente, getdbHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_JSON Is Nothing Then
        Set m_Col = m_JSON
        For i = 1 To m_Col.Count
            m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido_2")
            m_Nombre = TextoParsedoParaTxt(m_Valor)
            m_Valor = m_JSON(i)("DNI")
            m_DNI = TextoParsedoParaTxt(m_Valor)
            If m_Cadena = "" Then
                m_Cadena = m_Valor
            Else
                m_Cadena = m_Cadena & vbNewLine & m_Valor
            End If
        Next
    End If
    
    
    getHPSParaExpediente = m_Cadena
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHPSParaExpediente ha devuelto" & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function getSolicitudesEnEmpresa( _
                                            p_IDEmpresa As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudes " & _
            "WHERE " & _
            "IDEmpresaUsuario=" & p_IDEmpresa & " or " & _
            "IDEmpresaTramitadora=" & p_IDEmpresa & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesEnEmpresa Is Nothing Then
                   Set getSolicitudesEnEmpresa = New Scripting.Dictionary
                   getSolicitudesEnEmpresa.CompareMode = TextCompare
               End If
               If Not getSolicitudesEnEmpresa.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesEnEmpresa.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesEnEmpresa ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSolicitudesEnExpedientes( _
                                            p_IDExpediente As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Solicitud As solicitud
    
    
    
    On Error GoTo errores
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudes " & _
            "WHERE " & _
            "IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Solicitud = New solicitud
               For Each m_Campo In m_Solicitud.ColCampos
                  m_Solicitud.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getSolicitudesEnExpedientes Is Nothing Then
                   Set getSolicitudesEnExpedientes = New Scripting.Dictionary
                   getSolicitudesEnExpedientes.CompareMode = TextCompare
               End If
               If Not getSolicitudesEnExpedientes.Exists(m_Solicitud.IDSolicitud) Then
                   getSolicitudesEnExpedientes.Add m_Solicitud.IDSolicitud, m_Solicitud
               End If
               Set m_Solicitud = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesEnExpedientes ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUltimoCambio( _
                                Optional p_IDUsuario As String, _
                                Optional ByRef p_Error As String _
                                ) As UltimoCambio
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    
    On Error GoTo errores
    If p_IDUsuario <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUltimoCambio " & _
                "WHERE IDUsuarioCambio = " & p_IDUsuario & " " & _
                "ORDER BY FechaCambio DESC;"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbUltimoCambio " & _
                "ORDER BY FechaCambio DESC;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getUltimoCambio = New UltimoCambio
        For Each m_Campo In getUltimoCambio.ColCampos
            getUltimoCambio.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUltimoCambio ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getResponsable( _
                                Optional p_IDResponsable As String, _
                                Optional p_Correo As String, _
                                Optional ByRef p_Error As String _
                                ) As Responsable
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_IDResponsable = "" And p_Correo = "" Then
        Exit Function
    End If
    If p_IDResponsable <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbResponsables " & _
                "WHERE IDResponsable=" & p_IDResponsable & ";"
    ElseIf p_Correo <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbResponsables " & _
                "WHERE Correo='" & p_Correo & "';"
   
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getResponsable = New Responsable
        For Each m_Campo In getResponsable.ColCampos
            getResponsable.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getResponsable ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUsuariosPendientesSolicitud( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
                    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_UsuarioHPS As UsuarioEntidad
    
    On Error GoTo errores
    
    
    
    m_SQL = "TbUsuarios"
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
     With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_UsuarioHPS = New UsuarioHPS
               For Each m_Campo In m_UsuarioHPS.ColCampos
                  m_UsuarioHPS.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getUsuariosPendientesSolicitud Is Nothing Then
                   Set getUsuariosPendientesSolicitud = New Scripting.Dictionary
                   getUsuariosPendientesSolicitud.CompareMode = TextCompare
               End If
               If Not getUsuariosPendientesSolicitud.Exists(m_UsuarioHPS.id) Then
                   getUsuariosPendientesSolicitud.Add m_UsuarioHPS.id, m_UsuarioHPS
               End If
               Set m_UsuarioHPS = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosPendientesSolicitud ha devuelto el error: " & Err.Description
    End If
End Function



Public Function getUsuariosHPS( _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    
    On Error GoTo errores
    
    
    
    m_SQL = "TbUsuarios"
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
     With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_UsuarioHPS = New UsuarioHPS
               For Each m_Campo In m_UsuarioHPS.ColCampos
                  m_UsuarioHPS.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getUsuariosHPS Is Nothing Then
                   Set getUsuariosHPS = New Scripting.Dictionary
                   getUsuariosHPS.CompareMode = TextCompare
               End If
               If Not getUsuariosHPS.Exists(m_UsuarioHPS.id) Then
                   getUsuariosHPS.Add m_UsuarioHPS.id, m_UsuarioHPS
               End If
               Set m_UsuarioHPS = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPS ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getUsuarioHPS( _
                                Optional p_IDUsuario As String, _
                                Optional p_DNI As String, _
                                Optional p_Nombre As String, _
                                Optional ByRef p_Error As String _
                                ) As UsuarioHPS
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_IDUsuario = "" And p_DNI = "" Then
        Exit Function
    End If
    If p_IDUsuario <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuarios " & _
                "WHERE ID=" & p_IDUsuario & ";"
    ElseIf p_DNI <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuarios " & _
                "WHERE DNI='" & p_DNI & "';"
    End If
    
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuarioHPS = New UsuarioHPS
        For Each m_Campo In getUsuarioHPS.ColCampos
            getUsuarioHPS.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
             If p_Error <> "" Then
                 Err.Raise 1000
             End If
         Next
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioHPS ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUsuariosHPSPorPalabraClave( _
                                                p_PalabraClave As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    Dim m_ColPorNombre As Scripting.Dictionary
    Dim m_ColPorDNI As Scripting.Dictionary
    
    On Error GoTo errores
    If p_PalabraClave = "" Then
        Set getUsuariosHPSPorPalabraClave = m_ObjEntorno.colUsuariosHPS
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    Set m_ColPorNombre = getUsuariosHPSPorNombre(p_Nombre:=p_PalabraClave, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColPorDNI = getUsuariosHPSPorDNI(p_DNI:=p_PalabraClave, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColPorNombre Is Nothing Then
        For Each m_ID In m_ColPorNombre
            Set m_UsuarioHPS = m_ColPorNombre(m_ID)
            If getUsuariosHPSPorPalabraClave Is Nothing Then
                Set getUsuariosHPSPorPalabraClave = New Scripting.Dictionary
                getUsuariosHPSPorPalabraClave.CompareMode = TextCompare
            End If
            If Not getUsuariosHPSPorPalabraClave.Exists(m_UsuarioHPS.id) Then
                getUsuariosHPSPorPalabraClave.Add m_UsuarioHPS.id, m_UsuarioHPS
            End If
            Set m_UsuarioHPS = Nothing
        Next
    End If
    If Not m_ColPorDNI Is Nothing Then
        For Each m_ID In m_ColPorDNI
            Set m_UsuarioHPS = m_ColPorDNI(m_ID)
            If getUsuariosHPSPorPalabraClave Is Nothing Then
                Set getUsuariosHPSPorPalabraClave = New Scripting.Dictionary
                getUsuariosHPSPorPalabraClave.CompareMode = TextCompare
            End If
            If Not getUsuariosHPSPorPalabraClave.Exists(m_UsuarioHPS.id) Then
                getUsuariosHPSPorPalabraClave.Add m_UsuarioHPS.id, m_UsuarioHPS
            End If
            Set m_UsuarioHPS = Nothing
        Next
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPSPorPalabraClave ha devuelto el error: " & Err.Description
    End If
End Function
'
'Public Function getUsuariosHPS( _
'                                Optional ByRef p_Error As String _
'                                ) As Scripting.Dictionary
'
'
'    Dim rcdDatos As DAO.Recordset
'    Dim m_SQL As String
'    Dim m_Campo As Variant
'    Dim m_UsuarioHPS As UsuarioHPS
'
'    On Error GoTo errores
'
'
'
'    m_SQL = "TbUsuarios"
'    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
'     With rcdDatos
'         If Not .EOF Then
'             .MoveFirst
'            Do While Not .EOF
'               Set m_UsuarioHPS = New UsuarioHPS
'               For Each m_Campo In m_UsuarioHPS.ColCampos
'                  m_UsuarioHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
'                   If p_Error <> "" Then
'                       Err.Raise 1000
'                   End If
'               Next
'               If getUsuariosHPS Is Nothing Then
'                   Set getUsuariosHPS = New Scripting.Dictionary
'                   getUsuariosHPS.CompareMode = TextCompare
'               End If
'               If Not getUsuariosHPS.Exists(m_UsuarioHPS.ID) Then
'                   getUsuariosHPS.Add m_UsuarioHPS.ID, m_UsuarioHPS
'               End If
'               Set m_UsuarioHPS = Nothing
'               .MoveNext
'            Loop
'         End If
'    End With
'    rcdDatos.Close
'    Set rcdDatos = Nothing
'    Exit Function
'
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El método getUsuariosHPS ha devuelto el error: " & Err.Description
'    End If
'End Function

Public Function getUsuariosHPSPorNombre( _
                                        p_Nombre As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    
    On Error GoTo errores
    If p_Nombre = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuarios " & _
            "WHERE Nombre Like '*" & p_Nombre & "*' " & _
            "OR Apellido_1 Like '*" & p_Nombre & "*' " & _
            "OR Apellido_2 Like '*" & p_Nombre & "*';"
    
    
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
     With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_UsuarioHPS = New UsuarioHPS
               For Each m_Campo In m_UsuarioHPS.ColCampos
                  m_UsuarioHPS.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getUsuariosHPSPorNombre Is Nothing Then
                   Set getUsuariosHPSPorNombre = New Scripting.Dictionary
                   getUsuariosHPSPorNombre.CompareMode = TextCompare
               End If
               If Not getUsuariosHPSPorNombre.Exists(m_UsuarioHPS.id) Then
                   getUsuariosHPSPorNombre.Add m_UsuarioHPS.id, m_UsuarioHPS
               End If
               Set m_UsuarioHPS = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPSPorNombre ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUsuariosHPSPorDNI( _
                                        p_DNI As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    
    On Error GoTo errores
    If p_DNI = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuarios " & _
            "WHERE DNI Like '*" & p_DNI & "*';"
    
    
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
     With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_UsuarioHPS = New UsuarioHPS
               For Each m_Campo In m_UsuarioHPS.ColCampos
                  m_UsuarioHPS.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getUsuariosHPSPorDNI Is Nothing Then
                   Set getUsuariosHPSPorDNI = New Scripting.Dictionary
                   getUsuariosHPSPorDNI.CompareMode = TextCompare
               End If
               If Not getUsuariosHPSPorDNI.Exists(m_UsuarioHPS.id) Then
                   getUsuariosHPSPorDNI.Add m_UsuarioHPS.id, m_UsuarioHPS
               End If
               Set m_UsuarioHPS = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPSPorDNI ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getExpedientesBusqueda( _
                                        Optional p_PalabraClave As String, _
                                        Optional p_IDResponsableCalidad As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Expediente As Expediente
    
    On Error GoTo errores
    If p_IDResponsableCalidad <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbExpedientes " & _
                "WHERE IDResponsableCalidad =" & p_IDResponsableCalidad & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbExpedientes;"
    End If
       
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Expediente = New Expediente
                For Each m_Campo In m_Expediente.ColCampos
                    m_Expediente.setPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If p_PalabraClave <> "" Then
                    If p_PalabraClave <> m_Expediente.IDExpediente And _
                        InStr(1, m_Expediente.Nemotecnico, p_PalabraClave) = 0 And _
                        InStr(1, m_Expediente.Titulo, p_PalabraClave) = 0 And _
                        InStr(1, m_Expediente.CodExp, p_PalabraClave) = 0 And _
                        InStr(1, m_Expediente.CodExpLargo, p_PalabraClave) = 0 Then
                        GoTo siguiente
                    End If
                End If

                If getExpedientesBusqueda Is Nothing Then
                    Set getExpedientesBusqueda = New Scripting.Dictionary
                    getExpedientesBusqueda.CompareMode = TextCompare
                End If
                If Not getExpedientesBusqueda.Exists(CStr(m_Expediente.IDExpediente)) Then
                    getExpedientesBusqueda.Add CStr(m_Expediente.IDExpediente), m_Expediente
                End If
siguiente:
                .MoveNext
            Loop
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedientesBusqueda ha devuelto el error: " & Err.Description
    End If
End Function

