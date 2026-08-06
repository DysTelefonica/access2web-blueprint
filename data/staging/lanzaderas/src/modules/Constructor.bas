Attribute VB_Name = "Constructor"
Option Compare Database
Option Explicit

Public Function getUsuarioConectadoPorMaquina( _
                                                Optional ByRef p_Error As String _
                                                ) As usuario
    Dim objNetwork As Object
    On Error GoTo errores
    Set objNetwork = CreateObject("Wscript.Network")
    Set getUsuarioConectadoPorMaquina = Constructor.getUsuario(, objNetwork.UserName, , , p_Error)
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


Public Function getUsuario( _
                            Optional p_ID As String, _
                            Optional p_UsuarioRed As String, _
                            Optional p_Nombre As String, _
                            Optional p_Correo As String, _
                            Optional ByRef p_Error As String _
                            ) As usuario

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
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuario = New usuario
        For Each m_Campo In getUsuario.ColCampos
            getUsuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo), ""), p_Error
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

Public Function getAplicaciones( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_objAplicacion As Aplicacion
    Dim m_ID As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    p_Error = ""
    m_SQL = "TbAplicaciones"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objAplicacion = New Aplicacion
            For Each m_Campo In m_objAplicacion.ColCampos
                m_objAplicacion.SetPropiedad m_Campo, Nz(.Fields(m_Campo), ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getAplicaciones Is Nothing Then
                Set getAplicaciones = New Scripting.Dictionary
                getAplicaciones.CompareMode = TextCompare
            End If
           ' If m_ObjAplicacion.IDAplicacion = "51" Then Stop
            If Not getAplicaciones.Exists(CStr(m_objAplicacion.IDAplicacion)) Then
                getAplicaciones.Add CStr(m_objAplicacion.IDAplicacion), m_objAplicacion
            End If
            Set m_objAplicacion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAplicaciones ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAperturas( _
                                Optional p_Fecha As String, _
                                Optional p_IDAplicacion As String, _
                                Optional p_UsuarioRed As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Apertura As AplicacionApertura
    Dim m_ID As String
    Dim m_Campo As Variant
    Dim m_Where As String
    Dim m_WhereFecha As String
    Dim m_WhereAplicacion As String
    Dim m_WhereUsuario As String
    Dim m_SQLAlInicio As String
    
    On Error GoTo errores
    p_Error = ""
    If p_Fecha = "" And p_Fecha = "" And p_Fecha = "" Then
        Exit Function
    End If
    If IsDate(p_Fecha) Then
        m_WhereFecha = "FechaApertura Between #" & Format(p_Fecha, "mm/dd/yyyy") & "# And #" & Format(p_Fecha, "mm/dd/yyyy") & "#"
    Else
        m_WhereFecha = "FechaApertura Between #01/01/1900# And #01/01/2100#"
    End If
    If p_IDAplicacion <> "" Then
        m_WhereAplicacion = "IDAplicacion=" & p_IDAplicacion
    Else
        m_WhereAplicacion = "IDAplicacion Like '*'"
    End If
    If p_UsuarioRed <> "" Then
        m_WhereUsuario = "UsuarioRed='" & p_UsuarioRed & "'"
    Else
        m_WhereUsuario = "UsuarioRed Like '*'"
    End If
    m_SQLAlInicio = "SELECT TbAplicacionesAperturas.* " & _
                    "FROM TbAplicacionesAperturas LEFT JOIN TbUsuariosAplicaciones " & _
                    "ON TbAplicacionesAperturas.NombreUsuario = TbUsuariosAplicaciones.Nombre "
    m_Where = "WHERE " & _
            m_WhereFecha & " " & _
            "AND " & m_WhereAplicacion & " " & _
            "AND " & m_WhereUsuario & ";"
    m_SQL = m_SQLAlInicio & m_Where
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Apertura = New AplicacionApertura
            For Each m_Campo In m_Apertura.ColCampos
                m_Apertura.SetPropiedad m_Campo, Nz(.Fields(m_Campo), ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAperturas Is Nothing Then
                Set getAperturas = New Scripting.Dictionary
                getAperturas.CompareMode = TextCompare
            End If
           ' If m_Apertura.IDAplicacion = "51" Then Stop
            If Not getAperturas.Exists(CStr(m_Apertura.IDApertura)) Then
                getAperturas.Add CStr(m_Apertura.IDApertura), m_Apertura
            End If
            Set m_Apertura = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAperturas ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAperturasPorTipo( _
                                    p_Tipo As EnumApertura, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Apertura As AplicacionApertura
    Dim m_ID As String
    Dim m_Campo As Variant
    
    
    On Error GoTo errores
    p_Error = ""
    
    If p_Tipo = EnumApertura.Todas Then
        m_SQL = "TbAplicacionesAperturas"
    ElseIf p_Tipo = EnumApertura.HoyTodas Then
        m_SQL = "SELECT * " & _
                    "FROM TbAplicacionesAperturas  " & _
                    "WHERE FechaApertura=#" & Format(Date, "mm/dd/yyyy") & "#;"
    ElseIf p_Tipo = EnumApertura.TodasAbiertas Then
        m_SQL = "SELECT * " & _
                    "FROM TbAplicacionesAperturas  " & _
                    "WHERE FechaCierre Is Null;"
    ElseIf p_Tipo = EnumApertura.HoyAbiertas Then
        m_SQL = "SELECT * " & _
                    "FROM TbAplicacionesAperturas  " & _
                    "WHERE FechaApertura=#" & Format(Date, "mm/dd/yyyy") & "# " & _
                    "AND FechaCierre Is Null;"
    ElseIf p_Tipo = EnumApertura.HoyCerradas Then
        m_SQL = "SELECT * " & _
                    "FROM TbAplicacionesAperturas  " & _
                    "WHERE FechaApertura=#" & Format(Date, "mm/dd/yyyy") & "# " & _
                    "AND Not FechaCierre Is Null;"
    Else
        Exit Function
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Apertura = New AplicacionApertura
            For Each m_Campo In m_Apertura.ColCampos
                m_Apertura.SetPropiedad m_Campo, Nz(.Fields(m_Campo), ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAperturasPorTipo Is Nothing Then
                Set getAperturasPorTipo = New Scripting.Dictionary
                getAperturasPorTipo.CompareMode = TextCompare
            End If
           ' If m_Apertura.IDAplicacion = "51" Then Stop
            If Not getAperturasPorTipo.Exists(CStr(m_Apertura.IDApertura)) Then
                getAperturasPorTipo.Add CStr(m_Apertura.IDApertura), m_Apertura
            End If
            Set m_Apertura = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAperturasPorTipo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosAplicacionAbierta( _
                                            p_IDAplicacion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Apertura As AplicacionApertura
    Dim m_ID As String
    Dim m_Usuario As usuario
    Dim m_Campo As Variant
    
    
    On Error GoTo errores
    p_Error = ""
    
    If p_IDAplicacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT DISTINCT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones INNER JOIN TbAplicacionesAperturas " & _
            "ON TbUsuariosAplicaciones.Nombre = TbAplicacionesAperturas.NombreUsuario " & _
            "WHERE (((TbAplicacionesAperturas.IDAplicacion)=" & p_IDAplicacion & ") " & _
            "AND ((TbAplicacionesAperturas.FechaApertura)=Date()) " & _
            "AND ((TbAplicacionesAperturas.FechaCierre) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Usuario = New usuario
            For Each m_Campo In m_Usuario.ColCampos
                m_Usuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo), ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getUsuariosAplicacionAbierta Is Nothing Then
                Set getUsuariosAplicacionAbierta = New Scripting.Dictionary
                getUsuariosAplicacionAbierta.CompareMode = TextCompare
            End If
           ' If m_Apertura.IDAplicacion = "51" Then Stop
            If Not getUsuariosAplicacionAbierta.Exists(CStr(m_Usuario.ID)) Then
                getUsuariosAplicacionAbierta.Add CStr(m_Usuario.ID), m_Usuario
            End If
            Set m_Usuario = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosAplicacionAbierta ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAplicacionesActivas( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_objAplicacion As Aplicacion
    Dim m_ID As String
    Dim fld As DAO.Field
    
    On Error GoTo errores
    p_Error = ""
    m_SQL = "SELECT * " & _
            "FROM TbAplicaciones ;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objAplicacion = New Aplicacion
            For Each fld In rcdDatos.Fields
                m_objAplicacion.SetPropiedad fld.Name, Nz(fld.Value, "")
            Next
            If getAplicacionesActivas Is Nothing Then
                Set getAplicacionesActivas = New Scripting.Dictionary
                getAplicacionesActivas.CompareMode = TextCompare
            End If
           'If m_objAplicacion.IDAplicacion = "51" Then Stop
            If Not getAplicacionesActivas.Exists(CStr(m_objAplicacion.IDAplicacion)) Then
                getAplicacionesActivas.Add CStr(m_objAplicacion.IDAplicacion), m_objAplicacion
            End If
            Set m_objAplicacion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAplicacionesActivas ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAplicacionesVideosAyuda( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    
    Dim m_objVideo As Video
    Dim m_ID As Variant
    Dim m_objAplicacion As Aplicacion
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    p_Error = ""
    If m_ObjEntorno.ColAplicaciones Is Nothing Then
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(m_ID)
        Set m_Col = m_objAplicacion.ColVideos
        p_Error = m_objAplicacion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_Col Is Nothing Then
            If getAplicacionesVideosAyuda Is Nothing Then
                Set getAplicacionesVideosAyuda = New Scripting.Dictionary
                getAplicacionesVideosAyuda.CompareMode = TextCompare
            End If
            If Not getAplicacionesVideosAyuda.Exists(CStr(m_ID)) Then
                getAplicacionesVideosAyuda.Add CStr(m_ID), m_Col
            End If
            Set m_Col = Nothing
        End If
        Set m_objAplicacion = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAplicacionesVideosAyuda ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAplicacion( _
                                p_IDAplicacion As String, _
                                Optional ByRef p_Error As String _
                                ) As Aplicacion

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
        
    
    On Error GoTo errores
    If p_IDAplicacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbAplicaciones " & _
            "WHERE IDAplicacion=" & p_IDAplicacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
       Set getAplicacion = New Aplicacion
        For Each fld In rcdDatos.Fields
            'If fld.Name = "EjecucionEnOficina" Then Stop
            getAplicacion.SetPropiedad fld.Name, Nz(fld.Value, "")
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAplicacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAplicacionApertura( _
                                        p_IDApertura As String, _
                                        Optional ByRef p_Error As String _
                                        ) As AplicacionApertura
    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
        
    
    On Error GoTo errores
    If p_IDApertura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbAplicacionesAperturas " & _
            "WHERE IDApertura=" & p_IDApertura & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
       Set getAplicacionApertura = New AplicacionApertura
        For Each fld In rcdDatos.Fields
            'If fld.Name = "EjecucionEnOficina" Then Stop
            getAplicacionApertura.SetPropiedad fld.Name, Nz(fld.Value, ""), p_Error
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
        p_Error = "El método getAplicacionApertura ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getCorreo( _
                            p_IDCorreo As String, _
                            Optional ByRef p_Error As String _
                            ) As Correo

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
        
    On Error GoTo errores
    If p_IDCorreo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbUsuariosCorreosEnvio " & _
            "WHERE IDCorreo=" & p_IDCorreo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getCorreo = New Correo
        For Each fld In rcdDatos.Fields
            getCorreo.SetPropiedad fld.Name, Nz(fld.Value, "")
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



Public Function getConexion( _
                                Optional p_CorreoUsuario As String, _
                                Optional p_IDConexion As String, _
                                Optional ByRef p_Error As String _
                                ) As Conexion

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
        
    On Error GoTo errores
    If p_CorreoUsuario = "" And p_IDConexion = "" Then
        Exit Function
    End If
    If p_IDConexion <> "" Then
        m_SQL = "SELECT TbConexionesRegistro.* " & _
                "FROM TbConexionesRegistro " & _
                "WHERE IDConexion=" & p_IDConexion & ";"
    Else
        m_SQL = "SELECT TbConexionesRegistro.* " & _
            "FROM TbConexionesRegistro " & _
            "WHERE (((TbConexionesRegistro.IDConexion) In " & _
            "(SELECT Last(TbConexionesRegistro.IDConexion) AS ÚltimoDeIDConexion " & _
            "FROM TbConexionesRegistro " & _
            "WHERE (((TbConexionesRegistro.Usuario)='" & p_CorreoUsuario & "')))));"
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getConexion = New Conexion
        For Each fld In rcdDatos.Fields
            getConexion.SetPropiedad fld.Name, Nz(fld.Value, "")
        Next
           
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getConexion ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getConexiones( _
                                p_CorreoUsuario As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    Dim m_ObjConexion As Conexion
    Dim m_ID As String
        
    
    On Error GoTo errores
    If p_CorreoUsuario = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbConexionesRegistro.* " & _
            "FROM TbConexionesRegistro " & _
            "WHERE Usuario='" & p_CorreoUsuario & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjConexion = New Conexion
            For Each fld In rcdDatos.Fields
                m_ObjConexion.SetPropiedad fld.Name, Nz(fld.Value, "")
            Next
            
            If getConexiones Is Nothing Then
                Set getConexiones = New Scripting.Dictionary
                getConexiones.CompareMode = TextCompare
            End If
            If Not getConexiones.Exists(CStr(m_ObjConexion.IDConexion)) Then
                getConexiones.Add CStr(m_ObjConexion.IDConexion), m_ObjConexion
            End If
            Set m_ObjConexion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getConexiones ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUltimaConexion( _
                                    p_Correo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Conexion

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    
    
    On Error GoTo errores
    If p_Correo = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbConexionesRegistro.* " & _
            "FROM TbConexionesRegistro " & _
            "WHERE (((TbConexionesRegistro.Usuario)='" & p_Correo & "')) " & _
            "ORDER BY TbConexionesRegistro.FechaCierre DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getUltimaConexion = New Conexion
        For Each fld In rcdDatos.Fields
            getUltimaConexion.SetPropiedad fld.Name, Nz(fld.Value, "")
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUltimaConexion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUltimoCambioPass( _
                                        p_CorreoUsuario As String, _
                                        Optional ByRef p_Error As String _
                                        ) As HistoricoPass

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    
    On Error GoTo errores
    If p_CorreoUsuario = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbUsuariosHistoricoContrasenias.* " & _
            "FROM TbUsuariosHistoricoContrasenias " & _
            "WHERE (((TbUsuariosHistoricoContrasenias.Usuario)='" & p_CorreoUsuario & "')) " & _
            "ORDER BY FechaPass DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getUltimoCambioPass = New HistoricoPass
        For Each fld In rcdDatos.Fields
            getUltimoCambioPass.SetPropiedad fld.Name, Nz(fld.Value, "")
        Next
    
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUltimoCambioPass ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getCambiosPasss( _
                                p_CorreoUsuario As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    Dim m_ObjHistoricoPass As HistoricoPass
    Dim m_ID As String
        
    
    On Error GoTo errores
    If p_CorreoUsuario = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbUsuariosHistoricoContrasenias.* " & _
            "FROM TbUsuariosHistoricoContrasenias " & _
            "WHERE (((TbUsuariosHistoricoContrasenias.Usuario)='" & p_CorreoUsuario & "')) " & _
            "ORDER BY FechaPass DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjHistoricoPass = New Conexion
            For Each fld In rcdDatos.Fields
                m_ObjHistoricoPass.SetPropiedad fld.Name, Nz(fld.Value, "")
            Next
            
            If getCambiosPasss Is Nothing Then
                Set getCambiosPasss = New Scripting.Dictionary
                getCambiosPasss.CompareMode = TextCompare
            End If
            If Not getCambiosPasss.Exists(m_ObjHistoricoPass.FechaPass) Then
                getCambiosPasss.Add CStr(m_ObjHistoricoPass.FechaPass), m_ObjHistoricoPass
            End If
            Set m_ObjHistoricoPass = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCambiosPasss ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getListaUsuariosPorNombre( _
                                            Optional p_Nombre As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    Dim m_ObjUsuario As usuario
            
    
    On Error GoTo errores
    
    If p_Nombre = "" Then
        m_SQL = "TbUsuariosAplicaciones"
    Else
        m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE Nombre Like '*" & p_Nombre & "*';"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjUsuario = New usuario
            For Each fld In rcdDatos.Fields
                m_ObjUsuario.SetPropiedad fld.Name, Nz(fld.Value, "")
            Next
            
            If getListaUsuariosPorNombre Is Nothing Then
                Set getListaUsuariosPorNombre = New Scripting.Dictionary
                getListaUsuariosPorNombre.CompareMode = TextCompare
            End If
            If Not getListaUsuariosPorNombre.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                getListaUsuariosPorNombre.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
            End If
            Set m_ObjUsuario = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaUsuariosPorNombre ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuarios( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    
    
    On Error GoTo errores
    
    Set getUsuarios = getListaUsuariosPorNombre(, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarios ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getListaAplicacionesPorNombre( _
                                                Optional p_Nombre As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    Dim m_objAplicacion As Aplicacion
            
    
    On Error GoTo errores
    
    If p_Nombre = "" Then
        m_SQL = "TbAplicaciones"
    Else
        m_SQL = "SELECT TbAplicaciones.* " & _
            "FROM TbAplicaciones " & _
            "WHERE NombreAplicacion Like '*" & p_Nombre & "*';"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objAplicacion = New Aplicacion
            For Each fld In rcdDatos.Fields
                m_objAplicacion.SetPropiedad fld.Name, Nz(fld.Value, "")
            Next
            
            If getListaAplicacionesPorNombre Is Nothing Then
                Set getListaAplicacionesPorNombre = New Scripting.Dictionary
                getListaAplicacionesPorNombre.CompareMode = TextCompare
            End If
            If Not getListaAplicacionesPorNombre.Exists(CStr(m_objAplicacion.IDAplicacion)) Then
                getListaAplicacionesPorNombre.Add CStr(m_objAplicacion.IDAplicacion), m_objAplicacion
            End If
            Set m_objAplicacion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaAplicacionesPorNombre ha devuelto el error: " & Err.Description
    End If
End Function




Public Function getUsuarioAplicacionPermisos( _
                                            p_CorreoUsuario As String, _
                                            p_IDAplicacion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As UsuarioAplicacionPermisos

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    
    On Error GoTo errores
    If p_CorreoUsuario = "" Or p_IDAplicacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & p_CorreoUsuario & "' AND IDAplicacion=" & p_IDAplicacion & ";"
            
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getUsuarioAplicacionPermisos = New UsuarioAplicacionPermisos
        For Each fld In rcdDatos.Fields
            getUsuarioAplicacionPermisos.SetPropiedad fld.Name, Nz(fld.Value, "")
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioAplicacionPermisos ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getAplicacionesPermisos( _
                                            p_CorreoUsuario As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
            
    On Error GoTo errores
    If p_CorreoUsuario = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & p_CorreoUsuario & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjUsuarioAplicacionPermisos = New UsuarioAplicacionPermisos
            For Each fld In rcdDatos.Fields
                m_ObjUsuarioAplicacionPermisos.SetPropiedad fld.Name, Nz(fld.Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            'If m_ObjUsuarioAplicacionPermisos.IDAplicacion = "19" Then Stop
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


Public Function getMenu( _
                            frm As Form, _
                            Optional ByRef p_Error As String _
                            ) As Menu

    On Error GoTo errores
    
    Set getMenu = New Menu
    getMenu.EstablecerDatosIniciales frm, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMenu ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUltimaAplicacionAbierta( _
                                                p_CorreoUsuario As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Aplicacion
    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    On Error GoTo errores
    
    If p_CorreoUsuario = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbAplicaciones.* " & _
            "FROM (TbConexionesRegistro INNER JOIN TbConexionUltimaAppAbierta ON " & _
            "TbConexionesRegistro.IDConexion = TbConexionUltimaAppAbierta.IDConexion) " & _
            "INNER JOIN TbAplicaciones ON " & _
            "TbConexionUltimaAppAbierta.IDUltimaAplicacionAbierta = TbAplicaciones.IDAplicacion " & _
            "WHERE (((TbConexionesRegistro.Usuario)='" & p_CorreoUsuario & "')) " & _
            "ORDER BY TbConexionUltimaAppAbierta.IDConexion DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getUltimaAplicacionAbierta = New Aplicacion
        For Each fld In rcdDatos.Fields
            getUltimaAplicacionAbierta.SetPropiedad fld.Name, Nz(fld.Value, "")
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getIDAPlicacionUltimaAbierta ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getCategorias( _
                                p_IDVideo As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
    Dim m_ObjCategoria As Categoria
            
    
    On Error GoTo errores
    
    If p_IDVideo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbCategorias.* " & _
            "FROM TbCategorias INNER JOIN TbVideosCategorias ON " & _
            "TbCategorias.IDCategoria = TbVideosCategorias.IDCategoria " & _
            "WHERE TbVideosCategorias.IDVideo=" & p_IDVideo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjCategoria = New Categoria
            For Each fld In rcdDatos.Fields
                m_ObjCategoria.SetPropiedad fld.Name, Nz(fld.Value, "")
            Next
            If getCategorias Is Nothing Then
                Set getCategorias = New Scripting.Dictionary
                getCategorias.CompareMode = TextCompare
            End If
            If Not getCategorias.Exists(CStr(m_ObjCategoria.IDCategoria)) Then
                getCategorias.Add CStr(m_ObjCategoria.IDCategoria), m_ObjCategoria
            End If
            Set m_ObjCategoria = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCategorias ha devuelto el error: " & Err.Description
    End If
End Function




Public Function getVideos( _
                            Optional p_DeAyuda As EnumSino, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_objVideo As Video
    Dim m_ID As String
    Dim fld As Object
    Dim m_Where As String
    Dim m_SQLInicial As String
    On Error GoTo errores
    p_Error = ""
    If p_DeAyuda = 0 Then
        m_SQL = "TbVideos"
    Else
        m_SQLInicial = "SELECT TbVideos.* " & _
                        "FROM (TbVideos INNER JOIN TbVideosCategorias ON " & _
                        "TbVideos.IDVideo = TbVideosCategorias.IDVideo) " & _
                        "INNER JOIN TbCategorias ON " & _
                        "TbVideosCategorias.IDCategoria = TbCategorias.IDCategoria "
        If p_DeAyuda = EnumSino.Sí Then
            m_Where = "WHERE TbCategorias.NombreCategoria='Ayuda';"
        Else
            m_Where = "WHERE TbCategorias.NombreCategoria='Encuesta';"
        End If
        m_SQL = m_SQLInicial & m_Where
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objVideo = New Video
            For Each fld In rcdDatos.Fields
                m_objVideo.SetPropiedad fld.Name, Nz(fld.Value, "")
            Next
            If getVideos Is Nothing Then
                Set getVideos = New Scripting.Dictionary
                getVideos.CompareMode = TextCompare
            End If
            If Not getVideos.Exists(CStr(m_objVideo.IDVideo)) Then
                getVideos.Add CStr(m_objVideo.IDVideo), m_objVideo
            End If
            Set m_objVideo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getVideos ha devuelto el error: " & Err.Description
    End If
End Function



Public Function getVideo( _
                                m_IDVideo As String, _
                                Optional ByRef p_Error As String _
                                ) As Video

    Dim rcdDatos As DAO.Recordset
    Dim fld As Object
        
    
    On Error GoTo errores
    If m_IDVideo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbVideos " & _
            "WHERE IDVideo=" & m_IDVideo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
       Set getVideo = New Video
        For Each fld In rcdDatos.Fields
            'If fld.Name = "EjecucionEnOficina" Then Stop
            getVideo.SetPropiedad fld.Name, Nz(fld.Value, ""), p_Error
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
        p_Error = "El método getVideo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getVideoByNode( _
                                Optional Nodo As MSComctlLib.Node, _
                                Optional ByRef p_Error As String _
                                ) As Video

    
    Dim m_IDVideo As String
           
    On Error GoTo errores
    If Nodo Is Nothing Then
        Set Nodo = NodoActivo
    End If
    If Nodo Is Nothing Then
        Exit Function
    End If
    If Left(Nodo.key, 1) <> "V" Then
        Exit Function
    End If
    dato = Split(Nodo.key, "V")
    m_IDVideo = CStr(CInt(dato(1)))
    Set getVideoByNode = getVideo(m_IDVideo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getVideoByNode ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAplicacionByNode( _
                                    Optional Nodo As MSComctlLib.Node, _
                                    Optional ByRef p_Error As String _
                                    ) As Aplicacion

    
    Dim m_IDAplicacion As String
    Dim m_objVideo As Video
    
    On Error GoTo errores
    If Nodo Is Nothing Then
        Set Nodo = NodoActivo
    End If
    If Nodo Is Nothing Then
        Exit Function
    End If
    If Left(Nodo.key, 1) = "A" Then
        dato = Split(Nodo.key, "A")
        m_IDAplicacion = CStr(CInt(dato(1)))
        Set getAplicacionByNode = getAplicacion(m_IDAplicacion, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    Set m_objVideo = getVideoByNode(Nodo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objVideo Is Nothing Then
        Exit Function
    End If
    Set getAplicacionByNode = m_objVideo.Aplicacion
    p_Error = m_objVideo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAplicacionByNode ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getVideosDeAplicacion( _
                                        p_IDAplicacion As String, _
                                        Optional p_DeAyuda As EnumSino, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_objVideo As Video
    Dim m_ID As String
    Dim fld As Object
    Dim m_Where As String
    Dim m_SQLInicial As String
    
    On Error GoTo errores
    p_Error = ""
    If p_IDAplicacion = "" Then
        Exit Function
    End If
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        If p_DeAyuda = 0 Then
            m_SQL = "SELECT * FROM TbVideos " & _
                    "WHERE IDAplicacion=" & p_IDAplicacion & ";"
        Else
            m_SQLInicial = "SELECT TbVideos.* " & _
                            "FROM (TbVideos INNER JOIN TbVideosCategorias ON " & _
                            "TbVideos.IDVideo = TbVideosCategorias.IDVideo) " & _
                            "INNER JOIN TbCategorias ON " & _
                            "TbVideosCategorias.IDCategoria = TbCategorias.IDCategoria "
            If p_DeAyuda = EnumSino.Sí Then
                m_Where = "WHERE TbCategorias.NombreCategoria='Ayuda' AND TbVideos.IDAplicacion=" & p_IDAplicacion & ";"
            Else
                m_Where = "WHERE TbCategorias.NombreCategoria='Encuesta' AND TbVideos.IDAplicacion=" & p_IDAplicacion & ";"
            End If
            m_SQL = m_SQLInicial & m_Where
        End If
    Else
        If m_ObjUsuarioConectadoLogin.EsUsuarioCalidadCalculado = EnumSino.Sí Then
            If p_DeAyuda = 0 Then
                m_SQL = "SELECT * FROM TbVideos " & _
                        "WHERE IDAplicacion=" & p_IDAplicacion & " AND ParaCalidad='Sí';"
            Else
                m_SQLInicial = "SELECT TbVideos.* " & _
                                "FROM (TbVideos INNER JOIN TbVideosCategorias ON " & _
                                "TbVideos.IDVideo = TbVideosCategorias.IDVideo) " & _
                                "INNER JOIN TbCategorias ON " & _
                                "TbVideosCategorias.IDCategoria = TbCategorias.IDCategoria "
                If p_DeAyuda = EnumSino.Sí Then
                    m_Where = "WHERE TbCategorias.NombreCategoria='Ayuda' AND " & _
                            "TbVideos.IDAplicacion=" & p_IDAplicacion & " AND ParaCalidad='Sí';"
                Else
                    m_Where = "WHERE TbCategorias.NombreCategoria='Encuesta' AND " & _
                    "TbVideos.IDAplicacion=" & p_IDAplicacion & " AND ParaCalidad='Sí';"
                End If
                m_SQL = m_SQLInicial & m_Where
            End If
            
        Else
            If p_DeAyuda = 0 Then
                m_SQL = "SELECT * FROM TbVideos " & _
                        "WHERE IDAplicacion=" & p_IDAplicacion & " AND ParaTecnicos='Sí';"
            Else
                m_SQLInicial = "SELECT TbVideos.* " & _
                                "FROM (TbVideos INNER JOIN TbVideosCategorias ON " & _
                                "TbVideos.IDVideo = TbVideosCategorias.IDVideo) " & _
                                "INNER JOIN TbCategorias ON " & _
                                "TbVideosCategorias.IDCategoria = TbCategorias.IDCategoria "
                If p_DeAyuda = EnumSino.Sí Then
                    m_Where = "WHERE TbCategorias.NombreCategoria='Ayuda' AND " & _
                            "TbVideos.IDAplicacion=" & p_IDAplicacion & " AND ParaTecnicos='Sí';"
                Else
                    m_Where = "WHERE TbCategorias.NombreCategoria='Encuesta' AND " & _
                    "TbVideos.IDAplicacion=" & p_IDAplicacion & " AND ParaTecnicos='Sí';"
                End If
                m_SQL = m_SQLInicial & m_Where
            End If
        
        End If
        
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objVideo = New Video
            For Each fld In rcdDatos.Fields
                m_objVideo.SetPropiedad fld.Name, Nz(fld.Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getVideosDeAplicacion Is Nothing Then
                Set getVideosDeAplicacion = New Scripting.Dictionary
                getVideosDeAplicacion.CompareMode = TextCompare
            End If
            If Not getVideosDeAplicacion.Exists(CStr(m_objVideo.IDVideo)) Then
                getVideosDeAplicacion.Add CStr(m_objVideo.IDVideo), m_objVideo
            End If
            Set m_objVideo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getVideosDeAplicacion ha devuelto el error: " & Err.Description
    End If
End Function

'Public Function getComprobacionAccesoVideos( _
'                                            p_UsuarioRed As String, _
'                                            Optional ByRef p_Error As String _
'                                            ) As ComprobacionAccesoVideos
'
'    Dim rcdDatos As DAO.Recordset
'    Dim fld As Object
'
'
'    On Error GoTo errores
'    If p_UsuarioRed = "" Then
'        Exit Function
'    End If
'    m_SQL = "SELECT * FROM TbComprobacionCarpetasVideos " & _
'            "WHERE UsuarioRed='" & p_UsuarioRed & "';"
'    Set rcdDatos = getdb().OpenRecordset(m_SQL)
'    With rcdDatos
'        If .EOF Then
'            rcdDatos.Close
'            Set rcdDatos = Nothing
'            Exit Function
'        End If
'       Set getComprobacionAccesoVideos = New ComprobacionAccesoVideos
'        For Each fld In rcdDatos.Fields
'            'If fld.Name = "EjecucionEnOficina" Then Stop
'            getComprobacionAccesoVideos.SetPropiedad fld.Name, Nz(fld.Value, ""), p_Error
'            If p_Error <> "" Then
'                Err.Raise 1000
'            End If
'        Next
'    End With
'    rcdDatos.Close
'    Set rcdDatos = Nothing
'
'    Exit Function
'
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El método getComprobacionAccesoVideos ha devuelto el error: " & Err.description
'    End If
'End Function


