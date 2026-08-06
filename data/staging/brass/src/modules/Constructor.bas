Attribute VB_Name = "Constructor"
Option Compare Database
Option Explicit

Public Function getUsuarioConectadoPorMaquina( _
                                                Optional ByRef p_Error As String _
                                                ) As Usuario
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
            getUsuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
Public Function getUsuarios( _
                                Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    Dim m_NombreCampo As String
    
    On Error GoTo errores
    
    m_SQL = "TbUsuariosAplicaciones"
    
    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjUsuario = New Usuario
            For Each m_Campo In m_ObjUsuario.ColCampos
                m_ObjUsuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getUsuarios Is Nothing Then
                Set getUsuarios = New Scripting.Dictionary
                getUsuarios.CompareMode = TextCompare
            End If
            If Not getUsuarios.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                getUsuarios.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
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
        p_Error = "El método getUsuarios ha devuelto el error: " & Err.Description
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
                m_ObjUsuarioAplicacionPermisos.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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


Public Function getAnexo( _
                            p_IDAnexo As String, _
                            Optional ByRef p_Error As String _
                            ) As Anexo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDAnexo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDAnexo=" & p_IDAnexo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getAnexo = New Anexo
        For Each m_Campo In getAnexo.ColCampos
            getAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getAnexo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAnexosNoAlcanzablesBusqueda( _
                                                Optional p_IDEvento As String, _
                                                Optional p_BUI As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Anexo As Anexo
    
    Dim m_IDEvento As String
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    If p_IDEvento = "" And p_BUI = "" Then
        Exit Function
    End If
    p_Error = ""
    If p_IDEvento <> "" Then
        Set getAnexosNoAlcanzablesBusqueda = getAnexosNoAlcanzablesPorEvento(p_IDEvento, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    
    m_SQL = "SELECT IDEvento " & _
            "FROM TbEventos " & _
            "WHERE BUI='" & p_BUI & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_IDEvento = .Fields("IDEvento").Value
            Set m_Col = getAnexosNoAlcanzablesPorEvento(m_IDEvento, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_Col Is Nothing Then
                For Each m_ID In m_Col
                    Set m_Anexo = m_Col(m_ID)
                    If getAnexosNoAlcanzablesBusqueda Is Nothing Then
                        Set getAnexosNoAlcanzablesBusqueda = New Scripting.Dictionary
                        getAnexosNoAlcanzablesBusqueda.CompareMode = TextCompare
                    End If
                    If Not getAnexosNoAlcanzablesBusqueda.Exists(m_Anexo.IDAnexo) Then
                        getAnexosNoAlcanzablesBusqueda.Add m_Anexo.IDAnexo, m_Anexo
                    End If
                    Set m_Anexo = Nothing
                Next
            End If
            Set m_Col = Nothing
            
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosNoAlcanzablesBusqueda ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexosNoAlcanzablesPorEvento( _
                                                Optional p_IDEvento As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Anexo As Anexo
    Dim fso As New Scripting.FileSystemObject
    On Error GoTo errores
    
    If p_IDEvento = "" Then
        Exit Function
    End If
    p_Error = ""
    
    
    m_SQL = "SELECT * " & _
            "FROM TbAnexos " & _
            "WHERE IDEvento='" & p_IDEvento & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            Do While Not .EOF
                Set m_Anexo = New Anexo
                For Each m_Campo In m_Anexo.ColCampos
                    m_Anexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If Not fso.FileExists(m_Anexo.URLAnexo) Then
                    If getAnexosNoAlcanzablesPorEvento Is Nothing Then
                        Set getAnexosNoAlcanzablesPorEvento = New Scripting.Dictionary
                        getAnexosNoAlcanzablesPorEvento.CompareMode = TextCompare
                    End If
                    If Not getAnexosNoAlcanzablesPorEvento.Exists(m_Anexo.IDAnexo) Then
                        getAnexosNoAlcanzablesPorEvento.Add m_Anexo.IDAnexo, m_Anexo
                    End If
                End If
                
                Set m_Anexo = Nothing
                .MoveNext
            Loop
        End If
        
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbAnexos.* " & _
            "FROM TbAnexos INNER JOIN TbMaterial ON TbAnexos.IDMaterial = TbMaterial.IDMaterial;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            Do While Not .EOF
                Set m_Anexo = New Anexo
                For Each m_Campo In m_Anexo.ColCampos
                    m_Anexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If Not fso.FileExists(m_Anexo.URLAnexo) Then
                    If getAnexosNoAlcanzablesPorEvento Is Nothing Then
                        Set getAnexosNoAlcanzablesPorEvento = New Scripting.Dictionary
                        getAnexosNoAlcanzablesPorEvento.CompareMode = TextCompare
                    End If
                    If Not getAnexosNoAlcanzablesPorEvento.Exists(m_Anexo.IDAnexo) Then
                        getAnexosNoAlcanzablesPorEvento.Add m_Anexo.IDAnexo, m_Anexo
                    End If
                End If
                
                Set m_Anexo = Nothing
                .MoveNext
            Loop
        End If
        
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosNoAlcanzablesPorEvento ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAnexoAutorizacionDeMaterial( _
                                                p_IDMaterial As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Anexo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
   
    On Error GoTo errores
    
    If p_IDMaterial = "" Then
        Exit Function
    End If
    p_Error = ""
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDMaterial='" & p_IDMaterial & "' AND Titulo='Autorización Materiales';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
       Set getAnexoAutorizacionDeMaterial = New Anexo
        For Each m_Campo In getAnexoAutorizacionDeMaterial.ColCampos
            getAnexoAutorizacionDeMaterial.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getAnexoAutorizacionDeMaterial ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAnexosDeEvento( _
                                    p_IDEvento As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
    
    On Error GoTo errores
    
    If p_IDEvento = "" Then
        Exit Function
    End If
    p_Error = ""
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDEvento='" & p_IDEvento & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeEvento Is Nothing Then
                Set getAnexosDeEvento = New Scripting.Dictionary
                getAnexosDeEvento.CompareMode = TextCompare
            End If
            If Not getAnexosDeEvento.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeEvento.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosDeEvento ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexosDeActividad( _
                                    p_IDActividad As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
        
    
    On Error GoTo errores
    
    If p_IDActividad = "" Then
        Exit Function
    End If
    p_Error = ""
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDActividad='" & p_IDActividad & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeActividad Is Nothing Then
                Set getAnexosDeActividad = New Scripting.Dictionary
                getAnexosDeActividad.CompareMode = TextCompare
            End If
            If Not getAnexosDeActividad.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeActividad.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosDeActividad ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexosDeMaterial( _
                                    p_IDMaterial As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
   
    On Error GoTo errores
    
    If p_IDMaterial = "" Then
        Exit Function
    End If
    p_Error = ""
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDMaterial='" & p_IDMaterial & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeMaterial Is Nothing Then
                Set getAnexosDeMaterial = New Scripting.Dictionary
                getAnexosDeMaterial.CompareMode = TextCompare
            End If
            If Not getAnexosDeMaterial.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeMaterial.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosDeMaterial ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexosDeMaterialSeguimiento( _
                                                p_IDMaterialSeg As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
        
    
    On Error GoTo errores
    
    If p_IDMaterialSeg = "" Then
        Exit Function
    End If
    p_Error = ""
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDMaterialSeguimiento='" & p_IDMaterialSeg & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeMaterialSeguimiento Is Nothing Then
                Set getAnexosDeMaterialSeguimiento = New Scripting.Dictionary
                getAnexosDeMaterialSeguimiento.CompareMode = TextCompare
            End If
            If Not getAnexosDeMaterialSeguimiento.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeMaterialSeguimiento.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosDeMaterialSeguimiento ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexosDeGastos( _
                                    p_IDGasto As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
        
    
    On Error GoTo errores
    
    If p_IDGasto = "" Then
        Exit Function
    End If
    p_Error = ""
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDGasto=" & p_IDGasto & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeGastos Is Nothing Then
                Set getAnexosDeGastos = New Scripting.Dictionary
                getAnexosDeGastos.CompareMode = TextCompare
            End If
            If Not getAnexosDeGastos.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeGastos.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosDeGastos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexosDeSubContratacion( _
                                            p_IDSubContratacion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
        
    
    On Error GoTo errores
    
    If p_IDSubContratacion = "" Then
        Exit Function
    End If
    p_Error = ""
    m_SQL = "SELECT * FROM TbAnexos " & _
            "WHERE IDSubContratacion=" & p_IDSubContratacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeSubContratacion Is Nothing Then
                Set getAnexosDeSubContratacion = New Scripting.Dictionary
                getAnexosDeSubContratacion.CompareMode = TextCompare
            End If
            If Not getAnexosDeSubContratacion.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeSubContratacion.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosDeSubContratacion ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getFacturas( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjFactura As Factura
    Dim m_ID As String
    Dim m_Where As String
    Dim m_SQLInicial As String
    On Error GoTo errores
    
    p_Error = ""
    m_SQL = "SELECT * FROM TbFacturaPrincipal ORDER BY IDFactura Desc;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjFactura = New Factura
            For Each m_Campo In m_ObjFactura.ColCampos
                m_ObjFactura.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getFacturas Is Nothing Then
                Set getFacturas = New Scripting.Dictionary
                getFacturas.CompareMode = TextCompare
            End If
            If Not getFacturas.Exists(CStr(m_ObjFactura.IDFactura)) Then
                getFacturas.Add CStr(m_ObjFactura.IDFactura), m_ObjFactura
            End If
            Set m_ObjFactura = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFacturas ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getPartes( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_objParte As Parte
    
    On Error GoTo errores
    
    p_Error = ""
    m_SQL = "TbPartesPpal"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objParte = New Parte
            For Each m_Campo In m_objParte.ColCampos
                m_objParte.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPartes Is Nothing Then
                Set getPartes = New Scripting.Dictionary
                getPartes.CompareMode = TextCompare
            End If
            If Not getPartes.Exists(CStr(m_objParte.IDParte)) Then
                getPartes.Add CStr(m_objParte.IDParte), m_objParte
            End If
            Set m_objParte = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getPartes ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getPartesDeFactura( _
                                    p_IDFactura As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_objParte As Parte
   
    
    On Error GoTo errores
    
    p_Error = ""
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbPartesPpal.* " & _
            "FROM TbPartesPpal " & _
            "WHERE (((TbPartesPpal.IDFactura)=" & p_IDFactura & "));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objParte = New Parte
            For Each m_Campo In m_objParte.ColCampos
                m_objParte.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPartesDeFactura Is Nothing Then
                Set getPartesDeFactura = New Scripting.Dictionary
                getPartesDeFactura.CompareMode = TextCompare
            End If
            If Not getPartesDeFactura.Exists(CStr(m_objParte.IDParte)) Then
                getPartesDeFactura.Add CStr(m_objParte.IDParte), m_objParte
            End If
            Set m_objParte = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getPartesDeFactura ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getPerfilEnFactura( _
                                    Optional p_IDFacturaPerfil As String, _
                                    Optional p_IDFactura As String, _
                                    Optional p_Tipo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As FacturaPerfiles

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    p_Error = ""
    If p_IDFactura = "" Or p_Tipo = "" And p_IDFacturaPerfil = "" Then
        Exit Function
    End If
    If p_IDFacturaPerfil <> "" Then
        m_SQL = "SELECT TbFacturaPrincipalPerfiles.* " & _
            "FROM TbFacturaPrincipalPerfiles " & _
            "WHERE IDFacturaPerfil=" & p_IDFacturaPerfil & ";"
    Else
         If p_IDFactura = "" Or p_Tipo = "" Then
            Exit Function
        End If
        m_SQL = "SELECT TbFacturaPrincipalPerfiles.* " & _
            "FROM TbFacturaPrincipalPerfiles " & _
            "WHERE IDFactura=" & p_IDFactura & " AND Perfil='" & p_Tipo & "';"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPerfilEnFactura = New FacturaPerfiles
        For Each m_Campo In getPerfilEnFactura.ColCampos
            getPerfilEnFactura.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getPerfilEnFactura ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getListaPerfilesEnFactura( _
                                            p_IDFactura As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjFacturaPerfiles As FacturaPerfiles
    On Error GoTo errores
    
    p_Error = ""
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbFacturaPrincipalPerfiles.* " & _
            "FROM TbFacturaPrincipalPerfiles " & _
            "WHERE IDFactura=" & p_IDFactura & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjFacturaPerfiles = New FacturaPerfiles
            For Each m_Campo In m_ObjFacturaPerfiles.ColCampos
                m_ObjFacturaPerfiles.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getListaPerfilesEnFactura Is Nothing Then
                Set getListaPerfilesEnFactura = New Scripting.Dictionary
                getListaPerfilesEnFactura.CompareMode = TextCompare
            End If
            If Not getListaPerfilesEnFactura.Exists(m_ObjFacturaPerfiles.IDFacturaPerfil) Then
                getListaPerfilesEnFactura.Add m_ObjFacturaPerfiles.IDFacturaPerfil, m_ObjFacturaPerfiles
            End If
            Set m_ObjFacturaPerfiles = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaPerfilesEnFactura ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getSubContratacionesParaFactura( _
                                                p_FechaMax As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjSubsContratacion As Subcontratacion
    Dim m_SQLFacturados As String
    
    On Error GoTo errores
    
    p_Error = ""
    If Not IsDate(p_FechaMax) Then
        Exit Function
    End If
    m_SQLFacturados = "SELECT DISTINCT TbFacturaSubcontratacionesInvolucradas.IDSubcontratacion " & _
                    "FROM TbFacturaSubcontratacionesInvolucradas;"
    m_SQL = "SELECT  TbSubcontrataciones.* " & _
            "FROM TbSubcontrataciones " & _
            "WHERE FechaFin<=#" & Format(p_FechaMax, "mm/dd/yyyy") & _
            "# AND IDSubcontratacion Not In (" & m_SQLFacturados & ") " & _
            "ORDER BY TbSubcontrataciones.FechaAlta;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjSubsContratacion = New Subcontratacion
            For Each m_Campo In m_ObjSubsContratacion.ColCampos
                m_ObjSubsContratacion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getSubContratacionesParaFactura Is Nothing Then
                Set getSubContratacionesParaFactura = New Scripting.Dictionary
                getSubContratacionesParaFactura.CompareMode = TextCompare
            End If
            If Not getSubContratacionesParaFactura.Exists(m_ObjSubsContratacion.IDSubContratacion) Then
                getSubContratacionesParaFactura.Add m_ObjSubsContratacion.IDSubContratacion, m_ObjSubsContratacion
            End If
            Set m_ObjSubsContratacion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSubContratacionesParaFactura ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getGastosParaFactura( _
                                        p_FechaMax As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjGasto As Gasto
        
    
    On Error GoTo errores
    
    p_Error = ""
    If Not IsDate(p_FechaMax) Then
        Exit Function
    End If
    m_SQL = "SELECT TbGastos.* " & _
            "FROM TbGastos LEFT JOIN TbFacturaGastosInvolucrados ON TbGastos.IDGasto = TbFacturaGastosInvolucrados.IDGasto " & _
            "WHERE (((TbFacturaGastosInvolucrados.IDGasto) Is Null) AND " & _
            "((TbGastos.FechaImputacionGasto)<=#" & Format(p_FechaMax, "mm/dd/yyyy") & "#)) ORDER BY TbGastos.FechaImputacionGasto;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjGasto = New Gasto
            For Each m_Campo In m_ObjGasto.ColCampos
                m_ObjGasto.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getGastosParaFactura Is Nothing Then
                Set getGastosParaFactura = New Scripting.Dictionary
                getGastosParaFactura.CompareMode = TextCompare
            End If
            If Not getGastosParaFactura.Exists(m_ObjGasto.IDGasto) Then
                getGastosParaFactura.Add m_ObjGasto.IDGasto, m_ObjGasto
            End If
            Set m_ObjGasto = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getGastosParaFactura ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSQLEventosDeFactura( _
                                        p_FechaMax As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    On Error GoTo errores
    
    p_Error = ""
    
    If Not IsDate(p_FechaMax) Then
        Exit Function
    End If
    getSQLEventosDeFactura = "SELECT TbEventos.* " & _
            "FROM TbFacturaEventosInvolucrados RIGHT JOIN TbEventos ON " & _
            "TbFacturaEventosInvolucrados.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbEventos.FechaFinal) <= #" & Format(p_FechaMax, "mm/dd/yyyy") & _
            "#) And ((TbFacturaEventosInvolucrados.IDEVENTO) Is Null) AND ((TbEventos.Franqueado) = True)) " & _
            "ORDER BY TbEventos.IDEvento;"
'    getSQLEventosDeFactura = "SELECT TbEventos.* " & _
'                            "FROM TbEventos " & _
'                            "WHERE IDEvento='BRA2212002';"
                    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSQLEventosDeFactura ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEventosParaFactura( _
                                        Optional p_SQLEventos As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjEvento As Evento
    
    On Error GoTo errores
    
    p_Error = ""
    
    If p_SQLEventos = "" Then
        Exit Function
    End If
    
    m_SQL = p_SQLEventos
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            
            Set m_ObjEvento = New Evento
            For Each m_Campo In m_ObjEvento.ColCampos
                m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEventosParaFactura Is Nothing Then
                Set getEventosParaFactura = New Scripting.Dictionary
                getEventosParaFactura.CompareMode = TextCompare
            End If
            If Not getEventosParaFactura.Exists(m_ObjEvento.IDEVENTO) Then
                getEventosParaFactura.Add m_ObjEvento.IDEVENTO, m_ObjEvento
            End If
            Set m_ObjEvento = Nothing

            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosParaFactura ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getListaEventosNoFacturados( _
                                            Optional p_BUI As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Evento As Evento
    Dim m_Where As String
    
    On Error GoTo errores
    
    p_Error = ""
    '------------------------------------------
    ' Registros que no están franqueados
    '------------------------------------------
    If p_BUI = "" Then
        m_Where = "WHERE IDFactura Is Null;"
    Else
        m_Where = "WHERE IDFactura Is Null AND BUI='" & p_BUI & "';"
    End If
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos LEFT JOIN TbFacturaEventosInvolucrados ON " & _
            "TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento " & _
            m_Where
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No existe ningún registro con ese ID"
            Err.Raise 1000
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_Evento = New Evento
            For Each m_Campo In m_Evento.ColCampos
                m_Evento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getListaEventosNoFacturados Is Nothing Then
                Set getListaEventosNoFacturados = New Scripting.Dictionary
                getListaEventosNoFacturados.CompareMode = TextCompare
            End If
            If Not getListaEventosNoFacturados.Exists(m_Evento.IDEVENTO) Then
                getListaEventosNoFacturados.Add m_Evento.IDEVENTO, m_Evento
            End If
            Set m_Evento = Nothing
siguiente:
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaEventosNoFacturados ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getListaEventosDeTabla( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Evento As Evento
   
    
    
    On Error GoTo errores
    
    '------------------------------------------
    ' Registros que no están franqueados
    '------------------------------------------
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbAuxEventosParaInforme INNER JOIN TbEventos ON TbAuxEventosParaInforme.IDEvento = TbEventos.IDEvento " & _
            "ORDER BY TbAuxEventosParaInforme.IDEvento;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No existe ningún registro con ese ID"
            Err.Raise 1000
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_Evento = New Evento
            For Each m_Campo In m_Evento.ColCampos
                m_Evento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getListaEventosDeTabla Is Nothing Then
                Set getListaEventosDeTabla = New Scripting.Dictionary
                getListaEventosDeTabla.CompareMode = TextCompare
            End If
            If Not getListaEventosDeTabla.Exists(m_Evento.IDEVENTO) Then
                getListaEventosDeTabla.Add m_Evento.IDEVENTO, m_Evento
            End If
            Set m_Evento = Nothing
siguiente:
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaEventosDeTabla ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getActividadesNoDeEventosParaFactura( _
                                                    p_FechaMax As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjActividad As Actividad
      
    
    On Error GoTo errores
    
    p_Error = ""
    
    If Not IsDate(p_FechaMax) Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades LEFT JOIN TbFacturaActividadesInvolucradas ON " & _
            "TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad " & _
            "WHERE (((TbFacturaActividadesInvolucradas.IDActividad) Is Null) AND ((TbActividades.IDEvento) Is Null) AND " & _
            "((TbActividades.FechaAlta)<=#" & Format(p_FechaMax, "mm/dd/yyyy") & "#));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjActividad = New Actividad
            For Each m_Campo In m_ObjActividad.ColCampos
                m_ObjActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getActividadesNoDeEventosParaFactura Is Nothing Then
                Set getActividadesNoDeEventosParaFactura = New Scripting.Dictionary
                getActividadesNoDeEventosParaFactura.CompareMode = TextCompare
            End If
            If Not getActividadesNoDeEventosParaFactura.Exists(m_ObjActividad.IDActividad) Then
                getActividadesNoDeEventosParaFactura.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesNoDeEventosParaFactura ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getMaterialesParaFactura( _
                                        p_SQLEventos As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjMaterial As Material
    Dim m_SQLEventos As String
    On Error GoTo errores
    
    p_Error = ""
    If p_SQLEventos = "" Then
        Exit Function
    End If
    
    m_SQLEventos = Replace(p_SQLEventos, "SELECT TbEventos.* FROM", "SELECT TbEventos.IDEvento FROM")
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM (TbActividades INNER JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad) " & _
            "LEFT JOIN TbFacturaMaterialesInvolucrados ON TbMaterial.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial " & _
            "WHERE (((TbFacturaMaterialesInvolucrados.IDMaterial) Is Null) AND ((TbActividades.IDEvento) In (" & m_SQLEventos & "))) ORDER BY TbActividades.IDEvento;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjMaterial = New Material
            For Each m_Campo In m_ObjMaterial.ColCampos
                m_ObjMaterial.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getMaterialesParaFactura Is Nothing Then
                Set getMaterialesParaFactura = New Scripting.Dictionary
                getMaterialesParaFactura.CompareMode = TextCompare
            End If
            If Not getMaterialesParaFactura.Exists(m_ObjMaterial.IDMaterial) Then
                getMaterialesParaFactura.Add m_ObjMaterial.IDMaterial, m_ObjMaterial
            End If
            Set m_ObjMaterial = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMaterialesParaFactura ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getActividadesDeEventosParaFactura( _
                                                    p_SQLEventos As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjActividad As Actividad
    Dim m_SQLEventos As String
    
    
    On Error GoTo errores
    
    p_Error = ""
    
    If p_SQLEventos = "" Then
        Exit Function
    End If
    
    m_SQLEventos = Replace(p_SQLEventos, "SELECT TbEventos.* FROM", "SELECT TbEventos.IDEvento FROM")
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE IDEvento In (" & m_SQLEventos & ") " & _
            "ORDER BY TbActividades.IDEvento;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjActividad = New Actividad
            For Each m_Campo In m_ObjActividad.ColCampos
                m_ObjActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getActividadesDeEventosParaFactura Is Nothing Then
                Set getActividadesDeEventosParaFactura = New Scripting.Dictionary
                getActividadesDeEventosParaFactura.CompareMode = TextCompare
            End If
            If Not getActividadesDeEventosParaFactura.Exists(m_ObjActividad.IDActividad) Then
                getActividadesDeEventosParaFactura.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesDeEventosParaFactura ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getActividadesDeEvento( _
                                            p_IDEvento As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjActividad As Actividad
    
    On Error GoTo errores
    
    p_Error = ""
    
    If p_IDEvento = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE IDEVENTO='" & p_IDEvento & "' ORDER BY IDActividad;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjActividad = New Actividad
            For Each m_Campo In m_ObjActividad.ColCampos
                m_ObjActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getActividadesDeEvento Is Nothing Then
                Set getActividadesDeEvento = New Scripting.Dictionary
                getActividadesDeEvento.CompareMode = TextCompare
            End If
            If Not getActividadesDeEvento.Exists(m_ObjActividad.IDActividad) Then
                getActividadesDeEvento.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesDeEvento ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getActividadesDeEventos( _
                                            p_IDFactura As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjActividad As Actividad
    
    On Error GoTo errores
    
    p_Error = ""
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbFacturaActividadesInvolucradas INNER JOIN TbActividades ON " & _
            "TbFacturaActividadesInvolucradas.IDActividad = TbActividades.IDActividad " & _
            "WHERE (((TbFacturaActividadesInvolucradas.IDFactura)=" & p_IDFactura & _
            ") AND (Not(TbFacturaActividadesInvolucradas.IDEVENTO) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjActividad = New Actividad
            For Each m_Campo In m_ObjActividad.ColCampos
                m_ObjActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getActividadesDeEventos Is Nothing Then
                Set getActividadesDeEventos = New Scripting.Dictionary
                getActividadesDeEventos.CompareMode = TextCompare
            End If
            If Not getActividadesDeEventos.Exists(m_ObjActividad.IDActividad) Then
                getActividadesDeEventos.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesDeEventos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getActividadesNODeEventos( _
                                            p_IDFactura As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjActividad As Actividad
    
    On Error GoTo errores
    
    p_Error = ""
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbFacturaActividadesInvolucradas INNER JOIN TbActividades ON " & _
            "TbFacturaActividadesInvolucradas.IDActividad = TbActividades.IDActividad " & _
            "WHERE (((TbFacturaActividadesInvolucradas.IDFactura)=" & p_IDFactura & _
            ") AND ((TbFacturaActividadesInvolucradas.IDEVENTO) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjActividad = New Actividad
            For Each m_Campo In m_ObjActividad.ColCampos
                m_ObjActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getActividadesNODeEventos Is Nothing Then
                Set getActividadesNODeEventos = New Scripting.Dictionary
                getActividadesNODeEventos.CompareMode = TextCompare
            End If
            If Not getActividadesNODeEventos.Exists(m_ObjActividad.IDActividad) Then
                getActividadesNODeEventos.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesNODeEventos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getParte( _
                            Optional m_IDParte As String, _
                            Optional m_IDEvento As String, _
                            Optional ByRef p_Error As String _
                            ) As Parte

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If m_IDParte = "" And m_IDEvento = "" Then
        Exit Function
    End If
    If m_IDParte <> "" Then
        m_SQL = "SELECT * FROM TbPartesPpal " & _
            "WHERE IDParte='" & m_IDParte & "';"
    Else
        m_SQL = "SELECT TbPartesPpal.* " & _
                "FROM TbPartesDetalle INNER JOIN TbPartesPpal ON TbPartesDetalle.IDParte = TbPartesPpal.IDParte " & _
                "WHERE (((TbPartesDetalle.IDEvento)='" & m_IDEvento & "'));"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getParte = New Parte
        For Each m_Campo In getParte.ColCampos
            getParte.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getParte ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getMaterial( _
                            m_IDMaterial As String, _
                            Optional ByRef p_Error As String _
                            ) As Material

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If m_IDMaterial = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbMaterial " & _
            "WHERE IDMaterial='" & m_IDMaterial & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getMaterial = New Material
        For Each m_Campo In getMaterial.ColCampos
            getMaterial.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getMaterial ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getGasto( _
                            p_IDGasto As String, _
                            Optional ByRef p_Error As String _
                            ) As Gasto

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_IDGasto = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbGastos " & _
            "WHERE IDGasto=" & p_IDGasto & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getGasto = New Gasto
        For Each m_Campo In getGasto.ColCampos
            getGasto.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getGasto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getMaterialSeguimiento( _
                                        p_IDseguimiento As String, _
                                        Optional ByRef p_Error As String _
                                        ) As MaterialSeguimiento
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDseguimiento = "" Then
        Exit Function
    End If
    
    
    m_SQL = "SELECT TbMaterialSeguimiento.* " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE IDSeguimiento ='" & p_IDseguimiento & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getMaterialSeguimiento = New MaterialSeguimiento
        For Each m_Campo In getMaterialSeguimiento.ColCampos
            getMaterialSeguimiento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getMaterialSeguimiento ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getTecnico( _
                            p_Alias As String, _
                            Optional ByRef p_Error As String _
                            ) As Tecnico
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_Alias = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbTecnicos " & _
            "WHERE ALIAS='" & p_Alias & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getTecnico = New Tecnico
        For Each m_Campo In getTecnico.ColCampos
            getTecnico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "EL método getTecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function getTipoTecnicoPrecio( _
                                        Optional p_ID As String, _
                                        Optional p_Alias As String, _
                                        Optional p_Tipo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As TipoTecnicoPrecio
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_ID = "" And p_Alias = "" And p_Tipo = "" Then
        Exit Function
    End If
    If p_ID <> "" Then
        m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
                "FROM TbTipoTecnicoPrecios " & _
                "WHERE IDTipoTecnicoFecha=" & p_ID & ";"
    Else
        If p_Alias <> "" Then
            m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
                    "FROM TbTecnicos INNER JOIN TbTipoTecnicoPrecios ON TbTecnicos.TIPO = TbTipoTecnicoPrecios.TIPOTECNICO " & _
                    "WHERE ALIAS='" & p_Alias & "' AND FechaFinal Is Null;"
        Else
            m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
                "FROM TbTipoTecnicoPrecios " & _
                "WHERE TIPOTECNICO='" & p_Tipo & "' AND FechaFinal Is Null;"
        End If
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getTipoTecnicoPrecio = New TipoTecnicoPrecio
        For Each m_Campo In getTipoTecnicoPrecio.ColCampos
            getTipoTecnicoPrecio.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "EL método getTipoTecnicoPrecio ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getTipoTecnicoPrecios( _
                                        Optional p_Activos As EnumSino = EnumSino.Sí, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio
    
    On Error GoTo errores
    
    If p_Activos = EnumSino.Sí Then
        m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
                    "FROM TbTecnicos INNER JOIN TbTipoTecnicoPrecios ON TbTecnicos.TIPO = TbTipoTecnicoPrecios.TIPOTECNICO " & _
                    "WHERE  FechaFinal Is Null;"
    Else
        m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
                    "FROM TbTecnicos INNER JOIN TbTipoTecnicoPrecios ON TbTecnicos.TIPO = TbTipoTecnicoPrecios.TIPOTECNICO;"
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
            Set m_ObjTipoTecnicoPrecio = New TipoTecnicoPrecio
            For Each m_Campo In m_ObjTipoTecnicoPrecio.ColCampos
                m_ObjTipoTecnicoPrecio.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getTipoTecnicoPrecios Is Nothing Then
                Set getTipoTecnicoPrecios = New Scripting.Dictionary
                getTipoTecnicoPrecios.CompareMode = TextCompare
            End If
            If Not getTipoTecnicoPrecios.Exists(m_ObjTipoTecnicoPrecio.IDTipoTecnicoFecha) Then
                getTipoTecnicoPrecios.Add m_ObjTipoTecnicoPrecio.IDTipoTecnicoFecha, m_ObjTipoTecnicoPrecio
            End If
            Set m_ObjTipoTecnicoPrecio = Nothing
            .MoveNext
        Loop
        
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getTipoTecnicoPrecios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getTipoTecnicoPrecioAntPost( _
                                            Optional p_ID As String, _
                                            Optional p_Tipo As String, _
                                            Optional p_EsAnterior As EnumSino = EnumSino.Sí, _
                                            Optional ByRef p_Error As String _
                                            ) As TipoTecnicoPrecio
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_ID = "" And p_Tipo = "" Then
        Exit Function
    End If
    If p_EsAnterior = 0 Then
        p_EsAnterior = EnumSino.Sí
    End If
    If p_EsAnterior = EnumSino.Sí Then
        m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
                "FROM TbTipoTecnicoPrecios " & _
                "WHERE IDTipoTecnicoFecha<" & p_ID & _
                " AND TIPOTECNICO='" & p_Tipo & "';"
    Else
        m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
                "FROM TbTipoTecnicoPrecios " & _
                "WHERE IDTipoTecnicoFecha>" & p_ID & _
                " AND TIPOTECNICO='" & p_Tipo & "';"
    End If
        
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getTipoTecnicoPrecioAntPost = New TipoTecnicoPrecio
        For Each m_Campo In getTipoTecnicoPrecioAntPost.ColCampos
            getTipoTecnicoPrecioAntPost.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "EL método getTipoTecnicoPrecioAntPost ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getTipoTecnicoPrecioUltimo( _
                                            Optional p_Alias As String, _
                                            Optional p_Tipo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As TipoTecnicoPrecio
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_SQLLImitante As String
        
    On Error GoTo errores
    
    If p_Alias = "" And p_Tipo = "" Then
        Exit Function
    End If
    If p_Tipo <> "" Then
        m_SQLLImitante = "SELECT Last(TbTipoTecnicoPrecios.IDTipoTecnicoFecha) AS ÚltimoDeIDTipoTecnicoFecha " & _
                    "FROM TbTipoTecnicoPrecios " & _
                    "WHERE TIPOTECNICO='" & p_Tipo & "';"
    Else
        m_SQLLImitante = "SELECT Last(TbTipoTecnicoPrecios.IDTipoTecnicoFecha) AS ÚltimoDeIDTipoTecnicoFecha " & _
                        "FROM TbTipoTecnicoPrecios INNER JOIN TbTecnicos ON " & _
                        "TbTipoTecnicoPrecios.TIPOTECNICO = TbTecnicos.TIPO " & _
                        "WHERE (((TbTecnicos.ALIAS)='" & p_Alias & "'));"
    End If
    m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
            "FROM TbTipoTecnicoPrecios " & _
            "WHERE IDTipoTecnicoFecha In (" & m_SQLLImitante & ");"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getTipoTecnicoPrecioUltimo = New TipoTecnicoPrecio
        For Each m_Campo In getTipoTecnicoPrecioUltimo.ColCampos
            getTipoTecnicoPrecioUltimo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "EL método getTipoTecnicoPrecioUltimo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function getTipoTecnicoActivo( _
                                        Optional p_Tipo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As TipoTecnicoPrecio
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    On Error GoTo errores
    If p_Tipo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbTipoTecnicoPrecios.* " & _
            "FROM TbTipoTecnicoPrecios " & _
            "WHERE (((TbTipoTecnicoPrecios.TIPOTECNICO)='" & p_Tipo & _
            "') AND ((TbTipoTecnicoPrecios.FechaFinal) Is Null)) " & _
            "ORDER BY TbTipoTecnicoPrecios.FechaInicial DESC;"

    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getTipoTecnicoActivo = New TipoTecnicoPrecio
        For Each m_Campo In getTipoTecnicoActivo.ColCampos
            getTipoTecnicoActivo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "EL método getTipoTecnicoActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getEvento( _
                            p_IDEvento As String, _
                            Optional ByRef p_Error As String _
                            ) As Evento

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDEvento = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbEventos " & _
            "WHERE IDEvento='" & p_IDEvento & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEvento = New Evento
        For Each m_Campo In getEvento.ColCampos
            getEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getEvento ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getActividad( _
                            p_IDActividad As String, _
                            Optional ByRef p_Error As String _
                            ) As Actividad

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDActividad = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbActividades " & _
            "WHERE IDActividad='" & p_IDActividad & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getActividad = New Actividad
        For Each m_Campo In getActividad.ColCampos
            getActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getActividad ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEquipo( _
                            p_IDEquipo As String, _
                            Optional ByRef p_Error As String _
                            ) As Equipo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDEquipo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbEquipos " & _
            "WHERE IDEquipo=" & p_IDEquipo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEquipo = New Equipo
        For Each m_Campo In getEquipo.ColCampos
            getEquipo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getEquipo ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getSeguimientos( _
                                p_IDMaterial As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    Dim m_MaterialSeguimiento As MaterialSeguimiento
    
    On Error GoTo errores
    
    If p_IDMaterial = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbMaterialSeguimiento.* " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE IDMaterial ='" & p_IDMaterial & "' ORDER BY IDSeguimiento"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No existe ningún registro con ese ID"
            Err.Raise 1000
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_MaterialSeguimiento = New MaterialSeguimiento
            For Each m_Campo In m_MaterialSeguimiento.ColCampos
                m_MaterialSeguimiento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getSeguimientos Is Nothing Then
                Set getSeguimientos = New Scripting.Dictionary
                getSeguimientos.CompareMode = TextCompare
            End If
            If Not getSeguimientos.Exists(m_MaterialSeguimiento.IDSeguimiento) Then
                getSeguimientos.Add m_MaterialSeguimiento.IDSeguimiento, m_MaterialSeguimiento
            End If
            
            Set m_MaterialSeguimiento = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSeguimientos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getFactura( _
                            p_IDFactura As String, _
                            Optional ByRef p_Error As String _
                            ) As Factura

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
        
    
    On Error GoTo errores
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbFacturaPrincipal " & _
            "WHERE IDFactura=" & p_IDFactura & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getFactura = New Factura
        For Each m_Campo In getFactura.ColCampos
            getFactura.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getFactura ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEventosFiltrados( _
                                    p_EventoBusqueda As EventoBusqueda, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
   
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Evento As Evento
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    
    Dim m_WhereIDEvento As String
    Dim m_WhereBUI As String
    Dim m_WhereSubSistema As String
    Dim m_WhereEquipo As String
    Dim m_WhereTipoEvento As String
    Dim m_WhereFranqueo As String
    Dim m_WhereFechas As String
    Dim m_SQLInicial As String
    Dim m_ParteWhere As String
   
    On Error GoTo errores
    
    
    p_Error = ""
    
    m_SQLInicial = "SELECT  TbEventos.* " & _
                "FROM TbEventos INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo "
    
    If p_EventoBusqueda Is Nothing Then
        Exit Function
    End If
    With p_EventoBusqueda
        If .IDEVENTO = "" Then
            m_WhereIDEvento = "TbEventos.IDEvento Like '*' "
        Else
            m_WhereIDEvento = "TbEventos.IDEvento='" & .IDEVENTO & "'"
        End If
        If .BUI = "" Then
            m_WhereBUI = "TbEventos.BUI Like '*' "
        Else
            m_WhereBUI = "TbEventos.BUI='" & .BUI & "'"
        End If
        If .SubSistema = "" Then
            m_WhereSubSistema = "TbEventos.SUBSISTEMA Like '*' "
        Else
            m_WhereSubSistema = "TbEventos.SUBSISTEMA='" & .SubSistema & "'"
        End If
        If .Equipo = "" Then
            m_WhereEquipo = "(TbEquipos.Equipo Like '*' or TbEquipos.Equipo Is Null) "
        Else
            m_WhereEquipo = "TbEquipos.Equipo='" & .Equipo & "'"
        End If
        If .TipoEvento = "" Then
            m_WhereTipoEvento = "TbEventos.TIPOEVENTO Like '*' "
        Else
            m_WhereTipoEvento = "TbEventos.TIPOEVENTO='" & .TipoEvento & "'"
        End If
        If .Franqueado = "" Then
            m_WhereFranqueo = "(TbEventos.Franqueado=True Or TbEventos.Franqueado=False)"
        Else
            If .Franqueado = "Sí" Then
                 m_WhereFranqueo = "TbEventos.Franqueado=True"
            Else
                m_WhereFranqueo = "TbEventos.Franqueado=False"
            End If
        End If
        If Not IsDate(.FechaInicial) Then .FechaInicial = "01/01/1900"
        If Not IsDate(.FechaFinal) Then .FechaFinal = "01/01/2100"
        m_WhereFechas = "TbEventos.FECHAALTAEVENTO Between #" & _
                        Format(.FechaInicial, "mm/dd/yyyy") & _
                        "# And #" & Format(.FechaFinal, "mm/dd/yyyy") & "#"
    End With
    m_ParteWhere = "WHERE (" & _
                    m_WhereIDEvento & " AND " & _
                    m_WhereBUI & " AND " & _
                    m_WhereSubSistema & " AND " & _
                    m_WhereEquipo & " AND " & _
                    m_WhereTipoEvento & " AND " & _
                    m_WhereFranqueo & " AND " & _
                    m_WhereFechas & ") " & _
                    "ORDER BY TbEventos.FECHAALTAEVENTO DESC;"

    m_SQL = m_SQLInicial & m_ParteWhere
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Evento = New Evento
            For Each m_Campo In m_Evento.ColCampos
                m_Evento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEventosFiltrados Is Nothing Then
                Set getEventosFiltrados = New Scripting.Dictionary
                getEventosFiltrados.CompareMode = TextCompare
            End If
            If Not getEventosFiltrados.Exists(CStr(m_Evento.IDEVENTO)) Then
                getEventosFiltrados.Add CStr(m_Evento.IDEVENTO), m_Evento
            End If
            Set m_Evento = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosFiltrados ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function getEventosEnParte( _
                                    p_IDParte As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjEvento As Evento
    Dim m_ID As String
    Dim m_Where As String
    Dim m_SQLInicial As String
    
    
    On Error GoTo errores
    
    If p_IDParte = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbPartesDetalle INNER JOIN TbEventos ON TbPartesDetalle.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbPartesDetalle.IDParte)='" & p_IDParte & "'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjEvento = New Evento
            For Each m_Campo In m_ObjEvento.ColCampos
                m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEventosEnParte Is Nothing Then
                Set getEventosEnParte = New Scripting.Dictionary
                getEventosEnParte.CompareMode = TextCompare
            End If
            If Not getEventosEnParte.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                getEventosEnParte.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
            End If
            Set m_ObjEvento = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosEnParte ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEventosSinFechaInformeRAC( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjEvento As Evento
    Dim m_ID As String
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbEventos " & _
            "WHERE FechaEnInformeRAC  Is Null " & _
            "AND Franqueado=true;"

    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjEvento = New Evento
            For Each m_Campo In m_ObjEvento.ColCampos
                m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEventosSinFechaInformeRAC Is Nothing Then
                Set getEventosSinFechaInformeRAC = New Scripting.Dictionary
                getEventosSinFechaInformeRAC.CompareMode = TextCompare
            End If
            If Not getEventosSinFechaInformeRAC.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                getEventosSinFechaInformeRAC.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
            End If
            Set m_ObjEvento = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosSinFechaInformeRAC ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getPrimerEventoEnParte( _
                                        p_IDParte As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Evento

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDParte = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbPartesDetalle INNER JOIN TbEventos ON TbPartesDetalle.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbPartesDetalle.IDParte)='" & p_IDParte & "'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getPrimerEventoEnParte = New Evento
        For Each m_Campo In getPrimerEventoEnParte.ColCampos
            getPrimerEventoEnParte.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getPrimerEventoEnParte ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEventosEnFactura( _
                                    p_IDFactura As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjEvento As Evento
    Dim m_ID As String
    Dim m_Where As String
    Dim m_SQLInicial As String
    
    On Error GoTo errores
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbFacturaEventosInvolucrados INNER JOIN TbEventos ON TbFacturaEventosInvolucrados.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbFacturaEventosInvolucrados.IDFactura)=" & p_IDFactura & "))ORDER BY TbEventos.IDEvento;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjEvento = New Evento
            For Each m_Campo In m_ObjEvento.ColCampos
                m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEventosEnFactura Is Nothing Then
                Set getEventosEnFactura = New Scripting.Dictionary
                getEventosEnFactura.CompareMode = TextCompare
            End If
            If Not getEventosEnFactura.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                getEventosEnFactura.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
            End If
            Set m_ObjEvento = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosEnFactura ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getCentros( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_BUI As String
   
    
    On Error GoTo errores
    m_SQL = "SELECT BUI " & _
            "FROM TbBUI ;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_BUI = .Fields("BUI")
            
            If getCentros Is Nothing Then
                Set getCentros = New Scripting.Dictionary
                getCentros.CompareMode = TextCompare
            End If
            If Not getCentros.Exists(m_BUI) Then
                getCentros.Add m_BUI, m_BUI
            End If
            m_BUI = ""
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
        
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCentros ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getCentrosPorFactura( _
                                        p_Factura As Factura, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Centro As String
   
    Dim m_IDEnvento As Variant
    Dim m_Evento As Evento
    
    On Error GoTo errores
    
    If p_Factura Is Nothing Then
        Exit Function
    End If
    If p_Factura.IDFactura <> "" Then
        m_SQL = "SELECT distinct TbEventos.BUI " & _
                "FROM TbFacturaEventosInvolucrados INNER JOIN TbEventos ON TbFacturaEventosInvolucrados.IDEvento = TbEventos.IDEvento " & _
                "WHERE (((TbFacturaEventosInvolucrados.IDFactura)=" & p_Factura.IDFactura & "));"
        
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                rcdDatos.Close
                Set rcdDatos = Nothing
                Exit Function
            End If
            .MoveFirst
            Do While Not .EOF
                m_Centro = .Fields("BUI")
                
                If getCentrosPorFactura Is Nothing Then
                    Set getCentrosPorFactura = New Scripting.Dictionary
                    getCentrosPorFactura.CompareMode = TextCompare
                End If
                If Not getCentrosPorFactura.Exists(m_Centro) Then
                    getCentrosPorFactura.Add m_Centro, m_Centro
                End If
                m_Centro = ""
                .MoveNext
            Loop
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Exit Function
    
    End If
    If p_Factura.colListaEventosPorFacturar Is Nothing Then
        Exit Function
    End If
    For Each m_IDEnvento In p_Factura.colListaEventosPorFacturar
        Set m_Evento = p_Factura.colListaEventosPorFacturar(m_IDEnvento)
        If m_Evento.BUI <> "" Then
            If getCentrosPorFactura Is Nothing Then
                Set getCentrosPorFactura = New Scripting.Dictionary
                getCentrosPorFactura.CompareMode = TextCompare
            End If
            If Not getCentrosPorFactura.Exists(m_Evento.BUI) Then
                getCentrosPorFactura.Add m_Evento.BUI, m_Evento.BUI
            End If
        End If
        Set m_Evento = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCentrosPorFactura ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEventosEnFacturaYCentro( _
                                            p_Factura As Factura, _
                                            p_BUI As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjEvento As Evento
    Dim m_IDEvento As Variant
    On Error GoTo errores
    
    If p_Factura Is Nothing Or p_BUI = "" Then
        Exit Function
    End If
    If p_Factura.IDFactura <> "" Then
        m_SQL = "SELECT TbEventos.* " & _
                "FROM TbFacturaEventosInvolucrados INNER JOIN TbEventos ON TbFacturaEventosInvolucrados.IDEvento = TbEventos.IDEvento " & _
                "WHERE (((TbFacturaEventosInvolucrados.IDFactura)=" & p_Factura.IDFactura & _
                ") AND ((TbEventos.BUI)='" & p_BUI & "'));"
        
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                rcdDatos.Close
                Set rcdDatos = Nothing
                Exit Function
            End If
            .MoveFirst
            Do While Not .EOF
                Set m_ObjEvento = New Evento
                For Each m_Campo In m_ObjEvento.ColCampos
                    m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getEventosEnFacturaYCentro Is Nothing Then
                    Set getEventosEnFacturaYCentro = New Scripting.Dictionary
                    getEventosEnFacturaYCentro.CompareMode = TextCompare
                End If
                If Not getEventosEnFacturaYCentro.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                    getEventosEnFacturaYCentro.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
                End If
                Set m_ObjEvento = Nothing
                .MoveNext
            Loop
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Exit Function
    End If
    If p_Factura.colListaEventosPorFacturar Is Nothing Then
        Exit Function
    End If
    For Each m_IDEvento In p_Factura.colListaEventosPorFacturar
        Set m_ObjEvento = p_Factura.colListaEventosPorFacturar(m_IDEvento)
        If m_ObjEvento.BUI = p_BUI Then
            If getEventosEnFacturaYCentro Is Nothing Then
                Set getEventosEnFacturaYCentro = New Scripting.Dictionary
                getEventosEnFacturaYCentro.CompareMode = TextCompare
            End If
            If Not getEventosEnFacturaYCentro.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                getEventosEnFacturaYCentro.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
            End If
        End If
        Set m_ObjEvento = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosEnFacturaYCentro ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getFacturaDeEvento( _
                                    p_IDEvento As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Factura

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDEvento = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbFacturaPrincipal.* " & _
            "FROM TbFacturaEventosInvolucrados INNER JOIN TbFacturaPrincipal ON " & _
            "TbFacturaEventosInvolucrados.IDFactura = TbFacturaPrincipal.IDFactura " & _
            "WHERE (((TbFacturaEventosInvolucrados.IDEvento)='" & p_IDEvento & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getFacturaDeEvento = New Factura
        For Each m_Campo In getFacturaDeEvento.ColCampos
            getFacturaDeEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getFacturaDeEvento ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getFacturaDeActividad( _
                                    p_IDActividad As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Factura

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_IDActividad = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbFacturaPrincipal.* " & _
            "FROM TbFacturaActividadesInvolucradas INNER JOIN TbFacturaPrincipal ON " & _
            "TbFacturaActividadesInvolucradas.IDFactura = TbFacturaPrincipal.IDFactura " & _
            "WHERE (((TbFacturaActividadesInvolucradas.IDActividad)='" & p_IDActividad & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getFacturaDeActividad = New Factura
        For Each m_Campo In getFacturaDeActividad.ColCampos
            getFacturaDeActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getFacturaDeActividad ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getFacturaDeMaterial( _
                                    p_IDMaterial As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Factura

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_IDMaterial = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbFacturaPrincipal.* " & _
            "FROM TbFacturaMaterialesInvolucrados INNER JOIN TbFacturaPrincipal ON " & _
            "TbFacturaMaterialesInvolucrados.IDFactura = TbFacturaPrincipal.IDFactura " & _
            "WHERE (((TbFacturaMaterialesInvolucrados.IDMaterial)='" & p_IDMaterial & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getFacturaDeMaterial = New Factura
        For Each m_Campo In getFacturaDeMaterial.ColCampos
            getFacturaDeMaterial.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getFacturaDeMaterial ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEventosConEquipoMedida( _
                                        p_IDEquipoMedida As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_ObjEvento As Evento
    
    On Error GoTo errores
    
    If p_IDEquipoMedida = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventosEquipoMedida INNER JOIN TbEventos ON TbEventosEquipoMedida.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbEventosEquipoMedida.IDEquipoMedida)=" & p_IDEquipoMedida & "));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjEvento = New Evento
            For Each m_Campo In m_ObjEvento.ColCampos
                m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEventosConEquipoMedida Is Nothing Then
                Set getEventosConEquipoMedida = New Scripting.Dictionary
                getEventosConEquipoMedida.CompareMode = TextCompare
            End If
            If Not getEventosConEquipoMedida.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                getEventosConEquipoMedida.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
            End If
            Set m_ObjEvento = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosConEquipoMedida ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getFacturaDeSubContratacion( _
                                            p_IDSubContratacion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Factura

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_IDSubContratacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbFacturaPrincipal.* " & _
            "FROM TbFacturaSubcontratacionesInvolucradas INNER JOIN TbFacturaPrincipal ON " & _
            "TbFacturaSubcontratacionesInvolucradas.IDFactura = TbFacturaPrincipal.IDFactura " & _
            "WHERE (((TbFacturaSubcontratacionesInvolucradas.IDSubcontratacion)=" & p_IDSubContratacion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getFacturaDeSubContratacion = New Factura
        For Each m_Campo In getFacturaDeSubContratacion.ColCampos
            getFacturaDeSubContratacion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getFacturaDeSubContratacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getFacturaDeGasto( _
                                    p_IDGasto As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Factura

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    On Error GoTo errores
    If p_IDGasto = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbFacturaPrincipal.* " & _
            "FROM TbFacturaGastosInvolucrados INNER JOIN TbFacturaPrincipal ON " & _
            "TbFacturaGastosInvolucrados.IDFactura = TbFacturaPrincipal.IDFactura " & _
            "WHERE (((TbFacturaGastosInvolucrados.IDGasto)=" & p_IDGasto & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getFacturaDeGasto = New Factura
        For Each m_Campo In getFacturaDeGasto.ColCampos
            getFacturaDeGasto.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getFacturaDeGasto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getActividadesEnFactura( _
                                        p_IDFactura As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Actividad As Actividad
    
    
    On Error GoTo errores
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    
   
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbFacturaActividadesInvolucradas INNER JOIN TbActividades ON " & _
            "TbFacturaActividadesInvolucradas.IDActividad = TbActividades.IDActividad " & _
            "WHERE (((TbFacturaActividadesInvolucradas.IDFactura) =" & p_IDFactura & ")) " & _
            "ORDER BY TbActividades.IDEvento;"

    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_Actividad = New Actividad
            For Each m_Campo In m_Actividad.ColCampos
                m_Actividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getActividadesEnFactura Is Nothing Then
                Set getActividadesEnFactura = New Scripting.Dictionary
                getActividadesEnFactura.CompareMode = TextCompare
            End If
            If Not getActividadesEnFactura.Exists(m_Actividad.IDActividad) Then
                getActividadesEnFactura.Add m_Actividad.IDActividad, m_Actividad
            End If
            Set m_Actividad = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesEnFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getAnexosAutorizacionMaterialesEnFactura( _
                                                            p_IDFactura As String, _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_ID As String
    
    
    On Error GoTo errores
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbAnexos.* " & _
            "FROM TbAnexos " & _
            "WHERE ((Not(TbAnexos.IDMaterial) Is Null) " & _
            "AND ((TbAnexos.Descripcion)='Anexado automáticamente desde la factura " & p_IDFactura & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosAutorizacionMaterialesEnFactura Is Nothing Then
                Set getAnexosAutorizacionMaterialesEnFactura = New Scripting.Dictionary
                getAnexosAutorizacionMaterialesEnFactura.CompareMode = TextCompare
            End If
            If Not getAnexosAutorizacionMaterialesEnFactura.Exists(CStr(m_ObjAnexo.IDAnexo)) Then
                getAnexosAutorizacionMaterialesEnFactura.Add CStr(m_ObjAnexo.IDAnexo), m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosAutorizacionMaterialesEnFactura ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getMaterialesEnFactura( _
                                        p_IDFactura As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Material As Material
    
    
    On Error GoTo errores
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM (TbMaterial INNER JOIN TbFacturaMaterialesInvolucrados ON " & _
            "TbMaterial.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial) " & _
            "INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad " & _
            "WHERE (((TbFacturaMaterialesInvolucrados.IDFactura)=" & p_IDFactura & ")) " & _
            "ORDER BY TbActividades.IDEvento;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_Material = New Material
            For Each m_Campo In m_Material.ColCampos
                m_Material.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getMaterialesEnFactura Is Nothing Then
                Set getMaterialesEnFactura = New Scripting.Dictionary
                getMaterialesEnFactura.CompareMode = TextCompare
            End If
            If Not getMaterialesEnFactura.Exists(m_Material.IDMaterial) Then
                getMaterialesEnFactura.Add m_Material.IDMaterial, m_Material
            End If
            Set m_Material = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMaterialesEnFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getMaterialesEnActividad( _
                                        p_IDActividad As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Material As Material
    
    
    On Error GoTo errores
    
    If p_IDActividad = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDActividad)='" & p_IDActividad & "')) " & _
            "ORDER BY TbMaterial.IDMaterial;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_Material = New Material
            For Each m_Campo In m_Material.ColCampos
                m_Material.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getMaterialesEnActividad Is Nothing Then
                Set getMaterialesEnActividad = New Scripting.Dictionary
                getMaterialesEnActividad.CompareMode = TextCompare
            End If
            If Not getMaterialesEnActividad.Exists(m_Material.IDMaterial) Then
                getMaterialesEnActividad.Add m_Material.IDMaterial, m_Material
            End If
            Set m_Material = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMaterialesEnActividad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getGastosEnFactura( _
                                    p_IDFactura As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Gastos As Gasto
    
    
    On Error GoTo errores
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    
   
    m_SQL = "SELECT TbGastos.* " & _
            "FROM TbFacturaGastosInvolucrados INNER JOIN TbGastos ON TbFacturaGastosInvolucrados.IDGasto = TbGastos.IDGasto " & _
            "WHERE (((TbFacturaGastosInvolucrados.IDFactura) =" & p_IDFactura & ")) " & _
            "ORDER BY TbGastos.IDGasto;"


    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_Gastos = New Gasto
            For Each m_Campo In m_Gastos.ColCampos
                m_Gastos.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getGastosEnFactura Is Nothing Then
                Set getGastosEnFactura = New Scripting.Dictionary
                getGastosEnFactura.CompareMode = TextCompare
            End If
            If Not getGastosEnFactura.Exists(m_Gastos.IDGasto) Then
                getGastosEnFactura.Add m_Gastos.IDGasto, m_Gastos
            End If
            Set m_Gastos = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getGastosEnFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getSubContratacionesEnFactura( _
                                                p_IDFactura As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_SubContratacion As Subcontratacion
    
    
    On Error GoTo errores
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbSubcontrataciones.* " & _
            "FROM TbFacturaSubcontratacionesInvolucradas INNER JOIN TbSubcontrataciones ON " & _
            "TbFacturaSubcontratacionesInvolucradas.IDSubcontratacion = TbSubcontrataciones.IDSubcontratacion " & _
            "WHERE (((TbFacturaSubcontratacionesInvolucradas.IDFactura) =" & p_IDFactura & ")) " & _
            "ORDER BY TbSubcontrataciones.IDSubcontratacion;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_SubContratacion = New Subcontratacion
            For Each m_Campo In m_SubContratacion.ColCampos
                m_SubContratacion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getSubContratacionesEnFactura Is Nothing Then
                Set getSubContratacionesEnFactura = New Scripting.Dictionary
                getSubContratacionesEnFactura.CompareMode = TextCompare
            End If
            If Not getSubContratacionesEnFactura.Exists(m_SubContratacion.IDSubContratacion) Then
                getSubContratacionesEnFactura.Add m_SubContratacion.IDSubContratacion, m_SubContratacion
            End If
            Set m_SubContratacion = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSubContratacionesEnFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getAnexosAutorizacionMateriales( _
                                                p_IDMaterial As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_ID As String
    
    On Error GoTo errores
    
    If p_IDMaterial = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbAnexos.* " & _
            "FROM TbMaterial INNER JOIN TbAnexos ON TbMaterial.IDMaterial = TbAnexos.IDMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & p_IDMaterial & _
            "') AND ((TbAnexos.Descripcion) Like 'Anexado automáticamente desde la factura*'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosAutorizacionMateriales Is Nothing Then
                Set getAnexosAutorizacionMateriales = New Scripting.Dictionary
                getAnexosAutorizacionMateriales.CompareMode = TextCompare
            End If
            If Not getAnexosAutorizacionMateriales.Exists(CStr(m_ObjAnexo.IDAnexo)) Then
                getAnexosAutorizacionMateriales.Add CStr(m_ObjAnexo.IDAnexo), m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosAutorizacionMateriales ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEventosPorBUIEnFactura( _
                                            p_Factura As Factura, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    'sirve para poder agrupar la colección de eventos por BUI de una factura
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjEvento As Evento
    Dim m_ObjColEventosBUI As Scripting.Dictionary
    Dim m_BUI As Variant
    Dim m_SQLLimitanteFactura As String
    Dim m_ColCentrosPorFactura As Scripting.Dictionary
    
    
    
    On Error GoTo errores
    
    If p_Factura Is Nothing Then
        Exit Function
    End If
    If p_Factura.IDFactura = "" Then
        Set getEventosPorBUIEnFactura = getEventosPorBUIEnFacturaSimulada(p_Factura, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    m_SQLLimitanteFactura = "SELECT TbFacturaEventosInvolucrados.IDEvento " & _
                            "FROM TbFacturaEventosInvolucrados " & _
                            "WHERE (((TbFacturaEventosInvolucrados.IDFactura)=" & p_Factura.IDFactura & "));"
    Set m_ColCentrosPorFactura = getCentrosPorFactura(p_Factura, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColCentrosPorFactura Is Nothing Then
        Exit Function
    End If
    For Each m_BUI In m_ColCentrosPorFactura
        m_SQL = "SELECT TbEventos.* " & _
                "FROM TbEventos " & _
                "WHERE BUI='" & m_BUI & "' AND IDEvento In (" & m_SQLLimitanteFactura & ") ORDER BY IDEvento;"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            
            .MoveFirst
            Do While Not .EOF
                Set m_ObjEvento = New Evento
                For Each m_Campo In m_ObjEvento.ColCampos
                    m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                
                If m_ObjColEventosBUI Is Nothing Then
                    Set m_ObjColEventosBUI = New Scripting.Dictionary
                    m_ObjColEventosBUI.CompareMode = TextCompare
                End If
                If Not m_ObjColEventosBUI.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                    'm_BUI|m_FechaMinima|FechaMaxima,m_ObjColEventos
                    m_ObjColEventosBUI.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
                End If
                Set m_ObjEvento = Nothing
                
                .MoveNext
            Loop
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        If Not m_ObjColEventosBUI Is Nothing Then
            If getEventosPorBUIEnFactura Is Nothing Then
                Set getEventosPorBUIEnFactura = New Scripting.Dictionary
                getEventosPorBUIEnFactura.CompareMode = TextCompare
            End If
            If Not getEventosPorBUIEnFactura.Exists(m_BUI) Then
                getEventosPorBUIEnFactura.Add m_BUI, m_ObjColEventosBUI
            End If
            Set m_ObjColEventosBUI = Nothing
        End If
       
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosPorBUIEnFactura ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEventosPorBUIEnFacturaSimulada( _
                                                    p_Factura As Factura, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    'sirve para poder agrupar la colección de eventos por BUI de una factura
   
    Dim m_Evento As Evento
    Dim m_IDEvento As Variant
    Dim m_ColEventosPorBUI As Scripting.Dictionary
    Dim m_BUI As Variant
    Dim m_ColBuisDistintos As Scripting.Dictionary
    
    On Error GoTo errores
    
    If p_Factura Is Nothing Then
        Exit Function
    End If
    If p_Factura.colListaEventosPorFacturar Is Nothing Then
        Exit Function
    End If
    Set m_ColBuisDistintos = Constructor.getCentrosPorFactura(p_Factura, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColBuisDistintos Is Nothing Then
        Exit Function
    End If
    For Each m_BUI In m_ColBuisDistintos
        For Each m_IDEvento In p_Factura.colListaEventosPorFacturar
            Set m_Evento = p_Factura.colListaEventosPorFacturar(m_IDEvento)
            If m_Evento.BUI = CStr(m_BUI) Then
                If m_ColEventosPorBUI Is Nothing Then
                    Set m_ColEventosPorBUI = New Scripting.Dictionary
                    m_ColEventosPorBUI.CompareMode = TextCompare
                End If
                If Not m_ColEventosPorBUI.Exists(m_IDEvento) Then
                    m_ColEventosPorBUI.Add m_IDEvento, m_Evento
                End If
            End If
            
            
            Set m_Evento = Nothing
        Next
        If Not m_ColEventosPorBUI Is Nothing Then
            If getEventosPorBUIEnFacturaSimulada Is Nothing Then
                Set getEventosPorBUIEnFacturaSimulada = New Scripting.Dictionary
                getEventosPorBUIEnFacturaSimulada.CompareMode = TextCompare
            End If
            If Not getEventosPorBUIEnFacturaSimulada.Exists(m_BUI) Then
                getEventosPorBUIEnFacturaSimulada.Add m_BUI, m_ColEventosPorBUI
            End If
        End If
        Set m_ColEventosPorBUI = Nothing
    Next
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosPorBUIEnFacturaSimulada ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSubContratacion( _
                                    p_IDSubContratacion As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Subcontratacion

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_IDSubContratacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbSubcontrataciones " & _
            "WHERE IDSubcontratacion=" & p_IDSubContratacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getSubContratacion = New Subcontratacion
        For Each m_Campo In getSubContratacion.ColCampos
            getSubContratacion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getSubContratacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getFacturasBusqueda( _
                                    Optional p_IDFactura As String, _
                                    Optional p_FechaInicial As String, _
                                    Optional p_FechaFinal As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjFactura As Factura
    Dim m_ID As String
    Dim m_Where As String
    
    On Error GoTo errores
    
    p_Error = ""
    If p_IDFactura <> "" Then
        m_SQL = "SELECT * FROM TbFacturaPrincipal " & _
            "WHERE IDFactura=" & p_IDFactura & ";"
    Else
        If Not IsDate(p_FechaInicial) Or Not IsDate(p_FechaFinal) Then
            p_Error = "Se ha de indicar la fecha inicial o final"
            Err.Raise 1000
        End If
        If CDate(p_FechaInicial) < CDate(p_FechaFinal) Then
            p_Error = "Las fecha inicial ha de ser anterior a la final"
            Err.Raise 1000
        End If
        m_SQL = "SELECT TbFacturaPrincipal.* " & _
                "FROM TbFacturaPrincipal " & _
                "WHERE (((TbFacturaPrincipal.FechaFactura) Between #" & _
                        Format(p_FechaInicial, "mm/dd/yyyy") & "# And #" & _
                        Format(p_FechaFinal, "mm/dd/yyyy") & "#));"
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
            Set m_ObjFactura = New Factura
            For Each m_Campo In m_ObjFactura.ColCampos
                m_ObjFactura.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getFacturasBusqueda Is Nothing Then
                Set getFacturasBusqueda = New Scripting.Dictionary
                getFacturasBusqueda.CompareMode = TextCompare
            End If
            If Not getFacturasBusqueda.Exists(CStr(m_ObjFactura.IDFactura)) Then
                getFacturasBusqueda.Add CStr(m_ObjFactura.IDFactura), m_ObjFactura
            End If
            Set m_ObjFactura = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFacturasBusqueda ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getFiesta( _
                            m_Fecha As String, _
                            Optional ByRef p_Error As String _
                            ) As Fiesta

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If Not IsDate(m_Fecha) Then
        Exit Function
    End If
    m_Fecha = Format(m_Fecha, "mm/dd/yyyy")
    m_SQL = "SELECT * FROM TbTecnicosFiestas " & _
            "WHERE FechaFiesta=#" & m_Fecha & "# ;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getFiesta = New Fiesta
        For Each m_Campo In getFiesta.ColCampos
            getFiesta.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getFiesta ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getFiestas( _
                                Optional p_Año As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjFiesta As Fiesta
    Dim m_ID As String
    
    
    
    On Error GoTo errores
    
    p_Error = ""
    If IsNumeric(p_Año) Then
        m_SQL = "SELECT TbTecnicosFiestas.* " & _
                "FROM TbTecnicosFiestas " & _
                "WHERE (((Year([FechaFiesta]))=" & p_Año & "))ORDER BY FechaFiesta DESC;"
    Else
        m_SQL = "SELECT TbTecnicosFiestas.* " & _
                "FROM TbTecnicosFiestas " & _
                "ORDER BY FechaFiesta DESC;"
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
            Set m_ObjFiesta = New Fiesta
            For Each m_Campo In m_ObjFiesta.ColCampos
                m_ObjFiesta.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getFiestas Is Nothing Then
                Set getFiestas = New Scripting.Dictionary
                getFiestas.CompareMode = TextCompare
            End If
            If Not getFiestas.Exists(m_ObjFiesta.FechaFiesta) Then
                getFiestas.Add m_ObjFiesta.FechaFiesta, m_ObjFiesta
            End If
            Set m_ObjFiesta = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFiestas ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getLibranza( _
                            Optional p_IDAusencia As String, _
                            Optional p_Fecha As String, _
                            Optional ByRef p_Error As String _
                            ) As LIbranza

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If Not IsDate(p_Fecha) And p_IDAusencia = "" Then
        Exit Function
    End If
    If IsDate(p_Fecha) Then
        p_Fecha = Format(p_Fecha, "mm/dd/yyyy")
    End If
    If p_IDAusencia <> "" Then
        m_SQL = "SELECT * FROM TbTecnicosAusencias " & _
            "WHERE IDAusencia=" & p_IDAusencia & " ;"
    Else
        m_SQL = "SELECT * FROM TbTecnicosAusencias " & _
            "WHERE FechaLibranza=#" & p_Fecha & "# ;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getLibranza = New LIbranza
        For Each m_Campo In getLibranza.ColCampos
            getLibranza.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getLibranza ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getLibranzas( _
                                Optional p_Año As String, _
                                Optional p_Alias As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjLibranza As LIbranza
    Dim m_ID As String
    
    
    
    On Error GoTo errores
    
    p_Error = ""
    If p_Alias <> "" And p_Año <> "" Then
        m_SQL = "SELECT TbTecnicosAusencias.* " & _
                "FROM TbTecnicosAusencias " & _
                "WHERE " & _
                "Alias ='" & p_Alias & "' " & _
                "AND Year([FechaLibranza])=" & p_Año & " " & _
                "ORDER BY FechaLibranza DESC;"
    ElseIf p_Alias <> "" And p_Año = "" Then
        m_SQL = "SELECT TbTecnicosAusencias.* " & _
                "FROM TbTecnicosAusencias " & _
                "WHERE " & _
                "Alias ='" & p_Alias & "' " & _
                "ORDER BY FechaLibranza DESC;"
                
    ElseIf p_Alias = "" And p_Año <> "" Then
        m_SQL = "SELECT TbTecnicosAusencias.* " & _
                "FROM TbTecnicosAusencias " & _
                "WHERE " & _
                "Alias ='" & p_Alias & "' " & _
                "Year([FechaLibranza])=" & p_Año & " " & _
                "ORDER BY FechaLibranza DESC;"
    ElseIf p_Alias = "" And p_Año <> "" Then
        m_SQL = "SELECT TbTecnicosAusencias.* " & _
                "FROM TbTecnicosAusencias " & _
                "ORDER BY FechaLibranza DESC;"
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
            Set m_ObjLibranza = New LIbranza
            For Each m_Campo In m_ObjLibranza.ColCampos
                m_ObjLibranza.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getLibranzas Is Nothing Then
                Set getLibranzas = New Scripting.Dictionary
                getLibranzas.CompareMode = TextCompare
            End If
            If Not getLibranzas.Exists(m_ObjLibranza.IDAusencia) Then
                getLibranzas.Add m_ObjLibranza.IDAusencia, m_ObjLibranza
            End If
            Set m_ObjLibranza = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getLibranzas ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getTipoTecnico( _
                                Optional p_Tipo As String, _
                                Optional p_Alias As String, _
                                Optional ByRef p_Error As String _
                                ) As TIPOTECNICO

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_Tipo = "" And p_Alias = "" Then
        Exit Function
    End If
    If p_Tipo <> "" Then
        m_SQL = "SELECT * FROM TbTipoTecnico " & _
            "WHERE TIPO='" & p_Tipo & "';"
    Else
        m_SQL = "SELECT TbTipoTecnico.* " & _
                "FROM TbTecnicos INNER JOIN TbTipoTecnico ON TbTecnicos.TIPO = TbTipoTecnico.TIPO " & _
                "WHERE ALIAS='" & p_Alias & "';"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getTipoTecnico = New TIPOTECNICO
        For Each m_Campo In getTipoTecnico.ColCampos
            getTipoTecnico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getTipoTecnico ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getTecnicos( _
                                Optional p_EnTodos As EnumSino, _
                                Optional p_SoloEnActivo As EnumSino = EnumSino.Sí, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjTecnico As Tecnico
    
    On Error GoTo errores
    
    p_Error = ""
    
    If p_EnTodos <> EnumSino.Sí Then
        If p_SoloEnActivo = EnumSino.Sí Then
            m_SQL = "SELECT TbTecnicos.* " & _
                    "FROM TbTecnicos " & _
                    "WHERE FECHABAJA Is Null"
        Else
            m_SQL = "SELECT TbTecnicos.* " & _
                    "FROM TbTecnicos " & _
                    "WHERE Not FECHABAJA Is Null"
        End If
    Else
        m_SQL = "TbTecnicos"
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
            Set m_ObjTecnico = New Tecnico
            For Each m_Campo In m_ObjTecnico.ColCampos
                m_ObjTecnico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getTecnicos Is Nothing Then
                Set getTecnicos = New Scripting.Dictionary
                getTecnicos.CompareMode = TextCompare
            End If
            If Not getTecnicos.Exists(CStr(m_ObjTecnico.Alias)) Then
                getTecnicos.Add CStr(m_ObjTecnico.Alias), m_ObjTecnico
            End If
            Set m_ObjTecnico = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getTecnicos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getTecnicosPorTipo( _
                                    Optional p_Tipo As String, _
                                    Optional p_EnActivo As EnumSino = EnumSino.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjTecnico As Tecnico
    Dim m_WhereTipo As String
    Dim m_WhereEnActivo As String
    Dim m_Where As String
    Dim m_SQLAlInicio As String
    
    
    On Error GoTo errores
    
    p_Error = ""
    m_SQLAlInicio = "SELECT TbTecnicos.* " & _
                "FROM TbTecnicos "
                
    If p_Tipo = "" Then
        m_WhereTipo = "(TIPO Like '*' Or TIPO Is Null)"
    Else
        m_WhereTipo = "TIPO='" & p_Tipo & "'"
    End If
    If p_EnActivo = EnumSino.Sí Then
        m_WhereEnActivo = "FECHABAJA Is Null"
    ElseIf p_EnActivo = EnumSino.No Then
        m_WhereEnActivo = "Not FECHABAJA Is Null"
    End If
    If m_WhereEnActivo <> "" Then
        m_Where = "WHERE " & _
                m_WhereEnActivo & _
                " AND " & _
                m_WhereTipo & _
                ";"
    Else
        m_Where = "WHERE " & _
                    m_WhereTipo & _
                ";"
    End If
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
            Set m_ObjTecnico = New Tecnico
            For Each m_Campo In m_ObjTecnico.ColCampos
                m_ObjTecnico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getTecnicosPorTipo Is Nothing Then
                Set getTecnicosPorTipo = New Scripting.Dictionary
                getTecnicosPorTipo.CompareMode = TextCompare
            End If
            If Not getTecnicosPorTipo.Exists(CStr(m_ObjTecnico.Alias)) Then
                getTecnicosPorTipo.Add CStr(m_ObjTecnico.Alias), m_ObjTecnico
            End If
            Set m_ObjTecnico = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getTecnicosPorTipo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getActividadesTecnicoDia( _
                                        p_Alias As String, _
                                        p_Fecha As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjActividad As Actividad
    Dim m_ID As String
    
    
    On Error GoTo errores
    
    If p_Alias = "" And p_Fecha = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.ALIASTECNICO)='" & p_Alias & _
            "') AND ((TbActividades.FechaAlta)=#" & Format(p_Fecha, "mm/dd/yyyy") & "#));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjActividad = New Actividad
            For Each m_Campo In m_ObjActividad.ColCampos
                m_ObjActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getActividadesTecnicoDia Is Nothing Then
                Set getActividadesTecnicoDia = New Scripting.Dictionary
                getActividadesTecnicoDia.CompareMode = TextCompare
            End If
            If Not getActividadesTecnicoDia.Exists(CStr(m_ObjActividad.IDActividad)) Then
                getActividadesTecnicoDia.Add CStr(m_ObjActividad.IDActividad), m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesTecnicoDia ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getListaEventosEntreFechas( _
                                            Optional p_ObjColEventos As Scripting.Dictionary, _
                                            Optional p_FechaInicial As String, _
                                            Optional p_FechaFinal As String, _
                                            Optional p_SoloFranqueados As EnumSino = EnumSino.No, _
                                            Optional p_BUI As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Evento As Evento
    Dim m_IDEvento As Variant
    Dim m_FechaInicial As String
    Dim m_FechaFinal As String
    Dim m_Franqueado As Boolean
    Dim m_DesdeCol As Boolean
    
    On Error GoTo errores
    If p_ObjColEventos Is Nothing Then
        m_DesdeCol = False
    Else
        If p_ObjColEventos.count = 0 Then
            m_DesdeCol = False
        Else
            m_DesdeCol = True
        End If
    End If
    If Not IsDate(p_FechaInicial) Then
        p_FechaInicial = "01/01/1900"
    End If
    If Not IsDate(p_FechaFinal) Then
        p_FechaFinal = Date
    End If
    
    If m_DesdeCol = True Then
        For Each m_IDEvento In p_ObjColEventos
            Set m_Evento = p_ObjColEventos(m_IDEvento)
            If p_BUI <> "" Then
                If p_BUI <> m_Evento.BUI Then
                    GoTo siguiente1
                End If
            End If
            m_FechaInicial = m_Evento.FECHAALTAEVENTO
            m_FechaFinal = m_Evento.FechaFinal
            m_Franqueado = m_Evento.Franqueado
            If p_SoloFranqueados = EnumSino.No Then
                
                If IsDate(m_FechaFinal) Then
                    If CDate(m_FechaInicial) < CDate(p_FechaInicial) And CDate(m_FechaFinal) < CDate(p_FechaInicial) Then
                        GoTo siguiente1
                    End If
                    If CDate(m_FechaInicial) > CDate(p_FechaFinal) And CDate(m_FechaFinal) > CDate(p_FechaFinal) Then
                        GoTo siguiente1
                    End If
                End If
                
            Else
                If Not m_Franqueado Then
                    GoTo siguiente1
                End If
                If IsDate(m_FechaFinal) Then
                    If Not (CDate(m_FechaFinal) >= CDate(p_FechaInicial) And CDate(m_FechaFinal) <= CDate(p_FechaFinal)) Then
                        GoTo siguiente1
                    End If
                End If
            End If
            If getListaEventosEntreFechas Is Nothing Then
                Set getListaEventosEntreFechas = New Scripting.Dictionary
                getListaEventosEntreFechas.CompareMode = TextCompare
            End If
            If Not getListaEventosEntreFechas.Exists(m_Evento.IDEVENTO) Then
                getListaEventosEntreFechas.Add m_Evento.IDEVENTO, m_Evento
            End If
            Set m_Evento = Nothing
siguiente1:
        Next
        Exit Function
    End If
    
    '------------------------------------------
    ' Registros que no están franqueados
    '------------------------------------------
    If p_BUI = "" Then
        m_SQL = "SELECT TbEventos.* " & _
                "FROM TbEventos;"
    Else
        m_SQL = "SELECT TbEventos.* " & _
                "FROM TbEventos " & _
                "WHERE TbEventos.BUI='" & p_BUI & "';"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No existe ningún registro con ese ID"
            Err.Raise 1000
        End If
        .MoveFirst
        
        Do While Not .EOF
            m_IDEvento = Nz(.Fields("IDEvento"), "")
            m_FechaInicial = Nz(.Fields("FECHAALTAEVENTO"), "")
            m_FechaFinal = Nz(.Fields("FechaFinal"), "")
            m_Franqueado = .Fields("Franqueado")
            If p_SoloFranqueados = EnumSino.No Then
                
                If IsDate(m_FechaFinal) Then
                    If CDate(m_FechaInicial) < CDate(p_FechaInicial) And CDate(m_FechaFinal) < CDate(p_FechaInicial) Then
                        GoTo siguiente
                    End If
                    If CDate(m_FechaInicial) > CDate(p_FechaFinal) And CDate(m_FechaFinal) > CDate(p_FechaFinal) Then
                        GoTo siguiente
                    End If
                End If
                
            Else
                If Not m_Franqueado Then
                    GoTo siguiente
                End If
                If Not (CDate(m_FechaFinal) >= CDate(p_FechaInicial) And CDate(m_FechaFinal) <= CDate(p_FechaFinal)) Then
                    GoTo siguiente
                End If
            End If
            Set m_Evento = New Evento
            For Each m_Campo In m_Evento.ColCampos
                m_Evento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getListaEventosEntreFechas Is Nothing Then
                Set getListaEventosEntreFechas = New Scripting.Dictionary
                getListaEventosEntreFechas.CompareMode = TextCompare
            End If
            If Not getListaEventosEntreFechas.Exists(m_Evento.IDEVENTO) Then
                getListaEventosEntreFechas.Add m_Evento.IDEVENTO, m_Evento
            End If
            Set m_Evento = Nothing
siguiente:
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaEventosEntreFechas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getListaEventosNoFacturados2( _
                                                Optional p_SoloFranqueados As EnumSino = EnumSino.No, _
                                                Optional p_BUI As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Evento As Evento
    Dim m_IDEvento As String
    Dim m_Franqueado As Boolean
    
    Dim m_Where As String
    
    
    On Error GoTo errores
    '------------------------------------------
    ' Registros que no están facturados
    '------------------------------------------
    If p_BUI = "" Then
        m_Where = "WHERE TbFacturaEventosInvolucrados.IDFactura Is Null "
    Else
        m_Where = "WHERE TbFacturaEventosInvolucrados.IDFactura Is Null AND TbEventos.BUI='" & p_BUI & "' "
    End If
    
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos LEFT JOIN TbFacturaEventosInvolucrados ON " & _
            "TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento " & _
            m_Where & _
            "ORDER BY TbEventos.FechaRegistroAlta;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No existe ningún registro con ese ID"
            Err.Raise 1000
        End If
        .MoveFirst
        
        Do While Not .EOF
            Set m_Evento = New Evento
            m_IDEvento = Nz(.Fields("IDEvento"), "")
           
            m_Franqueado = .Fields("Franqueado")
            If p_SoloFranqueados = EnumSino.Sí Then
                If Not m_Franqueado Then
                    GoTo siguiente
                End If
                
            End If
            For Each m_Campo In m_Evento.ColCampos
                m_Evento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getListaEventosNoFacturados2 Is Nothing Then
                Set getListaEventosNoFacturados2 = New Scripting.Dictionary
                getListaEventosNoFacturados2.CompareMode = TextCompare
            End If
            If Not getListaEventosNoFacturados2.Exists(m_Evento.IDEVENTO) Then
                getListaEventosNoFacturados2.Add m_Evento.IDEVENTO, m_Evento
            End If
            Set m_Evento = Nothing
siguiente:
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaEventosNoFacturados2 ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getListaEventosEntreFechas2( _
                                            p_FechaInicial As String, _
                                            p_FechaFinal As String, _
                                            p_Franqueado As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_IDEvento As String
    Dim m_Evento As Evento
    Dim m_Where As String
    
    
    On Error GoTo errores
    If p_Franqueado = "Sí" Then
        m_Where = "WHERE (((TbEventos.Franqueado)=true) AND " & _
                        "((TbEventos.FechaFinal) Between #" & Format(p_FechaInicial, "mm/dd/yyyy") & _
                            "# And #" & Format(p_FechaFinal, "mm/dd/yyyy") & "#)) ORDER BY TbEventos.IDEvento;"
    ElseIf p_Franqueado = "No" Then
        m_Where = "WHERE (((TbEventos.Franqueado)=false) AND " & _
                        "((TbEventos.FECHAALTAEVENTO) Between #" & Format(p_FechaInicial, "mm/dd/yyyy") & _
                            "# And #" & Format(p_FechaFinal, "mm/dd/yyyy") & "#)) ORDER BY TbEventos.IDEvento;"
    Else
        m_Where = "WHERE (((TbEventos.FECHAALTAEVENTO) Between #" & Format(p_FechaInicial, "mm/dd/yyyy") & _
                            "# And #" & Format(p_FechaFinal, "mm/dd/yyyy") & "#)) ORDER BY TbEventos.IDEvento;"
    End If
    
    '------------------------------------------
    ' Registros que no están franqueados
    '------------------------------------------
    m_SQL = "SELECT TbEventos.* " & _
                "FROM TbEventos " & _
            m_Where
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Evento = New Evento
            For Each m_Campo In m_Evento.ColCampos
                m_Evento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getListaEventosEntreFechas2 Is Nothing Then
                Set getListaEventosEntreFechas2 = New Scripting.Dictionary
                getListaEventosEntreFechas2.CompareMode = TextCompare
            End If
            If Not getListaEventosEntreFechas2.Exists(m_Evento.IDEVENTO) Then
                getListaEventosEntreFechas2.Add m_Evento.IDEVENTO, m_Evento
            End If
            Set m_Evento = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaEventosEntreFechas2 ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getEquiposMedidaPorEvento( _
                                            p_IDEvento As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Equipo As EquipoMedida
    
    On Error GoTo errores
    If p_IDEvento = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbEquiposMedida.* " & _
            "FROM TbEventosEquipoMedida INNER JOIN TbEquiposMedida ON " & _
            "TbEventosEquipoMedida.IDEquipoMedida = TbEquiposMedida.IDEquipoMedida " & _
            "WHERE (((TbEventosEquipoMedida.IDEvento)='" & p_IDEvento & "'));"
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Equipo = New EquipoMedida
            
            For Each m_Campo In m_Equipo.ColCampos
                m_Equipo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEquiposMedidaPorEvento Is Nothing Then
                Set getEquiposMedidaPorEvento = New Scripting.Dictionary
                getEquiposMedidaPorEvento.CompareMode = TextCompare
            End If
            If Not getEquiposMedidaPorEvento.Exists(m_Equipo.IDEquipoMedida) Then
                getEquiposMedidaPorEvento.Add m_Equipo.IDEquipoMedida, m_Equipo
            End If
            Set m_Equipo = Nothing
            
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEquiposMedidaPorEvento ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEquiposMedida( _
                                Optional p_Nombre As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Equipo As EquipoMedida
    
    On Error GoTo errores
    
    If p_Nombre <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbEquiposMedida " & _
                "WHERE Nombre Like '*" & p_Nombre & "*';"
    Else
        m_SQL = "TbEquiposMedida"
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
            Set m_Equipo = New EquipoMedida
            For Each m_Campo In m_Equipo.ColCampos
                m_Equipo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEquiposMedida Is Nothing Then
                Set getEquiposMedida = New Scripting.Dictionary
                getEquiposMedida.CompareMode = TextCompare
            End If
            If Not getEquiposMedida.Exists(m_Equipo.IDEquipoMedida) Then
                getEquiposMedida.Add m_Equipo.IDEquipoMedida, m_Equipo
            End If
            Set m_Equipo = Nothing
            .MoveNext
        Loop
        
        
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEquiposMedida ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEquipoMedida( _
                                p_IDEquipo As String, _
                                Optional ByRef p_Error As String _
                                ) As EquipoMedida

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
   
    
    On Error GoTo errores
    If p_IDEquipo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbEquiposMedida " & _
            "WHERE IDEquipoMedida=" & p_IDEquipo & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEquipoMedida = New EquipoMedida
        For Each m_Campo In getEquipoMedida.ColCampos
            getEquipoMedida.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getEquiposMedida ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEventoEquipoMedida( _
                                        p_IDEventoEquipoMedida As String, _
                                        Optional ByRef p_Error As String _
                                        ) As EventoEquipoMedida

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
   
    
    On Error GoTo errores
    If p_IDEventoEquipoMedida = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbEventosEquipoMedida " & _
            "WHERE IDEventoEquipoMedida=" & p_IDEventoEquipoMedida & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEventoEquipoMedida = New EventoEquipoMedida
        For Each m_Campo In getEventoEquipoMedida.ColCampos
            getEventoEquipoMedida.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getEventoEquipoMedida ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEventoEquiposMedida( _
                                        p_IDEvento As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_EventoEquipoMedida As EventoEquipoMedida
    
    On Error GoTo errores
    If p_IDEvento = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbEventosEquipoMedida " & _
            "WHERE IDEvento='" & p_IDEvento & "';"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Do While Not .EOF
            Set m_EventoEquipoMedida = New EventoEquipoMedida
            For Each m_Campo In m_EventoEquipoMedida.ColCampos
                m_EventoEquipoMedida.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEventoEquiposMedida Is Nothing Then
                Set getEventoEquiposMedida = New Scripting.Dictionary
                getEventoEquiposMedida.CompareMode = TextCompare
            End If
            If Not getEventoEquiposMedida.Exists(m_EventoEquipoMedida.IDEventoEquipoMedida) Then
                getEventoEquiposMedida.Add m_EventoEquipoMedida.IDEventoEquipoMedida, m_EventoEquipoMedida
            End If
            
            Set m_EventoEquipoMedida = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventoEquipoMedida ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEquipoMedidaCalibracion( _
                                            Optional p_IDEquipoMedidaCalibracion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As EquipoMedidaCalibracion

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
        
    
    On Error GoTo errores
    If p_IDEquipoMedidaCalibracion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbEquiposMedidaCalibraciones " & _
                "WHERE IDCalibracion=" & p_IDEquipoMedidaCalibracion & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEquipoMedidaCalibracion = New EquipoMedidaCalibracion
        For Each m_Campo In getEquipoMedidaCalibracion.ColCampos
            getEquipoMedidaCalibracion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getEquipoMedidaCalibracion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEquipoMedidaCalibraciones( _
                                            p_IDEquipoMedidaCalibracion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_EquipoMedidaCalibracion As EquipoMedidaCalibracion
    
    On Error GoTo errores
    If p_IDEquipoMedidaCalibracion = "" Then
        Exit Function
    End If
   
    m_SQL = "SELECT * FROM TbEquiposMedidaCalibraciones " & _
            "WHERE IDEquipoMedida=" & p_IDEquipoMedidaCalibracion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_EquipoMedidaCalibracion = New EquipoMedidaCalibracion
            For Each m_Campo In m_EquipoMedidaCalibracion.ColCampos
                m_EquipoMedidaCalibracion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEquipoMedidaCalibraciones Is Nothing Then
                Set getEquipoMedidaCalibraciones = New Scripting.Dictionary
                getEquipoMedidaCalibraciones.CompareMode = TextCompare
            End If
            If Not getEquipoMedidaCalibraciones.Exists(m_EquipoMedidaCalibracion.IDCalibracion) Then
                getEquipoMedidaCalibraciones.Add m_EquipoMedidaCalibracion.IDCalibracion, m_EquipoMedidaCalibracion
            End If
            Set m_EquipoMedidaCalibracion = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEquipoMedidaCalibraciones ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEquiposMedidaActivos( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_SQLLImitante As String
    Dim m_EquipoMedida As EquipoMedida
    
    On Error GoTo errores
    
    m_SQLLImitante = "SELECT TbEquiposMedidaCalibraciones.IDEquipoMedida " & _
                    "FROM TbEquiposMedidaCalibraciones INNER JOIN TbEquiposMedida ON " & _
                    "TbEquiposMedidaCalibraciones.IDEquipoMedida = TbEquiposMedida.IDEquipoMedida " & _
                    "WHERE (((TbEquiposMedidaCalibraciones.FechaFinCalibracion)>=Date()) " & _
                    "AND ((TbEquiposMedida.FechaFinServicio) Is Null));"
                   
    m_SQL = "SELECT * " & _
            "FROM TbEquiposMedida " & _
            "WHERE  IDEquipoMedida In(" & m_SQLLImitante & ");"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_EquipoMedida = New EquipoMedida
            For Each m_Campo In m_EquipoMedida.ColCampos
                m_EquipoMedida.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEquiposMedidaActivos Is Nothing Then
                Set getEquiposMedidaActivos = New Scripting.Dictionary
                getEquiposMedidaActivos.CompareMode = TextCompare
            End If
            If Not getEquiposMedidaActivos.Exists(m_EquipoMedida.IDEquipoMedida) Then
                getEquiposMedidaActivos.Add m_EquipoMedida.IDEquipoMedida, m_EquipoMedida
            End If

            Set m_EquipoMedida = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEquiposMedidaActivos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getUltimaCalibracion( _
                                    p_IDEquipoMedida As String, _
                                    Optional ByRef p_Error As String _
                                    ) As EquipoMedidaCalibracion

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
        
    
    On Error GoTo errores
    If p_IDEquipoMedida = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbEquiposMedidaCalibraciones " & _
            "WHERE IDEquipoMedida =" & p_IDEquipoMedida & " " & _
            "ORDER BY FechaCalibracion DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getUltimaCalibracion = New EquipoMedidaCalibracion
        For Each m_Campo In getUltimaCalibracion.ColCampos
            getUltimaCalibracion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
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
        p_Error = "El método getUltimaCalibracion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getEventosGetronic( _
                                    p_ColEventos As Scripting.Dictionary, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim m_Evento As Evento
    Dim m_ID As Variant
    
    
    On Error GoTo errores
    
    If p_ColEventos Is Nothing Then
        Exit Function
    End If
    For Each m_ID In p_ColEventos
        Set m_Evento = p_ColEventos(m_ID)
        If Left(m_Evento, "PT-") <> 0 Then
            If getEventosGetronic Is Nothing Then
                Set getEventosGetronic = New Scripting.Dictionary
                getEventosGetronic.CompareMode = TextCompare
            End If
            If Not getEventosGetronic.Exists(m_Evento.IDEVENTO) Then
                getEventosGetronic.Add m_Evento.IDEVENTO, m_Evento
            End If
        End If
        Set m_Evento = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosGetronic ha devuelto el error: " & Err.Description
    End If
End Function

Public Function HayAlguienDeGetronicsEnEvento( _
                                                IDEVENTO As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSino

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
   
    
    On Error GoTo errores
    
    
    m_SQL = "SELECT TbActividades.ALIASTECNICO " & _
            "FROM TbActividades INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
            "WHERE (((TbActividades.IDEvento)='" & IDEVENTO & "') AND ((TbTecnicos.TIPO) Like '*GET*'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            HayAlguienDeGetronicsEnEvento = EnumSino.No
        Else
            HayAlguienDeGetronicsEnEvento = EnumSino.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HayAlguienDeGetronicsEnEvento ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getActividadesGetronic( _
                                        p_IDEvento As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String



    'INGENIERO DE SISTEMAS-->ANALISTA
    'TÉCNICO EN COMUNICACIONES-->TÉCNICO EN COMUNICACIONES
    'TÉCNICO DE SISTEMAS-->PROGRAMADOR
    
    'getActividadesGetronic="ANALISTA"; m_NHoras ; m_Coste & "|" & _
                            "TECCOM" ; m_NHoras ; m_Coste & "|" & _
                            "PROGRAMADOR" ; m_NHoras ; m_Coste & _
                            "TOTAL" ; m_NHoras ; m_Coste
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim rcdDatos As DAO.Recordset
    Dim m_Actividad As Actividad
    Dim m_ID As Variant
    Dim m_Texto As String
    Dim m_TextoTotal  As String
    Dim m_NHoras As String
    Dim m_Coste As String
    Dim m_PrecioHoraAnalista As String
    Dim m_PrecioHoraTecCom As String
    Dim m_PrecioHoraProgramador As String
    Dim m_TipoTecnicoPrecio As TipoTecnicoPrecio
    
    Dim dblNHoras As Double
    Dim dblPrecio As Double
    Dim dblCoste As Double
    Dim dblTotal As Double
    
    On Error GoTo errores
    
    If p_IDEvento = "" Then
        Exit Function
    End If
    
    ' "ANALISTA" INGENIERO DE SISTEMAS
    Set m_TipoTecnicoPrecio = getTipoTecnicoPrecioUltimo(, "INGENIERO DE SISTEMAS", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecio Is Nothing Then
        p_Error = "No se ha podido obtener el precio hora actual de un INGENIERO DE SISTEMAS"
        Err.Raise 1000
    End If
    m_PrecioHoraAnalista = m_TipoTecnicoPrecio.PrecioHoraLab
    If Not IsNumeric(m_PrecioHoraAnalista) Then
        p_Error = "No se ha podido obtener el precio hora actual de un INGENIERO DE SISTEMAS"
        Err.Raise 1000
    End If
    
    
    ' "TÉCNICO EN COMUNICACIONES"
    Set m_TipoTecnicoPrecio = getTipoTecnicoPrecioUltimo(, "TÉCNICO EN COMUNICACIONES", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecio Is Nothing Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO EN COMUNICACIONES"
        Err.Raise 1000
    End If
    m_PrecioHoraTecCom = m_TipoTecnicoPrecio.PrecioHoraLab
    If Not IsNumeric(m_PrecioHoraTecCom) Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO EN COMUNICACIONES"
        Err.Raise 1000
    End If
    
    ' "TÉCNICO DE SISTEMAS" PROGRAMADOR
    Set m_TipoTecnicoPrecio = getTipoTecnicoPrecioUltimo(, "TÉCNICO DE SISTEMAS", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecio Is Nothing Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO DE SISTEMAS"
        Err.Raise 1000
    End If
    m_PrecioHoraProgramador = m_TipoTecnicoPrecio.PrecioHoraLab
    If Not IsNumeric(m_PrecioHoraTecCom) Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO DE SISTEMAS"
        Err.Raise 1000
    End If
    
    '    getActividadesGetronic.Add "ANALISTA", m_NHoras & "|" & m_Coste
    m_SQL = "SELECT Sum(TbActividades.HorasLaborables) AS SumaDeHorasLaborables " & _
            "FROM TbActividades INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
            "WHERE (((TbActividades.IDEvento)='" & p_IDEvento & "') AND ((TbTecnicos.TIPO)='Ingeniero de sistemas'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_NHoras = Nz(.Fields("SumaDeHorasLaborables"), "")
            If IsNumeric(m_NHoras) Then
                m_Coste = CStr(CDbl(m_NHoras) * CDbl(m_PrecioHoraAnalista))
                dblTotal = CDbl(m_Coste)
                m_Texto = "ANALISTA" & ";" & m_NHoras & ";" & m_Coste
                m_TextoTotal = m_Texto
            Else
                m_Texto = "ANALISTA" & ";0;0"
                m_TextoTotal = m_Texto
            End If
           
        Else
            m_Texto = "ANALISTA" & ";0;0"
            m_TextoTotal = m_Texto
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    '    getActividadesGetronic.Add "TECCOM", , m_NHoras & "|" & m_Coste
    m_SQL = "SELECT Sum(TbActividades.HorasLaborables) AS SumaDeHorasLaborables " & _
            "FROM TbActividades INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
            "WHERE (((TbActividades.IDEvento)='" & p_IDEvento & "') AND ((TbTecnicos.TIPO)='TÉCNICO EN COMUNICACIONES'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            m_NHoras = Nz(.Fields("SumaDeHorasLaborables"), "")
            If IsNumeric(m_NHoras) Then
                m_Coste = CStr(CDbl(m_NHoras) * CDbl(m_PrecioHoraTecCom))
                dblTotal = dblTotal + CDbl(m_Coste)
                m_Texto = "TECCOM" & ";" & m_NHoras & ";" & m_Coste
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            Else
                 m_Texto = "TECCOM" & ";0;0"
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            End If
            
        Else
            m_Texto = "TECCOM" & ";0;0"
            m_TextoTotal = m_TextoTotal & "|" & m_Texto
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    '    getActividadesGetronic.Add "PROGRAMADOR", , m_NHoras & "|" & m_Coste
    m_SQL = "SELECT Sum(TbActividades.HorasLaborables) AS SumaDeHorasLaborables " & _
            "FROM TbActividades INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
            "WHERE (((TbActividades.IDEvento)='" & p_IDEvento & "') AND ((TbTecnicos.TIPO)='TÉCNICO DE SISTEMAS'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            m_NHoras = Nz(.Fields("SumaDeHorasLaborables"), "")
            If IsNumeric(m_NHoras) Then
                m_Coste = CStr(CDbl(m_NHoras) * CDbl(m_PrecioHoraProgramador))
                dblTotal = dblTotal + CDbl(m_Coste)
                m_Texto = "PROGRAMADOR" & ";" & m_NHoras & ";" & m_Coste
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            Else
                m_Texto = "PROGRAMADOR" & ";0;0"
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            End If
            
        Else
            m_Texto = "PROGRAMADOR" & ";0;0"
            m_TextoTotal = m_TextoTotal & "|" & m_Texto
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_TextoTotal = m_TextoTotal & "|" & dblTotal
    getActividadesGetronic = m_TextoTotal
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesGetronic ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getActividadesGetronic1( _
                                        p_IDEvento As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String



    'INGENIERO DE SISTEMAS-->ANALISTA
    'TÉCNICO EN COMUNICACIONES-->TÉCNICO EN COMUNICACIONES
    'TÉCNICO DE SISTEMAS-->PROGRAMADOR
    
    'getActividadesGetronic="ANALISTA"; m_NHoras ; m_Coste & "|" & _
                            "TECCOM" ; m_NHoras ; m_Coste & "|" & _
                            "PROGRAMADOR" ; m_NHoras ; m_Coste & _
                            "TOTAL" ; m_NHoras ; m_Coste
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim rcdDatos As DAO.Recordset
    Dim m_Actividad As Actividad
    Dim m_ID As Variant
    Dim m_Texto As String
    Dim m_TextoTotal  As String
    Dim m_NHoras As String
    Dim m_Coste As String
    Dim m_PrecioHoraAnalista As String
    Dim m_PrecioHoraTecCom As String
    Dim m_PrecioHoraProgramador As String
    Dim m_TipoTecnicoPrecio As TipoTecnicoPrecio
    
    Dim dblNHoras As Double
    Dim dblPrecio As Double
    Dim dblCoste As Double
    Dim dblTotal As Double
    
    On Error GoTo errores
    
    If p_IDEvento = "" Then
        Exit Function
    End If
    
    ' "ANALISTA" INGENIERO DE SISTEMAS
    Set m_TipoTecnicoPrecio = getTipoTecnicoPrecioUltimo(, "INGENIERO DE SISTEMAS", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecio Is Nothing Then
        p_Error = "No se ha podido obtener el precio hora actual de un INGENIERO DE SISTEMAS"
        Err.Raise 1000
    End If
    m_PrecioHoraAnalista = m_TipoTecnicoPrecio.PrecioHoraLab
    If Not IsNumeric(m_PrecioHoraAnalista) Then
        p_Error = "No se ha podido obtener el precio hora actual de un INGENIERO DE SISTEMAS"
        Err.Raise 1000
    End If
    
    
    ' "TÉCNICO EN COMUNICACIONES"
    Set m_TipoTecnicoPrecio = getTipoTecnicoPrecioUltimo(, "TÉCNICO EN COMUNICACIONES", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecio Is Nothing Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO EN COMUNICACIONES"
        Err.Raise 1000
    End If
    m_PrecioHoraTecCom = m_TipoTecnicoPrecio.PrecioHoraLab
    If Not IsNumeric(m_PrecioHoraTecCom) Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO EN COMUNICACIONES"
        Err.Raise 1000
    End If
    
    ' "TÉCNICO DE SISTEMAS" PROGRAMADOR
    Set m_TipoTecnicoPrecio = getTipoTecnicoPrecioUltimo(, "TÉCNICO DE SISTEMAS", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecio Is Nothing Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO DE SISTEMAS"
        Err.Raise 1000
    End If
    m_PrecioHoraProgramador = m_TipoTecnicoPrecio.PrecioHoraLab
    If Not IsNumeric(m_PrecioHoraTecCom) Then
        p_Error = "No se ha podido obtener el precio hora actual de un TÉCNICO DE SISTEMAS"
        Err.Raise 1000
    End If
    
    '    getActividadesGetronic1.Add "ANALISTA", m_NHoras & "|" & m_Coste
    m_SQL = "SELECT Sum(TbActividades.HorasLaborables) AS SumaDeHorasLaborables " & _
            "FROM TbActividades INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
            "WHERE (((TbActividades.IDEvento)='" & p_IDEvento & "') AND ((TbTecnicos.TIPO)='Ingeniero de sistemas'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_NHoras = Nz(.Fields("SumaDeHorasLaborables"), "")
            If IsNumeric(m_NHoras) Then
                m_Coste = CStr(CDbl(m_NHoras) * CDbl(m_PrecioHoraAnalista))
                dblTotal = CDbl(m_Coste)
                m_Texto = "ANALISTA" & ";" & m_NHoras & ";" & m_Coste
                m_TextoTotal = m_Texto
            Else
                m_Texto = "ANALISTA" & ";0;0"
                m_TextoTotal = m_Texto
            End If
           
        Else
            m_Texto = "ANALISTA" & ";0;0"
            m_TextoTotal = m_Texto
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    '    getActividadesGetronic1.Add "TECCOM", , m_NHoras & "|" & m_Coste
    m_SQL = "SELECT Sum(TbActividades.HorasLaborables) AS SumaDeHorasLaborables " & _
            "FROM TbActividades INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
            "WHERE (((TbActividades.IDEvento)='" & p_IDEvento & "') AND ((TbTecnicos.TIPO)='TÉCNICO EN COMUNICACIONES'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            m_NHoras = Nz(.Fields("SumaDeHorasLaborables"), "")
            If IsNumeric(m_NHoras) Then
                m_Coste = CStr(CDbl(m_NHoras) * CDbl(m_PrecioHoraTecCom))
                dblTotal = dblTotal + CDbl(m_Coste)
                m_Texto = "TECCOM" & ";" & m_NHoras & ";" & m_Coste
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            Else
                 m_Texto = "TECCOM" & ";0;0"
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            End If
            
        Else
            m_Texto = "TECCOM" & ";0;0"
            m_TextoTotal = m_TextoTotal & "|" & m_Texto
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    '    getActividadesGetronic1.Add "PROGRAMADOR", , m_NHoras & "|" & m_Coste
    m_SQL = "SELECT Sum(TbActividades.HorasLaborables) AS SumaDeHorasLaborables " & _
            "FROM TbActividades INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
            "WHERE (((TbActividades.IDEvento)='" & p_IDEvento & "') AND ((TbTecnicos.TIPO)='TÉCNICO DE SISTEMAS'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            m_NHoras = Nz(.Fields("SumaDeHorasLaborables"), "")
            If IsNumeric(m_NHoras) Then
                m_Coste = CStr(CDbl(m_NHoras) * CDbl(m_PrecioHoraProgramador))
                dblTotal = dblTotal + CDbl(m_Coste)
                m_Texto = "PROGRAMADOR" & ";" & m_NHoras & ";" & m_Coste
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            Else
                m_Texto = "PROGRAMADOR" & ";0;0"
                m_TextoTotal = m_TextoTotal & "|" & m_Texto
            End If
            
        Else
            m_Texto = "PROGRAMADOR" & ";0;0"
            m_TextoTotal = m_TextoTotal & "|" & m_Texto
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_TextoTotal = m_TextoTotal & "|" & dblTotal
    getActividadesGetronic1 = m_TextoTotal
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getActividadesGetronic1 ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEventosGetronics( _
                                    p_ColEventos As Scripting.Dictionary, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim m_Evento As Evento
    Dim m_ID As Variant
    
    On Error GoTo errores
    If p_ColEventos Is Nothing Then
        Exit Function
    End If
    For Each m_ID In p_ColEventos
        Set m_Evento = p_ColEventos(m_ID)
        If Left(m_Evento.Descripcion, 3) = "PT-" Then
            If getEventosGetronics Is Nothing Then
                Set getEventosGetronics = New Scripting.Dictionary
                getEventosGetronics.CompareMode = TextCompare
            End If
            If Not getEventosGetronics.Exists(CStr(m_Evento.IDEVENTO)) Then
                getEventosGetronics.Add CStr(m_Evento.IDEVENTO), m_Evento
            End If
        End If
        Set m_Evento = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEventosGetronics ha devuelto el error: " & Err.Description
    End If
End Function

