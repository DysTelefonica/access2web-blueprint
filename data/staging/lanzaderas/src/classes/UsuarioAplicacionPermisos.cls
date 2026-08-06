VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioAplicacionPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public CorreoUsuario As String
Public IDAplicacion As String
Public EsUsuarioAdministrador As String
Public EsUsuarioCalidad As String
Public EsUsuarioEconomia As String
Public EsUsuarioSecretaria As String
Public EsUsuarioTecnico As String
Public EsUsuarioSinAcceso As String
Public EsUsuarioCalidadAvisos As String

Private m_ObjUsuario As usuario
Private m_objAplicacion As Aplicacion

Private m_eEsUsuarioAdministradorCalculado As EnumSino
Private m_eEsUsuarioCalidadCalculado As EnumSino
Private m_eEsUsuarioEconomiaCalculado As EnumSino
Private m_eEsUsuarioSecretariaCalculado As EnumSino
Private m_eEsUsuarioTecnicoCalculado As EnumSino
Private m_eEsUsuarioSinAccesoCalculado As EnumSino
Private m_eEsUsuarioCalidadAvisosCalculado As EnumSino

Private m_objColCampos As Collection
Public Error As String


Public Property Get usuario() As usuario

    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjUsuario Is Nothing Then
        Set usuario = m_ObjUsuario
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_ObjUsuario = Constructor.getUsuario(, , , Me.CorreoUsuario, Me.Error)
    Set usuario = m_ObjUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método UsuarioAplicacionPermiso.Usuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Set usuario(ByVal objNewValue As usuario)

    Set m_ObjUsuario = objNewValue

End Property

Public Property Get Aplicacion() As Aplicacion

    On Error GoTo errores
    Me.Error = ""
    If Not m_objAplicacion Is Nothing Then
        Set Aplicacion = m_objAplicacion
        Exit Property
    End If
    If Me.IDAplicacion = "" Then
        Exit Property
    End If
    Set m_objAplicacion = Constructor.getAplicacion(Me.IDAplicacion, Me.Error)
    Set Aplicacion = m_objAplicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.Aplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Set Aplicacion(ByVal objNewValue As Aplicacion)

    Set m_objAplicacion = objNewValue

End Property

Public Property Get EsUsuarioAdministradorCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsUsuarioAdministrador = "Sí" Then
        m_eEsUsuarioAdministradorCalculado = EnumSino.Sí
    ElseIf Me.EsUsuarioAdministrador = "No" Then
        m_eEsUsuarioAdministradorCalculado = EnumSino.No
    End If
    
    EsUsuarioAdministradorCalculado = m_eEsUsuarioAdministradorCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.EsUsuarioAdministradorCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Let EsUsuarioAdministradorCalculado(ByVal eNewValue As EnumSino)

    m_eEsUsuarioAdministradorCalculado = eNewValue

End Property

Public Property Get EsUsuarioCalidadCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsUsuarioCalidad = "Sí" Then
        m_eEsUsuarioCalidadCalculado = EnumSino.Sí
    ElseIf Me.EsUsuarioCalidad = "No" Then
        m_eEsUsuarioCalidadCalculado = EnumSino.No
    End If
    
    EsUsuarioCalidadCalculado = m_eEsUsuarioCalidadCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.EsUsuarioCalidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Let EsUsuarioCalidadCalculado(ByVal eNewValue As EnumSino)

    m_eEsUsuarioCalidadCalculado = eNewValue

End Property

Public Property Get EsUsuarioEconomiaCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsUsuarioEconomia = "Sí" Then
        m_eEsUsuarioEconomiaCalculado = EnumSino.Sí
    ElseIf Me.EsUsuarioEconomia = "No" Then
        m_eEsUsuarioEconomiaCalculado = EnumSino.No
    End If
    
    EsUsuarioEconomiaCalculado = m_eEsUsuarioEconomiaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.EsUsuarioEconomiaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Let EsUsuarioEconomiaCalculado(ByVal eNewValue As EnumSino)

    m_eEsUsuarioEconomiaCalculado = eNewValue

End Property

Public Property Get EsUsuarioSecretariaCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsUsuarioSecretaria = "Sí" Then
        m_eEsUsuarioSecretariaCalculado = EnumSino.Sí
    ElseIf Me.EsUsuarioSecretaria = "No" Then
        m_eEsUsuarioSecretariaCalculado = EnumSino.No
    End If
    
    EsUsuarioSecretariaCalculado = m_eEsUsuarioSecretariaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.EsUsuarioSecretariaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Let EsUsuarioSecretariaCalculado(ByVal eNewValue As EnumSino)

    m_eEsUsuarioSecretariaCalculado = eNewValue

End Property

Public Property Get EsUsuarioTecnicoCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsUsuarioTecnico = "Sí" Then
        m_eEsUsuarioTecnicoCalculado = EnumSino.Sí
    ElseIf Me.EsUsuarioTecnico = "No" Then
        m_eEsUsuarioTecnicoCalculado = EnumSino.No
    End If
    
    EsUsuarioTecnicoCalculado = m_eEsUsuarioTecnicoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.EsUsuarioTecnicoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Let EsUsuarioTecnicoCalculado(ByVal eNewValue As EnumSino)

    m_eEsUsuarioTecnicoCalculado = eNewValue

End Property

Public Property Get EsUsuarioSinAccesoCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsUsuarioSinAcceso = "Sí" Then
        m_eEsUsuarioSinAccesoCalculado = EnumSino.Sí
    ElseIf Me.EsUsuarioSinAcceso = "No" Then
        m_eEsUsuarioSinAccesoCalculado = EnumSino.No
    End If
    
    EsUsuarioSinAccesoCalculado = m_eEsUsuarioSinAccesoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.EsUsuarioSinAccesoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Let EsUsuarioSinAccesoCalculado(ByVal eNewValue As EnumSino)

    m_eEsUsuarioSinAccesoCalculado = eNewValue

End Property

Public Property Get EsUsuarioCalidadAvisosCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsUsuarioCalidadAvisos = "Sí" Then
        m_eEsUsuarioCalidadAvisosCalculado = EnumSino.Sí
    ElseIf Me.EsUsuarioCalidadAvisos = "No" Then
        m_eEsUsuarioCalidadAvisosCalculado = EnumSino.No
    End If
    
    EsUsuarioCalidadAvisosCalculado = m_eEsUsuarioCalidadAvisosCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método AplicacionAplicacionPermiso.EsUsuarioCalidadAvisosCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Let EsUsuarioCalidadAvisosCalculado(ByVal eNewValue As EnumSino)

    m_eEsUsuarioCalidadAvisosCalculado = eNewValue

End Property

Public Function RegistrarPermiso( _
                                p_EsAlta As EnumSino, _
                                p_TipoPermiso As EnumPermisos, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_NombreCampo As String
    On Error GoTo errores
    Me.Error = ""
    If Me.CorreoUsuario = "" Then
        p_Error = "No se sabe el usuario"
        Err.Raise 1000
    End If
    If Me.IDAplicacion = "" Then
        p_Error = "No se sabe la aplicacion"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno.ColPermisosTipo.Exists(CStr(p_TipoPermiso)) Then
        p_Error = "Permiso no reconocido"
        Err.Raise 1000
    End If
    m_NombreCampo = m_ObjEntorno.ColPermisosTipo(CStr(p_TipoPermiso))
    
    
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & Me.CorreoUsuario & "' AND IDAplicacion=" & Me.IDAplicacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("CorreoUsuario") = Me.CorreoUsuario
                .Fields("IDAplicacion") = Me.IDAplicacion
                
                If p_EsAlta = EnumSino.Sí Then
                    .Fields(m_NombreCampo) = "Sí"
                Else
                    .Fields(m_NombreCampo) = "No"
                End If
            .Update
        Else
            .Edit
                .Fields(m_NombreCampo) = "Sí"
                If p_EsAlta = EnumSino.Sí Then
                    .Fields(m_NombreCampo) = "Sí"
                Else
                    .Fields(m_NombreCampo) = "No"
                End If
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioAplicacionPermisos.RegistrarPermiso ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RegistrarPermisosCompletos( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    Me.Error = ""
    If Me.CorreoUsuario = "" Then
        p_Error = "No se sabe el usuario"
        Err.Raise 1000
    End If
    If Me.IDAplicacion = "" Then
        p_Error = "No se sabe la aplicacion"
        Err.Raise 1000
    End If
    
    
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & Me.CorreoUsuario & "' AND IDAplicacion=" & Me.IDAplicacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("CorreoUsuario") = Me.CorreoUsuario
                .Fields("IDAplicacion") = Me.IDAplicacion
                .Fields("EsUsuarioAdministrador") = Me.EsUsuarioAdministrador
                .Fields("EsUsuarioCalidad") = Me.EsUsuarioCalidad
                .Fields("EsUsuarioEconomia") = Me.EsUsuarioEconomia
                .Fields("EsUsuarioSecretaria") = Me.EsUsuarioSecretaria
                .Fields("EsUsuarioTecnico") = Me.EsUsuarioTecnico
                .Fields("EsUsuarioCalidadAvisos") = Me.EsUsuarioCalidadAvisos
                
                .Fields("EsUsuarioSinAcceso") = "No"
                
            .Update
        Else
            .Edit
                If Me.EsUsuarioAdministradorCalculado = EnumSino.Sí Then
                    .Fields("EsUsuarioAdministrador") = "Sí"
                Else
                    .Fields("EsUsuarioAdministrador") = "No"
                End If
                If Me.EsUsuarioCalidadCalculado = EnumSino.Sí Then
                    .Fields("EsUsuarioCalidad") = "Sí"
                Else
                    .Fields("EsUsuarioCalidad") = "No"
                End If
                If Me.EsUsuarioEconomiaCalculado = EnumSino.Sí Then
                    .Fields("EsUsuarioEconomia") = "Sí"
                Else
                    .Fields("EsUsuarioEconomia") = "No"
                End If
                If Me.EsUsuarioSecretariaCalculado = EnumSino.Sí Then
                    .Fields("EsUsuarioSecretaria") = "Sí"
                Else
                    .Fields("EsUsuarioSecretaria") = "No"
                End If
                If Me.EsUsuarioTecnicoCalculado = EnumSino.Sí Then
                    .Fields("EsUsuarioTecnico") = "Sí"
                Else
                    .Fields("EsUsuarioTecnico") = "No"
                End If
                If Me.EsUsuarioCalidadAvisosCalculado = EnumSino.Sí Then
                    .Fields("EsUsuarioCalidadAvisos") = "Sí"
                Else
                    .Fields("EsUsuarioCalidadAvisos") = "No"
                End If
                .Fields("EsUsuarioSinAcceso") = "No"
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioAplicacionPermisos.RegistrarPermisosCompletos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    Me.Error = ""
    If Me.CorreoUsuario = "" Then
        p_Error = "No se sabe el usuario"
        Err.Raise 1000
    End If
    If Me.IDAplicacion = "" Then
        p_Error = "No se sabe la aplicacion"
        Err.Raise 1000
    End If
   
    
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & Me.CorreoUsuario & "' AND IDAplicacion=" & Me.IDAplicacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .Delete
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioAplicacionPermisos.Borrar ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Property Get ColCampos() As Collection
    
    Dim tdf As DAO.TableDef
    Dim fld As Object
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set tdf = getdb().TableDefs("TbUsuariosAplicacionesPermisos")
    For Each fld In tdf.Fields
        If m_objColCampos Is Nothing Then
            Set m_objColCampos = New Collection
        
        End If
        m_objColCampos.Add fld.Name
    Next
    Set tdf = Nothing
    
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El metodo UsuarioAplicacionPermisos.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "CorreoUsuario" Then
        getPropiedad = Me.CorreoUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        getPropiedad = Me.IDAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioAdministrador" Then
        getPropiedad = Me.EsUsuarioAdministrador
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidad" Then
        getPropiedad = Me.EsUsuarioCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioEconomia" Then
        getPropiedad = Me.EsUsuarioEconomia
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioSecretaria" Then
        getPropiedad = Me.EsUsuarioSecretaria
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioTecnico" Then
        getPropiedad = Me.EsUsuarioTecnico
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioSinAcceso" Then
        getPropiedad = Me.EsUsuarioSinAcceso
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidadAvisos" Then
        getPropiedad = Me.EsUsuarioCalidadAvisos
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en UsuarioAplicacionPermisos.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo UsuarioAplicacionPermisos.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "CorreoUsuario" Then
        Me.CorreoUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        Me.IDAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioAdministrador" Then
        Me.EsUsuarioAdministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidad" Then
        Me.EsUsuarioCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioEconomia" Then
        Me.EsUsuarioEconomia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioSecretaria" Then
        Me.EsUsuarioSecretaria = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioTecnico" Then
        Me.EsUsuarioTecnico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioSinAcceso" Then
        Me.EsUsuarioSinAcceso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidadAvisos" Then
        Me.EsUsuarioCalidadAvisos = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en UsuarioAplicacionPermisos.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo UsuarioAplicacionPermisos.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function











