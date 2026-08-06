VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MotivoHPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public MotivoHPS As String
Private m_objColUsuariosConMotivo As Scripting.Dictionary
Private m_objColSolicitudesConMotivo As Scripting.Dictionary
Private m_sCadenaUsuariosConMotivo As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Function EditarMotivo( _
                                p_MotivoAlInicio As String, _
                                p_Motivo As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    If p_MotivoAlInicio = "" Then
        Exit Function
    End If
    If m_EnOficina = EnumSiNo.no Then
        p_Error = "Esta acción sólo se puede hacer en la oficina"
        Err.Raise 1000
    End If
    m_SQL = "UPDATE TbMotivoHPS SET MotivoHPS = '" & p_Motivo & "' " & _
                "WHERE MotivoHPS='" & p_MotivoAlInicio & "';"
    
    getdb().Execute m_SQL
    ActualizarSolicitudesMotivo p_MotivoAlInicio, p_Motivo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ActualizarUsuariosHPSMotivo p_MotivoAlInicio, p_Motivo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EditarMotivo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function AltaMotivo( _
                            Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    'TODO:
    On Error GoTo errores
    If Me.MotivoHPS = "" Then
        p_Error = "Se ha de indicar un motivo"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMotivoHPS.MotivoHPS " & _
            "FROM TbMotivoHPS " & _
            "WHERE (((TbMotivoHPS.MotivoHPS)='" & Me.MotivoHPS & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Not rcdDatos.EOF Then
        p_Error = "Ese motivo ya estaba registrado"
        Err.Raise 1000
    End If
    rcdDatos.AddNew
        rcdDatos.Fields("MotivoHPS") = Me.MotivoHPS
    rcdDatos.Update
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaMotivo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EliminarMotivo( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_SQL As String
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Usuario As UsuarioHPS
    Dim m_Solicitud As solicitud
    Dim m_Resultado As String
    
    On Error GoTo errores
    If Me.MotivoHPS = "" Then
        p_Error = "Se ha de indicar un motivo"
        Err.Raise 1000
    End If
    Set m_Col = constructor.getUsuariosConMotivo(p_Motivo:=Me.MotivoHPS, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Usuario = m_Col(m_ID)
            If m_Resultado = "" Then
                m_Resultado = m_Usuario.Nombre
            Else
                m_Resultado = m_Resultado & vbNewLine & m_Usuario.Nombre
            End If
            Set m_Usuario = Nothing
        Next
        p_Error = "No se puede eliminar ya que los siguientes usuarios HPS lo tienen" & vbNewLine & m_Resultado
        Err.Raise 1000
    End If
    m_Resultado = ""
    Set m_Col = constructor.getSolicitudesConMotivo(p_Motivo:=Me.MotivoHPS, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Solicitud = m_Col(m_ID)
            If m_Resultado = "" Then
                m_Resultado = m_Solicitud.IDSolicitud
            Else
                m_Resultado = m_Resultado & vbNewLine & m_Solicitud.IDSolicitud
            End If
            Set m_Solicitud = Nothing
        Next
        p_Error = "No se puede eliminar ya que las siguientes solicitudes lo tienen" & vbNewLine & m_Resultado
        Err.Raise 1000
    End If
    
    'comprobar si está en otros sitios
    
    m_SQL = "DELETE TbMotivoHPS.MotivoHPS " & _
            "FROM TbMotivoHPS " & _
            "WHERE (((TbMotivoHPS.MotivoHPS)='" & Me.MotivoHPS & "'));"
    getdb().Execute m_SQL
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarMotivo ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ActualizarSolicitudesMotivo( _
                                            p_MotivoAlInicio As String, _
                                            p_Motivo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    If p_MotivoAlInicio = "" Then
        Exit Function
    End If
    m_SQL = "UPDATE TbSolicitudes SET Motivo_HPS = '" & p_Motivo & "' " & _
            "WHERE Motivo_HPS='" & p_MotivoAlInicio & "';"
    getdb().Execute m_SQL
    
        
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarSolicitudesMotivo ha devuelto el error: " & Err.Description
    End If
End Function
Private Function ActualizarUsuariosHPSMotivo( _
                                            p_MotivoAlInicio As String, _
                                            p_Motivo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    If p_MotivoAlInicio = "" Then
        Exit Function
    End If
    m_SQL = "UPDATE TbUsuarios SET Motivo_HPS = '" & p_Motivo & "' " & _
            "WHERE Motivo_HPS='" & p_MotivoAlInicio & "';"
    getdb().Execute m_SQL
    m_SQL = "UPDATE TbUsuariosHistoricos SET Motivo_HPS = '" & p_Motivo & "' " & _
            "WHERE Motivo_HPS='" & p_MotivoAlInicio & "';"
    getdb().Execute m_SQL
        
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarUsuariosHPSMotivo ha devuelto el error: " & Err.Description
    End If
End Function
Public Property Get ColUsuariosConMotivo() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosConMotivo Is Nothing Then
        Set ColUsuariosConMotivo = m_objColUsuariosConMotivo
        Exit Property
    End If
    Set m_objColUsuariosConMotivo = constructor.getUsuariosConMotivo(p_Motivo:=Me.MotivoHPS, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColUsuariosConMotivo = m_objColUsuariosConMotivo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método MotivoHPS.ColUsuariosConMotivo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get ColSolicitudesConMotivo() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColSolicitudesConMotivo Is Nothing Then
        Set ColSolicitudesConMotivo = m_objColSolicitudesConMotivo
        Exit Property
    End If
    Set m_objColSolicitudesConMotivo = constructor.getSolicitudesConMotivo(p_Motivo:=Me.MotivoHPS, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColSolicitudesConMotivo = m_objColSolicitudesConMotivo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método MotivoHPS.ColSolicitudesConMotivo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "MotivoHPS"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método MotivoHPS.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "MotivoHPS" Then
        getPropiedad = Me.MotivoHPS
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en MotivoHPS.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MotivoHPS.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "MotivoHPS" Then
        Me.MotivoHPS = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en MotivoHPS.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MotivoHPS.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

