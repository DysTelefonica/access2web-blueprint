VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FacturaPerfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDFacturaPerfil As String
Public IDFactura As String
Public Perfil As String
Public PrecioHoraLab As String
Public PrecioHoraExt As String
Public NHorasLab As String
Public NHorasExt As String

Private m_sImporteHorasLab As String
Private m_sImporteHorasExt As String
Private m_sImporte As String

Private m_ObjTipoTecnico As TIPOTECNICO
Private m_sPrecioHoraLabCalculado As String
Private m_sPrecioHoraExtCalculado As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get TIPOTECNICO() As TIPOTECNICO
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_ObjTipoTecnico Is Nothing Then
        Set TIPOTECNICO = m_ObjTipoTecnico
        Exit Property
    End If
    If Me.Perfil = "" Then
        Exit Property
    End If
    Set m_ObjTipoTecnico = New TIPOTECNICO
    m_ObjTipoTecnico.Tipo = Me.Perfil
   
    Set TIPOTECNICO = m_ObjTipoTecnico
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método FacturaPerfiles.TipoTecnico ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get ImporteHorasLab() As String

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If IsNumeric(m_sImporteHorasLab) Then
        ImporteHorasLab = m_sImporteHorasLab
        Exit Property
    End If
    If IsNumeric(Me.NHorasLab) And IsNumeric(Me.PrecioHoraLab) Then
        m_sImporteHorasLab = CStr(CDbl(Me.NHorasLab) * CDbl(Me.PrecioHoraLab))
    End If
    
    ImporteHorasLab = m_sImporteHorasLab
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método FacturaPerfiles.ImporteHorasLab ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Get ImporteHorasExt() As String

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If IsNumeric(m_sImporteHorasExt) Then
        ImporteHorasExt = m_sImporteHorasExt
        Exit Property
    End If
    If IsNumeric(Me.NHorasExt) And IsNumeric(Me.PrecioHoraExt) Then
        m_sImporteHorasExt = CStr(CDbl(Me.NHorasExt) * CDbl(Me.PrecioHoraExt))
    End If
    
    ImporteHorasExt = m_sImporteHorasExt
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método FacturaPerfiles.ImporteHorasExt ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Get Importe() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If IsNumeric(m_sImporte) Then
        Importe = m_sImporte
        Exit Property
    End If
    If IsNumeric(Me.ImporteHorasLab) And IsNumeric(Me.ImporteHorasExt) Then
        m_sImporte = CStr(CDbl(Me.ImporteHorasLab) + CDbl(Me.ImporteHorasExt))
    End If
    
    Importe = m_sImporte
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método FacturaPerfiles.Importe ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get PrecioHoraLabCalculado() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If IsNumeric(m_sPrecioHoraLabCalculado) Then
        PrecioHoraLabCalculado = m_sPrecioHoraLabCalculado
        Exit Property
    End If
    
    If Me.TIPOTECNICO Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_sPrecioHoraLabCalculado = Me.TIPOTECNICO.PrecioHoraLabActual
    m_Error = Me.TIPOTECNICO.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PrecioHoraLabCalculado = m_sPrecioHoraLabCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método FacturaPerfiles.PrecioHoraLabCalculado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get PrecioHoraExtCalculado() As String

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If IsNumeric(m_sPrecioHoraExtCalculado) Then
        PrecioHoraExtCalculado = m_sPrecioHoraExtCalculado
        Exit Property
    End If
    
    If Me.TIPOTECNICO Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_sPrecioHoraExtCalculado = Me.TIPOTECNICO.PrecioHoraLabActual
    m_Error = Me.TIPOTECNICO.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PrecioHoraExtCalculado = m_sPrecioHoraExtCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método FacturaPerfiles.PrecioHoraExtCalculado ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "IDFacturaPerfil"
        .Add "IDFactura"
        .Add "Perfil"
        .Add "PrecioHoraLab"
        .Add "PrecioHoraExt"
        .Add "NHorasLab"
        .Add "NHorasExt"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método FacturaPerfiles.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDFacturaPerfil" Then
        getPropiedad = Me.IDFacturaPerfil
        Exit Function
    End If
    If m_NombrePropiedad = "IDFactura" Then
        getPropiedad = Me.IDFactura
        Exit Function
    End If
    If m_NombrePropiedad = "Perfil" Then
        getPropiedad = Me.Perfil
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraLab" Then
        getPropiedad = Me.PrecioHoraLab
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraExt" Then
        getPropiedad = Me.PrecioHoraExt
        Exit Function
    End If
    If m_NombrePropiedad = "NHorasLab" Then
        getPropiedad = Me.NHorasLab
        Exit Function
    End If
    If m_NombrePropiedad = "NHorasExt" Then
        getPropiedad = Me.NHorasExt
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en FacturaPerfiles.getPropiedad"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FacturaPerfiles.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDFacturaPerfil" Then
        Me.IDFacturaPerfil = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDFactura" Then
        Me.IDFactura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Perfil" Then
        Me.Perfil = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraLab" Then
        Me.PrecioHoraLab = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraExt" Then
        Me.PrecioHoraExt = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NHorasLab" Then
        Me.NHorasLab = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NHorasExt" Then
        Me.NHorasExt = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en FacturaPerfiles.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FacturaPerfiles.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


