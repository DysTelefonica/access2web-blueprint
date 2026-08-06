VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventoEquipoMedidaAntes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDEventoEquipoMedida As String
Public IDEVENTO As String
Public IDEquipoMedida As String

Private m_sEstadoCalibracion As String
Private m_sIDEventoEquipoMedidaCalculado As String
Private m_objEquipoMedida As EquipoMedida
Private m_objCalibracionEnAltaEvento As EquipoMedidaCalibracion

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String
Public Property Get IDEventoEquipoMedidaCalculado() As String
    
     On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    m_sIDEventoEquipoMedidaCalculado = DameID("TbEquiposEquipoMedida", "IDEventoEquipoMedida")
   
    IDEventoEquipoMedidaCalculado = m_sIDEventoEquipoMedidaCalculado
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EventoEquipoMedida.IDEventoEquipoMedidaCalculado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get EquipoMedida() As EquipoMedida

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objEquipoMedida Is Nothing Then
        Set EquipoMedida = m_objEquipoMedida
        Exit Property
    End If
    Set m_objEquipoMedida = Constructor.getEquipoMedida(Me.IDEquipoMedida, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set EquipoMedida = m_objEquipoMedida
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EventoEquipoMedida.EquipoMedida ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set EquipoMedida(ByVal objNewValue As EquipoMedida)

    Set m_objEquipoMedida = objNewValue

End Property

Public Property Get CalibracionEnAltaEvento() As EquipoMedidaCalibracion
    
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objCalibracionEnAltaEvento Is Nothing Then
        Set CalibracionEnAltaEvento = m_objCalibracionEnAltaEvento
        Exit Property
    End If
    Set m_objCalibracionEnAltaEvento = getCalibracionPorFecha(Me.IDEquipoMedida, Me.IDEVENTO, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set CalibracionEnAltaEvento = m_objCalibracionEnAltaEvento
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EventoEquipoMedida.CalibracionEnAltaEvento ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "IDEventoEquipoMedida"
        .Add "IDEvento"
        .Add "IDEquipoMedida"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EventoEquipoMedida.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDEventoEquipoMedida" Then
        getPropiedad = Me.IDEventoEquipoMedida
        Exit Function
    End If
    If m_NombrePropiedad = "IDEvento" Then
        getPropiedad = Me.IDEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipoMedida" Then
        getPropiedad = Me.IDEquipoMedida
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EventoEquipoMedida.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EventoEquipoMedida.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDEventoEquipoMedida" Then
        Me.IDEventoEquipoMedida = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEvento" Then
        Me.IDEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipoMedida" Then
        Me.IDEquipoMedida = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EventoEquipoMedida.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EventoEquipoMedida.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


