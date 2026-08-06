VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Aspecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private m_Error As String

Private m_eEstado As EnumModoApertura
Private m_lCabeceraBackColor As Long
Private m_lListaBackColor As Long
Private m_lListaForeColor As Long
Private m_lListaBorderColor As Long
Private m_lEtiquetaBackColor As Long
Private m_lEtiquetaForeColor As Long
Private m_lEtiquetaBorderColor As Long
Private m_lCampoBackColor As Long
Private m_lCampoForeColor As Long
Private m_lCampoBorderColor As Long
Private m_sError As String
Private m_lEtiquetaBackColorSeleccionado As Long
Private m_lEtiquetaFontColorSeleccionado As Long


Public Property Get Estado() As EnumModoApertura
    
    On Error GoTo errores
    Me.Error = ""
    
    If m_eEstado <> 0 Then
        Estado = m_eEstado
        Exit Property
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_eEstado = EnumModoApertura.slave
    Else
        m_eEstado = EnumModoApertura.master
    End If
    If m_eEstado <> 0 Then
        Estado = m_eEstado
        Exit Property
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Aspecto.Estado ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get CabeceraBackColor() As Long

    CabeceraBackColor = m_lCabeceraBackColor

End Property

Public Property Let CabeceraBackColor(ByVal lNewValue As Long)

    m_lCabeceraBackColor = lNewValue

End Property

Public Property Get ListaBackColor() As Long

    ListaBackColor = m_lListaBackColor

End Property

Public Property Let ListaBackColor(ByVal lNewValue As Long)

    m_lListaBackColor = lNewValue

End Property

Public Property Get ListaForeColor() As Long

    ListaForeColor = m_lListaForeColor

End Property

Public Property Let ListaForeColor(ByVal lNewValue As Long)

    m_lListaForeColor = lNewValue

End Property

Public Property Get ListaBorderColor() As Long

    ListaBorderColor = m_lListaBorderColor

End Property

Public Property Let ListaBorderColor(ByVal lNewValue As Long)

    m_lListaBorderColor = lNewValue

End Property

Public Property Get EtiquetaBackColor() As Long

    EtiquetaBackColor = m_lEtiquetaBackColor

End Property

Public Property Let EtiquetaBackColor(ByVal lNewValue As Long)

    m_lEtiquetaBackColor = lNewValue

End Property

Public Property Get EtiquetaForeColor() As Long

    EtiquetaForeColor = m_lEtiquetaForeColor

End Property

Public Property Let EtiquetaForeColor(ByVal lNewValue As Long)

    m_lEtiquetaForeColor = lNewValue

End Property

Public Property Get EtiquetaBorderColor() As Long

    EtiquetaBorderColor = m_lEtiquetaBorderColor

End Property

Public Property Let EtiquetaBorderColor(ByVal lNewValue As Long)

    m_lEtiquetaBorderColor = lNewValue

End Property

Public Property Get CampoBackColor() As Long

    CampoBackColor = m_lCampoBackColor

End Property

Public Property Let CampoBackColor(ByVal lNewValue As Long)

    m_lCampoBackColor = lNewValue

End Property

Public Property Get CampoForeColor() As Long

    CampoForeColor = m_lCampoForeColor

End Property

Public Property Let CampoForeColor(ByVal lNewValue As Long)

    m_lCampoForeColor = lNewValue

End Property

Public Property Get CampoBorderColor() As Long

    CampoBorderColor = m_lCampoBorderColor

End Property

Public Property Let CampoBorderColor(ByVal lNewValue As Long)

    m_lCampoBorderColor = lNewValue

End Property

Public Property Get EtiquetaFontColorSeleccionado() As Long

    EtiquetaFontColorSeleccionado = m_lEtiquetaFontColorSeleccionado

End Property

Public Property Let EtiquetaFontColorSeleccionado(ByVal lNewValue As Long)

    m_lEtiquetaFontColorSeleccionado = lNewValue

End Property

Public Property Get EtiquetaBackColorSeleccionado() As Long

    EtiquetaBackColorSeleccionado = m_lEtiquetaBackColorSeleccionado

End Property

Public Property Let EtiquetaBackColorSeleccionado(ByVal lNewValue As Long)

    m_lEtiquetaBackColorSeleccionado = lNewValue

End Property

Public Property Get Error() As String

    Error = m_sError

End Property

Public Property Let Error(ByVal sNewValue As String)

    m_sError = sNewValue

End Property



