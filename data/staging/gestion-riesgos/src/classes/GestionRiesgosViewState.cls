VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestionRiesgosViewState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' GestionRiesgosViewState (issue #41)
' Pure result DTO for Form_FormGestionRiesgos view-state decisions.
' The form keeps object loading and all Access UI/control mutation.
' ============================================================

Private m_Titulo As String
Private m_NavResponsablesEnabled As Boolean
Private m_NavRiesgosOfertaEnabled As Boolean
Private m_NavSuministradoresEnabled As Boolean
Private m_NavDatosGeneralesEnabled As Boolean
Private m_ComandoGrabarEnabled As Boolean
Private m_PermitidoEditar As Boolean
Private m_Error As String

Public Property Get Titulo() As String
    Titulo = m_Titulo
End Property

Public Property Let Titulo(ByVal p_Value As String)
    m_Titulo = p_Value
End Property

Public Property Get NavResponsablesEnabled() As Boolean
    NavResponsablesEnabled = m_NavResponsablesEnabled
End Property

Public Property Let NavResponsablesEnabled(ByVal p_Value As Boolean)
    m_NavResponsablesEnabled = p_Value
End Property

Public Property Get NavRiesgosOfertaEnabled() As Boolean
    NavRiesgosOfertaEnabled = m_NavRiesgosOfertaEnabled
End Property

Public Property Let NavRiesgosOfertaEnabled(ByVal p_Value As Boolean)
    m_NavRiesgosOfertaEnabled = p_Value
End Property

Public Property Get NavSuministradoresEnabled() As Boolean
    NavSuministradoresEnabled = m_NavSuministradoresEnabled
End Property

Public Property Let NavSuministradoresEnabled(ByVal p_Value As Boolean)
    m_NavSuministradoresEnabled = p_Value
End Property

Public Property Get NavDatosGeneralesEnabled() As Boolean
    NavDatosGeneralesEnabled = m_NavDatosGeneralesEnabled
End Property

Public Property Let NavDatosGeneralesEnabled(ByVal p_Value As Boolean)
    m_NavDatosGeneralesEnabled = p_Value
End Property

Public Property Get ComandoGrabarEnabled() As Boolean
    ComandoGrabarEnabled = m_ComandoGrabarEnabled
End Property

Public Property Let ComandoGrabarEnabled(ByVal p_Value As Boolean)
    m_ComandoGrabarEnabled = p_Value
End Property

Public Property Get PermitidoEditar() As Boolean
    PermitidoEditar = m_PermitidoEditar
End Property

Public Property Let PermitidoEditar(ByVal p_Value As Boolean)
    m_PermitidoEditar = p_Value
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = m_Error
End Property

Public Property Let ErrorMessage(ByVal p_Value As String)
    m_Error = p_Value
End Property
