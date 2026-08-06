VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoEstadoCalidadViewState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' RiesgoEstadoCalidadViewState (issue #43)
' Pure result DTO for acceptance/retirement quality button state.
' Forms keep UI mutation, prompts, cache refreshes and persistence.
' ============================================================

Private m_ComandoAceptarVisible As Boolean
Private m_ComandoRechazarVisible As Boolean
Private m_ComandoAceptarEnabled As Boolean
Private m_ComandoRechazarEnabled As Boolean
Private m_ComandoAceptarCaption As String
Private m_ComandoRechazarCaption As String
Private m_HasButtonDecision As Boolean
Private m_Error As String

Public Property Get ComandoAceptarVisible() As Boolean
    ComandoAceptarVisible = m_ComandoAceptarVisible
End Property

Public Property Let ComandoAceptarVisible(ByVal p_Value As Boolean)
    m_ComandoAceptarVisible = p_Value
End Property

Public Property Get ComandoRechazarVisible() As Boolean
    ComandoRechazarVisible = m_ComandoRechazarVisible
End Property

Public Property Let ComandoRechazarVisible(ByVal p_Value As Boolean)
    m_ComandoRechazarVisible = p_Value
End Property

Public Property Get ComandoAceptarEnabled() As Boolean
    ComandoAceptarEnabled = m_ComandoAceptarEnabled
End Property

Public Property Let ComandoAceptarEnabled(ByVal p_Value As Boolean)
    m_ComandoAceptarEnabled = p_Value
End Property

Public Property Get ComandoRechazarEnabled() As Boolean
    ComandoRechazarEnabled = m_ComandoRechazarEnabled
End Property

Public Property Let ComandoRechazarEnabled(ByVal p_Value As Boolean)
    m_ComandoRechazarEnabled = p_Value
End Property

Public Property Get ComandoAceptarCaption() As String
    ComandoAceptarCaption = m_ComandoAceptarCaption
End Property

Public Property Let ComandoAceptarCaption(ByVal p_Value As String)
    m_ComandoAceptarCaption = p_Value
End Property

Public Property Get ComandoRechazarCaption() As String
    ComandoRechazarCaption = m_ComandoRechazarCaption
End Property

Public Property Let ComandoRechazarCaption(ByVal p_Value As String)
    m_ComandoRechazarCaption = p_Value
End Property

Public Property Get HasButtonDecision() As Boolean
    HasButtonDecision = m_HasButtonDecision
End Property

Public Property Let HasButtonDecision(ByVal p_Value As Boolean)
    m_HasButtonDecision = p_Value
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = m_Error
End Property

Public Property Let ErrorMessage(ByVal p_Value As String)
    m_Error = p_Value
End Property
