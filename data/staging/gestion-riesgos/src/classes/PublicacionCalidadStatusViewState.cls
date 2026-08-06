VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublicacionCalidadStatusViewState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' PublicacionCalidadStatusViewState (issue #46)
' Pure DTO for FormPublicacionCalidadPublicar status decisions.
' Forms keep UI mutation, prompts, DoCmd, cache and persistence.
' ============================================================

Private m_UteRequeridoCaption As String
Private m_UteComplementadoCaption As String
Private m_UteComplementadoColorRole As String
Private m_SuministradoresRequeridoCaption As String
Private m_SuministradoresComplementadoCaption As String
Private m_SuministradoresComplementadoColorRole As String
Private m_ComandoRechazarEnabled As Boolean
Private m_ComandoPublicarEnabled As Boolean
Private m_ComandoPublicarCaption As String
Private m_ComandoGenerarInformePublicabilidadEnabled As Boolean
Private m_ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled As Boolean
Private m_ComandoVerRechazoVisible As Boolean
Private m_HasComandoVerRechazoVisibleDecision As Boolean
Private m_PublicacionVisible As Boolean
Private m_PublicacionCaption As String
Private m_PublicacionColorRole As String

Public Property Get UteRequeridoCaption() As String
    UteRequeridoCaption = m_UteRequeridoCaption
End Property

Public Property Let UteRequeridoCaption(ByVal p_Value As String)
    m_UteRequeridoCaption = p_Value
End Property

Public Property Get UteComplementadoCaption() As String
    UteComplementadoCaption = m_UteComplementadoCaption
End Property

Public Property Let UteComplementadoCaption(ByVal p_Value As String)
    m_UteComplementadoCaption = p_Value
End Property

Public Property Get UteComplementadoColorRole() As String
    UteComplementadoColorRole = m_UteComplementadoColorRole
End Property

Public Property Let UteComplementadoColorRole(ByVal p_Value As String)
    m_UteComplementadoColorRole = p_Value
End Property

Public Property Get SuministradoresRequeridoCaption() As String
    SuministradoresRequeridoCaption = m_SuministradoresRequeridoCaption
End Property

Public Property Let SuministradoresRequeridoCaption(ByVal p_Value As String)
    m_SuministradoresRequeridoCaption = p_Value
End Property

Public Property Get SuministradoresComplementadoCaption() As String
    SuministradoresComplementadoCaption = m_SuministradoresComplementadoCaption
End Property

Public Property Let SuministradoresComplementadoCaption(ByVal p_Value As String)
    m_SuministradoresComplementadoCaption = p_Value
End Property

Public Property Get SuministradoresComplementadoColorRole() As String
    SuministradoresComplementadoColorRole = m_SuministradoresComplementadoColorRole
End Property

Public Property Let SuministradoresComplementadoColorRole(ByVal p_Value As String)
    m_SuministradoresComplementadoColorRole = p_Value
End Property

Public Property Get ComandoRechazarEnabled() As Boolean
    ComandoRechazarEnabled = m_ComandoRechazarEnabled
End Property

Public Property Let ComandoRechazarEnabled(ByVal p_Value As Boolean)
    m_ComandoRechazarEnabled = p_Value
End Property

Public Property Get ComandoPublicarEnabled() As Boolean
    ComandoPublicarEnabled = m_ComandoPublicarEnabled
End Property

Public Property Let ComandoPublicarEnabled(ByVal p_Value As Boolean)
    m_ComandoPublicarEnabled = p_Value
End Property

Public Property Get ComandoPublicarCaption() As String
    ComandoPublicarCaption = m_ComandoPublicarCaption
End Property

Public Property Let ComandoPublicarCaption(ByVal p_Value As String)
    m_ComandoPublicarCaption = p_Value
End Property

Public Property Get ComandoGenerarInformePublicabilidadEnabled() As Boolean
    ComandoGenerarInformePublicabilidadEnabled = m_ComandoGenerarInformePublicabilidadEnabled
End Property

Public Property Let ComandoGenerarInformePublicabilidadEnabled(ByVal p_Value As Boolean)
    m_ComandoGenerarInformePublicabilidadEnabled = p_Value
End Property

Public Property Get ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled() As Boolean
    ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled = m_ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled
End Property

Public Property Let ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled(ByVal p_Value As Boolean)
    m_ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled = p_Value
End Property

Public Property Get ComandoVerRechazoVisible() As Boolean
    ComandoVerRechazoVisible = m_ComandoVerRechazoVisible
End Property

Public Property Let ComandoVerRechazoVisible(ByVal p_Value As Boolean)
    m_ComandoVerRechazoVisible = p_Value
End Property

Public Property Get HasComandoVerRechazoVisibleDecision() As Boolean
    HasComandoVerRechazoVisibleDecision = m_HasComandoVerRechazoVisibleDecision
End Property

Public Property Let HasComandoVerRechazoVisibleDecision(ByVal p_Value As Boolean)
    m_HasComandoVerRechazoVisibleDecision = p_Value
End Property

Public Property Get PublicacionVisible() As Boolean
    PublicacionVisible = m_PublicacionVisible
End Property

Public Property Let PublicacionVisible(ByVal p_Value As Boolean)
    m_PublicacionVisible = p_Value
End Property

Public Property Get PublicacionCaption() As String
    PublicacionCaption = m_PublicacionCaption
End Property

Public Property Let PublicacionCaption(ByVal p_Value As String)
    m_PublicacionCaption = p_Value
End Property

Public Property Get PublicacionColorRole() As String
    PublicacionColorRole = m_PublicacionColorRole
End Property

Public Property Let PublicacionColorRole(ByVal p_Value As String)
    m_PublicacionColorRole = p_Value
End Property
