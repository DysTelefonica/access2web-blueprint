Attribute VB_Name = "modIndicadores"
Option Compare Database
Option Explicit

Private Const k_ModTag As String = "modIndicadores"

' Singleton — every WithEvents listener in the project attaches to this same instance.
Public m_Indicadores As New clsIndicadoresBus

' Thin wrapper. Best-effort, never sets caller's p_Error.
Public Sub Indicadores_NotificarMovimiento(p_Direction As String, p_UserId As String)
    On Error Resume Next
    m_Indicadores.RaiseLifecycleMoveCompleted p_Direction, p_UserId
    On Error GoTo 0
End Sub