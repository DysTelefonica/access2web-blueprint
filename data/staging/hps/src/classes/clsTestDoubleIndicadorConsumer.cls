VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestDoubleIndicadorConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' WithEvents bridge between modIndicadores.m_Indicadores (the singleton bus)
' and a clsTestDoubleIndicador instance (the recording target).
' A standard module cannot host WithEvents, so the consumer is a class.
Private WithEvents m_Bus As clsIndicadoresBus
Private m_Target As clsTestDoubleIndicador

Public Sub Init(ByVal target As clsTestDoubleIndicador)
    Set m_Target = target
    Set m_Bus = modIndicadores.m_Indicadores
End Sub

Private Sub m_Bus_LifecycleMoveCompleted(p_Direction As String, p_UserId As String)
    If Not m_Target Is Nothing Then
        m_Target.OnLifecycleMoveCompleted p_Direction, p_UserId
    End If
End Sub

Public Sub Detach()
    Set m_Target = Nothing
    Set m_Bus = Nothing
End Sub