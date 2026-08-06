VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestDoubleIndicador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Test double — records the lifecycle move notifications it receives.
' Wired by the consumer (clsTestDoubleIndicadorConsumer) via WithEvents on
' modIndicadores.m_Indicadores. The test atom asserts on m_HandlerCalls,
' m_LastDirection and m_LastUserId.
Public m_HandlerCalls As Long
Public m_LastDirection As String
Public m_LastUserId As String

Public Sub OnLifecycleMoveCompleted(ByVal direction As String, ByVal userId As String)
    m_HandlerCalls = m_HandlerCalls + 1
    m_LastDirection = direction
    m_LastUserId = userId
End Sub

Public Sub Reset()
    m_HandlerCalls = 0
    m_LastDirection = ""
    m_LastUserId = ""
End Sub