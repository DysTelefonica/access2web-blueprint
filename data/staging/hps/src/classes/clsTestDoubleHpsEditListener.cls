VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestDoubleHpsEditListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' WithEvents bridge for the parameterless EdicionRegistro event exposed by
' clsTestDoubleHpsEditPublisher. The handler increments m_HandlerCalls.
' The test atom asserts on m_HandlerCalls after DoRaise.
Private WithEvents m_Src As clsTestDoubleHpsEditPublisher
Public m_HandlerCalls As Long

Public Sub Init(ByVal src As clsTestDoubleHpsEditPublisher)
    Set m_Src = src
End Sub

Private Sub m_Src_EdicionRegistro()
    m_HandlerCalls = m_HandlerCalls + 1
End Sub
