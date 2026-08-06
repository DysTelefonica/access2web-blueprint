VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestDoubleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Test double — simulates a form's Form_Activate and requery entry points.
' Used by T5 (Form_ActivateFallbackFiresOnFocusReturn) and T10
' (HistoricosSubformRequeriesOnActivate). The atom invokes Form_Activate
' directly on a fresh instance and asserts the counters incremented.
Public m_FormActivateCalls As Long
Public m_RequeryCount As Long

Public Sub Form_Activate()
    m_FormActivateCalls = m_FormActivateCalls + 1
    m_RequeryCount = m_RequeryCount + 1
End Sub
