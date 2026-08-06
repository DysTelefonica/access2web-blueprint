VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIndicadoresBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Declaration ordering (vba-access §10.1): Private Const BEFORE first Public declaration.
Private Const k_BusTag As String = "clsIndicadoresBus"

Public Event LifecycleMoveCompleted(p_Direction As String, p_UserId As String)

Public Sub RaiseLifecycleMoveCompleted(p_Direction As String, p_UserId As String)
    On Error GoTo EH
    RaiseEvent LifecycleMoveCompleted(p_Direction, p_UserId)
    Exit Sub
EH:
    Debug.Print k_BusTag & ".Raise swallow "; Err.Number; "-"; Err.Description
End Sub