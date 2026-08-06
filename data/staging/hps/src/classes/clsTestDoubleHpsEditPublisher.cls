VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestDoubleHpsEditPublisher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Test double — exposes a parameterless EdicionRegistro event and a DoRaise
' helper. Used by T3 (HpsEditRaisesEdicionRegistro) and T9
' (SicaEditFiresSicaListenerNotHps) to verify the WithEvents listener
' mechanism without depending on a real form raise.
Public Event EdicionRegistro()

Public Sub DoRaise()
    RaiseEvent EdicionRegistro
End Sub
