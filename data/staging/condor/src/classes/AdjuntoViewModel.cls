VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdjuntoViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' MÓDULO: AdjuntoViewModel.cls (NUEVO)
' RESPONSABILIDAD: Contenedor de datos listos para mostrar en la UI de adjuntos.
' ==========================================================================
Option Compare Database
Option Explicit

' --- Propiedades del adjunto original ---
Public idAdjunto As Long
Public idSolicitud As Long
Public nombreArchivo As String
Public fechaSubida As Date
Public usuarioSubida As String
Public descripcion As String
Public tipoAccion As String
' --- Propiedad enriquecida por el servicio ---
Public etapaNombre As String

' --- API uniforme de entidades CONDOR (Opcional, pero buena práctica) ---
Public Function getPropiedad(ByVal nombrePropiedad As String) As Variant
    On Error GoTo Fin
    getPropiedad = CallByName(Me, nombrePropiedad, VbGet)
    Exit Function
Fin:
End Function

Public Sub SetPropiedad(ByVal nombrePropiedad As String, ByVal valorPropiedad As Variant)
    On Error Resume Next
    CallByName Me, nombrePropiedad, VbLet, valorPropiedad
End Sub
