VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Compare Database
Option Explicit


Public IDCorreo As Long
Public Aplicacion As String
Public Originador As String
Public Destinatarios As String
Public DestinatariosConCopia As String
Public DestinatariosConCopiaOculta As String
Public asunto As String
Public cuerpo As String
Public URLAdjunto As String
Public fechaEnvio As Variant ' Se deja Null, por eso es Variant
Public CuerpoHTML As Boolean
Public FechaGrabacion As Date ' <<< AÑADIR ESTA LÍNEA

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "IDCorreo"
            .Add "Aplicacion"
            .Add "Originador"
            .Add "Destinatarios"
            .Add "DestinatariosConCopia"
            .Add "DestinatariosConCopiaOculta"
            .Add "Asunto"
            .Add "Cuerpo"
            .Add "URLAdjunto"
            .Add "FechaEnvio"
            .Add "CuerpoHTML"
            .Add "FechaGrabacion" ' <<< AÑADIR ESTA LÍNEA
        End With
    End If
    Set ColCampos = m_objColCampos
End Property

' Sin manejo de errores. Cualquier fallo se propagará al llamador.
Public Function getPropiedad(ByVal nombrePropiedad As String) As Variant
    getPropiedad = CallByName(Me, nombrePropiedad, VbGet)
End Function

' Sin manejo de errores. Cualquier fallo se propagará al llamador.
Public Sub SetPropiedad(ByVal nombrePropiedad As String, ByVal valorPropiedad As Variant)
    CallByName Me, nombrePropiedad, VbLet, valorPropiedad
End Sub
