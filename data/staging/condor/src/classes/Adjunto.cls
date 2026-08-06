VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Adjunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Adjunto.cls (Normalizado como Entidad Pura - VERSIÓN FINAL)
' ==========================================================================

Public idAdjunto As Long
Public idSolicitud As Long
Public etapaWF As String
Public nombreArchivo As String
Public fechaSubida As Date
Public usuarioSubida As String
Public descripcion As String
Public tipoAccion As String
Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idAdjunto"
            .Add "idSolicitud"
            .Add "etapaWF"
            .Add "nombreArchivo"
            .Add "fechaSubida"
            .Add "usuarioSubida"
            .Add "descripcion"
            .Add "TipoAccion"
        End With
    End If
    Set ColCampos = m_objColCampos
End Property

' Sin manejo de errores. Cualquier fallo se propagará al llamador.
Public Function getPropiedad(ByVal m_NombrePropiedad As String) As Variant
    getPropiedad = CallByName(Me, m_NombrePropiedad, VbGet)
End Function

' Sin manejo de errores. Cualquier fallo se propagará al llamador.
Public Sub SetPropiedad(ByVal m_NombrePropiedad As String, ByVal m_ValorPropiedad As Variant)
    CallByName Me, m_NombrePropiedad, VbLet, m_ValorPropiedad
End Sub
