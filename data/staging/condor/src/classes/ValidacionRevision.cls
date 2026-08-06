VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidacionRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public id As Long
Public idSolicitud As Long
Public ordinal As Long
Public fechaEnvio As Date
Public fechaRecepcion As Date
Public resultado As String
Public comentarios As String
Public usuario As String
Public idAdjunto As Long
Public hashDatos As String

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "Id"
            .Add "idSolicitud"
            .Add "ordinal"
            .Add "FechaEnvio"
            .Add "FechaRecepcion"
            .Add "Resultado"
            .Add "Comentarios"
            .Add "Usuario"
            .Add "idAdjunto"
            .Add "HashDatos"
        End With
    End If
    Set ColCampos = m_objColCampos
End Property

Public Function getPropiedad(ByVal m_NombrePropiedad As String) As Variant
    getPropiedad = CallByName(Me, m_NombrePropiedad, VbGet)
End Function

Public Sub SetPropiedad(ByVal m_NombrePropiedad As String, ByVal m_ValorPropiedad As Variant)
    CallByName Me, m_NombrePropiedad, VbLet, m_ValorPropiedad
End Sub


