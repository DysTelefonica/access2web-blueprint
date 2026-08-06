VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public idSolicitud As Long
Public idExpediente As Long

Public tipoSolicitud As String

Public codigoSolicitud As String
Public idNCAsociada As Long
Public idEstadoInterno As Long
Public fechaCreacion As Date
Public usuarioCreacion As String

Public fechaModificacion As Date
Public usuarioModificacion As String
Public revisionCalidadEstado As String
Public revisionCalidadComentarios As String

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        ' CORRECCIÓN: Primero se crea la nueva colección.
        Set m_objColCampos = New Collection
        
        With m_objColCampos
            .Add "idSolicitud"
            .Add "idExpediente"
            .Add "tipoSolicitud"
            .Add "codigoSolicitud"
            .Add "idNCAsociada"
            .Add "idEstadoInterno"
            .Add "fechaCreacion"
            .Add "usuarioCreacion"
            .Add "fechaModificacion"
            .Add "usuarioModificacion"
            .Add "revisionCalidadEstado"
            .Add "revisionCalidadComentarios"
            
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

