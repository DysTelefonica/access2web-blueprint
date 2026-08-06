VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: LogEstado.cls (NUEVA ENTIDAD)
' RESPONSABILIDAD: Representa un registro en la tabla de historial de
'                  estados (tbLogEstados).
' ==========================================================================

Public idLogEstado As Long
Public idSolicitud As Long
Public idEstadoAnterior As Long
Public idEstadoNuevo As Long
Public fechaTransicion As Date
Public usuarioTransicion As String

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idLogEstado"
            .Add "idSolicitud"
            .Add "idEstadoAnterior"
            .Add "idEstadoNuevo"
            .Add "fechaTransicion"
            .Add "usuarioTransicion"
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
