VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: LogCambio.cls
' RESPONSABILIDAD: Entidad que representa un registro en tbLogCambios.
' ==========================================================================

Public idLogCambio As Long
Public fechaHora As Date
Public usuario As String
Public suplantadoPor As String
Public tabla As String
Public registro As Long
Public campo As String
Public valorAnterior As String
Public valorNuevo As String
Public tipoOperacion As String

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idLogCambio"
            .Add "fechaHora"
            .Add "usuario"
            .Add "suplantadoPor"
            .Add "tabla"
            .Add "registro"
            .Add "campo"
            .Add "valorAnterior"
            .Add "valorNuevo"
            .Add "tipoOperacion"
        End With
    End If
    Set ColCampos = m_objColCampos
End Property

Public Function getPropiedad(ByVal nombrePropiedad As String) As Variant
    
    getPropiedad = CallByName(Me, nombrePropiedad, VbGet)
End Function

Public Sub SetPropiedad(ByVal nombrePropiedad As String, ByVal valorPropiedad As Variant)
    
    CallByName Me, nombrePropiedad, VbLet, valorPropiedad
End Sub
