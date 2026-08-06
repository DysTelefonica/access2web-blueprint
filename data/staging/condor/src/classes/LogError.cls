VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: LogError.cls (Versión 2.0 con Suplantación)
' ==========================================================================

Public idLogError As Long
Public fechaHora As Date
Public usuario As String
Public suplantadoPor As String ' <-- NUEVA PROPIEDAD
Public modulo As String
Public procedimiento As String
Public numeroError As Long
Public descripcionError As String
Public contexto As String

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idLogError"
            .Add "fechaHora"
            .Add "usuario"
            .Add "suplantadoPor" ' <-- AÑADIDO A LA COLECCIÓN
            .Add "modulo"
            .Add "procedimiento"
            .Add "numeroError"
            .Add "descripcionError"
            .Add "contexto"
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
