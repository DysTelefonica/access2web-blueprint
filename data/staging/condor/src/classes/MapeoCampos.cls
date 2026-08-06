VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapeoCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: MapeoCampos.cls (SIMPLIFICADO COMO ENTIDAD PURA)
' ==========================================================================

Public idMapeo As Long
Public nombrePlantilla As String
Public nombreCampoTabla As String
Public valorAsociado As String
Public nombreCampoWord As String
Public numExtensiones As Long

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idMapeo"
            .Add "nombrePlantilla"
            .Add "nombreCampoTabla"
            .Add "valorAsociado"
            .Add "nombreCampoWord"
            .Add "numExtensiones"
        End With
    End If
    Set ColCampos = m_objColCampos
End Property

Public Function getPropiedad(ByVal m_NombrePropiedad As String) As Variant
    ' Sin manejo de errores. Cualquier fallo se propagará al llamador.
    getPropiedad = CallByName(Me, m_NombrePropiedad, VbGet)
End Function

Public Sub SetPropiedad(ByVal m_NombrePropiedad As String, ByVal m_ValorPropiedad As Variant)
    ' Sin manejo de errores. Cualquier fallo se propagará al llamador.
    CallByName Me, m_NombrePropiedad, VbLet, m_ValorPropiedad
End Sub
