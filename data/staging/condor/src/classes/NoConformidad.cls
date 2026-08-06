VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoConformidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==========================================================================
' MÓDULO: NoConformidad.cls (NUEVA ENTIDAD)
' ==========================================================================
Option Compare Database
Option Explicit

Public IDNoConformidad As Long
Public CodigoNoConformidad As String
Public CodConcesionAsociada As String
Public descripcion As String
Public idExpediente As Long

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "IDNoConformidad"
            .Add "CodigoNoConformidad"
            .Add "CodConcesionAsociada"
            .Add "Descripcion"
            .Add "idExpediente"
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

