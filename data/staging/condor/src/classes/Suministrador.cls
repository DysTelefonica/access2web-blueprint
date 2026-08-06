VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Suministrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Suministrador.cls (Versión con Propiedad Calculada)
' ==========================================================================

Public idSuministrador As Long
Public nombre As String
Public CIF As String
Public descripcion As String
Public TramitadoraHPS As String
Public Nemotecnico As String
Public Direccion As String
Public CP As String
Public Ciudad As String

Private m_objColCampos As Collection

' --- NUEVA PROPIEDAD DE SOLO LECTURA ---
Public Property Get NombreParaMostrar() As String
    ' Responsabilidad: Devuelve el Nemotécnico si existe, si no, devuelve el Nombre.
    If Nz(Me.Nemotecnico, "") <> "" Then
        NombreParaMostrar = Me.Nemotecnico
    Else
        NombreParaMostrar = Me.nombre
    End If
End Property

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "IDSuministrador"
            .Add "Nombre"
            .Add "CIF"
            .Add "Descripcion" ' Corregido, debe coincidir con el campo de la tabla
            .Add "TramitadoraHPS"
            .Add "Nemotecnico"
            .Add "Direccion"
            .Add "CP"
            .Add "Ciudad"
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

