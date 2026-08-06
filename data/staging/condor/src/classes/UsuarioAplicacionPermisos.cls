VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioAplicacionPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: UsuarioAplicacionPermisos.cls (CORREGIDO - Sin manejo de errores)
' ==========================================================================

Public CorreoUsuario As String
Public IDAplicacion As String
Public EsUsuarioAdministrador As String
Public EsUsuarioCalidad As String
Public EsUsuarioEconomia As String
Public EsUsuarioSecretaria As String
Public EsUsuarioTecnico As String
Public EsUsuarioSinAcceso As String
Public EsUsuarioCalidadAvisos As String
Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "CorreoUsuario"
            .Add "IDAplicacion"
            .Add "EsUsuarioAdministrador"
            .Add "EsUsuarioCalidad"
            .Add "EsUsuarioEconomia"
            .Add "EsUsuarioSecretaria"
            .Add "EsUsuarioTecnico"
            .Add "EsUsuarioSinAcceso"
            .Add "EsUsuarioCalidadAvisos"
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
