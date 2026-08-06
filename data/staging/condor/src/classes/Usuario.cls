VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Usuario.cls (CORREGIDO - Sin manejo de errores)
' ==========================================================================

Public CorreoUsuario As String
Public password As String
Public UsuarioRed As String
Public nombre As String
Public Matricula As String
Public FechaAlta As String
Public Activado As Boolean
Public FechaProximoCambioContrasenia As String
Public FechaUltimaConexion As String
Public TieneQueCambiarLaContrasenia As Boolean
Public Telefono As String
Public Movil As String
Public observaciones As String
Public UsuarioImborrable As Boolean
Public EsAdministrador As String
Public PermisosAsignados As Boolean
Public FechaBaja As String
Public PasswordNuncaCaduca As Boolean
Public MantenerLanzaderaAbierta As Boolean
Public PassIncialPlana As String
Public UsuarioSSID As String
Public id As String
Public JefeDelUsuario As String
Public PermisoPruebas As String
Public ParaTareasProgramadas As Boolean
Public FechaBloqueo As String
Public rol As rol
Public ColAplicacionesPermisos As Scripting.Dictionary
Public Permisos As UsuarioAplicacionPermisos
Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "CorreoUsuario", "CorreoUsuario"
            .Add "Password", "Password"
            .Add "UsuarioRed", "UsuarioRed"
            .Add "Nombre", "Nombre"
            .Add "Matricula", "Matricula"
            .Add "FechaAlta", "FechaAlta"
            .Add "Activado", "Activado"
            .Add "FechaProximoCambioContrasenia", "FechaProximoCambioContrasenia"
            .Add "FechaUltimaConexion", "FechaUltimaConexion"
            .Add "TieneQueCambiarLaContrasenia", "TieneQueCambiarLaContrasenia"
            .Add "Telefono", "Telefono"
            .Add "Movil", "Movil"
            .Add "Observaciones", "Observaciones"
            .Add "UsuarioImborrable", "UsuarioImborrable"
            .Add "EsAdministrador", "EsAdministrador"
            .Add "PermisosAsignados", "PermisosAsignados"
            .Add "FechaBaja", "FechaBaja"
            .Add "PasswordNuncaCaduca", "PasswordNuncaCaduca"
            .Add "MantenerLanzaderaAbierta", "MantenerLanzaderaAbierta"
            .Add "PassIncialPlana", "PassIncialPlana"
            .Add "UsuarioSSID", "UsuarioSSID"
            .Add "Id", "Id"
            .Add "JefeDelUsuario", "JefeDelUsuario"
            .Add "PermisoPruebas", "PermisoPruebas"
            .Add "ParaTareasProgramadas", "ParaTareasProgramadas"
            .Add "FechaBloqueo", "FechaBloqueo"
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
