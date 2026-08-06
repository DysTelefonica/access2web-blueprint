Attribute VB_Name = "UsuarioRepositorio"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: UsuarioRepositorio.bas (REFACTORIZADO COMPLETO CON CondorError)
' ==========================================================================

Public Function getUsuario(Optional p_ID As String, _
                           Optional p_UsuarioRed As String, _
                           Optional p_Nombre As String, _
                           Optional p_Correo As String, _
                           Optional ByRef db As DAO.Database) As usuario
    Dim sql As String
    Dim wheres As Object
    Dim params As Object
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    Set wheres = CreateObject("Scripting.Dictionary")
    Set params = CreateObject("Scripting.Dictionary")
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    If Nz(p_ID, "") <> "" Then
        wheres.Add "ID = [p_ID]", ""
        params.Add "p_ID", CLng(p_ID)
    End If
    If Nz(p_UsuarioRed, "") <> "" Then
        wheres.Add "UsuarioRed = [p_UsuarioRed]", ""
        params.Add "p_UsuarioRed", p_UsuarioRed
    End If
    If Nz(p_Nombre, "") <> "" Then
        wheres.Add "Nombre = [p_Nombre]", ""
        params.Add "p_Nombre", p_Nombre
    End If
    If Nz(p_Correo, "") <> "" Then
        wheres.Add "CorreoUsuario = [p_Correo]", ""
        params.Add "p_Correo", p_Correo
    End If
    
    If wheres.count = 0 Then Exit Function
    
    sql = "SELECT * FROM TbUsuariosAplicaciones WHERE " & Join(wheres.Keys, " AND ") & ";"
    Set getUsuario = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Usuario", dbConexion, params)

    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "UsuarioRepositorio.getUsuario"
    errObj.Raise
End Function

Public Function getUsuarioConectadoPorMaquina() As usuario
    Dim objNetwork As Object
    On Error GoTo Errores
    
    Set objNetwork = CreateObject("Wscript.Network")
    Set getUsuarioConectadoPorMaquina = getUsuario(, objNetwork.UserName)
    Set objNetwork = Nothing
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "UsuarioRepositorio.getUsuarioConectadoPorMaquina"
    errObj.Raise
End Function

