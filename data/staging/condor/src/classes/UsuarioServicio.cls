VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: UsuarioServicio.cls (Refactorizado con lógica de composición)
' ==========================================================================

Public Function getUsuarioConPermisos(Optional p_ID As String, _
                                      Optional p_UsuarioRed As String, _
                                      Optional p_Nombre As String, _
                                      Optional p_Correo As String, _
                                      Optional ByRef db As DAO.Database) As usuario
    Dim objUsuario As usuario
    On Error GoTo Errores
    
    ' 1. Obtenemos el objeto base del usuario desde el repositorio
    Set objUsuario = UsuarioRepositorio.getUsuario(p_ID, p_UsuarioRed, p_Nombre, p_Correo, db)
    
    If Not objUsuario Is Nothing Then
        ' 2. Si el usuario existe, obtenemos su colección de permisos
        Set objUsuario.ColAplicacionesPermisos = AplicacionRepositorio.getAplicacionesPermisos(objUsuario.CorreoUsuario)
        
        ' 3. Extraemos el permiso específico para esta aplicación (CONDOR)
        If objUsuario.ColAplicacionesPermisos.Exists(CStr(IDAplicacion)) Then
            Set objUsuario.Permisos = objUsuario.ColAplicacionesPermisos(CStr(IDAplicacion))
        End If
        
        ' 4. Determinamos y asignamos su rol
        objUsuario.rol = Me.DeterminarRol(objUsuario)
    End If
    
    Set getUsuarioConPermisos = objUsuario
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "UsuarioServicio.getUsuarioConPermisos"
    errObj.Raise
End Function

' ==========================================================================
' EN: UsuarioServicio.cls (REEMPLAZAR ESTA FUNCIÓN)
' ==========================================================================
Public Function getUsuarioConectadoConPermisos() As usuario
    Dim objNetwork As Object
    Dim usuarioDeRed As String
    On Error GoTo Errores
    
    ' 1. Crea el objeto para saber el usuario de la máquina
    Set objNetwork = CreateObject("Wscript.Network")
    usuarioDeRed = objNetwork.UserName
    
    ' 2. (NUEVO) Aplicamos las reglas de mapeo específicas
    Select Case LCase(usuarioDeRed)
        Case "local1", "adm1"
            usuarioDeRed = "adm"
    End Select
    
    ' 3. Llama a la función principal pasándole el nombre de usuario de red ya procesado
    Set getUsuarioConectadoConPermisos = getUsuarioConPermisos(, usuarioDeRed)
    Set objNetwork = Nothing
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "UsuarioServicio.getUsuarioConectadoConPermisos"
    errObj.Raise
End Function

Public Function DeterminarRol(ByRef objUsuario As usuario) As rol
    On Error GoTo Errores
    
    If objUsuario Is Nothing Then
        Err.Raise 513, , "Se intentó determinar el rol de un objeto de usuario Nulo."
    End If
    
    ' La jerarquía es clara: Administrador > Calidad > Técnico
    If objUsuario.EsAdministrador = "Sí" Then
        DeterminarRol = rol.Administrador
        Exit Function
    End If
    
    If objUsuario.Permisos Is Nothing Then
        DeterminarRol = rol.Tecnico
        Exit Function
    End If
    
    If objUsuario.Permisos.EsUsuarioCalidad = "Sí" Then
        DeterminarRol = rol.Calidad
    Else
        DeterminarRol = rol.Tecnico
    End If
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "UsuarioServicio.DeterminarRol"
    errObj.Raise
End Function

' --- (Las funciones de getResponsables... no cambian) ---

Public Function getResponsablesTecnicos() As Scripting.Dictionary
    Dim sql As String
    On Error GoTo Errores
    
    sql = "SELECT DISTINCT T_Usr.* " & _
          "FROM (TbExpedientes AS T_Exp INNER JOIN TbExpedientesResponsables AS T_Resp ON T_Exp.IDExpediente = T_Resp.IdExpediente) " & _
          "INNER JOIN TbUsuariosAplicaciones AS T_Usr ON T_Resp.IdUsuario = T_Usr.Id " & _
          "WHERE (((T_Resp.EsJefeProyecto)='Sí') AND ((T_Exp.Pecal)='Sí') AND ((T_Usr.FechaBaja) Is Null)) " & _
          "ORDER BY T_Usr.Nombre;"
    Set getResponsablesTecnicos = RepositorioComun.HidratarColeccionDesdeSQL(sql, "Usuario", "Id", getdb())

    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "UsuarioServicio.getResponsablesTecnicos"
    errObj.Raise
End Function

Public Function getResponsablesCalidad() As Scripting.Dictionary
    Dim sql As String
    On Error GoTo Errores
    
    sql = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones INNER JOIN TbUsuariosAplicacionesPermisos " & _
            "ON TbUsuariosAplicaciones.CorreoUsuario = TbUsuariosAplicacionesPermisos.CorreoUsuario " & _
            "WHERE (((TbUsuariosAplicaciones.FechaBaja) Is Null) " & _
            "AND ((TbUsuariosAplicacionesPermisos.IDAplicacion)=" & IDAplicacion & ") " & _
            "AND ((TbUsuariosAplicacionesPermisos.EsUsuarioCalidad)='Sí')) " & _
            "ORDER BY TbUsuariosAplicaciones.Nombre;"
    Set getResponsablesCalidad = RepositorioComun.HidratarColeccionDesdeSQL(sql, "Usuario", "Id", getdb())

    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "UsuarioServicio.getResponsablesCalidad"
    errObj.Raise
End Function

