Attribute VB_Name = "SuministradorRepositorio"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: SuministradorRepositorio.bas (REFACTORIZADO CON CondorError)
' RESPONSABILIDAD: Acceso a datos para la tabla de Suministradores.
' ==========================================================================

Public Function getSuministradorPorID(ByVal p_IDSuministrador As String, Optional ByRef db As DAO.Database) As suministrador
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM TbSuministradores WHERE IDSuministrador = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", CLng(p_IDSuministrador)
    
    Set getSuministradorPorID = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Suministrador", dbConexion, params)
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SuministradorRepositorio.getSuministradorPorID"
    errObj.Raise
End Function
