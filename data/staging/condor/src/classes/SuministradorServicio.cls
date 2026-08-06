VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuministradorServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO DE CLASE: SuministradorServicio (REFACTORIZADO CON CondorError)
' RESPONSABILIDAD: Orquesta la lógica de negocio para la entidad Suministrador.
' ==========================================================================

Public Function getSuministradores(Optional ByVal p_PalabraClave As String = "") As Scripting.Dictionary
    Dim sql As String
    Dim params As Object
    
    On Error GoTo Errores
    
    Set params = CreateObject("Scripting.Dictionary")
    sql = "SELECT * FROM TbSuministradores "
    
    If p_PalabraClave <> "" Then
        sql = sql & "WHERE (Nombre LIKE [p_PalabraClave] OR Nemotecnico LIKE [p_PalabraClave]) "
        params.Add "p_PalabraClave", "*" & p_PalabraClave & "*"
    End If
    
    sql = sql & "ORDER BY Nombre;"
    
    Set getSuministradores = RepositorioComun.HidratarColeccionDesdeSQL(sql, "Suministrador", "IDSuministrador", getdb(), params)

    Exit Function

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SuministradorServicio.getSuministradores"
    errObj.Raise
End Function

Public Function getSuministradorPorID(ByVal idSuministrador As String, Optional ByRef db As DAO.Database) As suministrador
    On Error GoTo Errores

    If Trim(Nz(idSuministrador, "")) = "" Or Not IsNumeric(idSuministrador) Then
        Err.Raise 513, , "Se requiere un ID de Suministrador válido."
    End If
    
    Set getSuministradorPorID = SuministradorRepositorio.getSuministradorPorID(idSuministrador, db)
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SuministradorServicio.getSuministradorPorID"
    errObj.Raise
End Function
