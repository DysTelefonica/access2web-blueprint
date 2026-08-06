Attribute VB_Name = "EstadoRepositorio"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: EstadoRepositorio.bas (REFACTORIZADO CON CondorError)
' RESPONSABILIDAD: Acceso a datos para la tabla tbEstados.
' ==========================================================================

Public Function getTodosLosEstados(Optional ByRef db As DAO.Database) As Scripting.Dictionary
    ' Responsabilidad: Obtiene todos los estados de la base de datos.
    Dim sql As String
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbEstados ORDER BY orden;"
    
    Set getTodosLosEstados = RepositorioComun.HidratarColeccionDesdeSQL(sql, "Estado", "idEstado", dbConexion)

    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "EstadoRepositorio.getTodosLosEstados"
    errObj.Raise
End Function
