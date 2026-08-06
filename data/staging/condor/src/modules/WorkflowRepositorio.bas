Attribute VB_Name = "WorkflowRepositorio"


' ==========================================================================
' MÓDULO: WorkflowRepositorio.bas
' RESPONSABILIDAD: Acceso a datos para la tabla de reglas de negocio tbTransiciones.
' ==========================================================================
Option Compare Database
Option Explicit

' ----- EN: WorkflowRepositorio.bas (REEMPLAZAR FUNCIÓN) -----

Public Function getPosiblesDestinos(ByVal idEstadoOrigen As Long, Optional ByRef db As DAO.Database) As Scripting.Dictionary
    ' Responsabilidad: Devuelve una colección simple (IDDestino, RolRequerido)
    '                  de todas las transiciones que parten de un estado origen.
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    Dim col As New Scripting.Dictionary
    Dim dbConexion As DAO.Database
    
    ' --- CAMBIO: Añadido manejo de errores robusto ---
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    col.CompareMode = TextCompare
    
    sql = "SELECT idEstadoDestino, rolRequerido FROM tbTransiciones WHERE idEstadoOrigen = [p_IDOrigen];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDOrigen", idEstadoOrigen
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then
            Do While Not rcd.EOF
                col.Add CStr(rcd.Fields("idEstadoDestino").value), rcd.Fields("rolRequerido").value
                rcd.MoveNext
            Loop
        End If
    End If
    
    Set getPosiblesDestinos = col
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
' --- CAMBIO: Nuevo bloque de errores que lanza un CondorError ---
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowRepositorio.getPosiblesDestinos"
    errObj.Raise
End Function


