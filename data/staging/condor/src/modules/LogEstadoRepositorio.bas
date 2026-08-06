Attribute VB_Name = "LogEstadoRepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: LogEstadoRepositorio.bas
' ACCIÓN: Reemplazar la subrutina Guardar para que acepte una DB transaccional.
' ==========================================================================
Public Sub Guardar(ByRef objLog As LogEstado, Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    On Error GoTo Errores
    
    ' --- Lógica de conexión ---
    ' Si nos pasan una conexión, la usamos. Si no, obtenemos la global.
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Como el ID es autonumérico, no necesitamos calcularlo.
    objLog.idLogEstado = 0
    
    Set rcd = dbConexion.OpenRecordset("tbLogEstados", dbOpenDynaset)
    rcd.AddNew
    
    ' Usamos nuestro ayudante robusto para rellenar los campos, omitiendo el ID autonumérico.
    Dim campo As Variant
    For Each campo In objLog.ColCampos
        'Debug.Print campo
        If CStr(campo) <> "idLogEstado" Then
            rcd.Fields(CStr(campo)).value = objLog.getPropiedad(CStr(campo))
        End If
    Next campo
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogEstadoRepositorio.Guardar"
    If Not campo Is Nothing Then errObj.AddToCallStack "Error probable en el campo: " & CStr(campo)
    errObj.Raise
End Sub
Public Function getHistorialPorIdSolicitud(ByVal idSol As Long, Optional ByRef db As DAO.Database) As Scripting.Dictionary
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    ' Usar la conexión provista o la default
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' La ordenación por fecha es crucial para calcular los tiempos correctamente.
    sql = "SELECT * FROM tbLogEstados WHERE idSolicitud = [p_ID] ORDER BY fechaTransicion ASC;"
    
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSol
    
    Set getHistorialPorIdSolicitud = RepositorioComun.HidratarColeccionDesdeSQL(sql, "LogEstado", "idLogEstado", dbConexion, params)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogEstadoRepositorio.getHistorialPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Function

' ==========================================================================
' MÓDULO: LogEstadoRepositorio.bas
' ACCIÓN: Añadir esta nueva subrutina.
' ==========================================================================

Public Sub EliminarPorIdSolicitud(ByVal idSol As Long, ByRef db As DAO.Database)
    ' RESPONSABILIDAD: Eliminar todos los registros de tbLogEstados asociados a una solicitud.
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = "DELETE FROM tbLogEstados WHERE idSolicitud = [p_IDSol];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogEstadoRepositorio.EliminarPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Sub
Public Function getUltimoEstadoAnterior(ByVal idSol As Long, Optional ByRef db As DAO.Database) As Long
    ' RESPONSABILIDAD: Encontrar el estado en el que estaba una solicitud
    '                  justo antes de su última transición.
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    ' Usar la conexión provista o la default
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Ordenamos por fecha descendente y cogemos el primer registro (el más reciente)
    sql = "SELECT TOP 1 idEstadoAnterior FROM tbLogEstados " & _
          "WHERE idSolicitud = [p_ID] ORDER BY fechaTransicion DESC;"
          
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSol
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd.EOF Then
        getUltimoEstadoAnterior = Nz(rcd!idEstadoAnterior, 0)
    Else
        ' Si no hay historial, devolvemos 0
        getUltimoEstadoAnterior = 0
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    ' Si falla la consulta, devolvemos 0 para que el servicio lo maneje.
    getUltimoEstadoAnterior = 0
    Resume LimpiarYSalir
End Function


