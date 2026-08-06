Attribute VB_Name = "ExpedienteRepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: ExpedienteRepositorio.bas (REFACTORIZADO CON CondorError)
' ==========================================================================

Public Function getExpedientePorID(ByVal p_IDExpediente As String, Optional ByRef db As DAO.Database) As Expediente
    Dim sql As String
    Dim params As Object
    Dim fromClause As String
    Dim dbLocal As DAO.Database

    On Error GoTo Errores
    
    fromClause = "FROM (((TbExpedientes AS T_Exp " & _
                 "LEFT JOIN TbUsuariosAplicaciones AS T_RC ON T_Exp.IDResponsableCalidad = T_RC.Id) " & _
                 "LEFT JOIN (SELECT T_Resp_RT.IdExpediente, FIRST(T_Usr_RT.Nombre) AS NombreRT, FIRST(T_Usr_RT.CorreoUsuario) AS EmailRT " & _
                           "FROM (TbExpedientesResponsables AS T_Resp_RT INNER JOIN TbUsuariosAplicaciones AS T_Usr_RT ON T_Resp_RT.IdUsuario = T_Usr_RT.Id) " & _
                             "WHERE T_Resp_RT.EsJefeProyecto = 'Sí' GROUP BY T_Resp_RT.IdExpediente) AS SubRT ON T_Exp.IDExpediente = SubRT.IdExpediente) " & _
                 "LEFT JOIN (SELECT T_ExpRACS.IDEXPEDIENTE, FIRST(T_RACS.RAC) AS NombreRAC " & _
                             "FROM (TbExpedientesRACS AS T_ExpRACS INNER JOIN TbRACS AS T_RACS ON T_ExpRACS.IDRAC = T_RACS.IDRAC) " & _
                             "GROUP BY T_ExpRACS.IDEXPEDIENTE) AS SubRAC ON T_Exp.IDExpediente = SubRAC.IDEXPEDIENTE) "

    sql = "SELECT " & _
            "T_Exp.IDExpediente, T_Exp.Nemotecnico, T_Exp.Titulo, T_Exp.CodExp, T_Exp.CodExpLargo,T_Exp.CodigoActividad, " & _
            "T_Exp.ContratistaPrincipal, T_Exp.FechaInicioContrato, T_Exp.FechaFinContrato, T_Exp.FechaFinGarantia,T_Exp.objetoContrato, " & _
            "T_RC.Nombre AS ResponsableCalidad, T_RC.CorreoUsuario AS EmailResponsableCalidad, " & _
            "SubRT.NombreRT AS ResponsableTecnico, " & _
            "SubRT.EmailRT AS EmailResponsableTecnico, " & _
            "SubRAC.NombreRAC AS RacAsignado " & _
          fromClause & _
          "WHERE T_Exp.IDExpediente = [p_ID];"
          
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", CLng(p_IDExpediente)
    
    If db Is Nothing Then Set db = getdb()
    Set dbLocal = db
    Set getExpedientePorID = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Expediente", dbLocal, params)
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteRepositorio.getExpedientePorID"
    errObj.Raise
End Function


Public Function getSuministradoresPorExpediente(ByVal idExpediente As Long, Optional ByRef db As DAO.Database) As Scripting.Dictionary
    Dim sql As String
    Dim dbConexion As DAO.Database
    On Error GoTo Errores

    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    ' La consulta SQL se actualiza para seleccionar de la tabla correcta.
    sql = "SELECT ES.* " & _
          "FROM TbExpedientesSuministradores AS ES " & _
          "WHERE ES.IDExpediente = " & idExpediente

    ' La llamada al hidratador ya es robusta y se adaptará a la nueva estructura de la clase
    ' siempre que los nombres de los campos y las propiedades coincidan.
    Set getSuministradoresPorExpediente = RepositorioComun.HidratarColeccionDesdeSQL( _
        sql, _
        "ExpedienteSuministrador", _
        "IDExpedienteSuministrador", _
        dbConexion _
    )

    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteRepositorio.getSuministradoresPorExpediente"
    errObj.AddToCallStack "ID Expediente: " & idExpediente
    errObj.Raise
End Function

Public Function GetRacAsignadoPorExpediente(ByVal idExpediente As Long, Optional ByRef db As DAO.Database) As String
    Dim sql As String
    Dim dbConexion As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT FIRST(TbRACS.RAC) AS NombreRAC " & _
          "FROM TbExpedientesRACS INNER JOIN TbRACS ON TbExpedientesRACS.IDRAC = TbRACS.IDRAC " & _
          "WHERE TbExpedientesRACS.IDExpediente = " & idExpediente
    
    Set rs = dbConexion.OpenRecordset(sql)
    If Not rs.EOF Then
        GetRacAsignadoPorExpediente = Nz(rs!nombreRAC, "")
    Else
        GetRacAsignadoPorExpediente = ""
    End If
    rs.Close
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteRepositorio.GetRacAsignadoPorExpediente"
    errObj.AddToCallStack "ID Expediente: " & idExpediente
    errObj.Raise
End Function



