Attribute VB_Name = "RiesgoMaterializacionContingenciaService"
Option Compare Database
Option Explicit

Private Const PLAN_CONTINGENCIA_TABLE As String = "TbRiesgosPlanContingenciaPpal"
Private Const PLAN_CONTINGENCIA_ID_FIELD As String = "IDContingencia"
Private Const RIESGO_ID_FIELD As String = "IDRiesgo"
Private Const PLAN_ESTADO_ACTIVO As String = "Activo"
Private Const MATERIALIZACION_TABLE As String = "TbRiesgosMaterializaciones"

Public Function DiagnosticarPlanesElegiblesMaterializacion( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_Error As String _
) As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim dbErr As String
    Dim sqlElegibles As String
    Dim idRiesgo As Long
    Dim eligibleCount As Long
    Dim missingSchema As String

    On Error GoTo EH

    p_Error = ""
    If Not TryResolveLongId(p_IDRiesgo, "riesgo", idRiesgo, p_Error) Then
        DiagnosticarPlanesElegiblesMaterializacion = _
            BuildDiagnosticJson(False, p_Error, 0, "")
        Exit Function
    End If

    Set db = getdb(dbErr)
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para diagnosticar planes elegibles"
        If dbErr <> "" Then p_Error = p_Error & ": " & dbErr
        DiagnosticarPlanesElegiblesMaterializacion = _
            BuildDiagnosticJson(False, p_Error, 0, "")
        GoTo Cleanup
    End If

    missingSchema = MissingEligibilitySchema(db)
    If missingSchema <> "" Then
        p_Error = "SCHEMA BLOCKED: " & missingSchema
        DiagnosticarPlanesElegiblesMaterializacion = _
            BuildDiagnosticJson(False, p_Error, 0, "")
        GoTo Cleanup
    End If

    sqlElegibles = SQLPlanesContingenciaElegiblesMaterializacion(p_IDRiesgo, p_Error)
    If p_Error <> "" Then
        DiagnosticarPlanesElegiblesMaterializacion = _
            BuildDiagnosticJson(False, p_Error, 0, "")
        GoTo Cleanup
    End If

    Set qdf = db.CreateQueryDef(vbNullString, GetEligiblePlanCountSql())
    qdf.Parameters("p_IDRiesgo").Value = idRiesgo
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    If Not rs.EOF Then
        eligibleCount = CLng(Nz(rs.Fields("EligibleCount").Value, 0))
    End If

    DiagnosticarPlanesElegiblesMaterializacion = _
        BuildDiagnosticJson(True, "", eligibleCount, sqlElegibles)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function

EH:
    p_Error = "DiagnosticarPlanesElegiblesMaterializacion: " & Err.Description
    DiagnosticarPlanesElegiblesMaterializacion = _
        BuildDiagnosticJson(False, p_Error, 0, "")
    Resume Cleanup
End Function

Public Function ActualizarMaterializacionVigentePlan( _
    ByVal p_IDRiesgo As String, _
    ByVal p_IDMaterializacion As String, _
    ByVal p_Fecha As String, _
    ByVal p_IDPlanContingencia As String, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim dbErr As String
    Dim idRiesgo As Long
    Dim idMaterializacion As Long
    Dim resolvedPlanContingencia As Long
    Dim idPlanContingencia As Variant
    Dim fechaMaterializacion As Date

    On Error GoTo EH

    ActualizarMaterializacionVigentePlan = EnumSiNo.No
    p_Error = ""

    If Not TryResolveLongId(p_IDRiesgo, "riesgo", idRiesgo, p_Error) Then
        Exit Function
    End If
    If Not TryResolveLongId(p_IDMaterializacion, "materialización", _
        idMaterializacion, p_Error) Then
        Exit Function
    End If
    idPlanContingencia = Null
    If Trim$(Nz(p_IDPlanContingencia, "")) <> "" Then
        If Not TryResolveLongId(p_IDPlanContingencia, "plan de contingencia", _
            resolvedPlanContingencia, p_Error) Then
            Exit Function
        End If
        idPlanContingencia = resolvedPlanContingencia
    End If
    If Not IsDate(p_Fecha) Then
        p_Error = "Debe indicar una fecha de materialización válida"
        Exit Function
    End If
    fechaMaterializacion = CDate(p_Fecha)

    If Trim$(Nz(p_IDPlanContingencia, "")) <> "" Then
        If ValidarYResolverPlanContingenciaMaterializacion(p_IDRiesgo, _
            p_IDPlanContingencia, p_Error) <> EnumSiNo.Sí Then
            Exit Function
        End If
    End If

    Set db = getdb(dbErr)
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para actualizar la materialización"
        If dbErr <> "" Then p_Error = p_Error & ": " & dbErr
        GoTo Cleanup
    End If

    If Not MaterializacionBelongsToRisk(db, idMaterializacion, idRiesgo) Then
        p_Error = "La materialización indicada no pertenece al riesgo indicado"
        GoTo Cleanup
    End If
    If Not IsCurrentMaterializacionForRisk(db, idMaterializacion, idRiesgo) Then
        p_Error = "Solo se puede actualizar la materialización vigente del riesgo"
        GoTo Cleanup
    End If

    Set qdf = db.CreateQueryDef(vbNullString, GetUpdateMaterializacionSql())
    qdf.Parameters("p_Fecha").Value = fechaMaterializacion
    qdf.Parameters("p_IDPlanContingencia").Value = idPlanContingencia
    qdf.Parameters("p_IDMaterializacion").Value = idMaterializacion
    qdf.Execute dbFailOnError

    If qdf.RecordsAffected <> 1 Then
        p_Error = "La actualización de la materialización vigente no afectó exactamente una fila"
        GoTo Cleanup
    End If

    ActualizarMaterializacionVigentePlan = EnumSiNo.Sí

Cleanup:
    On Error Resume Next
    Set qdf = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function

EH:
    p_Error = "ActualizarMaterializacionVigentePlan: " & Err.Description
    Resume Cleanup
End Function

Public Function EsMaterializacionVigenteEditable( _
    ByVal p_IDRiesgo As String, _
    ByVal p_IDMaterializacion As String, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    Dim db As DAO.Database
    Dim dbErr As String
    Dim idRiesgo As Long
    Dim idMaterializacion As Long

    On Error GoTo EH

    EsMaterializacionVigenteEditable = EnumSiNo.No
    p_Error = ""

    If Not TryResolveLongId(p_IDRiesgo, "riesgo", idRiesgo, p_Error) Then
        Exit Function
    End If
    If Not TryResolveLongId(p_IDMaterializacion, "materialización", _
        idMaterializacion, p_Error) Then
        Exit Function
    End If

    Set db = getdb(dbErr)
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para validar la materialización vigente"
        If dbErr <> "" Then p_Error = p_Error & ": " & dbErr
        GoTo Cleanup
    End If

    If Not MaterializacionBelongsToRisk(db, idMaterializacion, idRiesgo) Then
        p_Error = "La materialización indicada no pertenece al riesgo indicado"
        GoTo Cleanup
    End If
    If Not IsCurrentMaterializacionForRisk(db, idMaterializacion, idRiesgo) Then
        p_Error = "Solo se puede editar la materialización vigente del riesgo"
        GoTo Cleanup
    End If

    EsMaterializacionVigenteEditable = EnumSiNo.Sí

Cleanup:
    On Error Resume Next
    Set db = Nothing
    On Error GoTo 0
    Exit Function

EH:
    p_Error = "EsMaterializacionVigenteEditable: " & Err.Description
    Resume Cleanup
End Function

Public Function ResumenPlanContingenciaMaterializacion( _
    ByVal p_IDPlanContingencia As String, _
    Optional ByRef p_Error As String _
) As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim dbErr As String
    Dim idPlanContingencia As Long
    Dim codPlan As String
    Dim disparadorPlan As String

    On Error GoTo EH

    p_Error = ""
    If Trim$(Nz(p_IDPlanContingencia, "")) = "" Then Exit Function
    If Not TryResolveLongId(p_IDPlanContingencia, "plan de contingencia", _
        idPlanContingencia, p_Error) Then
        Exit Function
    End If

    Set db = getdb(dbErr)
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para resumir el plan de contingencia"
        If dbErr <> "" Then p_Error = p_Error & ": " & dbErr
        GoTo Cleanup
    End If

    Set qdf = db.CreateQueryDef(vbNullString, GetPlanContingenciaResumenSql())
    qdf.Parameters("p_IDContingencia").Value = idPlanContingencia
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    If rs.EOF Then GoTo Cleanup

    codPlan = Trim$(Nz(rs.Fields("CodContingencia").Value, ""))
    disparadorPlan = Trim$(Nz(rs.Fields("DisparadorDelPlan").Value, ""))
    ResumenPlanContingenciaMaterializacion = _
        BuildPlanContingenciaResumen(codPlan, disparadorPlan)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function

EH:
    p_Error = "ResumenPlanContingenciaMaterializacion: " & Err.Description
    Resume Cleanup
End Function

Public Function SQLPlanesContingenciaElegiblesMaterializacion( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_Error As String _
) As String
    Dim idRiesgo As Long

    On Error GoTo EH

    p_Error = ""
    If Not TryResolveLongId(p_IDRiesgo, "riesgo", idRiesgo, p_Error) Then
        Exit Function
    End If

    SQLPlanesContingenciaElegiblesMaterializacion = _
        "SELECT [" & PLAN_CONTINGENCIA_ID_FIELD & "], [CodContingencia], " & _
        "[DisparadorDelPlan] " & _
        "FROM [" & PLAN_CONTINGENCIA_TABLE & "] " & _
        "WHERE [" & RIESGO_ID_FIELD & "]=" & CStr(idRiesgo) & " " & _
        "AND [Estado]='" & PLAN_ESTADO_ACTIVO & "' " & _
        "AND ([FechaDeActivacion] Is Null OR [FechaDeActivacion] <= Date()) " & _
        "AND ([FechaDesactivacion] Is Null OR [FechaDesactivacion] > Date()) " & _
        "ORDER BY [FechaDeActivacion], [CodContingencia]"
    Exit Function

EH:
    p_Error = "SQLPlanesContingenciaElegiblesMaterializacion: " & Err.Description
End Function

Public Function ValidarYResolverPlanContingenciaMaterializacion( _
    ByVal p_IDRiesgo As String, _
    ByVal p_IDPlanContingencia As String, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim dbErr As String
    Dim idRiesgo As Long
    Dim idPlanContingencia As Long
    Dim planRiskId As Long
    Dim planEstado As String
    Dim fechaActivacion As Variant
    Dim fechaDesactivacion As Variant

    On Error GoTo EH

    ValidarYResolverPlanContingenciaMaterializacion = EnumSiNo.No
    p_Error = ""

    If Not TryResolveLongId(p_IDRiesgo, "riesgo", idRiesgo, p_Error) Then
        Exit Function
    End If
    If Not TryResolveLongId(p_IDPlanContingencia, "plan de contingencia", _
        idPlanContingencia, p_Error) Then
        Exit Function
    End If

    Set db = getdb(dbErr)
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para validar el plan de contingencia"
        If dbErr <> "" Then p_Error = p_Error & ": " & dbErr
        GoTo Cleanup
    End If

    Set qdf = db.CreateQueryDef(vbNullString, GetPlanContingenciaSql())
    qdf.Parameters("p_IDContingencia").Value = idPlanContingencia
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    If rs.EOF Then
        p_Error = "El plan de contingencia seleccionado no existe"
        GoTo Cleanup
    End If

    planRiskId = CLng(Nz(rs.Fields(RIESGO_ID_FIELD).Value, 0))
    If planRiskId <> idRiesgo Then
        p_Error = "El plan de contingencia seleccionado no pertenece al riesgo indicado"
        GoTo Cleanup
    End If

    planEstado = Trim$(Nz(rs.Fields("Estado").Value, ""))
    If StrComp(planEstado, PLAN_ESTADO_ACTIVO, vbTextCompare) <> 0 Then
        p_Error = "El plan de contingencia seleccionado no está activo"
        GoTo Cleanup
    End If

    fechaActivacion = rs.Fields("FechaDeActivacion").Value
    If Not IsNull(fechaActivacion) Then
        If CDate(fechaActivacion) > Date Then
            p_Error = "El plan de contingencia seleccionado todavía no está vigente"
            GoTo Cleanup
        End If
    End If

    fechaDesactivacion = rs.Fields("FechaDesactivacion").Value
    If Not IsNull(fechaDesactivacion) Then
        If CDate(fechaDesactivacion) <= Date Then
            p_Error = "El plan de contingencia seleccionado ya no está vigente"
            GoTo Cleanup
        End If
    End If

    ValidarYResolverPlanContingenciaMaterializacion = EnumSiNo.Sí

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function

EH:
    p_Error = "ValidarYResolverPlanContingenciaMaterializacion: " & Err.Description
    Resume Cleanup
End Function

Private Function TryResolveLongId( _
    ByVal rawValue As String, _
    ByVal fieldLabel As String, _
    ByRef resolvedValue As Long, _
    ByRef p_Error As String _
) As Boolean
    Dim normalizedValue As String

    normalizedValue = Trim$(Nz(rawValue, ""))

    If normalizedValue = "" Or normalizedValue = "0" Then
        p_Error = "Debe indicar el identificador del " & fieldLabel
        Exit Function
    End If

    If Not IsNumeric(normalizedValue) Then
        p_Error = "El identificador del " & fieldLabel & " debe ser numérico"
        Exit Function
    End If

    resolvedValue = CLng(normalizedValue)
    TryResolveLongId = True
End Function

Private Function GetPlanContingenciaSql() As String
    GetPlanContingenciaSql = _
        "PARAMETERS p_IDContingencia LONG; " & _
        "SELECT [" & PLAN_CONTINGENCIA_ID_FIELD & "], [" & RIESGO_ID_FIELD & "], " & _
        "[Estado], [FechaDeActivacion], [FechaDesactivacion] " & _
        "FROM [" & PLAN_CONTINGENCIA_TABLE & "] " & _
        "WHERE [" & PLAN_CONTINGENCIA_ID_FIELD & "] = [p_IDContingencia]"
End Function

Private Function GetPlanContingenciaResumenSql() As String
    GetPlanContingenciaResumenSql = _
        "PARAMETERS p_IDContingencia LONG; " & _
        "SELECT [CodContingencia], [DisparadorDelPlan] " & _
        "FROM [" & PLAN_CONTINGENCIA_TABLE & "] " & _
        "WHERE [" & PLAN_CONTINGENCIA_ID_FIELD & "] = [p_IDContingencia]"
End Function

Private Function BuildPlanContingenciaResumen( _
    ByVal codPlan As String, _
    ByVal disparadorPlan As String _
) As String
    Dim resumen As String

    resumen = Trim$(codPlan & " " & disparadorPlan)
    If Len(resumen) > 80 Then resumen = Left$(resumen, 77) & "..."
    BuildPlanContingenciaResumen = resumen
End Function

Private Function GetEligiblePlanCountSql() As String
    GetEligiblePlanCountSql = _
        "PARAMETERS p_IDRiesgo LONG; " & _
        "SELECT Count(*) AS EligibleCount " & _
        "FROM [" & PLAN_CONTINGENCIA_TABLE & "] " & _
        "WHERE [" & RIESGO_ID_FIELD & "]=[p_IDRiesgo] " & _
        "AND [Estado]='" & PLAN_ESTADO_ACTIVO & "' " & _
        "AND ([FechaDeActivacion] Is Null OR [FechaDeActivacion] <= Date()) " & _
        "AND ([FechaDesactivacion] Is Null OR [FechaDesactivacion] > Date())"
End Function

Private Function GetUpdateMaterializacionSql() As String
    GetUpdateMaterializacionSql = _
        "PARAMETERS p_Fecha DATETIME, p_IDPlanContingencia LONG, " & _
        "p_IDMaterializacion LONG; " & _
        "UPDATE [" & MATERIALIZACION_TABLE & "] " & _
        "SET [Fecha]=[p_Fecha], [IDPlanContingencia]=[p_IDPlanContingencia] " & _
        "WHERE [ID]=[p_IDMaterializacion]"
End Function

Private Function MaterializacionBelongsToRisk( _
    ByVal db As DAO.Database, _
    ByVal idMaterializacion As Long, _
    ByVal idRiesgo As Long _
) As Boolean
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset

    On Error GoTo Cleanup

    Set qdf = db.CreateQueryDef(vbNullString, GetMaterializacionRiskSql())
    qdf.Parameters("p_IDMaterializacion").Value = idMaterializacion
    qdf.Parameters("p_IDRiesgo").Value = idRiesgo
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    MaterializacionBelongsToRisk = Not rs.EOF

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    On Error GoTo 0
End Function

Private Function IsCurrentMaterializacionForRisk( _
    ByVal db As DAO.Database, _
    ByVal idMaterializacion As Long, _
    ByVal idRiesgo As Long _
) As Boolean
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim currentId As Long

    On Error GoTo Cleanup

    Set qdf = db.CreateQueryDef(vbNullString, GetCurrentMaterializacionSql())
    qdf.Parameters("p_IDRiesgo").Value = idRiesgo
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    If Not rs.EOF Then
        currentId = CLng(Nz(rs.Fields("ID").Value, 0))
        IsCurrentMaterializacionForRisk = (currentId = idMaterializacion)
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    On Error GoTo 0
End Function

Private Function GetMaterializacionRiskSql() As String
    GetMaterializacionRiskSql = _
        "PARAMETERS p_IDMaterializacion LONG, p_IDRiesgo LONG; " & _
        "SELECT m.[ID] " & _
        "FROM [" & MATERIALIZACION_TABLE & "] AS m " & _
        "INNER JOIN [TbRiesgos] AS r ON m.[IDEdicion]=r.[IDEdicion] " & _
        "AND m.[CodigoRiesgo]=r.[CodigoRiesgo] " & _
        "WHERE m.[ID]=[p_IDMaterializacion] " & _
        "AND r.[" & RIESGO_ID_FIELD & "]=[p_IDRiesgo]"
End Function

Private Function GetCurrentMaterializacionSql() As String
    GetCurrentMaterializacionSql = _
        "PARAMETERS p_IDRiesgo LONG; " & _
        "SELECT TOP 1 m.[ID] " & _
        "FROM [" & MATERIALIZACION_TABLE & "] AS m " & _
        "INNER JOIN [TbRiesgos] AS r ON m.[IDEdicion]=r.[IDEdicion] " & _
        "AND m.[CodigoRiesgo]=r.[CodigoRiesgo] " & _
        "WHERE r.[" & RIESGO_ID_FIELD & "]=[p_IDRiesgo] " & _
        "AND m.[EsMaterializacion]='Sí' " & _
        "ORDER BY m.[Fecha] DESC, m.[ID] DESC"
End Function

Private Function MissingEligibilitySchema(ByVal db As DAO.Database) As String
    MissingEligibilitySchema = JoinMissingSchema( _
        MissingTableFields(db, PLAN_CONTINGENCIA_TABLE, _
            "IDContingencia|IDRiesgo|Estado|FechaDeActivacion|" & _
            "FechaDesactivacion|CodContingencia|DisparadorDelPlan"))
End Function

Private Function MissingTableFields( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldNames As String _
) As String
    Dim fields As Variant
    Dim fieldName As Variant

    If Not TableExists(db, tableName) Then
        MissingTableFields = tableName & "(table missing);"
        Exit Function
    End If

    fields = Split(fieldNames, "|")
    For Each fieldName In fields
        If Not FieldExists(db, tableName, CStr(fieldName)) Then
            MissingTableFields = MissingTableFields & tableName & "." & _
                CStr(fieldName) & ";"
        End If
    Next fieldName
End Function

Private Function JoinMissingSchema(ByVal rawMissing As String) As String
    If Len(rawMissing) = 0 Then Exit Function
    If Right$(rawMissing, 1) = ";" Then
        JoinMissingSchema = Left$(rawMissing, Len(rawMissing) - 1)
    Else
        JoinMissingSchema = rawMissing
    End If
End Function

Private Function TableExists( _
    ByVal db As DAO.Database, _
    ByVal tableName As String _
) As Boolean
    Dim tdf As DAO.TableDef

    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf
End Function

Private Function FieldExists( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String _
) As Boolean
    Dim fld As DAO.Field

    For Each fld In db.TableDefs(tableName).Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fld
End Function

Private Function BuildDiagnosticJson( _
    ByVal schemaKnown As Boolean, _
    ByVal errorMessage As String, _
    ByVal eligibleCount As Long, _
    ByVal sqlElegibles As String _
) As String
    BuildDiagnosticJson = _
        "{""schemaKnown"":" & JsonBool(schemaKnown) & _
        ",""eligibleCount"":" & CStr(eligibleCount) & _
        ",""error"":""" & JsonEscape(errorMessage) & """" & _
        ",""sql"":""" & JsonEscape(sqlElegibles) & """}"
End Function

Private Function JsonBool(ByVal value As Boolean) As String
    If value Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function

Private Function JsonEscape(ByVal value As String) As String
    JsonEscape = Replace(value, Chr$(92), Chr$(92) & Chr$(92))
    JsonEscape = Replace(JsonEscape, Chr$(34), Chr$(92) & Chr$(34))
    JsonEscape = Replace(JsonEscape, vbCrLf, "\n")
    JsonEscape = Replace(JsonEscape, vbCr, "\n")
    JsonEscape = Replace(JsonEscape, vbLf, "\n")
End Function
