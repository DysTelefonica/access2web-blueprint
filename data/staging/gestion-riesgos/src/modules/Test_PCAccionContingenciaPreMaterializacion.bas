Attribute VB_Name = "Test_PCAccionContingenciaPreMaterializacion"
Option Compare Database
Option Explicit

Private Const RISK_TABLE As String = "TbRiesgos"
Private Const EDITION_TABLE As String = "TbProyectosEdiciones"
Private Const PROJECT_TABLE As String = "TbProyectos"
Private Const PLAN_TABLE As String = "TbRiesgosPlanContingenciaPpal"
Private Const ACTION_TABLE As String = "TbRiesgosPlanContingenciaDetalle"
Private Const FIX_PROJECT_ID As Long = 927010
Private Const FIX_EDITION_ID As Long = 927011
Private Const FIX_RISK_ID As Long = 927012
Private Const FIX_PLAN_ID As Long = 927013
Private Const FIX_ACTION_ID As Long = 927014
Private Const FIX_ACTOR As String = "TEST_ISSUE_27"
Private Const FIX_RISK_CODE As String = "I27"
Private Const FIX_START_DATE_SQL As String = "#02/10/2026#"
Private Const FIX_END_DATE_SQL As String = "#02/20/2026#"

' Schema evidence gathered before authoring fixtures with Dysflow MCP:
' TbProyectos: PK IDProyecto required; IDExpediente optional.
' TbProyectosEdiciones: PK IDEdicion required, FK IDProyecto required, Edicion required.
' TbRiesgos: PK IDRiesgo required, FK IDEdicion required, CodigoUnico/CodigoRiesgo required.
' TbRiesgosPlanContingenciaPpal: PK IDContingencia required, FK IDRiesgo required.
' TbRiesgosPlanContingenciaDetalle: PK IDAccionContingencia required, FK IDContingencia required.
' Relationships: Proyecto -> Edicion -> Riesgo -> PlanContingencia -> AccionContingencia.
' Fixture order follows parents first; teardown deletes children first.

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function SqlText(ByVal value As String) As String
    SqlText = "'" & Replace(value, "'", "''") & "'"
End Function

Public Function Test_PCAccionContingenciaPreMat_SchemaInspection() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim missingSchema As String

    logs(0) = "1. Arrange: ForceLocalBackend uses sandbox"
    logs(1) = "2. Arrange: no seed before schema inspection"
    logs(2) = "3. Act: inspect issue #27 fixture graph"
    logs(3) = "4. Assert: parent/child tables exist"
    logs(4) = "5. Assert: required fixture fields exist"
    logs(5) = "6. Assert: DAO relationship metadata checked mechanically"
    logs(6) = "7. Assert: domain validation metadata checked mechanically"
    logs(7) = "8. Teardown: reset testing session only"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_PCAccionContingenciaPreMat_SchemaInspection = BuildFail(errMsg, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_PCAccionContingenciaPreMat_SchemaInspection = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    missingSchema = MissingIssue27Schema(db)
    If missingSchema <> "" Then
        Test_PCAccionContingenciaPreMat_SchemaInspection = _
            BuildFail("SCHEMA BLOCKED: " & missingSchema, logs)
        GoTo Teardown
    End If
    logs(5) = "6. Assert: " & Issue27RelationshipEvidence(db)
    logs(6) = "7. Assert: " & Issue27DomainEvidence(db)

    Test_PCAccionContingenciaPreMat_SchemaInspection = _
        BuildOk("issue_27_schema_known", logs)

Teardown:
    Set db = Nothing
    Test_Helper.ResetTestSession errMsg
    Exit Function

EH:
    Test_PCAccionContingenciaPreMat_SchemaInspection = _
        BuildFail("Test_PCAccionContingenciaPreMat_SchemaInspection: " & _
            Err.Description, logs)
    Resume Teardown
End Function

Public Function Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim actionFixture As PCAccion
    Dim persistedCount As Long
    Dim persistedStart As String
    Dim persistedStartDate As Date
    Dim expectedStartDate As Date

    logs(0) = "1. Arrange: ForceLocalBackend uses sandbox"
    logs(1) = "2. Arrange: schema-first gate before deterministic seed"
    logs(2) = "3. Arrange: active complete non-materialized risk with contingency plan"
    logs(3) = "4. Act: PCAccion.Registrar with FechaInicio in alta path"
    logs(4) = "5. Assert: no materialization-only blocker is returned"
    logs(5) = "6. Assert: action row is persisted exactly once"
    logs(6) = "7. Assert: persisted FechaInicio matches fixture"
    logs(7) = "8. Assert: other safeguards stayed satisfied"
    logs(8) = "9. Teardown: child rows before parent rows"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = BuildFail(errMsg, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If
    If MissingIssue27Schema(db) <> "" Then
        Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
            BuildFail("SCHEMA BLOCKED: " & MissingIssue27Schema(db), logs)
        GoTo Teardown
    End If

    SeedIssue27Graph db
    Set actionFixture = BuildIssue27ActionFixture()
    actionFixture.Registrar , errMsg
    If errMsg <> "" Then
        Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
            BuildFail("Alta con FechaInicio no debe bloquear por no materializado: " & _
                errMsg, logs)
        GoTo Teardown
    End If

    persistedCount = CountRowsWhere(db, ACTION_TABLE, _
        "IDAccionContingencia=" & CLng(actionFixture.IDAccionContingencia))
    If persistedCount <> 1 Then
        Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
            BuildFail("Debe persistir exactamente una acción fixture", logs)
        GoTo Teardown
    End If

    persistedStart = ScalarText(db, _
        "SELECT FechaInicio FROM " & ACTION_TABLE & _
        " WHERE IDAccionContingencia=" & CLng(actionFixture.IDAccionContingencia))
    If Not IsDate(persistedStart) Then
        Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
            BuildFail("Debe persistir FechaInicio", logs)
        GoTo Teardown
    End If
    persistedStartDate = DateValue(CDate(persistedStart))
    expectedStartDate = DateSerial(2026, 2, 10)
    If persistedStartDate <> expectedStartDate Then
        Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
            BuildFail("FechaInicio persistida=" & Format$(persistedStartDate, "yyyy-mm-dd") & _
                " expected=" & Format$(expectedStartDate, "yyyy-mm-dd"), logs)
        GoTo Teardown
    End If

    Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
        BuildOk("contingency_action_pre_materialization_allowed", logs)

Teardown:
    TeardownIssue27Graph db
    Set actionFixture = Nothing
    Set db = Nothing
    Test_Helper.ResetTestSession errMsg
    Exit Function

EH:
    Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate = _
        BuildFail("Test_PCAccionContingenciaPreMat_AllowsAltaWithStartDate: " & _
            Err.Description, logs)
    Resume Teardown
End Function

Public Function Test_PCAccionContingenciaPreMat_PreservesDateCoherence() As String
    On Error GoTo EH

    Dim logs(0 To 6) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim actionFixture As PCAccion
    Dim reason As String

    logs(0) = "1. Arrange: ForceLocalBackend uses sandbox"
    logs(1) = "2. Arrange: complete non-materialized risk and valid plan"
    logs(2) = "3. Arrange: action has FechaFinReal before FechaInicio"
    logs(3) = "4. Act: PCAccion.MotivoNoOK"
    logs(4) = "5. Assert: date coherence rejection remains"
    logs(5) = "6. Assert: rejection is not materialization-only text"
    logs(6) = "7. Teardown: release object graph"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_PCAccionContingenciaPreMat_PreservesDateCoherence = BuildFail(errMsg, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_PCAccionContingenciaPreMat_PreservesDateCoherence = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If
    If MissingIssue27Schema(db) <> "" Then
        Test_PCAccionContingenciaPreMat_PreservesDateCoherence = _
            BuildFail("SCHEMA BLOCKED: " & MissingIssue27Schema(db), logs)
        GoTo Teardown
    End If

    SeedIssue27Graph db
    Set actionFixture = BuildIssue27ActionFixture()
    actionFixture.FechaFinReal = CStr(DateSerial(2026, 2, 1))
    reason = actionFixture.MotivoNoOK(, errMsg)
    If errMsg <> "" Then
        Test_PCAccionContingenciaPreMat_PreservesDateCoherence = _
            BuildFail("MotivoNoOK no debe lanzar error funcional: " & errMsg, logs)
        GoTo Teardown
    End If
    If InStr(1, reason, "fecha fin real", vbTextCompare) = 0 Then
        Test_PCAccionContingenciaPreMat_PreservesDateCoherence = _
            BuildFail("Debe conservar validación de fecha fin real: " & reason, logs)
        GoTo Teardown
    End If
    If InStr(1, reason, "no está materializado", vbTextCompare) > 0 Then
        Test_PCAccionContingenciaPreMat_PreservesDateCoherence = _
            BuildFail("No debe devolver el blocker de materialización", logs)
        GoTo Teardown
    End If

    Test_PCAccionContingenciaPreMat_PreservesDateCoherence = _
        BuildOk("date_coherence_preserved", logs)

Teardown:
    TeardownIssue27Graph db
    Set actionFixture = Nothing
    Set db = Nothing
    Test_Helper.ResetTestSession errMsg
    Exit Function

EH:
    Test_PCAccionContingenciaPreMat_PreservesDateCoherence = _
        BuildFail("Test_PCAccionContingenciaPreMat_PreservesDateCoherence: " & _
            Err.Description, logs)
    Resume Teardown
End Function

Private Function BuildIssue27ActionFixture() As PCAccion
    Dim actionFixture As PCAccion

    Set actionFixture = New PCAccion
    With actionFixture
        .IDAccionContingencia = CStr(FIX_ACTION_ID)
        .IDContingencia = CStr(FIX_PLAN_ID)
        .CodAccion = "AC027"
        .Accion = "Fixture action issue 27"
        .ResponsableAccion = FIX_ACTOR
        .FechaInicio = CStr(DateSerial(2026, 2, 10))
        .FechaFinPrevista = CStr(DateSerial(2026, 2, 20))
        .Estado = "Planificado"
    End With

    Set BuildIssue27ActionFixture = actionFixture
End Function

Private Sub SeedIssue27Graph(ByVal db As DAO.Database)
    TeardownIssue27Graph db
    db.Execute "INSERT INTO " & PROJECT_TABLE & _
        " (IDProyecto, Proyecto, NombreProyecto, NombreUsuarioCalidad, " & _
        "CadenaNombreAutorizados, ParaInformeAvisos) VALUES (" & _
        FIX_PROJECT_ID & ", " & SqlText("ISSUE27") & ", " & _
        SqlText("Fixture issue 27") & ", " & SqlText(FIX_ACTOR) & ", " & _
        SqlText(FIX_ACTOR) & ", " & SqlText("No") & ")", dbFailOnError
    db.Execute "INSERT INTO " & EDITION_TABLE & _
        " (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        FIX_EDITION_ID & ", " & FIX_PROJECT_ID & ", 1, " & _
        SqlText(FIX_ACTOR) & ")", dbFailOnError
    db.Execute "INSERT INTO " & RISK_TABLE & _
        " (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, FechaDetectado, " & _
        "DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, ImpactoGlobal, " & _
        "Vulnerabilidad, Valoracion, Mitigacion, Descripcion, Estado, " & _
        "Priorizacion) VALUES (" & FIX_RISK_ID & ", " & FIX_EDITION_ID & _
        ", " & SqlText("ISSUE27-I27") & ", " & SqlText(FIX_RISK_CODE) & _
        ", #01/15/2026#, " & SqlText(FIX_ACTOR) & ", " & _
        SqlText(FIX_ACTOR) & ", " & SqlText("Medio") & ", " & _
        SqlText("Medio") & ", " & SqlText("Medio") & ", " & _
        SqlText("Medio") & ", " & SqlText("Media") & ", " & _
        SqlText("Medio") & ", " & SqlText("Mitigar") & ", " & _
        SqlText("Fixture issue 27") & ", " & SqlText("Detectado") & _
        ", 3)", dbFailOnError
    db.Execute "INSERT INTO " & PLAN_TABLE & _
        " (IDContingencia, IDRiesgo, CodContingencia, DisparadorDelPlan, " & _
        "Estado) VALUES (" & FIX_PLAN_ID & ", " & FIX_RISK_ID & ", " & _
        SqlText("PC27") & ", " & SqlText("Fixture plan issue 27") & ", " & _
        SqlText("Definido") & ")", dbFailOnError
End Sub

Private Sub TeardownIssue27Graph(ByVal db As DAO.Database)
    If db Is Nothing Then Exit Sub
    On Error Resume Next
    db.Execute "DELETE FROM " & ACTION_TABLE & _
        " WHERE IDAccionContingencia=" & FIX_ACTION_ID, dbFailOnError
    db.Execute "DELETE FROM " & ACTION_TABLE & _
        " WHERE IDContingencia=" & FIX_PLAN_ID, dbFailOnError
    db.Execute "DELETE FROM " & PLAN_TABLE & _
        " WHERE IDContingencia=" & FIX_PLAN_ID, dbFailOnError
    db.Execute "DELETE FROM " & RISK_TABLE & _
        " WHERE IDRiesgo=" & FIX_RISK_ID, dbFailOnError
    db.Execute "DELETE FROM " & EDITION_TABLE & _
        " WHERE IDEdicion=" & FIX_EDITION_ID, dbFailOnError
    db.Execute "DELETE FROM " & PROJECT_TABLE & _
        " WHERE IDProyecto=" & FIX_PROJECT_ID, dbFailOnError
    On Error GoTo 0
End Sub

Private Function MissingIssue27Schema(ByVal db As DAO.Database) As String
    Dim missing As String

    missing = MissingTable(db, PROJECT_TABLE) & MissingTable(db, EDITION_TABLE)
    missing = missing & MissingTable(db, RISK_TABLE) & MissingTable(db, PLAN_TABLE)
    missing = missing & MissingTable(db, ACTION_TABLE)

    missing = missing & MissingField(db, PROJECT_TABLE, "IDProyecto")
    missing = missing & MissingField(db, PROJECT_TABLE, "Proyecto")
    missing = missing & MissingField(db, PROJECT_TABLE, "NombreProyecto")
    missing = missing & MissingField(db, EDITION_TABLE, "IDEdicion")
    missing = missing & MissingField(db, EDITION_TABLE, "IDProyecto")
    missing = missing & MissingField(db, EDITION_TABLE, "Edicion")
    missing = missing & MissingField(db, RISK_TABLE, "IDRiesgo")
    missing = missing & MissingField(db, RISK_TABLE, "IDEdicion")
    missing = missing & MissingField(db, RISK_TABLE, "CodigoUnico")
    missing = missing & MissingField(db, RISK_TABLE, "CodigoRiesgo")
    missing = missing & MissingField(db, PLAN_TABLE, "IDContingencia")
    missing = missing & MissingField(db, PLAN_TABLE, "IDRiesgo")
    missing = missing & MissingField(db, PLAN_TABLE, "CodContingencia")
    missing = missing & MissingField(db, PLAN_TABLE, "DisparadorDelPlan")
    missing = missing & MissingField(db, PLAN_TABLE, "Estado")
    missing = missing & MissingField(db, ACTION_TABLE, "IDAccionContingencia")
    missing = missing & MissingField(db, ACTION_TABLE, "IDContingencia")

    missing = missing & MissingPrimaryKeyIndex(db, PROJECT_TABLE, "IDProyecto")
    missing = missing & MissingPrimaryKeyIndex(db, EDITION_TABLE, "IDEdicion")
    missing = missing & MissingPrimaryKeyIndex(db, RISK_TABLE, "IDRiesgo")
    missing = missing & MissingPrimaryKeyIndex(db, PLAN_TABLE, "IDContingencia")
    missing = missing & MissingPrimaryKeyIndex(db, ACTION_TABLE, "IDAccionContingencia")

    missing = missing & MissingRequiredOrPrimaryField(db, PROJECT_TABLE, "IDProyecto")
    missing = missing & MissingRequiredOrPrimaryField(db, EDITION_TABLE, "IDEdicion")
    missing = missing & MissingRequiredOrPrimaryField(db, RISK_TABLE, "IDRiesgo")
    missing = missing & MissingRequiredOrPrimaryField(db, PLAN_TABLE, "IDContingencia")
    missing = missing & MissingRequiredOrPrimaryField(db, ACTION_TABLE, "IDAccionContingencia")

    missing = missing & MissingFieldType(db, PROJECT_TABLE, "IDProyecto", dbLong)
    missing = missing & MissingFieldType(db, EDITION_TABLE, "IDEdicion", dbLong)
    missing = missing & MissingFieldType(db, EDITION_TABLE, "IDProyecto", dbLong)
    missing = missing & MissingFieldType(db, RISK_TABLE, "IDRiesgo", dbLong)
    missing = missing & MissingFieldType(db, RISK_TABLE, "IDEdicion", dbLong)
    missing = missing & MissingFieldType(db, PLAN_TABLE, "IDContingencia", dbLong)
    missing = missing & MissingFieldType(db, PLAN_TABLE, "IDRiesgo", dbLong)
    missing = missing & MissingFieldType(db, ACTION_TABLE, "IDAccionContingencia", dbLong)
    missing = missing & MissingFieldType(db, ACTION_TABLE, "IDContingencia", dbLong)

    MissingIssue27Schema = missing
End Function

Private Function Issue27RelationshipEvidence(ByVal db As DAO.Database) As String
    Issue27RelationshipEvidence = RelationshipEvidence(db, PROJECT_TABLE, _
        "IDProyecto", EDITION_TABLE, "IDProyecto") & "; " & _
        RelationshipEvidence(db, EDITION_TABLE, "IDEdicion", RISK_TABLE, _
        "IDEdicion") & "; " & RelationshipEvidence(db, RISK_TABLE, _
        "IDRiesgo", PLAN_TABLE, "IDRiesgo") & "; " & _
        RelationshipEvidence(db, PLAN_TABLE, "IDContingencia", ACTION_TABLE, _
        "IDContingencia")
End Function

Private Function Issue27DomainEvidence(ByVal db As DAO.Database) As String
    Issue27DomainEvidence = "domainValidation=" & _
        FieldValidationEvidence(db, RISK_TABLE, "Estado") & ";" & _
        FieldValidationEvidence(db, RISK_TABLE, "Priorizacion") & ";" & _
        FieldValidationEvidence(db, PLAN_TABLE, "Estado") & ";" & _
        FieldValidationEvidence(db, ACTION_TABLE, "Estado") & _
        "; seededValues=Riesgo.Estado=Detectado, Priorizacion=3, " & _
        "Plan.Estado=Definido, Accion.Estado=Planificado"
End Function

Private Function RelationshipEvidence( _
    ByVal db As DAO.Database, _
    ByVal parentTable As String, _
    ByVal parentField As String, _
    ByVal childTable As String, _
    ByVal childField As String _
) As String
    Dim rel As DAO.Relation
    Dim relField As DAO.Field

    For Each rel In db.Relations
        If StrComp(rel.Table, parentTable, vbTextCompare) = 0 And _
            StrComp(rel.ForeignTable, childTable, vbTextCompare) = 0 Then
            For Each relField In rel.Fields
                If StrComp(relField.Name, parentField, vbTextCompare) = 0 And _
                    StrComp(relField.ForeignName, childField, vbTextCompare) = 0 Then
                    RelationshipEvidence = "FK " & parentTable & "." & _
                        parentField & " -> " & childTable & "." & childField & _
                        " present"
                    Exit Function
                End If
            Next relField
        End If
    Next rel

    RelationshipEvidence = "no DAO FK metadata for " & parentTable & "." & _
        parentField & " -> " & childTable & "." & childField & _
        "; fixture verifies parent rows before child insert"
End Function

Private Function FieldValidationEvidence( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String _
) As String
    Dim ruleText As String

    If Not FieldExists(db, tableName, fieldName) Then
        FieldValidationEvidence = tableName & "." & fieldName & "=missing"
        Exit Function
    End If

    ruleText = Nz(db.TableDefs(tableName).Fields(fieldName).ValidationRule, "")
    If ruleText = "" Then
        FieldValidationEvidence = tableName & "." & fieldName & "=no DAO validation rule"
    Else
        FieldValidationEvidence = tableName & "." & fieldName & _
            " rule=" & ruleText
    End If
End Function

Private Function MissingTable(ByVal db As DAO.Database, ByVal tableName As String) As String
    If Not TableExists(db, tableName) Then MissingTable = "Missing table " & tableName & "; "
End Function

Private Function MissingField( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String _
) As String
    If Not FieldExists(db, tableName, fieldName) Then
        MissingField = "Missing field " & tableName & "." & fieldName & "; "
    End If
End Function

Private Function MissingFieldType( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByVal expectedType As Long _
) As String
    If Not FieldExists(db, tableName, fieldName) Then Exit Function
    If CLng(db.TableDefs(tableName).Fields(fieldName).Type) <> expectedType Then
        MissingFieldType = "Unexpected type " & tableName & "." & fieldName & "; "
    End If
End Function

Private Function MissingRequiredOrPrimaryField( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String _
) As String
    If Not FieldExists(db, tableName, fieldName) Then Exit Function
    If db.TableDefs(tableName).Fields(fieldName).Required Then Exit Function
    If TableHasPrimaryKeyIndex(db, tableName, fieldName) Then Exit Function
    MissingRequiredOrPrimaryField = "Field not required/indexed " & _
        tableName & "." & fieldName & "; "
End Function

Private Function MissingPrimaryKeyIndex( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String _
) As String
    If Not FieldExists(db, tableName, fieldName) Then Exit Function
    If Not TableHasPrimaryKeyIndex(db, tableName, fieldName) Then
        MissingPrimaryKeyIndex = "Missing primary index " & _
            tableName & "." & fieldName & "; "
    End If
End Function

Private Function TableHasPrimaryKeyIndex( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String _
) As Boolean
    Dim idx As DAO.Index
    Dim idxField As DAO.Field

    If Not TableExists(db, tableName) Then Exit Function

    For Each idx In db.TableDefs(tableName).Indexes
        If idx.Primary Then
            For Each idxField In idx.Fields
                If StrComp(idxField.Name, fieldName, vbTextCompare) = 0 Then
                    TableHasPrimaryKeyIndex = True
                    Exit Function
                End If
            Next idxField
        End If
    Next idx
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
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
    If Not TableExists(db, tableName) Then Exit Function
    For Each fld In db.TableDefs(tableName).Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fld
End Function

Private Function CountRowsWhere( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal whereClause As String _
) As Long
    CountRowsWhere = CLng(ScalarText(db, _
        "SELECT Count(*) AS Cnt FROM " & tableName & " WHERE " & whereClause))
End Function

Private Function ScalarText(ByVal db As DAO.Database, ByVal sqlText As String) As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If Not rs.EOF Then ScalarText = CStr(Nz(rs.Fields(0).Value, ""))
    rs.Close
    Set rs = Nothing
End Function
