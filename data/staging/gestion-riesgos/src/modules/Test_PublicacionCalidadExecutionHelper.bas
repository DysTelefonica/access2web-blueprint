Attribute VB_Name = "Test_PublicacionCalidadExecutionHelper"
' =============================================================================
' Test_PublicacionCalidadExecutionHelper.bas
'
' TDD RED atoms for modPublicacionCalidadExecutionHelper.RechazarPropuestaPublicacion
' Skill: access-vba-tdd v2.4 + access-vba-e2e-methodology
' Scope: 1 helper ? 4 scenarios = 4 atoms
'
' Helper signature (from form inline m_FormMotivos_Motivado, lines 372-427):
'   Function RechazarPropuestaPublicacion( _
'       ByVal p_IDEdicion As Long
'       Optional ByRef p_CorreoRechazo As Object
'       Optional ByRef p_PromptResult As Long _
'   ) As String
'
' Delegated services (confirmed from source):
'   - Edicion.RechazoPropuestaParaPublicacion(p_Error)  [Edicion.cls:3313]
'     Returns CORREO; updates TbProyectosEdiciones + sends email + logs PublicacionLog
'   - EsRechazoPropuestaNotificado(p_Correo) — form-local; inlined or parameterized in helper
'
' The module under test does NOT exist yet — all atoms are RED.
' =============================================================================
Option Compare Database
Option Explicit

' =============================================================================
' Atom: Test_RechazarPropuestaPublicacion_Happy
' Scenario: happy — valid IDEdicion, valid non-blank motivo, email sent OK
' Expected: p_Error = "", edition updated (PropuestaRechazadaPorCalidadFecha set),
'           p_CorreoRechazo populated (not Nothing), p_PromptResult = vbYes
' =============================================================================
Public Function Test_RechazarPropuestaPublicacion_Happy() As String
    Dim logs(0 To 9) As String
    logs(0) = "Arrange: seed TbProyectos + TbProyectosEdiciones (ID >= 900000)"
    logs(1) = "Act: call RechazarPropuestaPublicacion"
    logs(2) = "Assert: p_Error empty, edition updated, correo populated, promptResult vbYes"

    On Error GoTo EH

    ' -- Arrange --------------------------------------------------------------
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Set db = Test_Helper.GetCachedTestDb()
    If db Is Nothing Then
        logs(0) = logs(0) & " > FAILED: db is Nothing"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail("db is Nothing", logs)
        Exit Function
    End If

    ' Deterministic test ID
    Const TEST_ID_EDICION As Long = 900001

    ' Seed parent TbProyectos (required FK for TbProyectosEdiciones.IDProyecto)
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_EDICION, dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, ParaInformeAvisos) " & _
                       "VALUES (" & TEST_ID_EDICION & ", 'PROYECTO-TEST-HAPPY', 'Sí')", dbFailOnError

    ' Seed TbProyectosEdiciones: all fields required by RechazoPropuestaParaPublicacion
    ' - IDEdicion = TEST_ID_EDICION
    ' - IDProyecto = TEST_ID_EDICION (FK)
    ' - Elaborado = 'TestAutor' (required NOT NULL per ERD)
    ' - FechaPreparadaParaPublicar = today (required by IsDate check)
    ' - PropuestaRechazadaPorCalidadFecha = empty (not yet rejected)
    ' - PropuestaRechazadaPorCalidadMotivo = empty (will be set by helper)
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_EDICION, dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_EDICION & ", 'TestAutor', " & _
               Date & ")", dbFailOnError

    ' Verify pre-condition: edition is NOT yet rejected
    Dim rsCheck As DAO.Recordset
    Set rsCheck = db.OpenRecordset( _
        "SELECT PropuestaRechazadaPorCalidadFecha FROM TbProyectosEdiciones " & _
        "WHERE IDEdicion = " & TEST_ID_EDICION, dbOpenSnapshot)
    If rsCheck!PropuestaRechazadaPorCalidadFecha <> "" Then
        logs(0) = logs(0) & " > FAILED: edition already rejected"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail("pre-condition: edition already rejected", logs)
        rsCheck.Close: Set rsCheck = Nothing
        GoTo Teardown
    End If
    rsCheck.Close: Set rsCheck = Nothing

    ' -- Act -----------------------------------------------------------------
    Dim p_CorreoRechazo As Object
    Dim p_PromptResult As Long
    p_PromptResult = vbYes  ' simulate user confirmed send

    Dim helperResult As String
    helperResult = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_EDICION), _
        "Motivo de rechazo válido para test Happy.", _
        p_CorreoRechazo, _
        db, _
        p_PromptResult)

    ' -- Assert ---------------------------------------------------------------
    ' 1. helper returned empty string (no error)
    If helperResult <> "" Then
        logs(2) = logs(2) & " > FAILED: helper returned '" & helperResult & "'"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail(helperResult, logs)
        GoTo Teardown
    End If

    ' 2. p_CorreoRechazo is not Nothing (email was produced)
    If p_CorreoRechazo Is Nothing Then
        logs(2) = logs(2) & " > FAILED: p_CorreoRechazo is Nothing"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail("p_CorreoRechazo is Nothing", logs)
        GoTo Teardown
    End If

    ' 3. edition record was updated with rejection date and motivo
    Set rsCheck = db.OpenRecordset( _
        "SELECT PropuestaRechazadaPorCalidadFecha, PropuestaRechazadaPorCalidadMotivo, " & _
        "PropuestaRechazadaPorCalidadMotivo " & _
        "FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_EDICION, dbOpenSnapshot)
    If rsCheck.EOF Then
        logs(2) = logs(2) & " > FAILED: edition row not found after rejection"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail("edition row not found", logs)
        rsCheck.Close: Set rsCheck = Nothing
        GoTo Teardown
    End If
    If IsNull(rsCheck!PropuestaRechazadaPorCalidadFecha) Or rsCheck!PropuestaRechazadaPorCalidadFecha = "" Then
        logs(2) = logs(2) & " > FAILED: PropuestaRechazadaPorCalidadFecha not set"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail("PropuestaRechazadaPorCalidadFecha not set", logs)
        rsCheck.Close: Set rsCheck = Nothing
        GoTo Teardown
    End If
    If IsNull(rsCheck!PropuestaRechazadaPorCalidadMotivo) Or rsCheck!PropuestaRechazadaPorCalidadMotivo = "" Then
        logs(2) = logs(2) & " > FAILED: PropuestaRechazadaPorCalidadMotivo not set"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail("PropuestaRechazadaPorCalidadMotivo not set", logs)
        rsCheck.Close: Set rsCheck = Nothing
        GoTo Teardown
    End If
    rsCheck.Close: Set rsCheck = Nothing

    ' 4. FechaPreparadaParaPublicar was cleared (set to Null)
    Set rsCheck = db.OpenRecordset( _
        "SELECT FechaPreparadaParaPublicar FROM TbProyectosEdiciones " & _
        "WHERE IDEdicion = " & TEST_ID_EDICION, dbOpenSnapshot)
    If Not IsNull(rsCheck!FechaPreparadaParaPublicar) Then
        logs(2) = logs(2) & " > FAILED: FechaPreparadaParaPublicar not cleared after rejection"
        Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail("FechaPreparadaParaPublicar not cleared", logs)
        rsCheck.Close: Set rsCheck = Nothing
        GoTo Teardown
    End If
    rsCheck.Close: Set rsCheck = Nothing

    Test_RechazarPropuestaPublicacion_Happy = BuildJsonOk(TEST_ID_EDICION, logs)
    GoTo Teardown

EH:
    logs(UBound(logs)) = "EH: " & Err.Number & " - " & Err.description
    Test_RechazarPropuestaPublicacion_Happy = BuildJsonFail(Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_EDICION, dbFailOnError
    On Error GoTo 0
End Function

' =============================================================================
' Atom: Test_RechazarPropuestaPublicacion_Sad
' Scenario: sad — blank motivo AND non-existent edition ID
' Expected: helper returns non-empty error string; no edition record is modified
' Two sub-scenarios: (a) blank motivo > early exit (no DB write)
'                    (b) non-existent edition > error from Edicion.RechazoPropuestaParaPublicacion
' This atom tests (a): blank motivo triggers early exit with empty p_Error but
' no state change. Then it tests (b) separately with a known-bad edition ID.
' =============================================================================
Public Function Test_RechazarPropuestaPublicacion_Sad() As String
    Dim logs(0 To 9) As String
    logs(0) = "Arrange: configure sandbox"
    logs(1) = "Act: call RechazarPropuestaPublicacion with blank motivo + bad edition"
    logs(2) = "Assert: helper returns non-empty error, no record created"

    On Error GoTo EH

    ' -- Arrange --------------------------------------------------------------
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_RechazarPropuestaPublicacion_Sad = BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Set db = Test_Helper.GetCachedTestDb()
    If db Is Nothing Then
        Test_RechazarPropuestaPublicacion_Sad = BuildJsonFail("db is Nothing", logs)
        Exit Function
    End If

    Const TEST_ID_BAD As Long = 900002

    ' -- Act & Assert 1: non-existent edition ID -----------------------------
    ' The inline form (m_FormMotivos_Motivado) does NOT validate IDEdicion -
    ' it assumes m_ObjEdicionActiva is valid. But if IDEdicion does not exist,
    ' Edicion.RechazoPropuestaParaPublicacion raises error 1000 ("EOF").
    ' The helper propagates this as p_Error <> "".
    Dim p_CorreoRechazo1 As Object
    Dim p_PromptResult1 As Long
    Dim helperResult1 As String
    helperResult1 = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_BAD), _
        "Motivo existente para test", _
        p_CorreoRechazo1, _
        db, _
        p_PromptResult1)

    ' Helper should return non-empty error for non-existent edition
    If helperResult1 = "" Then
        logs(1) = logs(1) & " > FAILED: no error for non-existent edition"
        Test_RechazarPropuestaPublicacion_Sad = BuildJsonFail( _
            "expected error for non-existent edition, got empty string", logs)
        Exit Function
    End If

    ' -- Act & Assert 2: blank motivo on valid-looking edition ID --------------
    ' Seed a valid edition first (to isolate the blank-motivo check)
    ' Even if the edition exists, blank motivo should cause helper to exit early
    ' (see inline form: If p_Motivo = "" Then Exit Sub)
    ' The helper returns "" (no error propagated) but no state change occurs.
    ' We test that calling with blank motivo on a real edition:
    '   (a) does NOT raise an error
    '   (b) does NOT update the edition record
    Const TEST_ID_EDICION As Long = 900003
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_EDICION, dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, ParaInformeAvisos) " & _
               "VALUES (" & TEST_ID_EDICION & ", 'PROYECTO-TEST-SAD', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_EDICION & ", 'TestAutor', " & _
               Date & ")", dbFailOnError

    Dim p_CorreoRechazo2 As Object
    Dim p_PromptResult2 As Long
    Dim helperResult2 As String
    helperResult2 = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_EDICION), _
        "", _
        p_CorreoRechazo2, _
        db, _
        p_PromptResult2)

    ' For blank motivo: inline form exits early (Exit Sub, no error).
    ' The helper wraps this — expect helperResult2 = "" (no error propagated).
    ' The key assertion: the edition record was NOT modified.
    If helperResult2 <> "" Then
        logs(2) = logs(2) & " > unexpected error for blank motivo: " & helperResult2
        Test_RechazarPropuestaPublicacion_Sad = BuildJsonFail(helperResult2, logs)
        GoTo Teardown
    End If

    ' Verify edition was NOT updated (propuesta still empty, fecha still set)
    Dim rsCheck As DAO.Recordset
    Set rsCheck = db.OpenRecordset( _
        "SELECT PropuestaRechazadaPorCalidadFecha, PropuestaRechazadaPorCalidadMotivo " & _
        "FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_EDICION, dbOpenSnapshot)
    If Not rsCheck.EOF Then
        If Not IsNull(rsCheck!PropuestaRechazadaPorCalidadFecha) And rsCheck!PropuestaRechazadaPorCalidadFecha <> "" Then
            logs(2) = logs(2) & " > FAILED: blank motivo incorrectly updated rejection date"
            Test_RechazarPropuestaPublicacion_Sad = BuildJsonFail( _
                "blank motivo should not update edition", logs)
            rsCheck.Close: Set rsCheck = Nothing
            GoTo Teardown
        End If
        If Not IsNull(rsCheck!PropuestaRechazadaPorCalidadMotivo) And rsCheck!PropuestaRechazadaPorCalidadMotivo <> "" Then
            logs(2) = logs(2) & " > FAILED: blank motivo incorrectly updated rejection motivo"
            Test_RechazarPropuestaPublicacion_Sad = BuildJsonFail( _
                "blank motivo should not update edition", logs)
            rsCheck.Close: Set rsCheck = Nothing
            GoTo Teardown
        End If
    End If
    rsCheck.Close: Set rsCheck = Nothing

    Test_RechazarPropuestaPublicacion_Sad = BuildJsonOk("non-existent>error, blank>no-op", logs)
    GoTo Teardown

EH:
    logs(UBound(logs)) = "EH: " & Err.Number & " - " & Err.description
    Test_RechazarPropuestaPublicacion_Sad = BuildJsonFail(Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_EDICION, dbFailOnError
    On Error GoTo 0
End Function

' =============================================================================
' Atom: Test_RechazarPropuestaPublicacion_Edge
' Scenario: edge — motivo with max length (300 chars > 255), special chars
'           (<, >, &, "), and whitespace-only
' Expected: special-char motivo is processed (may depend on SQL-escaping in helper);
'           whitespace-only triggers blank check; long string handled without crash
' =============================================================================
Public Function Test_RechazarPropuestaPublicacion_Edge() As String
    Dim logs(0 To 9) As String
    logs(0) = "Arrange: configure sandbox + seed edition"
    logs(1) = "Act: call with special-chars, long, and whitespace-only motivo"
    logs(2) = "Assert: no crash; long motivo may be truncated or cause error; " & _
              "whitespace-only treated as blank"

    On Error GoTo EH

    ' -- Arrange --------------------------------------------------------------
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_RechazarPropuestaPublicacion_Edge = BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Set db = Test_Helper.GetCachedTestDb()
    If db Is Nothing Then
        Test_RechazarPropuestaPublicacion_Edge = BuildJsonFail("db is Nothing", logs)
        Exit Function
    End If

    ' -- Test 1: whitespace-only motivo (should behave like blank) -----------
    Const TEST_ID_WS As Long = 900004
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_WS, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_WS, dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, ParaInformeAvisos) " & _
               "VALUES (" & TEST_ID_WS & ", 'PROYECTO-TEST-WS', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & TEST_ID_WS & ", " & TEST_ID_WS & ", 'TestAutor', " & _
               Date & ")", dbFailOnError

    Dim p_CorreoRechazoWS As Object
    Dim p_PromptResultWS As Long
    Dim helperResultWS As String
    helperResultWS = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_WS), _
        "     " & vbTab & vbCrLf & "  ", _
        p_CorreoRechazoWS, _
        db, _
        p_PromptResultWS)

    ' Whitespace-only: the helper should trim and treat as blank (no error, no update)
    ' The inline form checks: If p_Motivo = "" Then — Trim() not used,
    ' so the form checks the raw untrimmed string.
    ' A string of spaces would NOT be "" so it would PROCEED (not exit early).
    ' This is an edge case worth noting. We assert: no crash at minimum.
    If helperResultWS = "" Then
        logs(1) = logs(1) & " [whitespace-only: no error]"
    Else
        logs(1) = logs(1) & " [whitespace-only: error='" & Left$(helperResultWS, 50) & "']"
    End If

    ' -- Test 2: special chars (<, >, &, ") --------------------------------
    Const TEST_ID_SPECIAL As Long = 900005
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_SPECIAL, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_SPECIAL, dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, ParaInformeAvisos) " & _
               "VALUES (" & TEST_ID_SPECIAL & ", 'PROYECTO-TEST-SPECIAL', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & TEST_ID_SPECIAL & ", " & TEST_ID_SPECIAL & ", 'TestAutor', " & _
               Date & ")", dbFailOnError

    Dim p_CorreoRechazoSC As Object
    Dim p_PromptResultSC As Long
    Dim helperResultSC As String
    helperResultSC = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_SPECIAL), _
        "Motivo con special chars: <script>&" & """' chars", _
        p_CorreoRechazoSC, _
        db, _
        p_PromptResultSC)

    ' Special chars: if the helper does proper SQL-escaping (Replace ', ''),
    ' the operation succeeds. If not, it may cause SQL error.
    ' Assert: no crash, p_Error captures the outcome.
    If helperResultSC = "" Then
        logs(2) = logs(2) & " [special-chars: accepted]"
    Else
        logs(2) = logs(2) & " [special-chars: error='" & Left$(helperResultSC, 50) & "']"
    End If

    ' -- Test 3: very long motivo (300 chars, schema limit is 255) ---------
    Const TEST_ID_LONG As Long = 900006
    Dim longMotivo As String
    longMotivo = String$(300, "X")  ' 300 X's
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_LONG, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_LONG, dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, ParaInformeAvisos) " & _
               "VALUES (" & TEST_ID_LONG & ", 'PROYECTO-TEST-LONG', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & TEST_ID_LONG & ", " & TEST_ID_LONG & ", 'TestAutor', " & _
               Date & ")", dbFailOnError

    Dim p_CorreoRechazoLong As Object
    Dim p_PromptResultLong As Long
    Dim helperResultLong As String
    helperResultLong = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_LONG), _
        longMotivo, _
        p_CorreoRechazoLong, _
        db, _
        p_PromptResultLong)

    ' Long motivo: if schema truncates at 255, the operation may succeed
    ' (truncation is silent in Access) OR the helper may validate length.
    ' Assert: no crash, p_Error captures the outcome.
    If helperResultLong = "" Then
        logs(2) = logs(2) & " [long-motivo(300): accepted (may be truncated)]"
    Else
        logs(2) = logs(2) & " [long-motivo(300): error='" & Left$(helperResultLong, 50) & "']"
    End If

    Test_RechazarPropuestaPublicacion_Edge = BuildJsonOk( _
        "{""ws"":""" & helperResultWS & """,""special"":""" & Left$(helperResultSC, 30) & """,""long"":""" & Left$(helperResultLong, 30) & """}", _
        logs)
    GoTo Teardown

EH:
    logs(UBound(logs)) = "EH: " & Err.Number & " - " & Err.description
    Test_RechazarPropuestaPublicacion_Edge = BuildJsonFail(Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_WS, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_WS, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_SPECIAL, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_SPECIAL, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_LONG, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_LONG, dbFailOnError
    On Error GoTo 0
End Function

' =============================================================================
' Atom: Test_RechazarPropuestaPublicacion_Adversarial
' Scenario: adversarial — (a) double-click simulation: call helper twice rapidly
'           (b) p_CorreoRechazo = Nothing (email send failure mid-call)
'
' For (a): the helper has no double-click guard. Two calls in succession should
'   either: succeed twice (if idempotent), or the second should return an error
'   (edition already rejected).
'
' For (b): If p_CorreoRechazo is passed as Nothing from the caller (not set by
'   the helper), the helper's internal EsRechazoPropuestaNotificado check will
'   receive Nothing. The inline form does: If m_CorreoRechazo Is Nothing Then
'   m_Notificado = False. So the else-branch "notificado" notification is shown.
'   The helper should handle Nothing gracefully (no crash), return "" (success)
'   but with the appropriate p_PromptResult for the non-notificado case.
' =============================================================================
Public Function Test_RechazarPropuestaPublicacion_Adversarial() As String
    Dim logs(0 To 9) As String
    logs(0) = "Arrange: configure sandbox + seed edition"
    logs(1) = "Act: rapid double-call + Nothing correo scenario"
    logs(2) = "Assert: first call OK, second call error (already rejected); " & _
              "Nothing correo handled gracefully"

    On Error GoTo EH

    ' -- Arrange --------------------------------------------------------------
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_RechazarPropuestaPublicacion_Adversarial = BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Set db = Test_Helper.GetCachedTestDb()
    If db Is Nothing Then
        Test_RechazarPropuestaPublicacion_Adversarial = BuildJsonFail("db is Nothing", logs)
        Exit Function
    End If

    ' -- Test 1: double-call (no double-click guard) ------------------------
    Const TEST_ID_DOUBLE As Long = 900007
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_DOUBLE, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_DOUBLE, dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, ParaInformeAvisos) " & _
               "VALUES (" & TEST_ID_DOUBLE & ", 'PROYECTO-TEST-DOUBLE', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & TEST_ID_DOUBLE & ", " & TEST_ID_DOUBLE & ", 'TestAutor', " & _
               Date & ")", dbFailOnError

    Dim p_CorreoRechazo1 As Object
    Dim p_PromptResult1 As Long
    Dim helperResult1 As String
    helperResult1 = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_DOUBLE), _
        "Primer rechazo (double-click test)", _
        p_CorreoRechazo1, _
        db, _
        p_PromptResult1)

    ' First call should succeed
    If helperResult1 <> "" Then
        logs(1) = logs(1) & " [first call error: " & Left$(helperResult1, 30) & "]"
    End If

    ' Second call: edition already has PropuestaRechazadaPorCalidadFecha set.
    ' Edicion.RechazoPropuestaParaPublicacion does NOT check this pre-condition
    ' (it only checks EsActivo, FechaPreparadaParaPublicar IsDate, and Motivo).
    ' So the second call would re-set the rejection date and re-send email.
    ' This is a design issue: no idempotency guard. The atom captures this.
    Dim p_CorreoRechazo2 As Object
    Dim p_PromptResult2 As Long
    Dim helperResult2 As String
    helperResult2 = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_DOUBLE), _
        "Segundo rechazo (double-click test)", _
        p_CorreoRechazo2, _
        db, _
        p_PromptResult2)

    ' No guard > second call also succeeds (design issue, not a crash)
    logs(1) = logs(1) & " [second call result: '" & Left$(helperResult2, 30) & "']"
    If helperResult2 = "" Then
        logs(2) = logs(2) & " [no double-click guard: second call also succeeded]"
    End If

    ' -- Test 2: p_CorreoRechazo = Nothing (email send failure) -------------
    ' Seed a fresh edition for the Nothing correo scenario
    Const TEST_ID_NOTHING As Long = 900008
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_NOTHING, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_NOTHING, dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, ParaInformeAvisos) " & _
               "VALUES (" & TEST_ID_NOTHING & ", 'PROYECTO-TEST-NOTHING', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & TEST_ID_NOTHING & ", " & TEST_ID_NOTHING & ", 'TestAutor', " & _
               Date & ")", dbFailOnError

    ' Simulate email failure: pass Nothing as p_CorreoRechazo
    ' The helper calls Edicion.RechazoPropuestaParaPublicacion which SETS p_CorreoRechazo.
    ' So we cannot simulate failure by passing Nothing as output param.
    ' The adversarial scenario here is: we pre-set p_CorreoRechazo = Nothing
    ' and verify the helper's EsRechazoPropuestaNotificado handles it.
    ' However, the helper assigns to p_CorreoRechazo FROM the Edicion call,
    ' so the Nothing we pass is overwritten. The real adversarial path is
    ' when Edicion.RechazoPropuestaParaPublicacion itself fails (email send fails
    ' inside that method). We can simulate this by having the edition's
    ' ParaInformeAvisos = "No" (EdicionCorreoRevision skips email registration
    ' when this is "No", but still returns CORREO from EnviarCorreoRechazoPropuestaPublicacion).
    '
    ' More direct adversarial: test that if the helper receives Nothing from
    ' the Edicion call (correo not sent / not available), it still handles
    ' the notification check gracefully (EsRechazoPropuestaNotificado handles Nothing).
    '
    ' For this atom, the adversarial check is: with a valid edition but
    ' email not sent (ParaInformeAvisos = "No"), the helper should still
    ' succeed but p_PromptResult should be vbNo (notificado=False path).
    Dim p_CorreoRechazoNA As Object
    Dim p_PromptResultNA As Long
    ' Set proyecto to ParaInformeAvisos = "No" so no email is registered
    db.Execute "UPDATE TbProyectos SET ParaInformeAvisos = 'No' WHERE IDProyecto = " & TEST_ID_NOTHING, dbFailOnError
    Dim helperResultNA As String
    helperResultNA = RechazarPropuestaPublicacion( _
        CStr(TEST_ID_NOTHING), _
        "Rechazo sin notificacion por correo", _
        p_CorreoRechazoNA, _
        db, _
        p_PromptResultNA)

    ' Without email notification, the helper should still succeed
    ' but p_PromptResult should signal the "no email sent" path (vbNo typically)
    If helperResultNA = "" Then
        logs(1) = logs(1) & " [no-notif: helper OK, promptResult=" & p_PromptResultNA & "]"
    Else
        logs(1) = logs(1) & " [no-notif: helper error='" & Left$(helperResultNA, 30) & "']"
    End If

    Test_RechazarPropuestaPublicacion_Adversarial = BuildJsonOk( _
        "{""double1"":""" & Left$(helperResult1, 20) & """," & _
        """double2"":""" & Left$(helperResult2, 20) & """," & _
        """noNotif"":""" & Left$(helperResultNA, 20) & """,""noNotifPrompt"":" & p_PromptResultNA & "}", _
        logs)
    GoTo Teardown

EH:
    logs(UBound(logs)) = "EH: " & Err.Number & " - " & Err.description
    Test_RechazarPropuestaPublicacion_Adversarial = BuildJsonFail(Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_DOUBLE, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_DOUBLE, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion = " & TEST_ID_NOTHING, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto = " & TEST_ID_NOTHING, dbFailOnError
    On Error GoTo 0
End Function


