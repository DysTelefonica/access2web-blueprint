Attribute VB_Name = "Test_GuardadoAutomaticoPestanas"
Option Compare Database
Option Explicit

Private Const TEST_AUTOSAVE_ID_BASE As Long = 994100000

Private Function JsonOk(ByVal p_Value As String, ByRef p_Logs() As String) As String
    JsonOk = BuildJsonOk(p_Value, p_Logs)
End Function

Private Function JsonFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    JsonFail = BuildJsonFail(p_Error, p_Logs)
End Function

Public Function Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim m_Error As String
    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_Count As Long
    Dim m_IDExp As Long

    m_IDExp = TEST_AUTOSAVE_ID_BASE + 1
    logs(0) = "Arrange: force sandbox and seed deterministic TbExpedientes fixture twice"

    If Not EnsureSandboxBackend(m_Error) Then
        Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonFail("Sandbox guard: " & m_Error, logs)
        Exit Function
    End If

    If Not SeedAutosaveExpedienteFixture(m_IDExp, m_Error) Then
        Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonFail(m_Error, logs)
        Exit Function
    End If
    If Not SeedAutosaveExpedienteFixture(m_IDExp, m_Error) Then
        Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonFail(m_Error, logs)
        Exit Function
    End If

    logs(1) = "Act: read only the deterministic expediente row by ID"
    Set m_Db = getdb(m_Error)
    If m_Db Is Nothing Then
        Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonFail("No se pudo abrir sandbox: " & m_Error, logs)
        GoTo Cleanup
    End If

    Set m_Rs = m_Db.OpenRecordset( _
        "SELECT COUNT(*) AS RowCount" & _
        " FROM TbExpedientes" & _
        " WHERE IDExpediente=" & m_IDExp & _
        " AND CodExp='TEST-AUTOSAVE-GF'" & _
        " AND Nemotecnico='TEST-AUTOSAVE-GF';")
    m_Count = CLng(Nz(m_Rs!rowCount, 0))

    logs(2) = "Assert: fixture cardinality is exactly one"
    If m_Count <> 1 Then
        Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonFail("Fixture General/Fechas no idempotente: count=" & CStr(m_Count), logs)
        GoTo Cleanup
    End If

    logs(3) = "Teardown: delete only deterministic autosave fixture rows"
    Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonOk("autosave_general_fechas_seed_idempotent", logs)

Cleanup:
    On Error Resume Next
    If Not m_Rs Is Nothing Then m_Rs.Close
    Set m_Rs = Nothing
    Call TeardownAutosaveExpedienteFixture(m_IDExp, m_Error)
    logs(4) = "Cleanup attempted for IDExpediente=" & CStr(m_IDExp)
    Exit Function

HandleError:
    Test_AutosaveFixture_SeedGeneralFechas_IsIdempotent = JsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse() As String
    Dim logs(0 To 5) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim m_Error As String
    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_IDExp As Long
    Dim m_IDHito As Long
    Dim m_IDModificado As Long
    Dim m_HitosCount As Long
    Dim m_ModificadosCount As Long

    m_IDExp = TEST_AUTOSAVE_ID_BASE + 2
    m_IDHito = TEST_AUTOSAVE_ID_BASE + 102
    m_IDModificado = TEST_AUTOSAVE_ID_BASE + 202
    logs(0) = "Arrange: force sandbox, seed parent expediente, then child hito/modificado"

    If Not EnsureSandboxBackend(m_Error) Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail("Sandbox guard: " & m_Error, logs)
        Exit Function
    End If

    If Not SeedAutosaveExpedienteFixture(m_IDExp, m_Error) Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail(m_Error, logs)
        Exit Function
    End If
    If Not SeedAutosaveHitoFixture(m_IDExp, m_IDHito, m_Error) Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail(m_Error, logs)
        GoTo Cleanup
    End If
    If Not SeedAutosaveModificadoFixture(m_IDExp, m_IDModificado, m_Error) Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail(m_Error, logs)
        GoTo Cleanup
    End If

    logs(1) = "Act: read child rows by deterministic parent/child IDs"
    Set m_Db = getdb(m_Error)
    If m_Db Is Nothing Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail("No se pudo abrir sandbox: " & m_Error, logs)
        GoTo Cleanup
    End If

    Set m_Rs = m_Db.OpenRecordset("SELECT COUNT(*) AS RowCount FROM TbExpedientesHitos WHERE IDExpediente=" & m_IDExp & " AND IDHitoExpediente=" & m_IDHito & ";")
    m_HitosCount = CLng(Nz(m_Rs!rowCount, 0))
    m_Rs.Close
    Set m_Rs = m_Db.OpenRecordset("SELECT COUNT(*) AS RowCount FROM TbExpedientesModificados WHERE IDExpediente=" & m_IDExp & " AND IDExpedienteModificado=" & m_IDModificado & ";")
    m_ModificadosCount = CLng(Nz(m_Rs!rowCount, 0))

    logs(2) = "Assert: each child fixture has cardinality exactly one"
    If m_HitosCount <> 1 Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail("Fixture Hitos no determinista: count=" & CStr(m_HitosCount), logs)
        GoTo Cleanup
    End If
    If m_ModificadosCount <> 1 Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail("Fixture Modificados no determinista: count=" & CStr(m_ModificadosCount), logs)
        GoTo Cleanup
    End If

    logs(3) = "Teardown: delete children before parent using deterministic IDs"
    If Not TeardownAutosaveExpedienteFixture(m_IDExp, m_Error) Then
        Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail(m_Error, logs)
        Exit Function
    End If
    logs(4) = "Assert: reverse-order teardown completed without relying on existing data"
    Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonOk("autosave_children_fixture_order_ok", logs)
    Exit Function

Cleanup:
    On Error Resume Next
    If Not m_Rs Is Nothing Then m_Rs.Close
    Set m_Rs = Nothing
    Call TeardownAutosaveExpedienteFixture(m_IDExp, m_Error)
    logs(5) = "Cleanup attempted for IDExpediente=" & CStr(m_IDExp)
    Exit Function

HandleError:
    Test_AutosaveFixture_SeedChildren_UsesParentFirstAndTeardownReverse = JsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_AutosaveRule_GeneralPersistsWithoutOpeningForm() As String
    Dim logs(0 To 6) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim m_Error As String
    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_IDExp As Long
    Dim m_NewTitulo As String
    Dim m_NewAmbito As String

    m_IDExp = TEST_AUTOSAVE_ID_BASE + 11
    m_NewTitulo = "Fixture autosave General changed through extracted rule"
    m_NewAmbito = "HPS"
    logs(0) = "Arrange: force sandbox and seed deterministic expediente for General rule"

    If Not EnsureSandboxBackend(m_Error) Then
        Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail("Sandbox guard: " & m_Error, logs)
        Exit Function
    End If

    If Not SeedAutosaveExpedienteFixture(m_IDExp, m_Error) Then
        Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail(m_Error, logs)
        Exit Function
    End If

    logs(1) = "Act: call extracted General persistence rule without opening or navigating forms"
    Call GuardarPestanaGeneralValores(CStr(m_IDExp), m_NewTitulo, m_NewAmbito, m_Error)
    If m_Error <> "" Then
        Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail(m_Error, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: read deterministic expediente and verify General values persisted"
    Set m_Db = getdb(m_Error)
    If m_Db Is Nothing Then
        Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail("No se pudo abrir sandbox: " & m_Error, logs)
        GoTo Cleanup
    End If

    Set m_Rs = m_Db.OpenRecordset( _
        "SELECT Titulo, Ambito FROM TbExpedientes WHERE IDExpediente=" & m_IDExp & ";", _
        dbOpenSnapshot)
    If m_Rs.EOF Then
        Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail("No existe expediente fixture tras guardar", logs)
        GoTo Cleanup
    End If

    If Nz(m_Rs!Titulo, "") <> m_NewTitulo Then
        Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail("Titulo no persistido. Actual='" & Nz(m_Rs!Titulo, "") & "'", logs)
        GoTo Cleanup
    End If
    If Nz(m_Rs!Ambito, "") <> m_NewAmbito Then
        Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail("Ambito no persistido. Actual='" & Nz(m_Rs!Ambito, "") & "'", logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: extracted General rule persisted concrete seeded values"
    Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonOk("autosave_general_rule_persisted_without_form", logs)

Cleanup:
    On Error Resume Next
    If Not m_Rs Is Nothing Then m_Rs.Close
    Set m_Rs = Nothing
    Call TeardownAutosaveExpedienteFixture(m_IDExp, m_Error)
    logs(6) = "Cleanup attempted for IDExpediente=" & CStr(m_IDExp)
    Exit Function

HandleError:
    Test_AutosaveRule_GeneralPersistsWithoutOpeningForm = JsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_AutosaveRule_FechasPersistsWithoutOpeningForm() As String
    Dim logs(0 To 6) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim m_Error As String
    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_IDExp As Long
    Dim m_NewInicio As Date
    Dim m_NewAdjudicacion As Date

    m_IDExp = TEST_AUTOSAVE_ID_BASE + 12
    m_NewInicio = DateSerial(2026, 7, 15)
    m_NewAdjudicacion = DateSerial(2026, 8, 20)
    logs(0) = "Arrange: force sandbox and seed deterministic expediente for Fechas rule"

    If Not EnsureSandboxBackend(m_Error) Then
        Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail("Sandbox guard: " & m_Error, logs)
        Exit Function
    End If

    If Not SeedAutosaveExpedienteFixture(m_IDExp, m_Error) Then
        Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail(m_Error, logs)
        Exit Function
    End If

    logs(1) = "Act: call extracted Fechas persistence rule without opening or navigating forms"
    Call GuardarPestanaFechasValores(CStr(m_IDExp), m_NewInicio, m_NewAdjudicacion, m_Error)
    If m_Error <> "" Then
        Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail(m_Error, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: read deterministic expediente and verify Fechas values persisted"
    Set m_Db = getdb(m_Error)
    If m_Db Is Nothing Then
        Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail("No se pudo abrir sandbox: " & m_Error, logs)
        GoTo Cleanup
    End If

    Set m_Rs = m_Db.OpenRecordset( _
        "SELECT FECHAINICIOLICITACION, FECHAADJUDICACION FROM TbExpedientes WHERE IDExpediente=" & m_IDExp & ";", _
        dbOpenSnapshot)
    If m_Rs.EOF Then
        Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail("No existe expediente fixture tras guardar", logs)
        GoTo Cleanup
    End If

    If Not IsDate(m_Rs!FECHAINICIOLICITACION) Or CDate(m_Rs!FECHAINICIOLICITACION) <> m_NewInicio Then
        Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail("FECHAINICIOLICITACION no persistida", logs)
        GoTo Cleanup
    End If
    If Not IsDate(m_Rs!FECHAADJUDICACION) Or CDate(m_Rs!FECHAADJUDICACION) <> m_NewAdjudicacion Then
        Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail("FECHAADJUDICACION no persistida", logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: extracted Fechas rule persisted concrete seeded values"
    Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonOk("autosave_fechas_rule_persisted_without_form", logs)

Cleanup:
    On Error Resume Next
    If Not m_Rs Is Nothing Then m_Rs.Close
    Set m_Rs = Nothing
    Call TeardownAutosaveExpedienteFixture(m_IDExp, m_Error)
    logs(6) = "Cleanup attempted for IDExpediente=" & CStr(m_IDExp)
    Exit Function

HandleError:
    Test_AutosaveRule_FechasPersistsWithoutOpeningForm = JsonFail(Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_AutosaveDecision_GeneralFechasDirtyEditMode_RequestSave() As String
    Dim logs(0 To 3) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveDecision_GeneralFechasDirtyEditMode_RequestSave = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    logs(0) = "Arrange: use scalar tab names and edit flags without opening forms"

    logs(1) = "Act: evaluate General and Fechas dirty edit-mode decisions"
    If Not DebeGuardarPestanaGeneralFechas("tabGeneral", True, True, True) Then
        Test_AutosaveDecision_GeneralFechasDirtyEditMode_RequestSave = JsonFail("tabGeneral dirty edit-mode should request save", logs)
        Exit Function
    End If
    If Not DebeGuardarPestanaGeneralFechas("tabFechas", True, True, True) Then
        Test_AutosaveDecision_GeneralFechasDirtyEditMode_RequestSave = JsonFail("tabFechas dirty edit-mode should request save", logs)
        Exit Function
    End If

    logs(2) = "Assert: only General/Fechas dirty edit-mode requests save"
    Test_AutosaveDecision_GeneralFechasDirtyEditMode_RequestSave = JsonOk("autosave_decision_general_fechas_dirty_edit_mode", logs)
    Exit Function

HandleError:
    logs(3) = "Error: " & Err.Description
    Test_AutosaveDecision_GeneralFechasDirtyEditMode_RequestSave = JsonFail(Err.Description, logs)
End Function

Public Function Test_AutosaveDecision_OutOfScopeTabsNoOp() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    logs(0) = "Arrange: use dirty edit-mode flags for out-of-scope tabs"

    logs(1) = "Act: evaluate Hitos/Modificados and already-auto-saved tabs"
    If DebeGuardarPestanaGeneralFechas("tabHitos", True, True, True) Then
        Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonFail("tabHitos must remain no-op in this slice", logs)
        Exit Function
    End If
    If DebeGuardarPestanaGeneralFechas("tabModificados", True, True, True) Then
        Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonFail("tabModificados must remain no-op in this slice", logs)
        Exit Function
    End If
    If DebeGuardarPestanaGeneralFechas("tabSuministradores", True, True, True) Then
        Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonFail("tabSuministradores must remain no-op", logs)
        Exit Function
    End If
    If DebeGuardarPestanaGeneralFechas("tabAnexos", True, True, True) Then
        Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonFail("tabAnexos must remain no-op", logs)
        Exit Function
    End If
    If DebeGuardarPestanaGeneralFechas("tabEntidades", True, True, True) Then
        Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonFail("tabEntidades must remain no-op", logs)
        Exit Function
    End If

    logs(2) = "Assert: out-of-scope dirty tabs do not request save"
    Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonOk("autosave_decision_out_of_scope_noop", logs)
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_AutosaveDecision_OutOfScopeTabsNoOp = JsonFail(Err.Description, logs)
End Function

Public Function Test_AutosaveDecision_UnchangedAltaOrReadOnlyNoOp() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveDecision_UnchangedAltaOrReadOnlyNoOp = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    logs(0) = "Arrange: use General tab with scalar flags for no-op branches"

    logs(1) = "Act: evaluate unchanged, Alta-mode, and read-only branches"
    If DebeGuardarPestanaGeneralFechas("tabGeneral", True, True, False) Then
        Test_AutosaveDecision_UnchangedAltaOrReadOnlyNoOp = JsonFail("unchanged General tab must not request save", logs)
        Exit Function
    End If
    If DebeGuardarPestanaGeneralFechas("tabGeneral", False, True, True) Then
        Test_AutosaveDecision_UnchangedAltaOrReadOnlyNoOp = JsonFail("Alta/non-edit mode must not request edit autosave", logs)
        Exit Function
    End If
    If DebeGuardarPestanaGeneralFechas("tabGeneral", True, False, True) Then
        Test_AutosaveDecision_UnchangedAltaOrReadOnlyNoOp = JsonFail("read-only form must not request save", logs)
        Exit Function
    End If

    logs(2) = "Assert: unchanged, Alta-mode, and read-only branches are no-op"
    Test_AutosaveDecision_UnchangedAltaOrReadOnlyNoOp = JsonOk("autosave_decision_noop_branches", logs)
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_AutosaveDecision_UnchangedAltaOrReadOnlyNoOp = JsonFail(Err.Description, logs)
End Function

Public Function Test_AutosaveDecision_MapsSourceObjectsToTabs() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_AutosaveDecision_MapsSourceObjectsToTabs = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    logs(0) = "Arrange: use source-object/form names as scalar inputs"

    logs(1) = "Act: map source-object names to canonical tab names"
    If NombrePestanaDesdeSourceObject("FormExpedienteGeneral") <> "tabGeneral" Then
        Test_AutosaveDecision_MapsSourceObjectsToTabs = JsonFail("FormExpedienteGeneral should map to tabGeneral", logs)
        Exit Function
    End If
    If NombrePestanaDesdeSourceObject("FormExpedienteFechas") <> "tabFechas" Then
        Test_AutosaveDecision_MapsSourceObjectsToTabs = JsonFail("FormExpedienteFechas should map to tabFechas", logs)
        Exit Function
    End If
    If NombrePestanaDesdeSourceObject("FormExpedienteDocumentacion") <> "tabAnexos" Then
        Test_AutosaveDecision_MapsSourceObjectsToTabs = JsonFail("FormExpedienteDocumentacion should map to tabAnexos", logs)
        Exit Function
    End If
    If NombrePestanaDesdeSourceObject("tabGeneral") <> "tabGeneral" Then
        Test_AutosaveDecision_MapsSourceObjectsToTabs = JsonFail("canonical tab name should remain unchanged", logs)
        Exit Function
    End If

    logs(2) = "Assert: source-object names map deterministically without forms"
    Test_AutosaveDecision_MapsSourceObjectsToTabs = JsonOk("autosave_decision_sourceobject_mapping", logs)
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_AutosaveDecision_MapsSourceObjectsToTabs = JsonFail(Err.Description, logs)
End Function

