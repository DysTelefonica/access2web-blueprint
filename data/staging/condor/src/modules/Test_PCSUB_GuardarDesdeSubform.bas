Attribute VB_Name = "Test_PCSUB_GuardarDesdeSubform"
Option Compare Database
Option Explicit

' ==========================================================================
' Test_PCSUB_GuardarDesdeSubform — Slice 2 E2E methodology rollout
' ==========================================================================

Private Const TEST_ID_BASE As Long = 900300
Private Const TEST_ID_TOP As Long = 900309
Private Const ID_SOL_GENERALES As Long = 900301
Private Const ID_SOL_GENERALES_INVALID As Long = 900302

Private Function CountRowsWhere(ByVal p_Db As DAO.Database, ByVal p_TableName As String, ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM " & p_TableName & " WHERE " & p_Where, dbOpenSnapshot)
    CountRowsWhere = Nz(rs!n, 0)
    rs.Close
End Function

Private Function EstadoSolicitud(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As Long
    Dim rs As DAO.Recordset
    Set rs = p_Db.OpenRecordset("SELECT idEstadoInterno FROM tbSolicitudes WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then EstadoSolicitud = Nz(rs!idEstadoInterno, 0)
    rs.Close
End Function

Private Sub TeardownFixture(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbLogEstados WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbLogCambios WHERE idRegistro >= " & TEST_ID_BASE & " AND idRegistro <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbDatosPCSUB WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE & " AND IDExpediente <= " & TEST_ID_TOP, dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedSolicitudRegistro(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    p_Db.Execute "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_IdSolicitud & ")", dbFailOnError
    p_Db.Execute "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado) " & _
                 "VALUES (" & p_IdSolicitud & ", " & p_IdSolicitud & ", 'PC_SUB', 'PCSUB-FORM-" & p_IdSolicitud & "', " & estadoRegistro & ", Now(), 'TestPCSUBForm', 'PENDIENTE')", dbFailOnError
End Sub

Private Function BuildViewModelGenerales(ByVal p_IdSolicitud As Long) As DatosPCSUBViewModel
    Dim vm As New DatosPCSUBViewModel
    Dim sol As New Solicitud
    Dim datos As New DatosPCSUB

    sol.idSolicitud = p_IdSolicitud
    sol.idExpediente = p_IdSolicitud
    sol.tipoSolicitud = "PC_SUB"
    sol.codigoSolicitud = "PCSUB-FORM-" & p_IdSolicitud
    sol.idEstadoInterno = estadoRegistro
    sol.fechaCreacion = Now()
    sol.usuarioCreacion = "TestPCSUBForm"
    sol.revisionCalidadEstado = "PENDIENTE"

    datos.idSolicitud = p_IdSolicitud
    Set vm.Solicitud = sol
    Set vm.datos = datos
    vm.vm_refContratoInspeccionOficial = "REF-FORM-" & p_IdSolicitud
    vm.vm_refSubSuministrador = "SUB-FORM-" & p_IdSolicitud
    vm.vm_SubsuministradorNombreDir = "Subcontratista Slice 2"
    vm.vm_denominacionContrato = "Contrato Slice 2"
    vm.vm_objetoContrato = "Objeto Slice 2"
    Set BuildViewModelGenerales = vm
End Function

Private Function BuildViewModelGeneralesInvalid(ByVal p_IdSolicitud As Long) As DatosPCSUBViewModel
    Dim vm As DatosPCSUBViewModel

    Set vm = BuildViewModelGenerales(p_IdSolicitud)
    vm.vm_refContratoInspeccionOficial = vbNullString
    Set BuildViewModelGeneralesInvalid = vm
End Function

Public Function Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(8)
    Dim errMsg As String
    Dim tempRoot As String
    Dim db As DAO.Database
    Dim vm As DatosPCSUBViewModel
    Dim result As Object
    Dim countBefore As Long
    Dim countAfter As Long
    Dim logBefore As Long
    Dim logAfter As Long
    Dim rs As DAO.Recordset

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow = TestHelper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "pcsub_form_slice2", "PCSUB Form Test User")
    If errMsg <> "" Then Err.Raise 513, , errMsg
    rolUsuario = rol.Calidad
    rolUsuarioReal = rol.Calidad
    m_ObjUsuarioActivo.rol = rol.Calidad
    Set m_ObjUsuarioReal = New Usuario
    m_ObjUsuarioReal.CorreoUsuario = "pcsub.form.test@example.invalid"

    Set db = TestHelper.GetTestDb()
    Call TeardownFixture(db)
    Call SeedSolicitudRegistro(db, ID_SOL_GENERALES)
    Set vm = BuildViewModelGenerales(ID_SOL_GENERALES)
    logs(0) = "1. Fixture PCSUB en estadoRegistro preparado"

    countBefore = CountRowsWhere(db, "tbDatosPCSUB", "idSolicitud=" & ID_SOL_GENERALES)
    logBefore = CountRowsWhere(db, "tbLogEstados", "idSolicitud=" & ID_SOL_GENERALES)
    If countBefore <> 0 Then Err.Raise 513, , "Fixture esperado sin tbDatosPCSUB inicial"

    Set result = DatosPCSUBGuardarHelper_GuardarDesdeSubformPlan("subfrmDatosPCSUB_Generales", vm, db)
    logs(1) = "2. Helper ejecutado con DAO.Database inyectada; prompt diferido"

    If result Is Nothing Then Err.Raise 513, , "El helper no devolvió resultado"
    If Not CBool(result("ok")) Then Err.Raise 513, , "Helper devolvió ok=False: " & Nz(result("errorText"), "")
    If Not CBool(result("persisted")) Then Err.Raise 513, , "Helper no marcó persisted=True"
    If Not CBool(result("requiresUserPrompt")) Then Err.Raise 513, , "El helper no planificó prompt post-persistencia"
    If CLng(result("workflowTarget")) <> 0 Then Err.Raise 513, , "workflowTarget debe seguir vacío antes del prompt"

    countAfter = CountRowsWhere(db, "tbDatosPCSUB", "idSolicitud=" & ID_SOL_GENERALES)
    If countAfter <> 1 Then Err.Raise 513, , "Cardinalidad tbDatosPCSUB esperada 1, obtenida " & countAfter
    logs(2) = "3. Cardinalidad tbDatosPCSUB: " & countBefore & " -> " & countAfter

    Set rs = db.OpenRecordset("SELECT refContratoInspeccionOficial, refSubSuministrador, SubsuministradorNombreDir, denominacionContrato, objetoContrato FROM tbDatosPCSUB WHERE idSolicitud=" & ID_SOL_GENERALES, dbOpenSnapshot)
    If rs.EOF Then Err.Raise 513, , "No se encontró la fila PCSUB persistida"
    If Nz(rs!refContratoInspeccionOficial, "") <> vm.vm_refContratoInspeccionOficial Then Err.Raise 513, , "refContratoInspeccionOficial no persistido"
    If Nz(rs!refSubSuministrador, "") <> vm.vm_refSubSuministrador Then Err.Raise 513, , "refSubSuministrador no persistido"
    If Nz(rs!SubsuministradorNombreDir, "") <> vm.vm_SubsuministradorNombreDir Then Err.Raise 513, , "SubsuministradorNombreDir no persistido"
    If Nz(rs!denominacionContrato, "") <> vm.vm_denominacionContrato Then Err.Raise 513, , "denominacionContrato no persistido"
    If Nz(rs!objetoContrato, "") <> vm.vm_objetoContrato Then Err.Raise 513, , "objetoContrato no persistido"
    rs.Close: Set rs = Nothing
    logs(3) = "4. Los cinco campos generales se persistieron"

    If EstadoSolicitud(db, ID_SOL_GENERALES) <> estadoRegistro Then Err.Raise 513, , "La solicitud avanzó antes del prompt"
    Set result = DatosPCSUBGuardarHelper_AsignarTecnicoPostGuardado(vm, db)
    If result Is Nothing Then Err.Raise 513, , "El helper de asignación no devolvió resultado"
    If Not CBool(result("ok")) Then Err.Raise 513, , "Asignación devolvió ok=False: " & Nz(result("errorText"), "")
    If CLng(result("workflowTarget")) <> estadoDesarrolloTecnico Then Err.Raise 513, , "workflowTarget inesperado"
    If EstadoSolicitud(db, ID_SOL_GENERALES) <> estadoDesarrolloTecnico Then Err.Raise 513, , "La solicitud no avanzó a Desarrollo Técnico"
    logAfter = CountRowsWhere(db, "tbLogEstados", "idSolicitud=" & ID_SOL_GENERALES)
    If logAfter <= logBefore Then Err.Raise 513, , "No se registró log de workflow"
    logs(4) = "5. Workflow vía WorkflowServicio: estadoRegistro -> estadoDesarrolloTecnico"

    logs(5) = "6. PASS"
    Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then TeardownFixture db
    Call TestHelper.EndTestSession(logs, errMsg)
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_GuardarDesdeSubform_Generales_InvalidDoesNotPlanPromptOrWorkflow() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(8)
    Dim errMsg As String
    Dim tempRoot As String
    Dim db As DAO.Database
    Dim vm As DatosPCSUBViewModel
    Dim result As Object
    Dim countBefore As Long
    Dim countAfter As Long
    Dim logBefore As Long
    Dim logAfter As Long

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        Test_PCSUB_GuardarDesdeSubform_Generales_InvalidDoesNotPlanPromptOrWorkflow = TestHelper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "pcsub_form_slice2_invalid", "PCSUB Form Invalid Test User")
    If errMsg <> "" Then Err.Raise 513, , errMsg
    rolUsuario = rol.Calidad
    rolUsuarioReal = rol.Calidad
    m_ObjUsuarioActivo.rol = rol.Calidad
    Set m_ObjUsuarioReal = New Usuario
    m_ObjUsuarioReal.CorreoUsuario = "pcsub.form.invalid.test@example.invalid"

    Set db = TestHelper.GetTestDb()
    Call TeardownFixture(db)
    Call SeedSolicitudRegistro(db, ID_SOL_GENERALES_INVALID)
    Set vm = BuildViewModelGeneralesInvalid(ID_SOL_GENERALES_INVALID)
    logs(0) = "1. Fixture PCSUB inválido preparado sin referencia de contrato"

    countBefore = CountRowsWhere(db, "tbDatosPCSUB", "idSolicitud=" & ID_SOL_GENERALES_INVALID)
    logBefore = CountRowsWhere(db, "tbLogEstados", "idSolicitud=" & ID_SOL_GENERALES_INVALID)
    If countBefore <> 0 Then Err.Raise 513, , "Fixture esperado sin tbDatosPCSUB inicial"

    Set result = DatosPCSUBGuardarHelper_GuardarDesdeSubformPlan("subfrmDatosPCSUB_Generales", vm, db)
    logs(1) = "2. Helper ejecutado con datos inválidos antes de cualquier prompt"

    If result Is Nothing Then Err.Raise 513, , "El helper no devolvió resultado"
    If CBool(result("ok")) Then Err.Raise 513, , "Helper devolvió ok=True con datos inválidos"
    If CBool(result("persisted")) Then Err.Raise 513, , "Helper marcó persisted=True con datos inválidos"
    If CBool(result("requiresUserPrompt")) Then Err.Raise 513, , "Helper planificó prompt antes de validar/persistir"
    If CLng(result("workflowTarget")) <> 0 Then Err.Raise 513, , "Helper planificó workflow con datos inválidos"
    If Len(Trim$(Nz(result("errorText"), ""))) = 0 Then Err.Raise 513, , "Helper no devolvió errorText de validación"

    countAfter = CountRowsWhere(db, "tbDatosPCSUB", "idSolicitud=" & ID_SOL_GENERALES_INVALID)
    If countAfter <> countBefore Then Err.Raise 513, , "Cardinalidad tbDatosPCSUB cambió con datos inválidos"
    logs(2) = "3. Cardinalidad tbDatosPCSUB permanece " & countBefore

    If EstadoSolicitud(db, ID_SOL_GENERALES_INVALID) <> estadoRegistro Then Err.Raise 513, , "La solicitud cambió de estado con datos inválidos"
    logAfter = CountRowsWhere(db, "tbLogEstados", "idSolicitud=" & ID_SOL_GENERALES_INVALID)
    If logAfter <> logBefore Then Err.Raise 513, , "Se registró workflow con datos inválidos"
    logs(3) = "4. Sin prompt planificado, sin persistencia y sin workflow"

    logs(4) = "5. PASS"
    Test_PCSUB_GuardarDesdeSubform_Generales_InvalidDoesNotPlanPromptOrWorkflow = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixture db
    Call TestHelper.EndTestSession(logs, errMsg)
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_GuardarDesdeSubform_Generales_InvalidDoesNotPlanPromptOrWorkflow = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function
