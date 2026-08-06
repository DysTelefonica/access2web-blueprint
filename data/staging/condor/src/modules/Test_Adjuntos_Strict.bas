Attribute VB_Name = "Test_Adjuntos_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Adjuntos_Strict — CAP-008 strict TDD first slice for attachment service
'
' Touched tables and schema evidence (exploration + ERD/schema-first gate):
'   TbExpedientes: PK IDExpediente; required IDExpediente.
'   tbSolicitudes: PK idSolicitud; required idExpediente, tipoSolicitud,
'     codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion,
'     revisionCalidadEstado.
'   tbAdjuntos: PK idAdjunto; required idSolicitud, nombreArchivo,
'     fechaSubida, usuarioSubida; nombreArchivo 255, usuarioSubida 100,
'     etapaWF 50.
'
' Fixture graph: TbExpedientes -> tbSolicitudes -> tbAdjuntos.
' Teardown order: tbAdjuntos -> tbSolicitudes -> TbExpedientes, filtered to
' deterministic IDs >= 900800 and filenames prefixed with CAP008-STRICT-.
'
' Contract:
'   - Public Function Test_*() As String returning canonical JSON.
'   - No UI, Debug.Print, MsgBox, or TbConfiguracionBackends mutation.
'   - Explicit DAO.Database is passed to AdjuntosServicio methods.
'   - Filesystem side effects are isolated under TEMP and cleaned defensively.
'   - No VBA short-circuit assumptions: object Nothing checks are split.
' ============================================================================

Private Const TEST_ID_BASE As Long = 900800
Private Const TEST_ID_TOP As Long = 900899

Private Const ID_SOL_GUARDAR As Long = 900801
Private Const ID_SOL_ELIMINAR As Long = 900802
Private Const ID_SOL_ETAPAS As Long = 900803
Private Const ID_SOL_RUTA_INEXISTENTE As Long = 900804
Private Const ID_SOL_FINAL_NO_PDF As Long = 900805
Private Const ID_SOL_DEDUP As Long = 900806
Private Const ID_SOL_ACTUALIZAR_HAPPY As Long = 900807
Private Const ID_SOL_ACTUALIZAR_INVAL As Long = 900808
Private Const ID_SOL_ACTUALIZAR_NOEXISTE As Long = 900809
Private Const ID_SOL_RUTA_VACIA As Long = 900811
Private Const ID_SOL_ID_INVALIDO As Long = 900812
Private Const ID_SOL_SUBIR_Y_CERRAR As Long = 900820
Private Const ID_SOL_ROLLBACK As Long = 900850

Private Const ID_ADJ_ELIMINAR As Long = 900821
Private Const ID_ADJ_ETAPA_A As Long = 900831
Private Const ID_ADJ_ETAPA_B As Long = 900832
Private Const ID_ADJ_ETAPA_KEEP As Long = 900833
Private Const ID_ADJ_ACTUALIZAR_HAPPY As Long = 900840

Private Const FILE_PREFIX As String = "CAP008-STRICT-"

' ============================================================================
' Public atoms
' ============================================================================

Public Function Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(8)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim expectedDest As String
    Dim idAdjuntoOut As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim rs As DAO.Recordset

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_GUARDAR, "CAP008-GUARDAR")
    sourcePath = tempRoot & "source\" & FILE_PREFIX & "guardar.txt"
    Call WriteTextFile(sourcePath, "contenido guardar")
    logs(0) = "1. Seeded parent graph and source file under TEMP"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_GUARDAR)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    svc.GuardarAdjuntoDesdeArchivo ID_SOL_GUARDAR, "CAP008-Etapa-Guardar", sourcePath, "Adjuntado", db, idAdjuntoOut, "CAP008 descripcion"
    logs(2) = "3. GuardarAdjuntoDesdeArchivo called with explicit DAO.Database"

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_GUARDAR)
    If countAfter <> 1 Then
        logs(3) = "4. FAILED: expected one adjunto after act, got " & countAfter
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("wrong cardinality after guardar", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter

    If idAdjuntoOut <= 0 Then
        logs(4) = "5. FAILED: idAdjuntoOut was not assigned"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("idAdjuntoOut not assigned", logs)
        GoTo CleanExit
    End If

    expectedDest = GetDocDir() & FILE_PREFIX & "guardar.txt"
    If Not FileExistsSafe(expectedDest) Then
        logs(4) = "5. FAILED: destination file was not copied: " & expectedDest
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("destination file missing", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. File copied to isolated documentation folder"

    Set rs = db.OpenRecordset("SELECT idAdjunto, etapaWF, nombreArchivo, usuarioSubida, descripcion, TipoAccion FROM tbAdjuntos WHERE idSolicitud=" & ID_SOL_GUARDAR, dbOpenSnapshot)
    If rs.EOF Then
        logs(5) = "6. FAILED: row missing after count assertion"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("row missing", logs)
        GoTo CleanExit
    End If

    If CLng(rs!idAdjunto) <> idAdjuntoOut Then
        logs(5) = "6. FAILED: row idAdjunto differs from idAdjuntoOut"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("idAdjuntoOut mismatch", logs)
        GoTo CleanExit
    End If
    If Nz(rs!etapaWF, "") <> "CAP008-Etapa-Guardar" Or Nz(rs!nombreArchivo, "") <> FILE_PREFIX & "guardar.txt" Then
        logs(5) = "6. FAILED: persisted etapa/nombreArchivo mismatch"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("persisted fields mismatch", logs)
        GoTo CleanExit
    End If
    If Nz(rs!usuarioSubida, "") <> "CAP008 Test User" Or Nz(rs!descripcion, "") <> "CAP008 descripcion" Or Nz(rs!tipoAccion, "") <> "Adjuntado" Then
        logs(5) = "6. FAILED: persisted metadata mismatch"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail("persisted metadata mismatch", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. Row fields match source, etapa, user and metadata"

    logs(6) = "7. PASS"
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim storedFile As String
    Dim countBefore As Long
    Dim countAfter As Long

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_ELIMINAR, "CAP008-ELIMINAR")
    storedFile = FILE_PREFIX & "eliminar.txt"
    Call WriteTextFile(GetDocDir() & storedFile, "contenido eliminar")
    Call SeedAdjunto(db, ID_ADJ_ELIMINAR, ID_SOL_ELIMINAR, "CAP008-Etapa-Eliminar", storedFile)
    logs(0) = "1. Seeded parent graph, adjunto row and physical file"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_ELIMINAR)
    If countBefore <> 1 Then
        logs(1) = "2. FAILED: expected one adjunto before delete, got " & countBefore
        Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo = TestHelper.BuildJsonFail("wrong fixture cardinality", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    svc.EliminarAdjunto ID_ADJ_ELIMINAR, db
    logs(2) = "3. EliminarAdjunto called with explicit DAO.Database"

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_ELIMINAR)
    If countAfter <> 0 Then
        logs(3) = "4. FAILED: expected zero adjuntos after delete, got " & countAfter
        Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo = TestHelper.BuildJsonFail("row was not deleted", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter

    If FileExistsSafe(GetDocDir() & storedFile) Then
        logs(4) = "5. FAILED: physical file still exists after delete"
        Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo = TestHelper.BuildJsonFail("file was not deleted", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. Physical file deleted"

    logs(5) = "6. PASS"
    Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(9)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim etapas As Variant
    Dim countBefore As Long
    Dim countAfter As Long

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_ETAPAS, "CAP008-ETAPAS")
    Call SeedAdjuntoConArchivo(db, ID_ADJ_ETAPA_A, ID_SOL_ETAPAS, "CAP008-Etapa-A", FILE_PREFIX & "etapa-a.txt")
    Call SeedAdjuntoConArchivo(db, ID_ADJ_ETAPA_B, ID_SOL_ETAPAS, "CAP008-Etapa-B", FILE_PREFIX & "etapa-b.txt")
    Call SeedAdjuntoConArchivo(db, ID_ADJ_ETAPA_KEEP, ID_SOL_ETAPAS, "CAP008-Etapa-Keep", FILE_PREFIX & "etapa-keep.txt")
    logs(0) = "1. Seeded three adjuntos in two target stages and one keep stage"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_ETAPAS)
    If countBefore <> 3 Then
        logs(1) = "2. FAILED: expected three adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail("wrong fixture cardinality", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    etapas = Array("CAP008-Etapa-A", "CAP008-Etapa-B")
    svc.EliminarAdjuntosPorEtapas ID_SOL_ETAPAS, etapas, db
    logs(2) = "3. EliminarAdjuntosPorEtapas called with explicit DAO.Database"

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_ETAPAS)
    If countAfter <> 1 Then
        logs(3) = "4. FAILED: expected one remaining adjunto, got " & countAfter
        Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail("wrong cardinality after etapewise delete", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter

    If CountAdjuntosByEtapa(db, ID_SOL_ETAPAS, "CAP008-Etapa-A") <> 0 Or CountAdjuntosByEtapa(db, ID_SOL_ETAPAS, "CAP008-Etapa-B") <> 0 Then
        logs(4) = "5. FAILED: target stage row still exists"
        Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail("target stage row still exists", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. Target stage rows deleted"

    If CountAdjuntosByEtapa(db, ID_SOL_ETAPAS, "CAP008-Etapa-Keep") <> 1 Then
        logs(5) = "6. FAILED: keep stage row was deleted"
        Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail("keep stage row missing", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. Non-target stage row preserved"

    If FileExistsSafe(GetDocDir() & FILE_PREFIX & "etapa-a.txt") Or FileExistsSafe(GetDocDir() & FILE_PREFIX & "etapa-b.txt") Then
        logs(6) = "7. FAILED: at least one target stage file still exists"
        Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail("target file still exists", logs)
        GoTo CleanExit
    End If
    If Not FileExistsSafe(GetDocDir() & FILE_PREFIX & "etapa-keep.txt") Then
        logs(6) = "7. FAILED: keep stage file was deleted"
        Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail("keep file missing", logs)
        GoTo CleanExit
    End If
    logs(6) = "7. Target files deleted and keep file preserved"

    logs(7) = "8. PASS"
    Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim missingPath As String
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long
    Dim idAdjuntoOut As Long

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_RUTA_INEXISTENTE, "CAP008-MISSING")
    missingPath = tempRoot & "source\" & FILE_PREFIX & "missing.txt"
    logs(0) = "1. Seeded parent graph; source path intentionally missing"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_RUTA_INEXISTENTE)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    On Error Resume Next
    svc.GuardarAdjuntoDesdeArchivo ID_SOL_RUTA_INEXISTENTE, "CAP008-Etapa-Missing", missingPath, "Adjuntado", db, idAdjuntoOut, "missing"
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber = 0 Then
        logs(2) = "3. FAILED: expected error for nonexistent source path"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente = TestHelper.BuildJsonFail("expected missing source error", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. GuardarAdjuntoDesdeArchivo rejected nonexistent path with Err " & errNumber

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_RUTA_INEXISTENTE)
    If countAfter <> countBefore Then
        logs(3) = "4. FAILED: cardinality changed after rejected file path: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente = TestHelper.BuildJsonFail("rejected path changed cardinality", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (no side effect)"

    If idAdjuntoOut <> 0 Then
        logs(4) = "5. FAILED: idAdjuntoOut changed on rejected path"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente = TestHelper.BuildJsonFail("idAdjuntoOut changed", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. idAdjuntoOut unchanged"

    logs(5) = "6. PASS"
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(8)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim expectedDest As String
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long
    Dim errDescription As String
    Dim idAdjuntoOut As Long

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_FINAL_NO_PDF, "CAP008-FINAL-NOPDF")
    sourcePath = tempRoot & "source\" & FILE_PREFIX & "final-no-pdf.txt"
    expectedDest = GetDocDir() & FILE_PREFIX & "final-no-pdf.txt"
    Call WriteTextFile(sourcePath, "contenido no pdf")
    logs(0) = "1. Seeded parent graph and non-PDF source file for Documento Final Firmado"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_FINAL_NO_PDF)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    On Error Resume Next
    svc.GuardarAdjuntoDesdeArchivo ID_SOL_FINAL_NO_PDF, "Documento Final Firmado", sourcePath, "Adjuntado", db, idAdjuntoOut, "non-pdf final"
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo EH

    If errNumber = 0 Then
        logs(2) = "3. FAILED: expected error for non-PDF Documento Final Firmado"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("expected non-PDF final signed document error", logs)
        GoTo CleanExit
    End If
    If InStr(1, errDescription, "PDF", vbTextCompare) = 0 Then
        logs(2) = "3. FAILED: rejection error did not mention PDF: " & errDescription
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("wrong rejection error", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. GuardarAdjuntoDesdeArchivo rejected non-PDF final signed document with Err " & errNumber

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_FINAL_NO_PDF)
    If countAfter <> countBefore Then
        logs(3) = "4. FAILED: cardinality changed after rejected non-PDF: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("rejected non-PDF changed cardinality", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (no side effect)"

    If idAdjuntoOut <> 0 Then
        logs(4) = "5. FAILED: idAdjuntoOut changed on rejected non-PDF"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("idAdjuntoOut changed", logs)
        GoTo CleanExit
    End If
    If FileExistsSafe(expectedDest) Then
        logs(4) = "5. FAILED: rejected non-PDF was copied to documentation folder"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("rejected non-PDF copied", logs)
        GoTo CleanExit
    End If
    If Not FileExistsSafe(sourcePath) Then
        logs(4) = "5. FAILED: source file disappeared after rejected non-PDF"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail("source file missing after rejection", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. No row, no destination file, and source remained intact"

    logs(5) = "6. PASS"
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(10)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim originalDest As String
    Dim dedupDest As String
    Dim idAdjuntoOut As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim persistedName As String
    Dim rs As DAO.Recordset

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_DEDUP, "CAP008-DEDUP")
    sourcePath = tempRoot & "source\" & FILE_PREFIX & "dedup.pdf"
    originalDest = GetDocDir() & FILE_PREFIX & "dedup.pdf"
    Call WriteTextFile(sourcePath, "nuevo pdf")
    Call WriteTextFile(originalDest, "destino preexistente")
    logs(0) = "1. Seeded parent graph, source PDF and pre-existing destination filename"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_DEDUP)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    svc.GuardarAdjuntoDesdeArchivo ID_SOL_DEDUP, "Documento Final Firmado", sourcePath, "Adjuntado", db, idAdjuntoOut, "dedup final"
    logs(2) = "3. GuardarAdjuntoDesdeArchivo called with explicit DAO.Database and colliding destination"

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_DEDUP)
    If countAfter <> 1 Then
        logs(3) = "4. FAILED: expected one adjunto after act, got " & countAfter
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("wrong cardinality after dedup guardar", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter

    Set rs = db.OpenRecordset("SELECT idAdjunto, etapaWF, nombreArchivo, descripcion FROM tbAdjuntos WHERE idSolicitud=" & ID_SOL_DEDUP, dbOpenSnapshot)
    If rs.EOF Then
        logs(4) = "5. FAILED: row missing after count assertion"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("row missing", logs)
        GoTo CleanExit
    End If
    If CLng(rs!idAdjunto) <> idAdjuntoOut Then
        logs(4) = "5. FAILED: row idAdjunto differs from idAdjuntoOut"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("idAdjuntoOut mismatch", logs)
        GoTo CleanExit
    End If
    persistedName = Nz(rs!nombreArchivo, "")
    If persistedName = FILE_PREFIX & "dedup.pdf" Then
        logs(4) = "5. FAILED: persisted name did not deduplicate"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("filename not deduplicated", logs)
        GoTo CleanExit
    End If
    If Not persistedName Like FILE_PREFIX & "dedup_########_######.pdf" Then
        logs(4) = "5. FAILED: deduplicated name has unexpected format: " & persistedName
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("wrong deduplicated filename format", logs)
        GoTo CleanExit
    End If
    If Nz(rs!etapaWF, "") <> "Documento Final Firmado" Or Nz(rs!descripcion, "") <> "dedup final" Then
        logs(4) = "5. FAILED: persisted stage/description mismatch"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("persisted metadata mismatch", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. Persisted deduplicated filename: " & persistedName

    dedupDest = GetDocDir() & persistedName
    If Not FileExistsSafe(originalDest) Then
        logs(5) = "6. FAILED: original destination file disappeared"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("original destination missing", logs)
        GoTo CleanExit
    End If
    If ReadTextFile(originalDest) <> "destino preexistente" Then
        logs(5) = "6. FAILED: original destination content was overwritten"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("original destination overwritten", logs)
        GoTo CleanExit
    End If
    If Not FileExistsSafe(dedupDest) Then
        logs(5) = "6. FAILED: deduplicated destination file missing: " & dedupDest
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("deduplicated destination missing", logs)
        GoTo CleanExit
    End If
    If ReadTextFile(dedupDest) <> "nuevo pdf" Then
        logs(5) = "6. FAILED: deduplicated destination content mismatch"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("deduplicated destination content mismatch", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. Original destination preserved and deduplicated copy created"

    If Not FileExistsSafe(sourcePath) Then
        logs(6) = "7. FAILED: source file disappeared after deduplicated copy"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail("source file missing after copy", logs)
        GoTo CleanExit
    End If
    logs(6) = "7. Source file remained intact"

    logs(7) = "8. PASS"
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(10)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourceNewPath As String
    Dim storedFile As String
    Dim storedFull As String
    Dim persistedBefore As String
    Dim persistedDescBefore As String
    Dim persistedTipoAccionBefore As String
    Dim persistedFechaSubidaBefore As String
    Dim persistedEtapaBefore As String
    Dim persistedUsuarioBefore As String
    Dim countBefore As Long
    Dim countAfter As Long
    Dim rs As DAO.Recordset

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_ACTUALIZAR_HAPPY, "CAP008-ACTUALIZAR")
    storedFile = FILE_PREFIX & "actualizar.txt"
    storedFull = GetDocDir() & storedFile
    Call SeedAdjuntoConArchivo(db, ID_ADJ_ACTUALIZAR_HAPPY, ID_SOL_ACTUALIZAR_HAPPY, "CAP008-Etapa-Actualizar", storedFile)
    sourceNewPath = tempRoot & "source\" & FILE_PREFIX & "actualizar-nuevo.txt"
    Call WriteTextFile(sourceNewPath, "contenido nuevo")
    logs(0) = "1. Seeded parent graph, existing adjunto row and pre-existing physical file"

    ' Snapshot persisted fields BEFORE the act so we can assert they are untouched.
    Set rs = db.OpenRecordset("SELECT nombreArchivo, fechaSubida, usuarioSubida, descripcion, TipoAccion, etapaWF FROM tbAdjuntos WHERE idAdjunto=" & ID_ADJ_ACTUALIZAR_HAPPY, dbOpenSnapshot)
    If rs.EOF Then
        logs(1) = "2. FAILED: seeded row missing before act"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("seeded row missing", logs)
        GoTo CleanExit
    End If
    persistedBefore = Nz(rs!nombreArchivo, "")
    persistedFechaSubidaBefore = CStr(rs!fechaSubida)
    persistedUsuarioBefore = Nz(rs!usuarioSubida, "")
    persistedDescBefore = Nz(rs!descripcion, "")
    persistedTipoAccionBefore = Nz(rs!TipoAccion, "")
    persistedEtapaBefore = Nz(rs!etapaWF, "")
    rs.Close
    Set rs = Nothing
    logs(1) = "2. Snapshot taken: nombreArchivo=" & persistedBefore & " etapaWF=" & persistedEtapaBefore

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_ACTUALIZAR_HAPPY)
    If countBefore <> 1 Then
        logs(2) = "3. FAILED: expected one adjunto before act, got " & countBefore
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("wrong fixture cardinality", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countBefore=" & countBefore

    svc.ActualizarFicheroAdjunto ID_ADJ_ACTUALIZAR_HAPPY, sourceNewPath, db
    logs(3) = "4. ActualizarFicheroAdjunto called with explicit DAO.Database and new source file"

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_ACTUALIZAR_HAPPY)
    If countAfter <> countBefore Then
        logs(4) = "5. FAILED: cardinality changed after update: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("cardinality changed on update", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. countAfter=" & countAfter & " (row preserved)"

    If Not FileExistsSafe(storedFull) Then
        logs(5) = "6. FAILED: destination file disappeared after update: " & storedFull
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("destination file missing", logs)
        GoTo CleanExit
    End If
    If ReadTextFile(storedFull) <> "contenido nuevo" Then
        logs(5) = "6. FAILED: destination file was not overwritten with new content: " & ReadTextFile(storedFull)
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("destination not overwritten", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. Destination file overwritten with new content at " & storedFull

    Set rs = db.OpenRecordset("SELECT nombreArchivo, fechaSubida, usuarioSubida, descripcion, TipoAccion, etapaWF, idSolicitud FROM tbAdjuntos WHERE idAdjunto=" & ID_ADJ_ACTUALIZAR_HAPPY, dbOpenSnapshot)
    If rs.EOF Then
        logs(6) = "7. FAILED: row missing after update"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("row missing after update", logs)
        GoTo CleanExit
    End If
    If CLng(rs!idSolicitud) <> ID_SOL_ACTUALIZAR_HAPPY Then
        logs(6) = "7. FAILED: idSolicitud changed from " & ID_SOL_ACTUALIZAR_HAPPY & " to " & rs!idSolicitud
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("idSolicitud changed", logs)
        GoTo CleanExit
    End If
    If Nz(rs!nombreArchivo, "") <> persistedBefore Then
        logs(6) = "7. FAILED: nombreArchivo changed from '" & persistedBefore & "' to '" & Nz(rs!nombreArchivo, "") & "'"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("nombreArchivo changed", logs)
        GoTo CleanExit
    End If
    If Nz(rs!etapaWF, "") <> persistedEtapaBefore Then
        logs(6) = "7. FAILED: etapaWF changed"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("etapaWF changed", logs)
        GoTo CleanExit
    End If
    If Nz(rs!usuarioSubida, "") <> persistedUsuarioBefore Then
        logs(6) = "7. FAILED: usuarioSubida changed"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("usuarioSubida changed", logs)
        GoTo CleanExit
    End If
    If Nz(rs!descripcion, "") <> persistedDescBefore Then
        logs(6) = "7. FAILED: descripcion changed"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("descripcion changed", logs)
        GoTo CleanExit
    End If
    If Nz(rs!TipoAccion, "") <> persistedTipoAccionBefore Then
        logs(6) = "7. FAILED: TipoAccion changed"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("TipoAccion changed", logs)
        GoTo CleanExit
    End If
    If CStr(rs!fechaSubida) <> persistedFechaSubidaBefore Then
        logs(6) = "7. FAILED: fechaSubida changed"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("fechaSubida changed", logs)
        GoTo CleanExit
    End If
    rs.Close
    Set rs = Nothing
    logs(6) = "7. Row preserved: idAdjunto, idSolicitud, etapaWF, fechaSubida, usuarioSubida, descripcion, TipoAccion, nombreArchivo all unchanged"

    If Not FileExistsSafe(sourceNewPath) Then
        logs(7) = "8. FAILED: source file disappeared after copy"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail("source file missing after copy", logs)
        GoTo CleanExit
    End If
    logs(7) = "8. Source file remained intact at " & sourceNewPath

    logs(8) = "9. PASS"
    Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long
    Dim errDescription As String

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_ACTUALIZAR_INVAL, "CAP008-ACT-INVALID")
    sourcePath = tempRoot & "source\" & FILE_PREFIX & "actualizar-inval.txt"
    Call WriteTextFile(sourcePath, "contenido inval")
    logs(0) = "1. Seeded parent graph and source file; idAdjunto intentionally 0"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_ACTUALIZAR_INVAL)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    On Error Resume Next
    svc.ActualizarFicheroAdjunto 0, sourcePath, db
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo EH

    If errNumber = 0 Then
        logs(2) = "3. FAILED: expected error for idAdjunto=0"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail("expected invalid idAdjunto error", logs)
        GoTo CleanExit
    End If
    If InStr(1, errDescription, "Invalid attachment ID", vbTextCompare) = 0 Then
        logs(2) = "3. FAILED: rejection error did not mention invalid id: " & errDescription
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail("wrong rejection error", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. ActualizarFicheroAdjunto rejected idAdjunto=0 with Err " & errNumber

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_ACTUALIZAR_INVAL)
    If countAfter <> countBefore Then
        logs(3) = "4. FAILED: cardinality changed after rejected invalid id: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail("rejected invalid id changed cardinality", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (no side effect)"

    If FileExistsSafe(GetDocDir() & FILE_PREFIX & "actualizar-inval.txt") Then
        logs(4) = "5. FAILED: rejected update copied file to documentation folder"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail("file copied on invalid id", logs)
        GoTo CleanExit
    End If
    If Not FileExistsSafe(sourcePath) Then
        logs(4) = "5. FAILED: source file disappeared after rejected invalid id"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail("source file missing after rejection", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. No destination file, source file remained intact"

    logs(5) = "6. PASS"
    Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim missingId As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long
    Dim errDescription As String

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_ACTUALIZAR_NOEXISTE, "CAP008-ACT-NOEX")
    sourcePath = tempRoot & "source\" & FILE_PREFIX & "actualizar-noex.txt"
    Call WriteTextFile(sourcePath, "contenido noex")
    missingId = ID_ADJ_ACTUALIZAR_HAPPY + 100  ' fuera del rango, no sembrado
    logs(0) = "1. Seeded parent graph; idAdjunto=" & missingId & " intentionally absent"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_ACTUALIZAR_NOEXISTE)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    On Error Resume Next
    svc.ActualizarFicheroAdjunto missingId, sourcePath, db
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo EH

    If errNumber = 0 Then
        logs(2) = "3. FAILED: expected error for missing idAdjunto"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail("expected attachment not found error", logs)
        GoTo CleanExit
    End If
    If InStr(1, errDescription, "Attachment not found", vbTextCompare) = 0 Then
        logs(2) = "3. FAILED: rejection error did not mention 'Attachment not found': " & errDescription
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail("wrong rejection error", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. ActualizarFicheroAdjunto rejected missing idAdjunto with Err " & errNumber

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_ACTUALIZAR_NOEXISTE)
    If countAfter <> countBefore Then
        logs(3) = "4. FAILED: cardinality changed after rejected missing id: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail("rejected missing id changed cardinality", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (no side effect)"

    If FileExistsSafe(GetDocDir() & FILE_PREFIX & "actualizar-noex.txt") Then
        logs(4) = "5. FAILED: rejected update copied file to documentation folder"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail("file copied on missing id", logs)
        GoTo CleanExit
    End If
    If Not FileExistsSafe(sourcePath) Then
        logs(4) = "5. FAILED: source file disappeared after rejected missing id"
        Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail("source file missing after rejection", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. No destination file, source file remained intact"

    logs(5) = "6. PASS"
    Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long
    Dim errDescription As String
    Dim idAdjuntoOut As Long

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    ' Intentionally do NOT seed a solicitud: idSolicitud<=0 must be rejected before
    ' any FK or required-field check fires.
    sourcePath = tempRoot & "source\" & FILE_PREFIX & "id-invalido.txt"
    Call WriteTextFile(sourcePath, "contenido id invalido")
    logs(0) = "1. Sandbox ready; no parent graph seeded (idSolicitud=0 must be rejected)"

    countBefore = CountAdjuntosBySolicitud(db, -1)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    On Error Resume Next
    svc.GuardarAdjuntoDesdeArchivo 0, "CAP008-Etapa-IdInvalido", sourcePath, "Adjuntado", db, idAdjuntoOut, "id invalido"
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo EH

    If errNumber = 0 Then
        logs(2) = "3. FAILED: expected error for idSolicitud=0"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail("expected invalid idSolicitud error", logs)
        GoTo CleanExit
    End If
    If InStr(1, errDescription, "ID de la Solicitud", vbTextCompare) = 0 Then
        logs(2) = "3. FAILED: rejection error did not mention idSolicitud: " & errDescription
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail("wrong rejection error", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. GuardarAdjuntoDesdeArchivo rejected idSolicitud=0 with Err " & errNumber

    countAfter = CountAdjuntosBySolicitud(db, -1)
    If countAfter <> countBefore Then
        logs(3) = "4. FAILED: cardinality changed after rejected idSolicitud: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail("rejected idSolicitud changed cardinality", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (no side effect)"

    If idAdjuntoOut <> 0 Then
        logs(4) = "5. FAILED: idAdjuntoOut changed on rejected idSolicitud"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail("idAdjuntoOut changed", logs)
        GoTo CleanExit
    End If
    If Not FileExistsSafe(sourcePath) Then
        logs(4) = "5. FAILED: source file disappeared after rejected idSolicitud"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail("source file missing after rejection", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. No row, source file remained intact, idAdjuntoOut unchanged"

    logs(5) = "6. PASS"
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long
    Dim errDescription As String
    Dim idAdjuntoOut As Long

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_RUTA_VACIA, "CAP008-RUTA-VACIA")
    logs(0) = "1. Seeded parent graph; source path intentionally empty string"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_RUTA_VACIA)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    On Error Resume Next
    svc.GuardarAdjuntoDesdeArchivo ID_SOL_RUTA_VACIA, "CAP008-Etapa-RutaVacia", "", "Adjuntado", db, idAdjuntoOut, "ruta vacia"
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo EH

    If errNumber = 0 Then
        logs(2) = "3. FAILED: expected error for empty source path"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonFail("expected empty path error", logs)
        GoTo CleanExit
    End If
    If InStr(1, errDescription, "ruta del archivo de origen", vbTextCompare) = 0 Then
        logs(2) = "3. FAILED: rejection error did not mention empty source path: " & errDescription
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonFail("wrong rejection error", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. GuardarAdjuntoDesdeArchivo rejected empty source path with Err " & errNumber

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_RUTA_VACIA)
    If countAfter <> countBefore Then
        logs(3) = "4. FAILED: cardinality changed after rejected empty path: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonFail("rejected empty path changed cardinality", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (no side effect)"

    If idAdjuntoOut <> 0 Then
        logs(4) = "5. FAILED: idAdjuntoOut changed on rejected empty path"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonFail("idAdjuntoOut changed", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. idAdjuntoOut unchanged, no row inserted"

    logs(5) = "6. PASS"
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre() As String
    ' Slice A5a (2026-06-15): verifies that AdjuntosServicio.SubirYCerrar invokes
    ' the atomic approval callback ONLY when etapaNombre = "Documento Final Firmado".
    ' For any other stage, the seam must behave exactly like GuardarAdjuntoDesdeArchivo:
    ' the adjunto is persisted, the callback is NOT invoked, and no transition is
    ' attempted. This is the simple half of the seam contract. The happy-path
    ' atom (A5b, "Adjuntar PDF y aprobar solicitud PC") requires the full
    ' WorkflowServicio.EjecutarTransicion preconditions to be met and is left as
    ' WIP because those preconditions are not yet expressible in the sandbox.
    On Error GoTo EH
    Dim logs(0 To 9) As String
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim expectedDest As String
    Dim countBefore As Long
    Dim countAfter As Long
    Dim idAdjuntoOut As Long
    Dim tipoSolicitudBefore As String
    Dim idEstadoInternoBefore As Long

    logs(0) = "0. entered function"
    Call SetupAdjuntosSandbox(tempRoot, setupError)
    logs(1) = "1. setupError=[" & setupError & "]"
    If setupError <> "" Then
        logs(2) = "2. TESTS BLOCKED: " & setupError
        Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_SUBIR_Y_CERRAR, "CAP008-SUBIR-Y-CERRAR")
    tipoSolicitudBefore = GetTipoSolicitud(db, ID_SOL_SUBIR_Y_CERRAR)
    idEstadoInternoBefore = GetIdEstadoInterno(db, ID_SOL_SUBIR_Y_CERRAR)
    logs(2) = "2. seeded solicitud: tipoSolicitud=[" & tipoSolicitudBefore & "] idEstadoInterno=" & idEstadoInternoBefore

    sourcePath = tempRoot & "source\" & FILE_PREFIX & "subir-y-cerrar.pdf"
    expectedDest = GetDocDir() & FILE_PREFIX & "subir-y-cerrar.pdf"
    Call WriteTextFile(sourcePath, "subir-y-cerrar test content")
    logs(3) = "3. created source file (non-cerrar stage)"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_SUBIR_Y_CERRAR)
    logs(4) = "4. countBefore=" & countBefore

    ' Act: SubirYCerrar with a non-cerrar stage must NOT invoke the callback.
    svc.SubirYCerrar idSolicitud:=ID_SOL_SUBIR_Y_CERRAR, _
        etapaNombre:="CAP008-Etapa-NoCerrar", _
        rutaArchivoOrigen:=sourcePath, _
        tipoAccion:="Adjuntado", _
        descripcion:="subir sin cerrar", _
        dbTransaccional:=db, _
        idAdjuntoOut:=idAdjuntoOut
    logs(5) = "5. SubirYCerrar returned idAdjuntoOut=" & idAdjuntoOut

    If idAdjuntoOut <= 0 Then
        logs(6) = "6. FAILED: idAdjuntoOut not assigned"
        Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonFail("idAdjuntoOut not assigned", logs)
        GoTo CleanExit
    End If

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_SUBIR_Y_CERRAR)
    If countAfter <> countBefore + 1 Then
        logs(6) = "6. FAILED: cardinality wrong: before=" & countBefore & " after=" & countAfter
        Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonFail("cardinality wrong", logs)
        GoTo CleanExit
    End If
    logs(6) = "6. countAfter=" & countAfter & " (one row inserted)"

    ' Verify the callback was NOT invoked: idEstadoInterno and tipoSolicitud must be unchanged.
    If GetIdEstadoInterno(db, ID_SOL_SUBIR_Y_CERRAR) <> idEstadoInternoBefore Then
        logs(7) = "7. FAILED: idEstadoInterno changed: before=" & idEstadoInternoBefore & " after=" & GetIdEstadoInterno(db, ID_SOL_SUBIR_Y_CERRAR)
        Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonFail("callback was invoked unexpectedly", logs)
        GoTo CleanExit
    End If
    If GetTipoSolicitud(db, ID_SOL_SUBIR_Y_CERRAR) <> tipoSolicitudBefore Then
        logs(7) = "7. FAILED: tipoSolicitud changed"
        Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonFail("tipoSolicitud changed", logs)
        GoTo CleanExit
    End If
    logs(7) = "7. idEstadoInterno and tipoSolicitud unchanged (callback not invoked)"

    If Not FileExistsSafe(expectedDest) Then
        logs(8) = "8. FAILED: destination file missing: " & expectedDest
        Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonFail("destination file missing", logs)
        GoTo CleanExit
    End If
    logs(8) = "8. destination file copied"

    logs(9) = "9. PASS"
    Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(9) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre = TestHelper.BuildJsonFail(Err.Number & " - " & Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(8)
    Dim setupError As String
    Dim db As DAO.Database
    Dim svc As New AdjuntosServicio
    Dim tempRoot As String
    Dim sourcePath As String
    Dim expectedDest As String
    Dim idAdjuntoOut As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim fsoLocal As Object
    Dim fileExistsAfter As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    Call SetupAdjuntosSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_ROLLBACK, "CAP008-ROLLBACK-50")
    sourcePath = tempRoot & "source\" & FILE_PREFIX & "rollback.txt"
    Call WriteTextFile(sourcePath, "contenido rollback")
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    expectedDest = GetDocDir() & FILE_PREFIX & "rollback.txt"
    If fsoLocal.FileExists(expectedDest) Then fsoLocal.DeleteFile expectedDest, True
    logs(0) = "1. Seeded parent graph and source file under TEMP"

    countBefore = CountAdjuntosBySolicitud(db, ID_SOL_ROLLBACK)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected no adjuntos before act, got " & countBefore
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonFail("unexpected fixture rows before act", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore

    On Error Resume Next
    svc.GuardarAdjuntoDesdeArchivo idSolicitud:=ID_SOL_ROLLBACK, _
                                    etapaNombre:="CAP008-Etapa-Rollback", _
                                    rutaArchivoOrigen:=sourcePath, _
                                    tipoAccion:="Adjuntado", _
                                    db:=db, _
                                    idAdjuntoOut:=idAdjuntoOut, _
                                    descripcion:="BR-006 test", _
                                    forceFailAfterFileCopy:=True
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo EH
    logs(2) = "3. GuardarAdjuntoDesdeArchivo called with forceFailAfterFileCopy=True (named args)"

    If errorNumber = 0 Then
        logs(3) = "4. FAILED: expected CondorError Raise (513) on forceFail, but no error fired"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonFail("expected Raise from forceFail", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. Caught error " & errorNumber & ": " & errorDescription

    countAfter = CountAdjuntosBySolicitud(db, ID_SOL_ROLLBACK)
    If countAfter <> 0 Then
        logs(4) = "5. FAILED: BR-006 rollback failed: countAfter=" & countAfter & " (expected 0)"
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonFail("rollback failed: row persisted", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. countAfter=" & countAfter & " (BR-006 atomicity verified: no row persisted after Rollback)"

    fileExistsAfter = fsoLocal.FileExists(expectedDest)
    If fileExistsAfter Then
        logs(5) = "6. FAILED: destination file remained on disk after rollback: " & expectedDest
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonFail("file compensation failed", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. destination file removed (compensación de archivo ok)"

    If Not fsoLocal.FileExists(sourcePath) Then
        logs(6) = "7. FAILED: source file was unexpectedly deleted: " & sourcePath
        Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonFail("source file deleted", logs)
        GoTo CleanExit
    End If
    logs(6) = "7. source file intact at " & sourcePath

    logs(7) = "8. PASS"
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    Set fsoLocal = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Adjuntos_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(14)
    Dim result As String

    result = Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_CreaFilaYCopiaArchivo()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(0) = "1. GuardarAdjuntoDesdeArchivo happy path passed"

    result = Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(1) = "2. EliminarAdjunto passed"

    result = Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(2) = "3. EliminarAdjuntosPorEtapas passed"

    result = Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(3) = "4. Missing path sad path passed"

    result = Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(4) = "5. Non-PDF final signed document rejection passed"

    result = Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(5) = "6. Destination filename deduplication passed"

    result = Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(6) = "7. ActualizarFicheroAdjunto happy path passed"

    result = Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInvalido()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(7) = "8. ActualizarFicheroAdjunto rejects invalid idAdjunto"

    result = Test_Adjuntos_Strict_ActualizarFicheroAdjunto_RechazaIdAdjuntoInexistente()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(8) = "9. ActualizarFicheroAdjunto rejects missing idAdjunto"

    result = Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(9) = "10. GuardarAdjuntoDesdeArchivo rejects idSolicitud<=0"

    result = Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaOrigenVacia()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(10) = "11. GuardarAdjuntoDesdeArchivo rejects empty source path"

    result = Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(11) = "12. SubirYCerrar does not invoke callback for non-cerrar stages"

    result = Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_Rollback_SiFallaDespuesDeCopia()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(12) = "13. GuardarAdjuntoDesdeArchivo triggers BR-006 rollback (countAfter=0 + file compensado) when forceFail=True"

    logs(13) = "14. PASS"
    Test_Adjuntos_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

Failed:
    Test_Adjuntos_Strict_RunAll = TestHelper.BuildJsonFail("one CAP-008 adjuntos atom failed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Adjuntos_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_Adjuntos_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

' ============================================================================
' Setup / teardown helpers
' ============================================================================

Private Sub SetupAdjuntosSandbox(ByRef p_TempRoot As String, ByRef p_Error As String)
    On Error GoTo EH
    Dim runError As String
    Dim logs() As String
    logs = TestHelper.NewLogsArray(3)

    p_Error = ""
    p_TempRoot = ""

    ' v2.4.3 harness: hardened production guard + sandbox URL validation
    ' moved to TestHelper.BeginTestSession (UNC, fingerprint, FSO+DAO).
    ' SetupProdGlobalsForTest covers the 4 prod globals + tempRoot base.
    If Not TestHelper.BeginTestSession(logs, runError) Then
        p_Error = runError
        Exit Sub
    End If

    Call TestHelper.SetupProdGlobalsForTest(runError, p_TempRoot, "cap008_adjuntos", "CAP008 Test User")
    If runError <> "" Then
        p_Error = runError
        Call TestHelper.ResetTestSession
        Exit Sub
    End If

    ' CAP008-specific: crear subcarpetas source/ y app/ANEXOS/ que el
    ' servicio de adjuntos necesita ademas del app\ base.
    Call EnsureFolder(p_TempRoot & "source\")
    Call EnsureFolder(p_TempRoot & "app\ANEXOS\")
    Exit Sub

EH:
    p_Error = "TESTS BLOCKED: SetupAdjuntosSandbox failed: " & Err.Number & " - " & Err.Description
    Call TestHelper.ResetTestSession
End Sub

Private Sub TeardownFixtures(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbAdjuntos WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbAdjuntos WHERE idAdjunto >= " & TEST_ID_BASE & " AND idAdjunto <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE & " AND IDExpediente <= " & TEST_ID_TOP, dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedSolicitud(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long, ByVal p_CodeSuffix As String)
    Dim sql As String
    sql = "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_IdSolicitud & ")"
    p_Db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", 'CDCA', '" & p_CodeSuffix & "', 2, Now(), " & _
          "'TestAdjuntosStrict', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedAdjuntoConArchivo(ByVal p_Db As DAO.Database, ByVal p_IdAdjunto As Long, ByVal p_IdSolicitud As Long, ByVal p_Etapa As String, ByVal p_NombreArchivo As String)
    Call WriteTextFile(GetDocDir() & p_NombreArchivo, "contenido " & p_NombreArchivo)
    Call SeedAdjunto(p_Db, p_IdAdjunto, p_IdSolicitud, p_Etapa, p_NombreArchivo)
End Sub

Private Sub SeedAdjunto(ByVal p_Db As DAO.Database, ByVal p_IdAdjunto As Long, ByVal p_IdSolicitud As Long, ByVal p_Etapa As String, ByVal p_NombreArchivo As String)
    Dim sql As String
    sql = "INSERT INTO tbAdjuntos (idAdjunto, idSolicitud, etapaWF, nombreArchivo, fechaSubida, usuarioSubida, descripcion, TipoAccion) VALUES (" & _
          p_IdAdjunto & ", " & p_IdSolicitud & ", " & TestHelper.SqlStr(p_Etapa) & ", " & _
          TestHelper.SqlStr(p_NombreArchivo) & ", Now(), 'TestAdjuntosStrict', 'CAP008 fixture', 'Adjuntado')"
    p_Db.Execute sql, dbFailOnError
End Sub

' ============================================================================
' Assertion / filesystem helpers
' ============================================================================

Private Function CountAdjuntosBySolicitud(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbAdjuntos WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs Is Nothing Then
        CountAdjuntosBySolicitud = Nz(rs!n, 0)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function CountTotalAdjuntos(ByVal p_Db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbAdjuntos", dbOpenSnapshot)
    If Not rs Is Nothing Then
        CountTotalAdjuntos = Nz(rs!n, 0)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function GetTipoSolicitud(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As String
    Dim rs As DAO.Recordset
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT tipoSolicitud FROM tbSolicitudes WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then
        GetTipoSolicitud = Nz(rs!tipoSolicitud, "")
    End If
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function GetIdEstadoInterno(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT idEstadoInterno FROM tbSolicitudes WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then
        GetIdEstadoInterno = Nz(rs!idEstadoInterno, 0)
    End If
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

' ============================================================================
' BR-006 deferred — not testable without refactor
' ============================================================================
'
' BR-006: "Si la persistencia falla tras la copia, se borra el archivo destino
' (fso.DeleteFile destinoFull, True)".
'
' Investigation 2026-06-15 found two blocking constraints:
'
' (a) PK collision via pre-inserted dummy row does NOT work. The service
'     computes idAdjunto = MAX(idAdjunto) + 1 in getSiguienteIDAdjunto() AFTER
'     the test inserts the dummy, so its MAX+1 always points one past the dummy
'     (the dummy is the new MAX), never colliding with the dummy itself.
'     Single-threaded VBA makes it impossible to insert the dummy between
'     getSiguienteIDAdjunto() and the actual INSERT.
'
' (b) Forcing the failure via Validar rule 3 (nombreArchivo > 255 chars) requires
'     creating a physical file with a 256+ char name. Windows MAX_PATH (260 chars)
'     blocks this when the %TEMP%\source\ path is already ~85 chars. FSO COM
'     does not support the \\?\ prefix natively.
'
' Resolution: BR-006 stays Verified-static. To unblock, the service must be
' refactored to expose a seam for forcing persistence failure (e.g., a
' dependency-injected DAO.Database that can be configured to fail on INSERT,
' or an internal "force fail" flag for tests). After that refactor, this test
' module can be extended with a real atom.

Private Function CountAdjuntosByEtapa(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long, ByVal p_Etapa As String) As Long
    Dim rs As DAO.Recordset
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbAdjuntos WHERE idSolicitud=" & p_IdSolicitud & " AND etapaWF=" & TestHelper.SqlStr(p_Etapa), dbOpenSnapshot)
    If Not rs Is Nothing Then
        CountAdjuntosByEtapa = Nz(rs!n, 0)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function GetDocDir() As String
    Dim result As String
    result = m_ObjEntorno.URLDirectorioDocumentacion
    If Right$(result, 1) <> "\" Then result = result & "\"
    GetDocDir = result
End Function

Private Sub EnsureFolder(ByVal p_Path As String)
    Dim fsoLocal As Object
    Dim parentPath As String
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If fsoLocal.FolderExists(p_Path) Then GoTo CleanExit
    parentPath = fsoLocal.GetParentFolderName(Left$(p_Path, Len(p_Path) - 1))
    If Len(parentPath) > 0 Then
        If Not fsoLocal.FolderExists(parentPath) Then
            Call EnsureFolder(parentPath & "\")
        End If
    End If
    fsoLocal.CreateFolder p_Path

CleanExit:
    Set fsoLocal = Nothing
End Sub

Private Sub WriteTextFile(ByVal p_Path As String, ByVal p_Content As String)
    Dim fsoLocal As Object
    Dim ts As Object
    Dim parentPath As String

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    parentPath = fsoLocal.GetParentFolderName(p_Path)
    If Len(parentPath) > 0 Then
        If Not fsoLocal.FolderExists(parentPath) Then
            Call EnsureFolder(parentPath & "\")
        End If
    End If
    Set ts = fsoLocal.CreateTextFile(p_Path, True, False)
    ts.Write p_Content
    ts.Close
    Set ts = Nothing
    Set fsoLocal = Nothing
End Sub

Private Function FileExistsSafe(ByVal p_Path As String) As Boolean
    Dim fsoLocal As Object
    On Error GoTo CleanExit
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    FileExistsSafe = fsoLocal.FileExists(p_Path)

CleanExit:
    Set fsoLocal = Nothing
End Function

Private Function ReadTextFile(ByVal p_Path As String) As String
    Dim fsoLocal As Object
    Dim ts As Object
    On Error GoTo CleanExit
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If fsoLocal.FileExists(p_Path) Then
        Set ts = fsoLocal.OpenTextFile(p_Path, 1, False)
        ReadTextFile = ts.ReadAll
    End If

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fsoLocal = Nothing
End Function
