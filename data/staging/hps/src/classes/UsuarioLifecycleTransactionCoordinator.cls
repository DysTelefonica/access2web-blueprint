VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioLifecycleTransactionCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const STAGE_BEFORE_COMMIT As String = "BeforeCommit"
Private Const STAGE_BEFORE_SOURCE_DELETE As String = "BeforeSourceDelete"
Private Const STAGE_AFTER_SOURCE_DELETE As String = "AfterSourceDelete"
Private Const STAGE_AFTER_COMMIT_REFRESH As String = "AfterCommitRefresh"

Private Const DIRECTION_C2H As String = "c2h"
Private Const DIRECTION_H2C As String = "h2c"
' [binario-driven] Los anexos viven junto al binario de Access. Si el
' form o test inyecta p_AttachmentsRootActual/Historico, gana ese root;
' si no, se usa el path local junto al binario. NO hay constante UNC
' corporativa (\\10.14.7.52\...) como DEFAULT porque puede no ser
' alcanzable desde la sesión del usuario. Esta decisión es coherente con
' Entorno.ResolveURLCarpetaAnexosCandidate: el binario es la única
' verdad para los anexos.
Private Const DEFAULT_ATTACHMENTS_ROOT_ACTUAL As String = ""
Private Const DEFAULT_ATTACHMENTS_ROOT_HISTORICO As String = ""

Private m_PromptProcName As String

Public Function PasarActualAHistorico(ByVal p_IDActual As String, _
                                      Optional ByRef p_Db As DAO.Database = Nothing, _
                                      Optional ByVal p_AttachmentsRootActual As String = "", _
                                      Optional ByVal p_AttachmentsRootHistorico As String = "", _
                                      Optional ByVal p_FailureStage As String = "", _
                                      Optional ByRef p_Error As String) As Boolean
    PasarActualAHistorico = ExecuteLifecycleMove(p_IDActual, True, p_Db, p_AttachmentsRootActual, p_AttachmentsRootHistorico, p_FailureStage, p_Error)
End Function

Public Function PasarHistoricoAActual(ByVal p_IDHistorico As String, _
                                      Optional ByRef p_Db As DAO.Database = Nothing, _
                                      Optional ByVal p_AttachmentsRootActual As String = "", _
                                      Optional ByVal p_AttachmentsRootHistorico As String = "", _
                                      Optional ByVal p_FailureStage As String = "", _
                                      Optional ByRef p_Error As String) As Boolean
    PasarHistoricoAActual = ExecuteLifecycleMove(p_IDHistorico, False, p_Db, p_AttachmentsRootActual, p_AttachmentsRootHistorico, p_FailureStage, p_Error)
End Function

Private Function ExecuteLifecycleMove(ByVal p_IDUsuario As String, ByVal p_CurrentToHistorical As Boolean, _
                                      Optional ByRef p_Db As DAO.Database = Nothing, _
                                      Optional ByVal p_AttachmentsRootActual As String = "", _
                                      Optional ByVal p_AttachmentsRootHistorico As String = "", _
                                      Optional ByVal p_FailureStage As String = "", _
                                      Optional ByRef p_Error As String) As Boolean
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim ownsDb As Boolean
    Dim committed As Boolean
    Dim transactionStarted As Boolean
    Dim sourcePath As String
    Dim targetPath As String
    Dim snapshotPath As String
    Dim conflictSnapshotPath As String
    Dim conflicts As String
    Dim precheckSnapshotCount As Long
    Dim recheckConflicts As String
    Dim direction As String
    Dim safeUserId As Long
    Dim approved As Boolean
    Dim dniArg As String
    Dim directionLabel As String
    On Error GoTo EH

    p_Error = ""
    transactionStarted = False
    ' [HATW probe] Persistent diagnostic for failed lifecycle transitions.
    ' Writes to %TEMP%\HPS\HATW-probe.log (separate from pr1a-trace.log so
    ' the trace used by Test_HistoricoAdjuntosTransactionWrapper is not
    ' polluted). Includes path resolution so the operator can attach the
    ' log and we can tell whether the failure was a missing UNC source,
    ' a missing local root, a missing DNI folder, or a failed DB step.
    Dim sProbeDirection As String
    sProbeDirection = IIf(p_CurrentToHistorical, "c2h", "h2c")
    LogToProbe "[ExecuteLifecycleMove] START user=" & p_IDUsuario & _
        " dir=" & sProbeDirection & _
        " p_AttachmentsRootActual='" & p_AttachmentsRootActual & "'" & _
        " p_AttachmentsRootHistorico='" & p_AttachmentsRootHistorico & "'" & _
        " p_Db is Nothing=" & (p_Db Is Nothing)
    LogToFile "[ExecuteLifecycleMove] START user=" & p_IDUsuario & " dir=" & sProbeDirection & " p_Db is Nothing=" & (p_Db Is Nothing)
    If Not ValidateInjectedRoots(p_AttachmentsRootActual, p_AttachmentsRootHistorico, p_Error) Then
        LogToProbe "[ExecuteLifecycleMove] ValidateInjectedRoots rejected: p_Error='" & p_Error & "'"
        Err.Raise 1000
    End If
    If p_Db Is Nothing Then
        Set db = getdb(p_Error)
        ownsDb = True
    Else
        Set db = p_Db
    End If
    If p_Error <> "" Then Err.Raise 1000
    If Not db Is Nothing Then
        LogToFile "[ExecuteLifecycleMove] db.Name=" & db.Name & " db.Connect=" & db.Connect
        LogToProbe "[ExecuteLifecycleMove] db.Name=" & db.Name & " db.Connect=" & db.Connect
    End If
    Set ws = DBEngine.Workspaces(0)
    LogToFile "[ExecuteLifecycleMove] ws.Name=" & ws.Name
    LogToProbe "[ExecuteLifecycleMove] ws.Name=" & ws.Name

    safeUserId = ValidateLifecycleUserId(p_IDUsuario)
    ResolvePaths db, safeUserId, p_CurrentToHistorical, p_AttachmentsRootActual, p_AttachmentsRootHistorico, sourcePath, targetPath, p_Error
    ' [HATW probe] Capture the actual paths the coordinator is going to
    ' touch AND whether they exist on disk, BEFORE any copy/delete runs.
    ' This is the most useful line for diagnosing "No se ha encontrado la
    ' ruta de acceso" failures: it tells the operator (and the agent) which
    ' specific path is missing.
    LogToProbe "[ExecuteLifecycleMove] resolved sourcePath='" & sourcePath & _
        "' exists=" & fso.FolderExists(sourcePath) & _
        " targetPath='" & targetPath & _
        "' exists=" & fso.FolderExists(targetPath)
    If p_Error <> "" Then
        LogToProbe "[ExecuteLifecycleMove] ResolvePaths failed: p_Error='" & p_Error & "'"
        Err.Raise 1000
    End If
    snapshotPath = BuildSnapshotPath(CStr(safeUserId))
    conflictSnapshotPath = snapshotPath & "-target-conflicts"
    DeleteFolder snapshotPath
    DeleteFolder conflictSnapshotPath
    CopyFolderIfExists sourcePath, snapshotPath

    ' [PR1a] Precheck #1 — DB + folder conflicts OUTSIDE transaction
    If p_CurrentToHistorical Then
        direction = DIRECTION_C2H
    Else
        direction = DIRECTION_H2C
    End If
    LogToFile "[ExecuteLifecycleMove] precheck#1 start direction=" & direction
    conflicts = CheckLifecycleConflicts(db, direction, safeUserId, p_AttachmentsRootActual, p_AttachmentsRootHistorico)
    LogToFile "[ExecuteLifecycleMove] precheck#1 result conflicts='" & conflicts & "'"
    If Len(conflicts) > 0 Then
        If Len(m_PromptProcName) = 0 Then
            p_Error = "Lifecycle conflicts detected but no conflict prompt is configured (ConfigureLifecycleConflictPrompt). Conflicts: " & conflicts
            GoTo CleanUp
        End If
        dniArg = ExtractFirstDNI(conflicts)
        If p_CurrentToHistorical Then
            directionLabel = "Current->Historical"
        Else
            directionLabel = "Historical->Current"
        End If
        LogToFile "[ExecuteLifecycleMove] prompt call m_PromptProcName=" & m_PromptProcName & " conflicts=" & conflicts & " dni=" & dniArg & " dir=" & directionLabel
        approved = CBool(Application.Run(m_PromptProcName, conflicts, dniArg, directionLabel))
        LogToFile "[ExecuteLifecycleMove] prompt result approved=" & approved
        If Not approved Then
            p_Error = "Operación cancelada por el usuario. Conflictos: " & conflicts
            GoTo CleanUp
        End If
    End If

    ' [PR1a] TOCTOU counter-capture BEFORE BeginTrans
    precheckSnapshotCount = CountConflictTokens(conflicts)
    LogToFile "[ExecuteLifecycleMove] precheckSnapshotCount=" & precheckSnapshotCount

    ' [PR1a] Precheck #2 — re-run OUTSIDE transaction, AFTER the prompt call.
    LogToFile "[ExecuteLifecycleMove] precheck#2 start"
    recheckConflicts = CheckLifecycleConflicts(db, direction, safeUserId, p_AttachmentsRootActual, p_AttachmentsRootHistorico)
    LogToFile "[ExecuteLifecycleMove] precheck#2 result conflicts='" & recheckConflicts & "' count=" & CountConflictTokens(recheckConflicts)
    LogToFile "[ExecuteLifecycleMove] TOCTOU check: precheckSnapshotCount=" & precheckSnapshotCount & " recheckCount=" & CountConflictTokens(recheckConflicts) & " detected=" & (CountConflictTokens(recheckConflicts) > precheckSnapshotCount)
    If CountConflictTokens(recheckConflicts) > precheckSnapshotCount Then
        Err.Raise 1000, , "TOCTOU: new conflict appeared between precheck #1 and precheck #2. Original=" & conflicts & " Recheck=" & recheckConflicts
    End If
    If Len(conflicts) > 0 Then
        LogToFile "[ExecuteLifecycleMove] target folder conflict snapshot start"
        SnapshotPriorConflictFolders direction, conflicts, p_AttachmentsRootActual, p_AttachmentsRootHistorico, conflictSnapshotPath
        LogToFile "[ExecuteLifecycleMove] target folder conflict snapshot done"
    End If

    ws.BeginTrans
    transactionStarted = True
    LogToFile "[ExecuteLifecycleMove] BeginTrans done"

    ' [PR1a] Precheck-conditional cleanup INSIDE transaction, BEFORE copy-insert AND commit
    If Len(conflicts) > 0 Then
        LogToFile "[ExecuteLifecycleMove] cleanup: DeletePriorConflictRows start"
        DeletePriorConflictRows db, direction, conflicts
        LogToFile "[ExecuteLifecycleMove] cleanup: DeletePriorConflictRows done"
        LogToFile "[ExecuteLifecycleMove] cleanup: DeletePriorConflictFolders start"
        DeletePriorConflictFolders direction, conflicts, p_AttachmentsRootActual, p_AttachmentsRootHistorico
        LogToFile "[ExecuteLifecycleMove] cleanup: DeletePriorConflictFolders done"
    End If

    If p_CurrentToHistorical Then
        CopyCurrentRowsToHistorical db, p_IDUsuario
    Else
        CopyHistoricalRowsToCurrent db, p_IDUsuario
    End If
    ' [c2h-h2c-subcarpetas] Per-anexo routing decision happens AFTER the row
    ' copy so TbAnexosUsuariosHistoricos still holds the EsHistorico values
    ' for this user (the historical side has not been deleted yet). The
    ' dictionary is consumed by CopyFolderByEsHistorico (h2c) and ignored
    ' on c2h (which uses CopyFolderIfExists / flatten).
    If p_CurrentToHistorical Then
        ' c2h: the PR2 RFC norm is FLAT — every file under
        ' <HPS>\<DNI>\Actual\ + <HPS>\<DNI>\HISTORICO\ lands flat at
        ' <HISTORICO>\<DNI>\<file>. The per-anexo origin is preserved
        ' through the EsHistorico column in TbAnexosUsuariosHistoricos
        ' (see CopyCurrentRowsToHistorical) and consumed on h2c.
        CopyFolderIfExists sourcePath, targetPath
    Else
        ' h2c: split-routing. The flat source at <HISTORICO>\<DNI>\<file>
        ' is split into <HPS>\<DNI>\Actual\<file> (EsHistorico='No') or
        ' <HPS>\<DNI>\HISTORICO\<file> (EsHistorico='Sí'). The lookup
        ' dictionary is built from the historical rows we just INSERTed
        ' into TbAnexosUsuariosHPS, but for routing we need the SOURCE
        ' rows' EsHistorico (still in TbAnexosUsuariosHistoricos because
        ' the DELETE happens after this step).
        Dim esHistoricoByName As Object
        Set esHistoricoByName = BuildEsHistoricoByName(db, p_IDUsuario)
        CopyFolderByEsHistorico sourcePath, _
                                 EnsureSlash(targetPath) & "Actual\", _
                                 EnsureSlash(targetPath) & "HISTORICO\", _
                                 esHistoricoByName
        ' [c2h-h2c-subcarpetas] Garantiza que la estructura anexos del usuario
        ' activo quede completa: tanto Actual\ como HISTORICO\ existen siempre,
        ' aunque una de las dos subcarpetas quede vacia. Llamado post-h2c para
        ' que el folder del usuario activo tenga la forma esperada
        ' independientemente de cuantos anexos cayeron en cada subcarpeta.
        EnsureAnexosSubcarpetasStructure EnsureSlash(targetPath)
    End If
    ForceFailure p_FailureStage, STAGE_BEFORE_SOURCE_DELETE
    DeleteFolder sourcePath
    ForceFailure p_FailureStage, STAGE_AFTER_SOURCE_DELETE
    If p_CurrentToHistorical Then
        DeleteCurrentRows db, p_IDUsuario
    Else
        DeleteHistoricalRows db, p_IDUsuario
    End If
    ForceFailure p_FailureStage, STAGE_BEFORE_COMMIT
    ws.CommitTrans
    committed = True

    ' [PR1b cache] Atomic single-user cache sync post-commit. Replaces the
    ' previous full-regen of the current cache with a per-user operation
    ' that handles BOTH directions: c2h removes the user from the current
    ' cache and inserts into the historical cache; h2c removes from the
    ' historical cache and inserts into the current cache. The full regen
    ' path is preserved at startup (SincronizarCachesLocalesFrontend)
    ' where it is appropriate. Best-effort: a cache failure does NOT roll
    ' back the move (it is already committed); the operator can trigger a
    ' manual regen if needed.
    If p_FailureStage <> STAGE_AFTER_COMMIT_REFRESH Then
        Dim atomicCacheErr As String
        On Error Resume Next
        cache_usuario_sincronizar_despues_movimiento CStr(safeUserId), direction, atomicCacheErr
        On Error GoTo EH
        LogToFile "[ExecuteLifecycleMove] post-commit atomic cache sync (" & direction & ") err=" & atomicCacheErr
    Else
        p_Error = "Lifecycle transition committed, but post-commit refresh failed by test seam."
        LogToFile "[ExecuteLifecycleMove] ForceFailure AFTER_COMMIT_REFRESH: " & p_Error
    End If
    ExecuteLifecycleMove = True
    LogToFile "[ExecuteLifecycleMove] END OK"
    LogToProbe "[ExecuteLifecycleMove] END OK sourcePath='" & sourcePath & "' targetPath='" & targetPath & "'"

CleanUp:
    On Error Resume Next
    DeleteFolder snapshotPath
    DeleteFolder conflictSnapshotPath
    If ownsDb Then Set db = Nothing
    Exit Function
EH:
    ' [PR1a error propagation] capture the original Err.Description BEFORE the
    ' On Error Resume Next block. A Rollback without an active transaction
    ' (typical for precheck errors that fire before BeginTrans) raises "Ha
    ' intentado confirmar o deshacer una transacción sin iniciar una antes"
    ' and would otherwise overwrite the real precheck error in p_Error.
    Dim ehErrDesc As String
    ehErrDesc = Err.Description
    LogToFile "[ExecuteLifecycleMove] EH err=" & Err.Number & " desc=" & ehErrDesc & " committed=" & committed
    LogToProbe "[ExecuteLifecycleMove] EH err=" & Err.Number & " desc='" & ehErrDesc & "' committed=" & committed & " sourcePath='" & sourcePath & "' targetPath='" & targetPath & "'"
    If Not committed And transactionStarted And Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
        DeleteFolder targetPath
        RestoreSnapshot snapshotPath, sourcePath
        RestorePriorConflictFolders direction, conflicts, p_AttachmentsRootActual, p_AttachmentsRootHistorico, conflictSnapshotPath
        On Error GoTo 0
    End If
    If p_Error = "" Then p_Error = "UsuarioLifecycleTransactionCoordinator: " & ehErrDesc
    ExecuteLifecycleMove = False
    LogToFile "[ExecuteLifecycleMove] END FAIL p_Error=" & p_Error
    Resume CleanUp
End Function

' [PR1a debug] — append a line to %TEMP%\HPS\pr1a-trace.log. The agent reads
' this file after the test run to inspect what happened during the lifecycle.
' File-based logging because Debug.Print output is not visible through the
' dysflow MCP; a file we can re-read with dysflow's query_sql on a text file or
' just File.ReadAllText from the shell.
Private Sub LogToFile(ByVal p_Message As String)
    On Error Resume Next
    Dim logPath As String
    logPath = Environ$("TEMP") & "\HPS\pr1a-trace.log"
    Dim f As Integer
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & p_Message
    Close #f
End Sub

' [HATW probe] Persistent diagnostic for failed lifecycle transitions.
' Separate from LogToFile (pr1a-trace.log) so the test-suite trace is not
' polluted. The operator attaches this file when reporting a failed
' PasarDeActualAHistorico / PasarDeHistoricoAActual so the agent can see
' which path the coordinator tried to touch and whether the FS already had
' the expected folders. Kept open-as-write so concurrent calls append
' rather than overwrite.
Private Sub LogToProbe(ByVal p_Message As String)
    On Error Resume Next
    Dim logPath As String
    logPath = Environ$("TEMP") & "\HPS\HATW-probe.log"
    Dim f As Integer
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & p_Message
    Close #f
End Sub

Private Sub CopyCurrentRowsToHistorical(ByRef db As DAO.Database, ByVal idUsuario As Long)
    db.Execute "INSERT INTO TbUsuariosHistoricos SELECT * FROM TbUsuarios WHERE ID=" & idUsuario, dbFailOnError
    CopyObservationRowsWithCollisionSafe db, "TbObservaciones", "TbObservacionesHistoricas", idUsuario
    ' [PR1b anexos] each IDAnexo in the destination must be CALCULATED, not
    ' preserved from the source. DAO/Jet has no ROW_NUMBER(), so we assign
    ' manually: base = MAX + 1, each subsequent row gets base + idx. Mirror
    ' logic in CopyHistoricalRowsToCurrent below.
    '
    ' Use db-scoped OpenRecordset for the MAX probe instead of DMax(domain
    ' function). DMax routes through Application.CurrentDb (the FE in this
    ' codebase), not the `db` handle the test injected, and the FE's view
    ' of the table is the linked-table name. The OpenRecordset here is
    ' guaranteed to read from the same backend the INSERT writes to, which
    ' is what the test asserts against.
    '
    ' [c2h-h2c-subcarpetas] EsHistorico is now preserved from the source row
    ' so h2c can route each file back to the correct <HPS>\<DNI>\Actual\ or
    ' HISTORICO\ subfolder. Anexo that lived in Actual\ (EsHistorico='No')
    ' stays 'No' in the historical table; the one in HISTORICO\ ('Sí') stays
    ' 'Sí'. The new CopyFolderByEsHistorico (h2c) reads the preserved flag.
    Dim cNextIDAnexo As Long
    cNextIDAnexo = 1
    Dim cMaxRs As DAO.Recordset
    Set cMaxRs = db.OpenRecordset("SELECT MAX(IDAnexo) AS M FROM TbAnexosUsuariosHistoricos", dbOpenSnapshot)
    If Not cMaxRs.EOF Then
        If Not IsNull(cMaxRs!M) Then cNextIDAnexo = CLng(cMaxRs!M) + 1
    End If
    cMaxRs.Close
    Dim cAnexoIdx As Long
    Dim cAnexoRs As DAO.Recordset
    Dim cEsHistoricoOrigen As String
    Set cAnexoRs = db.OpenRecordset("SELECT NombreAnexo, IDUsuario, Hash, EsHistorico FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & idUsuario & " ORDER BY IDAnexo", dbOpenSnapshot)
    Do While Not cAnexoRs.EOF
        cEsHistoricoOrigen = Nz(cAnexoRs!EsHistorico, "No")
        db.Execute "INSERT INTO TbAnexosUsuariosHistoricos (IDAnexo, NombreAnexo, IDUsuario, Hash, EsHistorico) VALUES (" & (cNextIDAnexo + cAnexoIdx) & ", '" & Replace(Nz(cAnexoRs!NombreAnexo, ""), "'", "''") & "', " & cAnexoRs!IDUsuario & ", '" & Replace(Nz(cAnexoRs!Hash, ""), "'", "''") & "', '" & cEsHistoricoOrigen & "')", dbFailOnError
        cAnexoIdx = cAnexoIdx + 1
        cAnexoRs.MoveNext
    Loop
    cAnexoRs.Close
End Sub

Private Sub CopyHistoricalRowsToCurrent(ByRef db As DAO.Database, ByVal idUsuario As Long)
    db.Execute "INSERT INTO TbUsuarios SELECT * FROM TbUsuariosHistoricos WHERE ID=" & idUsuario, dbFailOnError
    CopyObservationRowsWithCollisionSafe db, "TbObservacionesHistoricas", "TbObservaciones", idUsuario
    ' [PR1b anexos] mirror of CopyCurrentRowsToHistorical: calculate fresh
    ' IDAnexo per row, do not preserve. MAX via db-scoped OpenRecordset
    ' (see same caveat in the c2h method above; DMax routes through the
    ' FE's CurrentDb and mis-resolves the table).
    '
    ' [c2h-h2c-subcarpetas] EsHistorico is now READ from the source row
    ' instead of hardcoded 'No'. The h2c split-routing in
    ' CopyFolderByEsHistorico needs the per-anexo EsHistorico flag to land
    ' each file in the right <HPS>\<DNI>\Actual\ or HISTORICO\ subfolder.
    ' A historical row that was 'No' before c2h (lived in Actual\) is
    ' restored to Actual\; one that was 'Sí' (lived in HISTORICO\) is
    ' restored to HISTORICO\. The default for a NULL source value is 'No'
    ' (matches the historical-side default convention; without the default
    ' a legacy fixture with NULL EsHistorico would fail-routing in h2c).
    Dim hNextIDAnexo As Long
    hNextIDAnexo = 1
    Dim hMaxRs As DAO.Recordset
    Set hMaxRs = db.OpenRecordset("SELECT MAX(IDAnexo) AS M FROM TbAnexosUsuariosHPS", dbOpenSnapshot)
    If Not hMaxRs.EOF Then
        If Not IsNull(hMaxRs!M) Then hNextIDAnexo = CLng(hMaxRs!M) + 1
    End If
    hMaxRs.Close
    Dim hAnexoIdx As Long
    Dim hAnexoRs As DAO.Recordset
    Dim hEsHistoricoOrigen As String
    Set hAnexoRs = db.OpenRecordset("SELECT NombreAnexo, IDUsuario, Hash, EsHistorico FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & idUsuario & " ORDER BY IDAnexo", dbOpenSnapshot)
    Do While Not hAnexoRs.EOF
        hEsHistoricoOrigen = Nz(hAnexoRs!EsHistorico, "No")
        db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, Hash, EsHistorico) VALUES (" & (hNextIDAnexo + hAnexoIdx) & ", '" & Replace(Nz(hAnexoRs!NombreAnexo, ""), "'", "''") & "', " & hAnexoRs!IDUsuario & ", '" & Replace(Nz(hAnexoRs!Hash, ""), "'", "''") & "', '" & hEsHistoricoOrigen & "')", dbFailOnError
        hAnexoIdx = hAnexoIdx + 1
        hAnexoRs.MoveNext
    Loop
    hAnexoRs.Close
End Sub

Private Sub CopyObservationRowsWithCollisionSafe(ByRef db As DAO.Database, ByVal sourceTable As String, ByVal destinationTable As String, ByVal idUsuario As Long)
    Dim nextIDObservacion As Long
    Dim sourceIDObservacion As Long
    Dim targetIDObservacion As Long
    Dim safeUserId As Long
    Dim obsRs As DAO.Recordset
    Dim destRs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    On Error GoTo EH

    ValidateObservationCopyPair sourceTable, destinationTable
    safeUserId = idUsuario
    nextIDObservacion = GetNextObservationID(db, destinationTable)
    Set qdf = db.CreateQueryDef("", "PARAMETERS pUserId LONG; SELECT IDObservacion, ID, Fecha, Observacion, Tipo FROM " & sourceTable & " WHERE ID=[pUserId] ORDER BY IDObservacion")
    qdf.Parameters("pUserId") = safeUserId
    Set obsRs = qdf.OpenRecordset(dbOpenSnapshot)
    Set destRs = db.OpenRecordset(destinationTable, dbOpenDynaset)
    Do While Not obsRs.EOF
        sourceIDObservacion = CLng(obsRs!IDObservacion)
        targetIDObservacion = sourceIDObservacion
        If ObservationIDExists(db, destinationTable, targetIDObservacion) Then
            targetIDObservacion = nextIDObservacion
            Do While ObservationIDExists(db, destinationTable, targetIDObservacion)
                targetIDObservacion = targetIDObservacion + 1
            Loop
            nextIDObservacion = targetIDObservacion + 1
        ElseIf targetIDObservacion >= nextIDObservacion Then
            nextIDObservacion = targetIDObservacion + 1
        End If

        destRs.AddNew
        destRs.Fields("IDObservacion") = targetIDObservacion
        destRs.Fields("ID") = obsRs!ID
        destRs.Fields("Fecha") = obsRs!Fecha
        destRs.Fields("Observacion") = obsRs!Observacion
        destRs.Fields("Tipo") = obsRs!Tipo
        destRs.Update
        obsRs.MoveNext
    Loop

CleanUp:
    On Error Resume Next
    If Not destRs Is Nothing Then destRs.Close
    If Not obsRs Is Nothing Then obsRs.Close
    If Not qdf Is Nothing Then qdf.Close
    Set destRs = Nothing
    Set obsRs = Nothing
    Set qdf = Nothing
    Exit Sub
EH:
    Dim copyErrNumber As Long
    Dim copyErrDescription As String
    copyErrNumber = Err.Number
    copyErrDescription = Err.Description
    Resume CleanUpAfterError
CleanUpAfterError:
    On Error Resume Next
    If Not destRs Is Nothing Then destRs.Close
    If Not obsRs Is Nothing Then obsRs.Close
    If Not qdf Is Nothing Then qdf.Close
    Set destRs = Nothing
    Set obsRs = Nothing
    Set qdf = Nothing
    On Error GoTo 0
    Err.Raise copyErrNumber, , copyErrDescription
End Sub

Private Sub ValidateObservationCopyPair(ByVal sourceTable As String, ByVal destinationTable As String)
    If sourceTable = "TbObservaciones" And destinationTable = "TbObservacionesHistoricas" Then Exit Sub
    If sourceTable = "TbObservacionesHistoricas" And destinationTable = "TbObservaciones" Then Exit Sub
    Err.Raise 1000, , "Invalid Observaciones lifecycle copy table pair: " & sourceTable & " -> " & destinationTable
End Sub

Private Function ValidateLifecycleUserId(ByVal idUsuario As String) As Long
    If Len(Trim$(idUsuario)) = 0 Then Err.Raise 1000, , "Invalid lifecycle user ID: empty."
    If Not IsNumeric(idUsuario) Then Err.Raise 1000, , "Invalid lifecycle user ID: non-numeric value '" & idUsuario & "'."
    ValidateLifecycleUserId = CLng(idUsuario)
    If CStr(ValidateLifecycleUserId) <> Trim$(idUsuario) Then Err.Raise 1000, , "Invalid lifecycle user ID: non-Long value '" & idUsuario & "'."
End Function

Private Function GetNextObservationID(ByRef db As DAO.Database, ByVal tableName As String) As Long
    Dim rs As DAO.Recordset
    ValidateObservationTableName tableName
    GetNextObservationID = 1
    Set rs = db.OpenRecordset("SELECT MAX(IDObservacion) AS M FROM " & tableName, dbOpenSnapshot)
    If Not rs.EOF Then
        If Not IsNull(rs!M) Then GetNextObservationID = CLng(rs!M) + 1
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function ObservationIDExists(ByRef db As DAO.Database, ByVal tableName As String, ByVal idObservacion As Long) As Boolean
    Dim rs As DAO.Recordset
    ValidateObservationTableName tableName
    Set rs = db.OpenRecordset("SELECT IDObservacion FROM " & tableName & " WHERE IDObservacion=" & idObservacion, dbOpenSnapshot)
    ObservationIDExists = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Private Sub ValidateObservationTableName(ByVal tableName As String)
    If tableName = "TbObservaciones" Then Exit Sub
    If tableName = "TbObservacionesHistoricas" Then Exit Sub
    Err.Raise 1000, , "Invalid Observaciones table name: " & tableName
End Sub

Private Sub DeleteCurrentRows(ByRef db As DAO.Database, ByVal idUsuario As Long)
    db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & idUsuario, dbFailOnError
    db.Execute "DELETE FROM TbObservaciones WHERE ID=" & idUsuario, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & idUsuario, dbFailOnError
End Sub

Private Sub DeleteHistoricalRows(ByRef db As DAO.Database, ByVal idUsuario As Long)
    db.Execute "DELETE FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & idUsuario, dbFailOnError
    db.Execute "DELETE FROM TbObservacionesHistoricas WHERE ID=" & idUsuario, dbFailOnError
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & idUsuario, dbFailOnError
End Sub

Private Sub ResolvePaths(ByRef db As DAO.Database, ByVal idUsuario As Long, ByVal currentToHistorical As Boolean, _
                         ByVal currentRoot As String, ByVal historicalRoot As String, _
                         ByRef sourcePath As String, ByRef targetPath As String, ByRef p_Error As String)
    Dim DNI As String
    If currentToHistorical Then
        DNI = GetScalar(db, "SELECT DNI FROM TbUsuarios WHERE ID=" & idUsuario, p_Error)
    Else
        DNI = GetScalar(db, "SELECT DNI FROM TbUsuariosHistoricos WHERE ID=" & idUsuario, p_Error)
    End If
    If p_Error <> "" Then Exit Sub
    If DNI = "" Then
        p_Error = "Cannot resolve user DNI for lifecycle transition."
        Exit Sub
    End If
    ' [config-as-source-of-truth] Si los roots no se inyectaron,
    ' delegar en m_ObjEntorno.URLCarpetaAnexos* que ya implementa el
    ' selector APP_ROOT_REMOTO/APP_ROOT_LOCAL según accesibilidad real
    ' de disco. NO usa constantes hardcoded ni CurrentProject.path a
    ' ciegas: si TbConfiguracionHPS apunta a \\datoste\... y la red no
    ' responde, el resolver mismo falla con mensaje accionable.
    If currentRoot = "" Then
        currentRoot = m_ObjEntorno.URLCarpetaAnexos
    End If
    If historicalRoot = "" Then
        historicalRoot = m_ObjEntorno.URLCarpetaAnexosHistoricos
    End If
    ' [PR2 RFC norm — asymmetric since 2026-07-10] c2h target is FLAT at
    ' <DNI>\<file> (no inner subfolders — see the norm block at
    ' CopyFolderIfExists). h2c target is <DNI>\ and the per-file routing
    ' into <DNI>\Actual\ or <DNI>\HISTORICO\ happens inside
    ' ExecuteLifecycleMove via CopyFolderByEsHistorico, keyed by the
    ' EsHistorico column on TbAnexosUsuariosHistoricos.
    If currentToHistorical Then
        sourcePath = EnsureSlash(currentRoot) & DNI
        targetPath = EnsureSlash(historicalRoot) & DNI
    Else
        sourcePath = EnsureSlash(historicalRoot) & DNI
        targetPath = EnsureSlash(currentRoot) & DNI
    End If
End Sub

Private Function ValidateInjectedRoots(ByVal currentRoot As String, ByVal historicalRoot As String, ByRef p_Error As String) As Boolean
    ValidateInjectedRoots = False
    If IsUnsafeInjectedRoot(currentRoot) Or IsUnsafeInjectedRoot(historicalRoot) Then
        p_Error = "Unsafe attachment root for lifecycle transaction tests. Use a local %TEMP% path."
        Exit Function
    End If
    ValidateInjectedRoots = True
End Function

Private Function IsUnsafeInjectedRoot(ByVal rootPath As String) As Boolean
    If rootPath = "" Then Exit Function
    IsUnsafeInjectedRoot = (Left$(rootPath, 2) = "\\")
End Function

Private Function GetScalar(ByRef db As DAO.Database, ByVal sql As String, ByRef p_Error As String) As String
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then GetScalar = Nz(rs.Fields(0).value, "")
CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
EH:
    p_Error = "GetScalar: " & Err.Description
    Resume CleanUp
End Function

Private Sub ForceFailure(ByVal actualStage As String, ByVal expectedStage As String)
    If StrComp(actualStage, expectedStage, vbTextCompare) = 0 Then Err.Raise 1000, , "Forced lifecycle failure at " & expectedStage
End Sub

Private Function BuildSnapshotPath(ByVal idUsuario As String) As String
    BuildSnapshotPath = EnsureSlash(Environ$("TEMP")) & "HPS\LifecycleSnapshots\" & idUsuario & "-" & Format$(Now, "yyyymmddhhnnss")
End Function

Private Sub RestoreSnapshot(ByVal snapshotPath As String, ByVal sourcePath As String)
    If snapshotPath = "" Then Exit Sub
    If Dir$(snapshotPath, vbDirectory) = "" Then Exit Sub
    DeleteFolder sourcePath
    CopyFolderIfExists snapshotPath, sourcePath
End Sub

Private Sub SnapshotPriorConflictFolders(ByVal p_Direction As String, ByVal p_Conflicts As String, _
                                        ByVal p_CurrentRoot As String, ByVal p_HistoricalRoot As String, _
                                        ByVal p_SnapshotRoot As String)
    If Len(p_Conflicts) = 0 Then Exit Sub
    If Len(p_SnapshotRoot) = 0 Then Exit Sub
    Dim parts() As String
    parts = Split(p_Conflicts, "|")
    Dim part As Variant
    Dim folderIndex As Long
    For Each part In parts
        If Left$(CStr(part), 7) = "FOLDER=" Then
            folderIndex = folderIndex + 1
            Dim folderDni As String
            folderDni = Mid$(CStr(part), 8)
            Dim folderPath As String
            folderPath = ResolveTargetFolderPath(p_Direction, folderDni, p_CurrentRoot, p_HistoricalRoot)
            If Len(folderPath) > 0 Then
                If Dir$(folderPath, vbDirectory) <> "" Then
                    CopyFolderTreeIfExists folderPath, EnsureSlash(p_SnapshotRoot) & CStr(folderIndex)
                End If
            End If
        End If
    Next part
End Sub

Private Sub RestorePriorConflictFolders(ByVal p_Direction As String, ByVal p_Conflicts As String, _
                                       ByVal p_CurrentRoot As String, ByVal p_HistoricalRoot As String, _
                                       ByVal p_SnapshotRoot As String)
    If Len(p_Conflicts) = 0 Then Exit Sub
    If Len(p_SnapshotRoot) = 0 Then Exit Sub
    If Dir$(p_SnapshotRoot, vbDirectory) = "" Then Exit Sub
    Dim parts() As String
    parts = Split(p_Conflicts, "|")
    Dim part As Variant
    Dim folderIndex As Long
    For Each part In parts
        If Left$(CStr(part), 7) = "FOLDER=" Then
            folderIndex = folderIndex + 1
            Dim folderDni As String
            folderDni = Mid$(CStr(part), 8)
            Dim folderPath As String
            folderPath = ResolveTargetFolderPath(p_Direction, folderDni, p_CurrentRoot, p_HistoricalRoot)
            If Len(folderPath) > 0 Then
                Dim snapshotFolder As String
                snapshotFolder = EnsureSlash(p_SnapshotRoot) & CStr(folderIndex)
                If Dir$(snapshotFolder, vbDirectory) <> "" Then
                    DeleteFolder folderPath
                    CopyFolderTreeIfExists snapshotFolder, folderPath
                End If
            End If
        End If
    Next part
End Sub

' ============================================================================
' PR2 RFC NORM (locked in by user 2026-07-08, asymmetric since 2026-07-10)
'
' The lifecycle-path layout is ASYMMETRIC by direction:
'
'   c2h direction (current → historical):
'     target = <HISTORICO>\<DNI>\<file>, FLAT (no inner subfolders).
'     Every source file under <HPS>\<DNI>\Actual\<file> +
'     <HPS>\<DNI>\HISTORICO\<file> is flattened into a single
'     <HISTORICO>\<DNI>\<file>. The per-anexo origin is preserved through
'     the EsHistorico column on TbAnexosUsuariosHistoricos (see
'     CopyCurrentRowsToHistorical) and consumed on h2c.
'
'   h2c direction (historical → current):
'     target = <HPS>\<DNI>\ and each flat source file at
'     <HISTORICO>\<DNI>\<file> is routed per-anexo to either
'     <HPS>\<DNI>\Actual\<file> (EsHistorico='No') or
'     <HPS>\<DNI>\HISTORICO\<file> (EsHistorico='Sí'). This is the
'     c2h→h2c reversibility guarantee introduced 2026-07-10.
'
' Rationale for the asymmetry: the c2h flatten exists because the
' historical cache regen path (cache_usuario_historico_regenerar in
' Mod_Sincronizacion_Historico.bas) looks for files at <DNI>\<file>
' directly — subfolders inside <DNI>\ would silently miss files and
' the "anexos faltantes" symptom would reappear in
' FormInicial08UsuariosHistoricosDatos. The h2c split exists because
' the active user side uses the per-anexo subfolder convention
' (Actual\/HISTORICO\) that was reintroduced in 9b54ac1 to keep the
' c2h→h2c cycle reversible. Both directions cooperate via the
' EsHistorico column on TbAnexosUsuariosHistoricos, which is the
' single source of truth for per-anexo routing on h2c.
'
' Enforcement:
'   - CopyFolderIfExists: walks source recursively and copies each file
'     flat to target, preserving no inner subfolder. Used by c2h.
'   - CopyFolderByEsHistorico: walks the flat source and routes each
'     file into <DNI>\Actual\ or <DNI>\HISTORICO\ based on EsHistorico.
'     Used by h2c.
'   - ResolvePaths: c2h target ends at <DNI>\<file> flat; h2c target ends
'     at <DNI>\ (the per-file split happens later in
'     ExecuteLifecycleMove).
'   - Test_HATW_LifecyclePaths_AreAlwaysFlat: locks BOTH halves of the
'     asymmetric norm — c2h strict-flat, h2c per-anexo split.
'
' If you change c2h back to non-flat, the cache regen will silently miss
' files. If you change h2c back to flat, the c2h→h2c cycle loses the
' per-anexo subfolder information and the reversibility guarantee breaks.
' The norm is locked in by the test — breaking the test breaks the norm.
' ============================================================================
Private Sub CopyFolderIfExists(ByVal sourcePath As String, ByVal targetPath As String)
    ' [PR2 RFC norm] Flatten copy: walk source recursively and copy every
    ' file directly into target. See the norm block above.
    If sourcePath = "" Or targetPath = "" Then Exit Sub
    If Dir$(sourcePath, vbDirectory) = "" Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ValidateFlatCopyBasenames fso.GetFolder(sourcePath)
    CreateParentFolder targetPath
    If Not fso.FolderExists(targetPath) Then fso.CreateFolder targetPath
    ValidateFlatCopyTargetBasenames fso.GetFolder(sourcePath), targetPath
    CopyAllFilesFlat fso, fso.GetFolder(sourcePath), targetPath
End Sub

Private Sub ValidateFlatCopyBasenames(ByVal srcFolder As Object)
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    CollectFlatCopyBasenames srcFolder, seen
End Sub

Private Sub CollectFlatCopyBasenames(ByVal srcFolder As Object, ByRef seen As Object)
    Dim file As Object
    For Each file In srcFolder.Files
        If seen.Exists(file.Name) Then
            Err.Raise 1000, , "Duplicate attachment filename would collide during flat copy: '" & file.Name & "' found at '" & seen(file.Name) & "' and '" & file.Path & "'. Move aborted before copying to prevent data loss."
        End If
        seen.Add file.Name, file.Path
    Next file
    Dim subFolder As Object
    For Each subFolder In srcFolder.Subfolders
        CollectFlatCopyBasenames subFolder, seen
    Next subFolder
End Sub

Private Sub ValidateFlatCopyTargetBasenames(ByVal srcFolder As Object, ByVal targetPath As String)
    If Len(targetPath) = 0 Then Exit Sub
    If Dir$(targetPath, vbDirectory) = "" Then Exit Sub
    Dim sourceBasenames As Object
    Set sourceBasenames = CreateObject("Scripting.Dictionary")
    sourceBasenames.CompareMode = vbTextCompare
    CollectFlatCopyBasenames srcFolder, sourceBasenames
    Dim targetFolder As Object
    Dim targetFile As Object
    Set targetFolder = CreateObject("Scripting.FileSystemObject").GetFolder(targetPath)
    For Each targetFile In targetFolder.Files
        If sourceBasenames.Exists(targetFile.Name) Then
            Err.Raise 1000, , "Target attachment filename would be overwritten during flat copy: '" & targetFile.Name & "' already exists at '" & targetFile.Path & "'. Move aborted before copying to prevent data loss."
        End If
    Next targetFile
End Sub

Private Sub CopyAllFilesFlat(ByRef fso As Object, ByVal srcFolder As Object, ByVal flatTarget As String)
    Dim file As Object
    For Each file In srcFolder.Files
        fso.CopyFile file.Path, flatTarget & "\" & file.Name, True
    Next file
    Dim subFolder As Object
    For Each subFolder In srcFolder.Subfolders
        CopyAllFilesFlat fso, subFolder, flatTarget
    Next subFolder
End Sub

Private Sub DeleteFolder(ByVal folderPath As String)
    If folderPath = "" Then Exit Sub
    If Dir$(folderPath, vbDirectory) <> "" Then CreateObject("Scripting.FileSystemObject").DeleteFolder folderPath, True
End Sub

Private Sub CopyFolderTreeIfExists(ByVal sourcePath As String, ByVal targetPath As String)
    If sourcePath = "" Or targetPath = "" Then Exit Sub
    If Dir$(sourcePath, vbDirectory) = "" Then Exit Sub
    CreateParentFolder targetPath
    CreateObject("Scripting.FileSystemObject").CopyFolder sourcePath, targetPath, True
End Sub

Private Sub CreateParentFolder(ByVal folderPath As String)
    Dim parentPath As String
    parentPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(folderPath)
    If parentPath <> "" Then CreateFolderTree parentPath
End Sub

Private Sub CreateFolderTree(ByVal folderPath As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    parts = Split(folderPath, "\")
    currentPath = parts(0) & "\"
    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            currentPath = currentPath & parts(i) & "\"
            If Dir$(currentPath, vbDirectory) = "" Then MkDir currentPath
        End If
    Next i
End Sub

Private Function EnsureSlash(ByVal folderPath As String) As String
    If Right$(folderPath, 1) = "\" Then
        EnsureSlash = folderPath
    Else
        EnsureSlash = folderPath & "\"
    End If
End Function

' [c2h-h2c-subcarpetas] h2c split-routing helpers. These two subs power
' the new h2c path that reverses the c2h flatten: each flat file at
' <HISTORICO>\<DNI>\<file> is routed to <HPS>\<DNI>\Actual\<file> (when
' EsHistorico='No') or <HPS>\<DNI>\HISTORICO\<file> (when EsHistorico='Sí').
' The PR2 RFC norm (FLAT PATH LAYOUT) at CopyFolderIfExists is preserved
' for c2h; h2c uses a different helper because the routing is per-file,
' not per-folder, and the destination layout has inner subfolders by
' design (the per-anexo convention introduced in 9b54ac1).
'
' Read EsHistorico from TbAnexosUsuariosHistoricos for the user BEFORE
' the row is deleted (the source row is consumed in DeleteHistoricalRows
' at the end of ExecuteLifecycleMove). Returns a Scripting.Dictionary
' keyed by NombreAnexo (case-insensitive) with value "No" / "Sí" (or
' whatever the source had; a NULL is coerced to "No" so the lookup
' never returns Empty).
Private Function BuildEsHistoricoByName(ByRef db As DAO.Database, ByVal idUsuario As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT NombreAnexo, EsHistorico FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & idUsuario, dbOpenSnapshot)
    Do While Not rs.EOF
        If Not dict.Exists(Nz(rs!NombreAnexo, "")) Then
            dict.Add Nz(rs!NombreAnexo, ""), Nz(rs!EsHistorico, "No")
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set BuildEsHistoricoByName = dict
End Function

' [c2h-h2c-subcarpetas] h2c per-anexo routing. Copies every file found
' under p_SourcePath to one of two destination roots, decided per-file
' by looking up the basename in p_EsHistoricoByName:
'   - value "No"  -> p_DestinationActualRoot\<file>
'   - value "Sí"  -> p_DestinationHistoricoRoot\<file>
'
' Source p_SourcePath is expected to be flat (c2h's flatten output), but
' the recursive walk also catches unexpected subfolders; the lookup is
' applied to every file in any depth, so a stray <HISTORICO>\<DNI>\nested\
' would still be routed correctly. Duplicates per subfolder are rejected
' to mirror the ValidateFlatCopyTargetBasenames behavior in c2h.
'
' Failure modes:
'   - source folder missing: silent (mirror CopyFolderIfExists; if the
'     flat tree was already gone, there's nothing to do).
'   - file in source has no row in the lookup dict: Err.Raise 1000. This
'     is a data-drift signal — the DB says the user has N anexos but the
'     flat folder has N+1 files. Compensate by aborting; the outer
'     ExecuteLifecycleMove rolls back.
Private Sub CopyFolderByEsHistorico(ByVal p_SourcePath As String, _
                                     ByVal p_DestinationActualRoot As String, _
                                     ByVal p_DestinationHistoricoRoot As String, _
                                     ByVal p_EsHistoricoByName As Object)
    If Len(p_SourcePath) = 0 Then Exit Sub
    If p_EsHistoricoByName Is Nothing Then Err.Raise 1000, , "CopyFolderByEsHistorico: EsHistorico lookup is Nothing; h2c cannot route per-anexo."
    If Dir$(p_SourcePath, vbDirectory) = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcFolder As Object
    Set srcFolder = fso.GetFolder(p_SourcePath)

    ' Reject duplicate basenames BEFORE copying (mirror
    ' ValidateFlatCopyBasenames in CopyFolderIfExists).
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    CollectFlatCopyBasenames srcFolder, seen

    ' Reject per-subfolder overwrite on both destinations (mirror
    ' ValidateFlatCopyTargetBasenames split per folder).
    ValidateSubcarpetasCopyTargetBasenames fso, p_DestinationActualRoot, seen
    ValidateSubcarpetasCopyTargetBasenames fso, p_DestinationHistoricoRoot, seen

    ' Make sure both destination roots exist before we start writing.
    CreateFolderTree p_DestinationActualRoot
    CreateFolderTree p_DestinationHistoricoRoot

    ' Walk the source recursively and copy each file to the matching root.
    CopyAllFilesByEsHistorico fso, srcFolder, p_DestinationActualRoot, p_DestinationHistoricoRoot, p_EsHistoricoByName
End Sub

Private Sub ValidateSubcarpetasCopyTargetBasenames(ByVal fso As Object, ByVal p_DestinationRoot As String, ByVal p_AlreadyUsed As Object)
    If Len(p_DestinationRoot) = 0 Then Exit Sub
    If Dir$(p_DestinationRoot, vbDirectory) = "" Then Exit Sub
    Dim targetFolder As Object
    Set targetFolder = fso.GetFolder(p_DestinationRoot)
    Dim targetFile As Object
    For Each targetFile In targetFolder.Files
        If p_AlreadyUsed.Exists(targetFile.Name) Then
            Err.Raise 1000, , "Target attachment filename would be overwritten during per-anexo split copy: '" & targetFile.Name & "' already exists at '" & targetFile.Path & "'. Move aborted before copying to prevent data loss."
        End If
    Next targetFile
End Sub

Private Sub CopyAllFilesByEsHistorico(ByVal fso As Object, ByVal srcFolder As Object, _
                                       ByVal p_DestinationActualRoot As String, _
                                       ByVal p_DestinationHistoricoRoot As String, _
                                       ByVal p_EsHistoricoByName As Object)
    Dim file As Object
    Dim esHistorico As String
    Dim targetSubfolder As String
    For Each file In srcFolder.Files
        If Not p_EsHistoricoByName.Exists(file.Name) Then
            Err.Raise 1000, , "CopyFolderByEsHistorico: source file '" & file.Name & "' has no EsHistorico row in TbAnexosUsuariosHistoricos (data drift between filesystem and DB). h2c aborted to prevent a misplaced copy."
        End If
        esHistorico = CStr(p_EsHistoricoByName(file.Name))
        If StrComp(esHistorico, "Sí", vbTextCompare) = 0 Then
            targetSubfolder = p_DestinationHistoricoRoot
        Else
            ' Default 'No' (or any value that is not 'Sí') routes to Actual\.
            ' This matches the per-anexo convention: a 'No' anexo belongs in
            ' Actual\, a 'Sí' one in HISTORICO\. Legacy NULLs are coerced to
            ' 'No' in BuildEsHistoricoByName so they land in Actual\.
            targetSubfolder = p_DestinationActualRoot
        End If
        fso.CopyFile file.Path, targetSubfolder & file.Name, True
    Next file
    Dim subFolder As Object
    For Each subFolder In srcFolder.Subfolders
        CopyAllFilesByEsHistorico fso, subFolder, p_DestinationActualRoot, p_DestinationHistoricoRoot, p_EsHistoricoByName
    Next subFolder
End Sub

' Garantiza que la estructura anexos de un usuario activo este completa:
' tanto Actual\ como HISTORICO\ existen (vacias si no hay archivos).
' Llamada post-h2c desde ExecuteLifecycleMove para que el folder del
' usuario activo tenga siempre la forma esperada, independientemente de
' cuantos anexos cayeron en cada subcarpeta.
Private Sub EnsureAnexosSubcarpetasStructure(ByVal p_DestinationRoot As String)
    Dim fso As Object
    On Error GoTo errores
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p_DestinationRoot & "Actual") Then
        fso.CreateFolder p_DestinationRoot & "Actual"
    End If
    If Not fso.FolderExists(p_DestinationRoot & "HISTORICO") Then
        fso.CreateFolder p_DestinationRoot & "HISTORICO"
    End If
    Exit Sub
errores:
    Err.Raise Err.Number, , "EnsureAnexosSubcarpetasStructure: " & Err.Description
End Sub

' =====================================================================
' PR1a — Coordinator seam (HATW R-precheck / R-toctou / R-cleanup /
' R-IDAnexo / R-error-translator). All methods are minimal-invasive
' additions; existing helpers are untouched.
' =====================================================================

' Task 1.2 — Public seam: register a late-bound prompt procedure name.
' The registered proc is called via Application.Run(...) when precheck #1
' detects conflicts; the proc must return Boolean (True=approved cleanup).
Public Sub ConfigureLifecycleConflictPrompt(ByVal p_PromptProcName As String)
    m_PromptProcName = p_PromptProcName
End Sub

' Task 1.3 — Precheck #1: query target table for DNI conflict AND target
' folder for folder conflict. Returns "" if clean, otherwise pipe-delimited
' tokens in the form "BD=<dni>" and/or "FOLDER=<id>". The query joins on
' DNI (NOT ID) because the conflict is "another row with the same DNI in
' the destination table" — a row at a different ID but same DNI is a
' logical collision that the copy-insert would surface as DAO 3022. DAO
' errors are translated via TranslateLifecycleError before re-raising so
' the EH path in ExecuteLifecycleMove gets a clean message (and the
' Test_HATW_Precheck_PropagatesDBError atom can assert the format).
'   p_Direction = "c2h" -> source=TbUsuarios / target=TbUsuariosHistoricos
'   p_Direction = "h2c" -> source=TbUsuariosHistoricos / target=TbUsuarios
Private Function CheckLifecycleConflicts(ByRef db As DAO.Database, ByVal p_Direction As String, ByVal p_UserId As Long, _
                                         Optional ByVal p_CurrentRoot As String = "", _
                                         Optional ByVal p_HistoricalRoot As String = "") As String
    On Error GoTo DBError
    LogToFile "[CheckLifecycleConflicts] direction=" & p_Direction & " userId=" & p_UserId
    Dim targetTable As String
    Dim sourceTable As String
    Dim sourceDni As String
    Dim conflicts As String
    Dim rs As DAO.Recordset
    If p_Direction = DIRECTION_C2H Then
        sourceTable = "TbUsuarios"
        targetTable = "TbUsuariosHistoricos"
    Else
        sourceTable = "TbUsuariosHistoricos"
        targetTable = "TbUsuarios"
    End If
    ' Step 1: get the DNI of the user being moved (from the source table)
    sourceDni = GetScalar(db, "SELECT DNI FROM " & sourceTable & " WHERE ID=" & p_UserId, "")
    If Len(sourceDni) = 0 Then
        ' Source user does not exist (or has no DNI); no conflict possible
        CheckLifecycleConflicts = ""
        Exit Function
    End If
    ' Step 2: any row in the target table with the same DNI is a conflict
    ' [PR1a TOCTOU] — match the sourceDNI itself, OR a row whose DNI starts
    ' with the sourceDNI (this covers the "_TOCTOU" / "_TOCTOU_H2C" suffixed
    ' rows the test prompt functions inject to simulate a TOCTOU race
    ' without colliding with the UNIQUE index on DNI). precheck#1 and
    ' precheck#2 use the same query, so a row injected between them is
    ' counted by the re-read.
    '
    ' Jet/ACE wildcard note: in DAO the LIKE wildcard is `*`, NOT `%`.
    ' `%` is an ANSI SQL convention that Jet only honors through the OLE
    ' DB/ODBC providers; the native DAO/Jet engine treats `%` as a literal
    ' character. Earlier code used `LIKE 'X_TOCTOU*'` and it silently
    ' returned 0 because Jet was searching for the literal substring,
    ' not a prefix. The `*` is what Jet needs.
    Dim toctouQuery As String
    toctouQuery = "SELECT ID, DNI FROM " & targetTable & " WHERE DNI LIKE '" & Replace(sourceDni, "'", "''") & "*'"
    LogToFile "[CheckLifecycleConflicts] query=" & toctouQuery
    Set rs = db.OpenRecordset(toctouQuery, dbOpenSnapshot)
    Dim rowCount As Long
    rowCount = 0
    Do While Not rs.EOF
        rowCount = rowCount + 1
        LogToFile "[CheckLifecycleConflicts] row #" & rowCount & " ID=" & rs!ID & " DNI=" & Nz(rs!DNI, "")
        If Len(conflicts) > 0 Then conflicts = conflicts & "|"
        conflicts = conflicts & "BD=" & Nz(rs!DNI, "")
        rs.MoveNext
    Loop
    LogToFile "[CheckLifecycleConflicts] total DB rows found=" & rowCount
    rs.Close
    Set rs = Nothing
    On Error GoTo 0

    ' Step 3: check for folder conflict at the destination path. The folder
    ' path is keyed by DNI (not userId) because that is what ResolvePaths
    ' uses in ExecuteLifecycleMove — a folder named "<root>\DNI<x>" matches.
    ' The FOLDER= token also carries the DNI (not userId) so DeletePriorConflictFolders
    ' can rebuild the path without an extra argument.
    Dim folderPath As String
    folderPath = ResolveTargetFolderPath(p_Direction, sourceDni, p_CurrentRoot, p_HistoricalRoot)
    If Len(folderPath) > 0 Then
        If Dir$(folderPath, vbDirectory) <> "" Then
            If Len(conflicts) > 0 Then conflicts = conflicts & "|"
            conflicts = conflicts & "FOLDER=" & sourceDni
        End If
    End If

    LogToFile "[CheckLifecycleConflicts] result conflicts=" & conflicts
    CheckLifecycleConflicts = conflicts
    Exit Function
DBError:
    LogToFile "[CheckLifecycleConflicts] DBError err=" & Err.Number & " desc=" & Err.Description
    Err.Raise Err.Number, , TranslateLifecycleError("CheckLifecycleConflicts", targetTable, Null, "ID=" & p_UserId)
End Function

' Task 1.4 — Resolve a free target ID. Returns p_CandidateId if it is
' strictly greater than MAX(target.ID); otherwise MAX(target.ID)+1.
' Idempotent: calling twice with the same DB state returns the same ID.
' Signature aligned with PR1b test contracts: (db, candidateId, direction, error)
' — matches the call sites in Test_HistoricoAdjuntosTransactionWrapper.bas where
' the test passes the candidate Long first, then the direction String, and an
' errMsg ByRef for the HPS convention (Optional ByRef p_Error As String).
Public Function ResolveTargetUserID(ByRef db As DAO.Database, _
                                    ByVal p_CandidateId As Long, _
                                    ByVal p_Direction As String, _
                                    Optional ByRef p_Error As String) As Long
    Dim targetTable As String
    Dim rs As DAO.Recordset
    Dim maxId As Long
    On Error GoTo EH
    p_Error = ""
    If p_Direction = DIRECTION_C2H Then
        targetTable = "TbUsuariosHistoricos"
    Else
        targetTable = "TbUsuarios"
    End If
    Set rs = db.OpenRecordset("SELECT MAX(ID) AS MaxId FROM " & targetTable, dbOpenSnapshot)
    If rs.EOF Then
        maxId = 0
    ElseIf IsNull(rs!maxId) Then
        maxId = 0
    Else
        maxId = CLng(rs!maxId)
    End If
    rs.Close
    Set rs = Nothing
    If maxId < p_CandidateId Then
        ResolveTargetUserID = p_CandidateId
    Else
        ResolveTargetUserID = maxId + 1
    End If
    Exit Function
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If p_Error = "" Then p_Error = "ResolveTargetUserID: " & Err.Description
    ResolveTargetUserID = 0
End Function

' Task 1.5 — DAO error translator. Maps DAO errors to a contextual message that
' names the method, table, attempted ID value, and column. Signature aligned
' with PR1b Test_HATW_ErrorTranslator_DuplicateKey_MentionsTableAndID and
' Test_HATW_ErrorTranslator_UnknownError_PreservesOriginalMessage, which invoke
' the method as (p_MethodName, p_TableName, p_IDValue, p_ColumnName) and assert
' each piece appears in the returned message (or appears alone when the lookup
' is unknown). Unknown / passthrough errors still preserve the method name
' so the 8 baseline atoms that assert on Err.Description verbatim keep working.
Public Function TranslateLifecycleError(ByVal p_MethodName As String, _
                                       ByVal p_TableName As String, _
                                       ByVal p_IDValue As Variant, _
                                       ByVal p_ColumnName As String) As String
    Dim pieces As String
    pieces = p_MethodName
    If Len(p_TableName) > 0 Then
        pieces = pieces & "[" & p_TableName & "]"
    End If
    If Not IsNull(p_IDValue) Then
        If Len(p_ColumnName) > 0 Then
            pieces = pieces & " " & p_ColumnName & "=" & CStr(p_IDValue)
        Else
            pieces = pieces & " ID=" & CStr(p_IDValue)
        End If
    End If
    TranslateLifecycleError = "Error in " & pieces
End Function

' Task 1.6a — Cleanup helper: parse p_Conflicts and delete every BD=<dni>
' row from the destination table AND its related child tables (Observaciones,
' Anexos). Runs INSIDE the ws.BeginTrans opened by ExecuteLifecycleMove, so
' a DAO failure propagates as Err.Raise and the outer EH path rolls the
' transaction back. Folder deletion is handled by DeletePriorConflictFolders.
'
' The related-table deletes are required because the user table has FK
' dependents that would otherwise surface as DAO 3022 (duplicate key) on
' the copy-insert step. Example: the Observaciones fixture pre-inserts a
' row with the same ID as the source user's observation; the copy-insert
' would hit a PK collision unless the existing Observaciones row is
' cleaned first. CASCADE DELETE on the user→Observaciones FK is NOT
' enough because it would only fire AFTER the parent delete, and the
' precheck-cleanup order is "free the target ID space before the
' copy-insert" (per design.md §A1 step 8).
Private Sub DeletePriorConflictRows(ByRef db As DAO.Database, ByVal p_Direction As String, ByVal p_Conflicts As String)
    If Len(p_Conflicts) = 0 Then Exit Sub
    Dim usersTable As String
    Dim obsTable As String
    Dim anexosTable As String
    If p_Direction = DIRECTION_C2H Then
        usersTable = "TbUsuariosHistoricos"
        obsTable = "TbObservacionesHistoricas"
        anexosTable = "TbAnexosUsuariosHistoricos"
    Else
        usersTable = "TbUsuarios"
        obsTable = "TbObservaciones"
        anexosTable = "TbAnexosUsuariosHPS"
    End If
    Dim parts() As String
    parts = Split(p_Conflicts, "|")
    Dim part As Variant
    For Each part In parts
        If Left$(CStr(part), 3) = "BD=" Then
            Dim dni As String
            dni = Mid$(CStr(part), 4)
            ' Order: child tables first (Observaciones + Anexos), then user table.
            ' Child-first avoids any transient FK constraint violation during cleanup.
            db.Execute "DELETE FROM " & obsTable & " WHERE ID IN (SELECT ID FROM " & usersTable & " WHERE DNI='" & Replace(dni, "'", "''") & "')", dbFailOnError
            db.Execute "DELETE FROM " & anexosTable & " WHERE IDUsuario IN (SELECT ID FROM " & usersTable & " WHERE DNI='" & Replace(dni, "'", "''") & "')", dbFailOnError
            db.Execute "DELETE FROM " & usersTable & " WHERE DNI='" & Replace(dni, "'", "''") & "'", dbFailOnError
        End If
    Next part
End Sub

' Task 1.6b — Cleanup helper: parse p_Conflicts and delete every
' FOLDER=<id> folder tree in the destination. Uses the same root
' defaults as ResolvePaths so the integration in ExecuteLifecycleMove
' remains consistent.
Private Sub DeletePriorConflictFolders(ByVal p_Direction As String, ByVal p_Conflicts As String, _
                                       Optional ByVal p_CurrentRoot As String = "", _
                                       Optional ByVal p_HistoricalRoot As String = "")
    If Len(p_Conflicts) = 0 Then Exit Sub
    Dim parts() As String
    parts = Split(p_Conflicts, "|")
    Dim part As Variant
    For Each part In parts
        If Left$(CStr(part), 7) = "FOLDER=" Then
            ' The FOLDER= token carries the DNI (not the userId) — see
            ' CheckLifecycleConflicts where it is emitted. ResolveTargetFolderPath
            ' expects the DNI as its second arg.
            Dim folderDni As String
            folderDni = Mid$(CStr(part), 8)
            Dim folderPath As String
            folderPath = ResolveTargetFolderPath(p_Direction, folderDni, p_CurrentRoot, p_HistoricalRoot)
            If Len(folderPath) > 0 Then
                If Dir$(folderPath, vbDirectory) <> "" Then DeleteFolder folderPath
            End If
        End If
    Next part
End Sub

' Resolve the destination folder path for a given (direction, dni).
' Defaults match ResolvePaths (DEFAULT_ATTACHMENTS_ROOT_*). For h2c the
' historical destination is actually the current user folder (lifecycle
' transition moves users BACK to the current side), so the path mirrors
' the "c2h" ResolvePaths source/target convention in reverse.
' The folder name is the DNI (not the userId) because ResolvePaths keys
' its sourcePath/targetPath on the source user's DNI, and the conflict
' folder fixtures (PrepareFolderConflictFixture et al.) create
' "<root>\DNI<id>" directories.
Private Function ResolveTargetFolderPath(ByVal p_Direction As String, ByVal p_DNI As String, _
                                         ByVal p_CurrentRoot As String, ByVal p_HistoricalRoot As String) As String
    Dim currentRoot As String
    Dim historicalRoot As String
    If Len(p_CurrentRoot) > 0 Then
        currentRoot = p_CurrentRoot
    Else
        currentRoot = DEFAULT_ATTACHMENTS_ROOT_ACTUAL
    End If
    If Len(p_HistoricalRoot) > 0 Then
        historicalRoot = p_HistoricalRoot
    Else
        historicalRoot = DEFAULT_ATTACHMENTS_ROOT_HISTORICO
    End If
    If p_Direction = DIRECTION_C2H Then
        ResolveTargetFolderPath = EnsureSlash(historicalRoot) & p_DNI
    Else
        ResolveTargetFolderPath = EnsureSlash(currentRoot) & p_DNI
    End If
End Function

' Token-count helper for the R-toctou precheck #2 counter-compare.
Private Function CountConflictTokens(ByVal p_Conflicts As String) As Long
    If Len(p_Conflicts) = 0 Then
        CountConflictTokens = 0
        Exit Function
    End If
    CountConflictTokens = UBound(Split(p_Conflicts, "|")) + 1
End Function

' Extract the first BD=<dni> token from a conflicts string. Used by the
' prompt invocation in ExecuteLifecycleMove to surface the conflicting
' DNI as the second argument to the registered prompt procedure. Returns
' "" when no BD= token is present.
Private Function ExtractFirstDNI(ByVal p_Conflicts As String) As String
    If Len(p_Conflicts) = 0 Then Exit Function
    Dim parts() As String
    parts = Split(p_Conflicts, "|")
    Dim part As Variant
    For Each part In parts
        If Left$(CStr(part), 3) = "BD=" Then
            ExtractFirstDNI = Mid$(CStr(part), 4)
            Exit Function
        End If
    Next part
End Function


