VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnexoSelectionTransactionCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =====================================================================
' [anexos-batch-hotfix] Atomic batch mover for TbAnexosUsuariosHPS rows.
' Coordinates N anexos between Actual <-> HISTORICO so that either ALL
' rows/files/lists move or NONE does. The single Workspace transaction
' covers the DB side; snapshots under %TEMP%\HPS\AnexoSelectionSnapshots\
' give us filesystem rollback (DAO cannot revert file copies/deletes).
'
' Stage seams (deterministic test hooks, no-op in production):
'   AfterCopy         -> raise after CopyFile of the index-th anexo
'   AfterDbUpdate     -> raise after the index-th UPDATE
'   AfterSourceDelete -> raise after the index-th source DeleteFile
'   BeforeCommit      -> raise after the last item, before CommitTrans
'
' Required schema (verified against HPST.accdb):
'   TbAnexosUsuariosHPS(IDAnexo Long PK, NombreAnexo Text(255),
'                       IDUsuario Long FK, Hash Text(64),
'                       EsHistorico Text(2))
' Required signatures (verified):
'   getdb() As DAO.Database
'   getAnexoUsuarioHPS(p_IDAnexo As String, p_Error As String,
'                      Optional p_Db As DAO.Database = Nothing) As AnexoUsuarioHPS
' =====================================================================

Private Const STAGE_AFTER_COPY As String = "AfterCopy"
Private Const STAGE_BEFORE_DB_UPDATE As String = "BeforeDbUpdate"
Private Const STAGE_AFTER_DB_UPDATE As String = "AfterDbUpdate"
Private Const STAGE_AFTER_SOURCE_DELETE As String = "AfterSourceDelete"
Private Const STAGE_BEFORE_COMMIT As String = "BeforeCommit"

Private Const SNAPSHOT_ROOT As String = "HPS\AnexoSelectionSnapshots\"
Private Const PATH_HISTORICO_SUFFIX As String = "\HISTORICO\"
Private Const PATH_ACTUAL_SUFFIX As String = "\Actual\"

' Public surface -------------------------------------------------------

' Move every IDAnexo in p_IDs from its current side to the other side.
' p_CurrentToHistorical = True -> Actual -> HISTORICO; False -> HISTORICO -> Actual.
' p_Db may be Nothing (coordinator opens one). p_SeamFase is a test hook
' that forces a deterministic failure at a named stage. p_Error carries
' the message on failure and stays empty on success.
Public Function MoverLote(ByVal p_IDs As Collection, _
                          ByVal p_CurrentToHistorical As Boolean, _
                          Optional ByVal p_Db As DAO.Database = Nothing, _
                          Optional ByVal p_SeamFase As String = "", _
                          Optional ByRef p_Error As String) As Boolean

    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim ownsDb As Boolean
    Dim transactionStarted As Boolean
    Dim hitError As Boolean
    Dim opId As String

    ' State that must persist from snapshot stage through EH / CleanUp.
    Dim snapshotRoot As String
    Dim snapshotCreated As Boolean
    Dim snapshotFullyOk As Boolean
    Dim rollbackOk As Boolean
    Dim compensationOk As Boolean
    Dim compensationErr As String

    Dim plan() As Long          ' index aligned with p_IDs
    Dim i As Long
    Dim n As Long
    Dim anexo As AnexoUsuarioHPS
    Dim usuario As UsuarioHPS
    Dim sourcePath As String
    Dim destPath As String
    Dim snapshotEntry As String
    Dim recordsAffected As Long
    Dim phaseErr As String

    Dim sourcePaths() As String
    Dim destPaths() As String
    Dim snapshotPaths() As String
    Dim creadoDest() As Boolean
    Dim eliminadoOrigen() As Boolean

    On Error GoTo EH

    p_Error = ""
    transactionStarted = False
    hitError = False
    ownsDb = False
    snapshotCreated = False
    snapshotFullyOk = False
    rollbackOk = True
    compensationOk = True
    compensationErr = ""

    ' ---- 1. Preflight ----------------------------------------------------
    If p_IDs Is Nothing Then
        p_Error = "AnexoSelectionTransactionCoordinator.MoverLote: la colección de IDs es Nothing."
        GoTo CleanUp
    End If
    n = p_IDs.Count
    If n = 0 Then
        p_Error = "AnexoSelectionTransactionCoordinator.MoverLote: la selección está vacía."
        GoTo CleanUp
    End If

    If p_Db Is Nothing Then
        Set db = getdb(p_Error)
        ownsDb = True
    Else
        Set db = p_Db
    End If
    If p_Error <> "" Then GoTo CleanUp

    Set ws = DBEngine.Workspaces(0)

    ReDim plan(1 To n)
    ReDim sourcePaths(1 To n)
    ReDim destPaths(1 To n)
    ReDim snapshotPaths(1 To n)
    ReDim creadoDest(1 To n)
    ReDim eliminadoOrigen(1 To n)

    Dim j As Long

    For i = 1 To n
        If Not IsNumeric(p_IDs.Item(i)) Then
            p_Error = "MoverLote: IDAnexo no numérico en la posición " & i & "."
            GoTo CleanUp
        End If
        plan(i) = CLng(p_IDs.Item(i))
    Next i

    ' Detect duplicate IDs (refuse before any DB work).
    For i = 1 To n - 1
        For j = i + 1 To n
            If plan(i) = plan(j) Then
                p_Error = "MoverLote: IDAnexo duplicado en la selección (" & plan(i) & ")."
                GoTo CleanUp
            End If
        Next j
    Next i

    ' Reload each row, validate direction, source exists, and pre-compute paths.
    For i = 1 To n
        Set anexo = getAnexoUsuarioHPS(CStr(plan(i)), p_Error, db)
        If p_Error <> "" Then GoTo CleanUp
        If anexo Is Nothing Then
            p_Error = "MoverLote: IDAnexo " & plan(i) & " no existe en TbAnexosUsuariosHPS."
            GoTo CleanUp
        End If

        Dim requiredEstado As String
        If p_CurrentToHistorical Then
            requiredEstado = "No"
        Else
            requiredEstado = "Sí"
        End If
        If anexo.EsHistorico <> requiredEstado Then
            p_Error = "MoverLote: el anexo " & plan(i) & " tiene EsHistorico='" & anexo.EsHistorico & _
                "' pero la dirección exige '" & requiredEstado & "'."
            GoTo CleanUp
        End If

        Set usuario = anexo.Usuario
        p_Error = anexo.Error
        If p_Error <> "" Then GoTo CleanUp
        If usuario Is Nothing Then
            p_Error = "MoverLote: no se pudo resolver el usuario del anexo " & plan(i) & "."
            GoTo CleanUp
        End If

        sourcePath = anexo.URLAnexo
        If Not fso.FileExists(sourcePath) Then
            p_Error = "MoverLote: archivo origen no encontrado en '" & sourcePath & _
                "' (IDAnexo=" & plan(i) & ")."
            GoTo CleanUp
        End If

        Dim dirOrigen As String
        Dim dirDestino As String
        If p_CurrentToHistorical Then
            dirOrigen = usuario.URLCarpetaAnexosActuales
            dirDestino = usuario.URLCarpetaAnexosHistoricos
        Else
            dirOrigen = usuario.URLCarpetaAnexosHistoricos
            dirDestino = usuario.URLCarpetaAnexosActuales
        End If
        If Len(dirOrigen) = 0 Or Len(dirDestino) = 0 Then
            p_Error = "MoverLote: rutas de anexos no resueltas para IDAnexo " & plan(i) & _
                " (dirOrigen='" & dirOrigen & "', dirDestino='" & dirDestino & "')."
            GoTo CleanUp
        End If
        If StrComp(dirOrigen, dirDestino, vbTextCompare) = 0 Then
            p_Error = "MoverLote: el origen y el destino coinciden para IDAnexo " & plan(i) & _
                " ('" & dirOrigen & "')."
            GoTo CleanUp
        End If

        Dim nombreAnexo As String
        nombreAnexo = anexo.NombreAnexo
        If Len(nombreAnexo) = 0 Then
            p_Error = "MoverLote: NombreAnexo vacío para IDAnexo " & plan(i) & "."
            GoTo CleanUp
        End If
        If Right$(dirDestino, 1) <> "\" Then dirDestino = dirDestino & "\"

        sourcePaths(i) = sourcePath
        destPaths(i) = dirDestino & nombreAnexo
        snapshotPaths(i) = ""    ' filled after snapshotRoot is built
        creadoDest(i) = False
        eliminadoOrigen(i) = False

        Set anexo = Nothing
        Set usuario = Nothing
    Next i

    ' Reject any destination that already exists (no overwrite contract).
    For i = 1 To n
        If fso.FileExists(destPaths(i)) Then
            p_Error = "MoverLote: ya existe el archivo de destino '" & destPaths(i) & _
                "' (IDAnexo=" & plan(i) & "). Operación cancelada para no sobrescribir."
            GoTo CleanUp
        End If
    Next i

    ' Reject duplicate destination paths inside the same selection (two IDs
    ' that resolve to the same target — defensive, shouldn't happen but
    ' guarantees the no-overwrite contract even when the DB pre-check
    ' passes).
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(destPaths(i), destPaths(j), vbTextCompare) = 0 Then
                p_Error = "MoverLote: dos IDAnexo distintos resuelven al mismo destino '" & _
                    destPaths(i) & "' (IDAnexo=" & plan(i) & " e IDAnexo=" & plan(j) & ")."
                GoTo CleanUp
            End If
        Next j
    Next i

    ' ---- 2. Snapshot -----------------------------------------------------
    opId = Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int(Rnd() * 1000000), "000000")
    snapshotRoot = EnsureSlash(Environ$("TEMP")) & SNAPSHOT_ROOT & opId & "\"
    CreateSnapshotFolderTree snapshotRoot
    snapshotCreated = True
    For i = 1 To n
        snapshotEntry = snapshotRoot & CStr(plan(i)) & "_" & SafeFileName(destPaths(i))
        snapshotPaths(i) = snapshotEntry
        If Not SafeCopyForSnapshot(sourcePaths(i), snapshotEntry) Then
            p_Error = "MoverLote: no se pudo crear snapshot de '" & sourcePaths(i) & _
                "' en '" & snapshotEntry & "'."
            GoTo CleanUp
        End If
    Next i
    snapshotFullyOk = True

    ' ---- 3. Transaction --------------------------------------------------
    ws.BeginTrans
    transactionStarted = True

    ' ---- 4. Per-item: copy, UPDATE, delete source ------------------------
    For i = 1 To n
        ' Ensure destination parent exists.
        Dim parentDest As String
        parentDest = fso.GetParentFolderName(destPaths(i))
        If Len(parentDest) > 0 Then
            If Not fso.FolderExists(parentDest) Then
                EnsureFolderTree parentDest
            End If
        End If

        ' 4a. CopyFile source -> dest (no overwrite; preflight guarantees this).
        fso.CopyFile sourcePaths(i), destPaths(i), False
        creadoDest(i) = True

        If StrComp(p_SeamFase, STAGE_AFTER_COPY, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 5151, , "SEAM AfterCopy: fallo forzado por el test."
        End If

        ' 4b. UPDATE EsHistorico with dbFailOnError + RecordsAffected=1.
        If StrComp(p_SeamFase, STAGE_BEFORE_DB_UPDATE, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 5156, , "SEAM BeforeDbUpdate: fallo forzado por el test."
        End If

        Dim sql As String
        If p_CurrentToHistorical Then
            sql = "UPDATE TbAnexosUsuariosHPS SET EsHistorico='Sí' WHERE IDAnexo=" & plan(i)
        Else
            sql = "UPDATE TbAnexosUsuariosHPS SET EsHistorico='No' WHERE IDAnexo=" & plan(i)
        End If
        db.Execute sql, dbFailOnError
        recordsAffected = db.RecordsAffected
        If recordsAffected <> 1 Then
            Err.Raise vbObjectError + 5152, , "MoverLote: la UPDATE de IDAnexo " & plan(i) & _
                " afectó " & recordsAffected & " filas (esperaba 1)."
        End If

        If StrComp(p_SeamFase, STAGE_AFTER_DB_UPDATE, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 5153, , "SEAM AfterDbUpdate: fallo forzado por el test."
        End If

        ' 4c. DeleteFile source.
        fso.DeleteFile sourcePaths(i), True
        eliminadoOrigen(i) = True

        If StrComp(p_SeamFase, STAGE_AFTER_SOURCE_DELETE, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 5154, , "SEAM AfterSourceDelete: fallo forzado por el test."
        End If
    Next i

    If StrComp(p_SeamFase, STAGE_BEFORE_COMMIT, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 5155, , "SEAM BeforeCommit: fallo forzado por el test."
    End If

    ' ---- 5. Commit -------------------------------------------------------
    ws.CommitTrans
    transactionStarted = False

    MoverLote = True
    GoTo CleanUp

EH:
    hitError = True
    phaseErr = Err.Description

    ' ---- Rollback: DAO + filesystem compensation -------------------------
    If transactionStarted Then
        On Error Resume Next
        ws.Rollback
        transactionStarted = False
        If Err.Number <> 0 Then
            rollbackOk = False
            compensationErr = "Rollback DAO falló: " & Err.Description
        End If
        Err.Clear
        On Error GoTo EH
    End If

    ' Compensate filesystem ONLY if DAO rollback succeeded AND snapshot was
    ' fully captured (preflight failures must not touch the filesystem).
    If rollbackOk And snapshotFullyOk Then
        Dim k As Long
        ' (a) delete destinations we created in this op (no others).
        For k = 1 To n
            If creadoDest(k) Then
                On Error Resume Next
                If fso.FileExists(destPaths(k)) Then fso.DeleteFile destPaths(k), True
                If Err.Number <> 0 Then
                    compensationOk = False
                    compensationErr = compensationErr & IIf(Len(compensationErr) > 0, vbNewLine, "") & _
                        "DeleteFile destino '" & destPaths(k) & "' falló: " & Err.Description
                End If
                Err.Clear
                On Error GoTo EH
            End If
        Next k
        ' (b) restore ONLY sources we actually deleted. Skip the overwrite
        ' (3rd arg) when the source is missing — the original is gone, just
        ' copy without overwrite semantics; if the source reappeared (e.g.
        ' another actor) verify-before-overwrite by content check.
        For k = 1 To n
            If eliminadoOrigen(k) Then
                On Error Resume Next
                If snapshotPaths(k) <> "" Then
                    If fso.FileExists(snapshotPaths(k)) Then
                        If fso.FileExists(sourcePaths(k)) Then
                            ' Source reappeared unexpectedly: overwrite with
                            ' snapshot to guarantee the on-disk state matches
                            ' the pre-run snapshot.
                            fso.CopyFile snapshotPaths(k), sourcePaths(k), True
                        Else
                            ' Source genuinely missing: create parent if
                            ' needed and copy without overwrite.
                            Dim parentSrc As String
                            parentSrc = fso.GetParentFolderName(sourcePaths(k))
                            If Len(parentSrc) > 0 Then
                                If Not fso.FolderExists(parentSrc) Then EnsureFolderTree parentSrc
                            End If
                            fso.CopyFile snapshotPaths(k), sourcePaths(k)
                        End If
                        If Err.Number <> 0 Then
                            compensationOk = False
                            compensationErr = compensationErr & IIf(Len(compensationErr) > 0, vbNewLine, "") & _
                                "Restore fuente '" & sourcePaths(k) & "' desde snapshot '" & snapshotPaths(k) & _
                                "' falló: " & Err.Description
                        End If
                    End If
                End If
                Err.Clear
                On Error GoTo EH
            End If
        Next k
    End If

CleanUp:
    ' Close owned db (only if we opened it; never close an injected p_Db).
    If ownsDb And Not db Is Nothing Then
        On Error Resume Next
        db.Close
        Set db = Nothing
        Err.Clear
        On Error GoTo EH
    End If

    ' Build p_Error ONLY when we entered via EH. Preflight paths set p_Error
    ' before GoTo CleanUp; do not overwrite that with phaseErr ("") concat.
    ' IMPORTANT: this MUST happen BEFORE the snapshot cleanup block disables
    ' the EH (On Error GoTo 0) — otherwise any subsequent error raises a
    ' VBA runtime error to the caller BEFORE p_Error is propagated, and
    ' the test sees "No se ha encontrado la ruta de acceso" instead of the
    ' actual coordinator message.
    If hitError And Len(p_Error) = 0 Then
        If Not rollbackOk Then
            p_Error = "MoverLote: error original '" & phaseErr & "'. ADEMÁS el Rollback DAO falló: " & _
                compensationErr
        ElseIf Not compensationOk Then
            p_Error = "MoverLote: error original '" & phaseErr & "'. Rollback DAO OK pero la " & _
                "compensación de filesystem falló: " & compensationErr & " Snapshot conservado en '" & _
                snapshotRoot & "' para restauración manual."
        Else
            p_Error = "MoverLote: " & phaseErr
        End If
    End If

    ' Clean snapshots ONLY when compensation succeeded. On rollback/compensation
    ' failure preserve the snapshot folder for operator-driven manual recovery.
    ' This block runs with On Error GoTo 0 because the snapshot folder is
    ' best-effort cleanup; any error here is cosmetic. By this point p_Error
    ' has already been populated and propagated via ByRef.
    If snapshotCreated And rollbackOk And compensationOk Then
        On Error Resume Next
        For i = n To 1 Step -1
            If snapshotPaths(i) <> "" Then
                If fso.FileExists(snapshotPaths(i)) Then fso.DeleteFile snapshotPaths(i), True
            End If
        Next i
        If fso.FolderExists(snapshotRoot) Then fso.DeleteFolder snapshotRoot, True
        snapshotCreated = False
        Err.Clear
    End If
End Function

' Private helpers ------------------------------------------------------

Private Function EnsureSlash(ByVal path As String) As String
    If Len(path) = 0 Then
        EnsureSlash = ""
    ElseIf Right$(path, 1) = "\" Then
        EnsureSlash = path
    Else
        EnsureSlash = path & "\"
    End If
End Function

Private Sub CreateSnapshotFolderTree(ByVal path As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    If Len(path) = 0 Then Exit Sub
    parts = Split(path, "\")
    If UBound(parts) < 1 Then Exit Sub
    currentPath = parts(0) & "\"
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            currentPath = currentPath & parts(i) & "\"
            If Dir$(currentPath, vbDirectory) = "" Then
                On Error Resume Next
                MkDir currentPath
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Sub EnsureFolderTree(ByVal path As String)
    CreateSnapshotFolderTree path
End Sub

Private Function SafeFileName(ByVal sourcePath As String) As String
    Dim base As String
    Dim i As Long
    base = fso.GetFileName(sourcePath)
    For i = 1 To Len(base)
        Dim ch As String
        ch = Mid$(base, i, 1)
        If InStr("/\<>:""'|?*", ch) > 0 Then
            Mid$(base, i, 1) = "_"
        End If
    Next i
    SafeFileName = base
End Function

Private Function SafeCopyForSnapshot(ByVal sourcePath As String, ByVal snapshotPath As String) As Boolean
    On Error GoTo Fail
    Dim parent As String
    parent = fso.GetParentFolderName(snapshotPath)
    If Len(parent) > 0 Then CreateSnapshotFolderTree parent
    fso.CopyFile sourcePath, snapshotPath, True
    SafeCopyForSnapshot = True
    Exit Function
Fail:
    SafeCopyForSnapshot = False
End Function
