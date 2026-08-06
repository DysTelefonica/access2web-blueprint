Attribute VB_Name = "Test_AnexoSelectionTransaction"
Option Compare Database
Option Explicit

' ============================================================
' ANEXOS-HISTORICOS hotfix tests for the ATOMIC BATCH coordinator
' (AnexoSelectionTransactionCoordinator.MoverLote). Wraps and
' replaces the per-row AnexoUsuarioHPS.PasarAnexo path with
' whole-selection atomicity: preflight -> snapshot -> single
' Workspace transaction -> per-item CopyFile/UPDATE/DeleteFile
' -> CommitTrans, with filesystem compensation via snapshot on
' any precommit failure.
'
' Test ID ranges (>= 900300, disjoint from Test_PerAnexoMove's
' 900201-900202):
'   C2H users start at 900300, H2C users start at 900340.
'   Anexo IDs derive from user ID + 1 / + 2.
' ============================================================

' Constants ----------------------------------------------------------
Private Const TEST_USER_C2H_BASE As Long = 900300
Private Const TEST_USER_H2C_BASE As Long = 900340
Private Const TEST_USER_REJECTS_BASE As Long = 900360
Private Const TEST_DNI_PREFIX As String = "DNI9003"
Private Const TEST_ROOT_SUBDIR As String = "AnexoSelectionTests"
Private Const TEST_CONFIG_KEY_HPS As String = "ANEXOS_HPS_PATH"
Private Const TEST_CONFIG_KEY_HIST As String = "ANEXOS_HISTORICO_PATH"

' Module-level state ------------------------------------------------
' Saves of globals that tests mutate, so the teardown can restore them.
Private m_SavedEntorno As entorno
Private m_bHadEntorno As Boolean
Private m_SavedEnOficina As EnumSiNo
Private m_SavedAccesoDatosTE As EnumSiNo
Private m_SavedDatosEnLocal As String
Private m_bHadDatosEnLocal As Boolean
' Snapshot of TbConfiguracionHPS rows BEFORE the test writes the Local paths,
' so the teardown can restore exact prior state. Dictionary(Clave -> row dict).
Private m_SavedConfigSnap As Object

' JSON wrappers -----------------------------------------------------

Private Function JsonOk(ByVal value As String, ByRef logs As Collection) As String
    JsonOk = "{""ok"":true,""value"":""" & EscapeJson(value) & """,""payload"":null,""error"":null,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function JsonFail(ByVal message As String, ByRef logs As Collection) As String
    JsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJson(message) & """,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function LogsJson(ByRef logs As Collection) As String
    Dim i As Long
    Dim result As String
    result = "["
    If Not logs Is Nothing Then
        For i = 1 To logs.Count
            If i > 1 Then result = result & ","
            result = result & """" & EscapeJson(CStr(logs(i))) & """"
        Next i
    End If
    LogsJson = result & "]"
End Function

Private Function EscapeJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", Chr$(92) & Chr$(34))
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    EscapeJson = value
End Function

' Path helpers ------------------------------------------------------

Private Function EnsureSlash(ByVal path As String) As String
    If Len(path) = 0 Then Exit Function
    If Right$(path, 1) = "\" Then
        EnsureSlash = path
    Else
        EnsureSlash = path & "\"
    End If
End Function

Private Function TempTestRoot(ByVal suffix As String) As String
    TempTestRoot = EnsureSlash(Environ$("TEMP")) & "HPS\" & TEST_ROOT_SUBDIR & "\" & suffix & "\"
End Function

Private Sub CreateFolderTree(ByVal path As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    If Len(path) = 0 Then Exit Sub
    parts = Split(path, "\")
    currentPath = parts(0) & "\"
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            currentPath = currentPath & parts(i) & "\"
            If Dir$(currentPath, vbDirectory) = "" Then MkDir currentPath
        End If
    Next i
End Sub

Private Sub DeleteFolderIfExists(ByVal path As String)
    If Len(path) = 0 Then Exit Sub
    Dim trimmed As String
    trimmed = path
    Do While Len(trimmed) > 0 And Right$(trimmed, 1) = "\"
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    Loop
    If Len(trimmed) = 0 Then Exit Sub
    If Dir$(trimmed, vbDirectory) <> "" Then
        CreateObject("Scripting.FileSystemObject").DeleteFolder trimmed, True
    End If
End Sub

Private Sub CreateTextFile(ByVal path As String, ByVal content As String)
    Dim parentPath As String
    parentPath = Left$(path, InStrRev(path, "\") - 1)
    CreateFolderTree parentPath
    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").CreateTextFile(path, True)
    ts.Write content
    ts.Close
End Sub

' Backend helpers ---------------------------------------------------

Private Function OpenLocalTestBackend(ByRef p_Error As String) As DAO.Database
    Dim backendPath As String
    Dim connectString As String
    On Error GoTo EH
    p_Error = ""
    ' [config-as-source-of-truth] Si el test setup borra TbConfiguracionHPS
    ' y luego no la siembra (escenario comun en la suite), sembramos la
    ' clave HPST_BACKEND_PATH con el path canónico de sandbox antes de
    ' resolver. Asi la regla "config gobierna" no requiere que el test
    ' setup siembre la config para que el test pueda abrir el backend.
    EnsureHPSTBackendKeySeeded
    backendPath = ResolveCanonicalBackendPath()
    If Dir$(backendPath) = "" Then
        p_Error = "Local test backend not found: " & backendPath
        Exit Function
    End If
    connectString = BackendConnectString()
    Set OpenLocalTestBackend = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, False, connectString)
    Exit Function
EH:
    p_Error = "OpenLocalTestBackend: " & Err.Description
End Function

' [config-as-source-of-truth] Helper: si HPST_BACKEND_PATH no está
' sembrado en TbConfiguracionHPS, lo siembra con el path canónico
' (C:\00repos\datos\HPST.accdb en sandbox). Se ejecuta ANTES de
' resolver el path para que el resolver encuentre la clave.
Private Sub EnsureHPSTBackendKeySeeded()
    Dim db As DAO.Database
    Dim sExisting As String
    Dim sCanonical As String
    On Error Resume Next
    Set db = CurrentDb
    sExisting = CStr(db.OpenRecordset( _
        "SELECT Valor FROM TbConfiguracionHPS WHERE Clave='HPST_BACKEND_PATH'", _
        dbOpenSnapshot).Fields("Valor").Value)
    On Error GoTo 0
    If Len(sExisting) > 0 Then Exit Sub
    sCanonical = "C:\00repos\datos\HPST.accdb"
    If Dir$(sCanonical) = "" Then
        ' Si el .accdb canónico no existe, sembramos CurrentProject.path
        ' y notificamos al operador.
        sCanonical = EnsureSlash(CurrentProject.path) & "HPST.accdb"
    End If
    On Error Resume Next
    db.Execute "INSERT INTO TbConfiguracionHPS (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('HPST_BACKEND_PATH', '" & Replace(sCanonical, "'", "''") & "', True, Now(), 'Test_AnexoSelection')", _
        dbFailOnError
    On Error GoTo 0
End Sub

Private Function ResolveCanonicalBackendPath() As String
    Dim backendPath As String
    On Error Resume Next
    backendPath = CStr(CurrentDb.OpenRecordset( _
        "SELECT Valor FROM TbConfiguracionHPS WHERE Clave='HPST_BACKEND_PATH'", _
        dbOpenSnapshot).Fields("Valor").Value)
    On Error GoTo 0
    If Len(backendPath) > 0 Then
        ResolveCanonicalBackendPath = backendPath
        Exit Function
    End If
    ResolveCanonicalBackendPath = EnsureSlash(CurrentProject.path) & "HPST.accdb"
End Function

Private Function BackendConnectString() As String
    Dim password As String
    password = Environ$("DYSFLOW_BACKEND_PASSWORD")
    If Len(password) = 0 Then password = Environ$("HPS_BACKEND_PASSWORD")
    If Len(password) = 0 Then password = Environ$("ACCESS_VBA_PASSWORD")
    If Len(password) = 0 Then
        Dim path As String
        path = EnsureSlash(CurrentProject.path) & ".dysflow\backend.pwd"
        If Dir$(path) <> "" Then password = Trim$(CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 1, False).ReadAll)
    End If
    If Len(password) = 0 Then Err.Raise 1000, , "BackendConnectString: env password or .dysflow\backend.pwd required for AnexoSelection tests."
    BackendConnectString = "MS Access;PWD=" & password
End Function

' Entorno / config setup -------------------------------------------

' Save the current entorno + flags + DatosEnLocal TempVar so the teardown
' can restore them. Builds a fresh entorno pinned to p_Root so the
' coordinator's path resolution lands under <p_Root>\<DNI>\Actual\ /
' <p_Root>\<DNI>\HISTORICO\. TbConfiguracionHPS lives in the FRONTEND
' (HPS.accdb / CurrentDb), not in the backend p_Db — we snapshot the
' full table state up-front and restore exact rows on teardown so we
' never mutate production config without restoration.
Private Sub SetupEntornoForTempRoot(ByRef p_Db As DAO.Database, ByVal p_Root As String, ByRef logs As Collection)
    Dim frontendDb As DAO.Database
    Dim oldEntorno As entorno

    On Error Resume Next
    Set oldEntorno = m_ObjEntorno
    m_bHadEntorno = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If m_bHadEntorno Then
        Set m_SavedEntorno = oldEntorno
    Else
        Set m_SavedEntorno = Nothing
    End If

    m_SavedEnOficina = m_EnOficina
    m_SavedAccesoDatosTE = m_AccesoADatosTE

    On Error Resume Next
    m_SavedDatosEnLocal = CStr(Application.TempVars("DatosEnLocal"))
    m_bHadDatosEnLocal = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    Application.TempVars.Remove "DatosEnLocal"
    Application.TempVars.Add "DatosEnLocal", "S" & Chr$(237)

    ' Snapshot the FULL TbConfiguracionHPS table from the frontend BEFORE
    ' we write the test paths so teardown can restore every row exactly.
    Set frontendDb = CurrentDb
    Set m_SavedConfigSnap = SnapshotTbConfiguracionHPS(frontendDb)

    ' [config-as-source-of-truth] El nuevo resolver compone los
    ' anexos desde APP_ROOT_REMOTO/APP_ROOT_LOCAL con sufijo
    ' "HPS\ANEXOS\HPS\". Sembramos APP_ROOT_LOCAL con p_Root para que
    ' el resolver use el path temporal. APP_ROOT_REMOTO lo apuntamos a
    ' un path NO alcanzable (p_Root\__noexiste__\REMOTO) para forzar al
    ' resolver a preferir APP_ROOT_LOCAL. Las claves ANEXOS_HPS_PATH /
    ' ANEXOS_HISTORICO_PATH ya no se leen.
    WriteConfigValue frontendDb, "APP_ROOT_LOCAL", p_Root
    WriteConfigValue frontendDb, "APP_ROOT_REMOTO", p_Root & "\_REMOTO_NO_EXISTE_\"

    ' Mantenemos las legacy por si el binario tiene tests legacy que
    ' todavía las lean (sirven como breadcrumbs si la arquitectura
    ' cambia de nuevo).
    WriteConfigValue frontendDb, TEST_CONFIG_KEY_HPS, p_Root
    WriteConfigValue frontendDb, TEST_CONFIG_KEY_HIST, p_Root

    ' [config-as-source-of-truth] El resolver ahora usa fso.FolderExists
    ' para decidir entre APP_ROOT_REMOTO y APP_ROOT_LOCAL. Forzamos
    ' m_EnOficina y m_AccesoADatosTE para que el resolver NO bifurque por
    ' flags: queremos que lea siempre de la config.
    m_EnOficina = EnumSiNo.Sí
    m_AccesoADatosTE = EnumSiNo.No

    Dim freshEntorno As New entorno
    freshEntorno.CargarConfiguracion
    Set m_ObjEntorno = freshEntorno
    logs.Add "Entorno pinneado a temp root: " & p_Root
End Sub

Private Sub TeardownEntorno()
    On Error Resume Next
    Set m_ObjEntorno = Nothing
    If m_bHadEntorno Then
        Set m_ObjEntorno = m_SavedEntorno
        Set m_SavedEntorno = Nothing
    End If
    m_EnOficina = m_SavedEnOficina
    m_AccesoADatosTE = m_SavedAccesoDatosTE
    Application.TempVars.Remove "DatosEnLocal"
    If m_bHadDatosEnLocal And Len(m_SavedDatosEnLocal) > 0 Then
        Application.TempVars.Add "DatosEnLocal", m_SavedDatosEnLocal
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ReadConfigValue(ByRef p_Db As DAO.Database, ByVal p_Clave As String) As String
    Dim rs As DAO.Recordset
    On Error Resume Next
    Set rs = p_Db.OpenRecordset("SELECT Valor FROM TbConfiguracionHPS WHERE Clave='" & Replace(p_Clave, "'", "''") & "'", dbOpenSnapshot)
    If rs Is Nothing Then Exit Function
    If Not rs.EOF Then ReadConfigValue = CStr(Nz(rs!Valor, ""))
    rs.Close
    Err.Clear
    On Error GoTo 0
End Function

Private Sub WriteConfigValue(ByRef p_Db As DAO.Database, ByVal p_Clave As String, ByVal p_Valor As String)
    ' UPSERT in TbConfiguracionHPS. The production schema requires Valor,
    ' Activo, and FechaModificacion (NOT NULL) — the legacy 2-column form
    ' would raise 3078/validation errors. Mirror UpsertConfig from
    ' modConfiguracionHPS so tests stay schema-compatible.
    Dim sql As String
    Dim existing As String
    Dim usuario As String
    existing = ReadConfigValue(p_Db, p_Clave)
    usuario = Replace(Environ("Username"), "'", "''")
    If Len(existing) > 0 Then
        sql = "UPDATE TbConfiguracionHPS SET Valor='" & Replace(p_Valor, "'", "''") & _
              "', Activo=True, FechaModificacion=Now(), UsuarioModificacion='" & usuario & _
              "' WHERE Clave='" & Replace(p_Clave, "'", "''") & "'"
    Else
        sql = "INSERT INTO TbConfiguracionHPS (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) VALUES ('" & _
              Replace(p_Clave, "'", "''") & "','" & Replace(p_Valor, "'", "''") & _
              "',True,Now(),'" & usuario & "')"
    End If
    p_Db.Execute sql, dbFailOnError
End Sub

' Captures every row of TbConfiguracionHPS from p_Db into a Dictionary
' (Clave -> Scripting.Dictionary with Valor/Activo/FechaModificacion/
' UsuarioModificacion). Used to snapshot exact production config state
' before a test mutates it, so teardown can restore byte-for-byte. The
' production schema has 4 columns plus Clave (verified against modConfiguracionHPS).
Private Function SnapshotTbConfiguracionHPS(ByRef p_Db As DAO.Database) As Object
    Dim snap As Object
    Dim rs As DAO.Recordset
    Dim row As Object
    On Error GoTo Fail
    Set snap = CreateObject("Scripting.Dictionary")
    snap.CompareMode = 1  ' TextCompare
    Set rs = p_Db.OpenRecordset( _
        "SELECT Clave, Valor, Activo, FechaModificacion, UsuarioModificacion FROM TbConfiguracionHPS", _
        dbOpenSnapshot)
    Do While Not rs.EOF
        Set row = CreateObject("Scripting.Dictionary")
        row.Add "Valor", CStr(Nz(rs!Valor, ""))
        row.Add "Activo", Nz(rs!Activo, False)
        row.Add "FechaModificacion", CDate(Nz(rs!FechaModificacion, CDate(0)))
        row.Add "UsuarioModificacion", CStr(Nz(rs!UsuarioModificacion, ""))
        If Not snap.Exists(CStr(rs!Clave)) Then
            snap.Add CStr(rs!Clave), row
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set SnapshotTbConfiguracionHPS = snap
    Exit Function
Fail:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set SnapshotTbConfiguracionHPS = Nothing
End Function

' Restores TbConfiguracionHPS from a snapshot captured with
' SnapshotTbConfiguracionHPS. Wipes current rows then re-inserts each row
' from the snapshot, preserving Valor / Activo / FechaModificacion /
' UsuarioModificacion exactly. If p_Snap is Nothing, this is a no-op.
Private Sub RestoreTbConfiguracionHPS(ByRef p_Db As DAO.Database, ByRef p_Snap As Object)
    Dim k As Variant
    Dim row As Object
    Dim d As Date
    Dim fechaSql As String
    If p_Snap Is Nothing Then Exit Sub
    On Error Resume Next
    p_Db.Execute "DELETE FROM TbConfiguracionHPS", dbFailOnError
    For Each k In p_Snap.Keys
        Set row = p_Snap(k)
        d = CDate(row("FechaModificacion"))
        fechaSql = Format$(d, "yyyy-mm-dd hh:nn:ss")
        p_Db.Execute _
            "INSERT INTO TbConfiguracionHPS (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) VALUES ('" & _
            Replace(CStr(k), "'", "''") & "','" & _
            Replace(CStr(row("Valor")), "'", "''") & "'," & _
            IIf(CBool(row("Activo")), "True", "False") & ",#" & _
            fechaSql & "#,'" & _
            Replace(CStr(row("UsuarioModificacion")), "'", "''") & "')", _
            dbFailOnError
    Next k
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub RestoreConfigValues(ByRef p_Db As DAO.Database)
    ' p_Db is the BACKEND (HPST.accdb). TbConfiguracionHPS lives in the
    ' FRONTEND (HPS.accdb), so we restore via CurrentDb, not p_Db.
    Dim frontendDb As DAO.Database
    On Error Resume Next
    Set frontendDb = CurrentDb
    If Not frontendDb Is Nothing Then
        RestoreTbConfiguracionHPS frontendDb, m_SavedConfigSnap
    End If
    Set m_SavedConfigSnap = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' Fixture helpers ---------------------------------------------------

' Seeds one TbUsuarios row (TEST_DNI) and n TbAnexosUsuariosHPS rows
' with consecutive IDAnexo = p_UsuarioId + 1 .. p_UsuarioId + n, all
' pointing at p_EsHistorico, with file fixtures on disk under
' <p_Root>\<DNI>\<Actual|Hist>\. Returns the list of
' IDAnexo that were inserted (Collection of Long).
Private Sub SeedC2HFixture(ByRef p_Db As DAO.Database, _
                          ByVal p_UsuarioId As Long, _
                          ByVal p_DNI As String, _
                          ByVal p_Root As String, _
                          ByVal p_NombreAnexoA As String, _
                          ByVal p_NombreAnexoB As String, _
                          ByRef p_IDAnexoA As Long, _
                          ByRef p_IDAnexoB As Long)
    Dim actualDir As String
    actualDir = EnsureSlash(p_Root) & p_DNI & "\Actual\"
    CreateFolderTree actualDir

    p_IDAnexoA = p_UsuarioId + 1
    p_IDAnexoB = p_UsuarioId + 2

    On Error Resume Next
    p_Db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDAnexo IN (" & p_IDAnexoA & "," & p_IDAnexoB & ")", dbFailOnError
    p_Db.Execute "DELETE FROM TbUsuarios WHERE ID=" & p_UsuarioId, dbFailOnError
    Err.Clear
    On Error GoTo 0

    p_Db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & p_UsuarioId & ", '" & _
                 Replace(p_DNI, "'", "''") & "', 'AnexoSelTest', 'C2H')", dbFailOnError
    p_Db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & _
                 p_IDAnexoA & ", '" & Replace(p_NombreAnexoA, "'", "''") & "', " & p_UsuarioId & ", 'No')", dbFailOnError
    p_Db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & _
                 p_IDAnexoB & ", '" & Replace(p_NombreAnexoB, "'", "''") & "', " & p_UsuarioId & ", 'No')", dbFailOnError

    CreateTextFile actualDir & p_NombreAnexoA, "contenido A para IDAnexo " & p_IDAnexoA
    CreateTextFile actualDir & p_NombreAnexoB, "contenido B para IDAnexo " & p_IDAnexoB
End Sub

' Seeds an H2C fixture: anexos already in HISTORICO subfolder, EsHistorico='Sí'.
Private Sub SeedH2CFixture(ByRef p_Db As DAO.Database, _
                           ByVal p_UsuarioId As Long, _
                           ByVal p_DNI As String, _
                           ByVal p_Root As String, _
                           ByVal p_NombreAnexoA As String, _
                           ByVal p_NombreAnexoB As String, _
                           ByRef p_IDAnexoA As Long, _
                           ByRef p_IDAnexoB As Long)
    Dim historicoDir As String
    historicoDir = EnsureSlash(p_Root) & p_DNI & "\HISTORICO\"
    CreateFolderTree historicoDir

    p_IDAnexoA = p_UsuarioId + 1
    p_IDAnexoB = p_UsuarioId + 2

    On Error Resume Next
    p_Db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDAnexo IN (" & p_IDAnexoA & "," & p_IDAnexoB & ")", dbFailOnError
    p_Db.Execute "DELETE FROM TbUsuarios WHERE ID=" & p_UsuarioId, dbFailOnError
    Err.Clear
    On Error GoTo 0

    p_Db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & p_UsuarioId & ", '" & _
                 Replace(p_DNI, "'", "''") & "', 'AnexoSelTest', 'H2C')", dbFailOnError
    p_Db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & _
                 p_IDAnexoA & ", '" & Replace(p_NombreAnexoA, "'", "''") & "', " & p_UsuarioId & ", 'Sí')", dbFailOnError
    p_Db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & _
                 p_IDAnexoB & ", '" & Replace(p_NombreAnexoB, "'", "''") & "', " & p_UsuarioId & ", 'Sí')", dbFailOnError

    CreateTextFile historicoDir & p_NombreAnexoA, "contenido hist A para IDAnexo " & p_IDAnexoA
    CreateTextFile historicoDir & p_NombreAnexoB, "contenido hist B para IDAnexo " & p_IDAnexoB
End Sub

Private Sub CleanupFixture(ByRef p_Db As DAO.Database, _
                            ByVal p_UsuarioId As Long, _
                            ByVal p_IDAnexoA As Long, _
                            ByVal p_IDAnexoB As Long)
    On Error Resume Next
    p_Db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDAnexo IN (" & p_IDAnexoA & "," & p_IDAnexoB & ")", dbFailOnError
    p_Db.Execute "DELETE FROM TbUsuarios WHERE ID=" & p_UsuarioId, dbFailOnError
    Err.Clear
    On Error GoTo 0
End Sub

Private Function BuildIDsCollection(ByVal idA As Long, ByVal idB As Long) As Collection
    Dim c As New Collection
    c.Add idA
    c.Add idB
    Set BuildIDsCollection = c
End Function

' Cardinality / state assertion helpers -----------------------------

Private Function EsHistoricoInDb(ByRef p_Db As DAO.Database, ByVal p_IDAnexo As Long) As String
    Dim rs As DAO.Recordset
    On Error Resume Next
    Set rs = p_Db.OpenRecordset("SELECT EsHistorico FROM TbAnexosUsuariosHPS WHERE IDAnexo=" & p_IDAnexo, dbOpenSnapshot)
    If rs Is Nothing Then Exit Function
    If rs.EOF Then
        EsHistoricoInDb = "<missing>"
    Else
        EsHistoricoInDb = CStr(Nz(rs!EsHistorico, ""))
    End If
    rs.Close
    Err.Clear
    On Error GoTo 0
End Function

Private Function CountRows(ByRef p_Db As DAO.Database, ByVal p_Table As String, ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    On Error Resume Next
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS N FROM " & p_Table & " WHERE " & p_Where, dbOpenSnapshot)
    If rs Is Nothing Then Exit Function
    If Not rs.EOF Then CountRows = CLng(Nz(rs!N, 0))
    rs.Close
    Err.Clear
    On Error GoTo 0
End Function

' Asserts: each IDAnexo exists exactly once, EsHistorico matches p_Expected,
' source path is gone (or remains if rollback), dest path exists (or absent
' if rollback). Returns "" on success, error message on failure.
Private Function AssertBatchState(ByRef p_Db As DAO.Database, _
                                  ByVal p_DNI As String, _
                                  ByVal p_Root As String, _
                                  ByVal p_IDAnexoA As Long, _
                                  ByVal p_IDAnexoB As Long, _
                                  ByVal p_NombreAnexoA As String, _
                                  ByVal p_NombreAnexoB As String, _
                                  ByVal p_EsActual As Boolean, _
                                  ByVal p_AfterCommit As Boolean) As String
    Dim actualDir As String
    Dim histDir As String
    Dim sourceA As String
    Dim sourceB As String
    Dim destA As String
    Dim destB As String
    Dim expectedEstado As String

    actualDir = EnsureSlash(p_Root) & p_DNI & "\Actual\"
    histDir = EnsureSlash(p_Root) & p_DNI & "\HISTORICO\"

    If p_EsActual Then
        sourceA = actualDir & p_NombreAnexoA
        sourceB = actualDir & p_NombreAnexoB
        destA = histDir & p_NombreAnexoA
        destB = histDir & p_NombreAnexoB
        expectedEstado = "No"
    Else
        sourceA = histDir & p_NombreAnexoA
        sourceB = histDir & p_NombreAnexoB
        destA = actualDir & p_NombreAnexoA
        destB = actualDir & p_NombreAnexoB
        expectedEstado = "Sí"
    End If

    If p_AfterCommit Then
        ' After commit: source gone, dest present, EsHistorico = OPPOSITE of origin.
        If fso.FileExists(sourceA) Then AssertBatchState = "source A still present after commit: " & sourceA: Exit Function
        If fso.FileExists(sourceB) Then AssertBatchState = "source B still present after commit: " & sourceB: Exit Function
        If Not fso.FileExists(destA) Then AssertBatchState = "dest A missing after commit: " & destA: Exit Function
        If Not fso.FileExists(destB) Then AssertBatchState = "dest B missing after commit: " & destB: Exit Function
        ' EsHistorico must equal NOT p_EsActual.
        Dim postEstado As String
        If p_EsActual Then
            postEstado = "Sí"
        Else
            postEstado = "No"
        End If
        If EsHistoricoInDb(p_Db, p_IDAnexoA) <> postEstado Then AssertBatchState = "IDAnexo " & p_IDAnexoA & " EsHistorico='" & EsHistoricoInDb(p_Db, p_IDAnexoA) & "', expected '" & postEstado & "'": Exit Function
        If EsHistoricoInDb(p_Db, p_IDAnexoB) <> postEstado Then AssertBatchState = "IDAnexo " & p_IDAnexoB & " EsHistorico='" & EsHistoricoInDb(p_Db, p_IDAnexoB) & "', expected '" & postEstado & "'": Exit Function
    Else
        ' After rollback: source back, dest absent, EsHistorico = original.
        If Not fso.FileExists(sourceA) Then AssertBatchState = "source A not restored after rollback: " & sourceA: Exit Function
        If Not fso.FileExists(sourceB) Then AssertBatchState = "source B not restored after rollback: " & sourceB: Exit Function
        If fso.FileExists(destA) Then AssertBatchState = "dest A leaked after rollback: " & destA: Exit Function
        If fso.FileExists(destB) Then AssertBatchState = "dest B leaked after rollback: " & destB: Exit Function
        If EsHistoricoInDb(p_Db, p_IDAnexoA) <> expectedEstado Then AssertBatchState = "IDAnexo " & p_IDAnexoA & " EsHistorico='" & EsHistoricoInDb(p_Db, p_IDAnexoA) & "', expected original '" & expectedEstado & "'": Exit Function
        If EsHistoricoInDb(p_Db, p_IDAnexoB) <> expectedEstado Then AssertBatchState = "IDAnexo " & p_IDAnexoB & " EsHistorico='" & EsHistoricoInDb(p_Db, p_IDAnexoB) & "', expected original '" & expectedEstado & "'": Exit Function
    End If
End Function

' Shared logic for the AfterCopy / AfterDbUpdate / AfterSourceDelete /
' BeforeCommit rollback atoms. We use a 2-anexo batch so the seam (which
' fires on item i=1) leaves item i=2 with no CopyFile/UPDATE/DeleteFile
' at all — the test asserts that NEITHER item moved AND NEITHER was left
' half-moved.
Private Function RunRollbackAtSeamTest(ByVal p_Seam As String, ByVal p_CardinalidadBase As String) As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Dim preCount As Long, postCount As Long
    Dim usuarioId As Long
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    usuarioId = TEST_USER_C2H_BASE + CLng(p_CardinalidadBase)
    dni = TEST_DNI_PREFIX & p_CardinalidadBase
    root = TempTestRoot("c2h-" & LCase(p_Seam))
    logs.Add "RB step: root='" & root & "'"
    DeleteFolderIfExists root
    logs.Add "RB step: deleted previous root"
    CreateFolderTree root
    logs.Add "RB step: created root tree"

    SetupEntornoForTempRoot db, root, logs
    logs.Add "RB step: entorno set"

    SeedC2HFixture db, usuarioId, dni, root, _
        "a-" & p_Seam & ".pdf", "b-" & p_Seam & ".pdf", idA, idB
    logs.Add "RB step: seed ok idA=" & idA & " idB=" & idB
    preCount = CountRows(db, "TbAnexosUsuariosHPS", "IDUsuario=" & usuarioId)
    logs.Add "RB step: preCount=" & preCount

    preCount = CountRows(db, "TbAnexosUsuariosHPS", "IDUsuario=" & usuarioId)

    Set coordinator = New AnexoSelectionTransactionCoordinator
    logs.Add "RB step: coordinator instantiated"
    Set ids = BuildIDsCollection(idA, idB)
    logs.Add "RB step: ids built count=" & ids.Count
    On Error Resume Next
    Dim mRet As Boolean
    mRet = coordinator.MoverLote(ids, True, db, p_Seam, errMsg)
    If Err.Number <> 0 Then
        logs.Add "RB step: MoverLote raised VBA err=" & Err.Number & " desc=" & Err.Description
        Err.Clear
    End If
    On Error GoTo EH
    If mRet Then
        RunRollbackAtSeamTest = JsonFail("Se esperaba fallo por seam '" & p_Seam & "', pero MoverLote devolvió éxito.", logs)
        GoTo CleanUp
    End If
    logs.Add "RB step: MoverLote returned errMsg='" & errMsg & "'"
    If Len(errMsg) = 0 Then
        RunRollbackAtSeamTest = JsonFail("Se esperaba p_Error no vacío para seam '" & p_Seam & "'.", logs)
        GoTo CleanUp
    End If

    Dim stateErr As String
    stateErr = AssertBatchState(db, dni, root, idA, idB, _
        "a-" & p_Seam & ".pdf", "b-" & p_Seam & ".pdf", True, False)
    If Len(stateErr) > 0 Then
        RunRollbackAtSeamTest = JsonFail("Rollback incompleto tras seam '" & p_Seam & "': " & stateErr, logs)
        GoTo CleanUp
    End If

    postCount = CountRows(db, "TbAnexosUsuariosHPS", "IDUsuario=" & usuarioId)
    If postCount <> preCount Then
        RunRollbackAtSeamTest = JsonFail("Cardinalidad cambió tras rollback (" & preCount & " -> " & postCount & ") tras seam '" & p_Seam & "'.", logs)
        GoTo CleanUp
    End If

    logs.Add "Seam '" & p_Seam & "' disparó rollback atómico: DB sin cambios, ficheros origen restaurados, sin residuos en destino."
    RunRollbackAtSeamTest = JsonOk("rollback-" & LCase(p_Seam), logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, usuarioId, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    RunRollbackAtSeamTest = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, usuarioId, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

' ============================================================
' PUBLIC TEST ATOMS — happy path
' ============================================================

Public Function Test_ASTC_C2H_BatchMovesTwoAnexos() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    dni = TEST_DNI_PREFIX & "01"
    root = TempTestRoot("c2h-success")
    DeleteFolderIfExists root
    CreateFolderTree root

    SetupEntornoForTempRoot db, root, logs

    SeedC2HFixture db, TEST_USER_C2H_BASE, dni, root, "alpha.pdf", "beta.pdf", idA, idB

    Set coordinator = New AnexoSelectionTransactionCoordinator
    Set ids = BuildIDsCollection(idA, idB)
    If Not coordinator.MoverLote(ids, True, db, "", errMsg) Then
        Test_ASTC_C2H_BatchMovesTwoAnexos = JsonFail("MoverLote c2h no completó: " & errMsg, logs)
        GoTo CleanUp
    End If
    If Len(errMsg) <> 0 Then
        Test_ASTC_C2H_BatchMovesTwoAnexos = JsonFail("MoverLote c2h devolvió p_Error no vacío en éxito: " & errMsg, logs)
        GoTo CleanUp
    End If

    Dim stateErr As String
    stateErr = AssertBatchState(db, dni, root, idA, idB, "alpha.pdf", "beta.pdf", True, True)
    If Len(stateErr) > 0 Then
        Test_ASTC_C2H_BatchMovesTwoAnexos = JsonFail("Estado post-commit inconsistente: " & stateErr, logs)
        GoTo CleanUp
    End If

    logs.Add "C2H batch de 2 anexos: ambos ficheros en HISTORICO, ambos EsHistorico='Sí', ninguno en Actual."
    Test_ASTC_C2H_BatchMovesTwoAnexos = JsonOk("c2h-batch-success", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, TEST_USER_C2H_BASE, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    Test_ASTC_C2H_BatchMovesTwoAnexos = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, TEST_USER_C2H_BASE, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

Public Function Test_ASTC_H2C_BatchMovesTwoAnexos() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    dni = TEST_DNI_PREFIX & "02"
    root = TempTestRoot("h2c-success")
    DeleteFolderIfExists root
    CreateFolderTree root

    SetupEntornoForTempRoot db, root, logs

    SeedH2CFixture db, TEST_USER_H2C_BASE, dni, root, "alpha-h.txt", "beta-h.txt", idA, idB

    Set coordinator = New AnexoSelectionTransactionCoordinator
    Set ids = BuildIDsCollection(idA, idB)
    If Not coordinator.MoverLote(ids, False, db, "", errMsg) Then
        Test_ASTC_H2C_BatchMovesTwoAnexos = JsonFail("MoverLote h2c no completó: " & errMsg, logs)
        GoTo CleanUp
    End If

    Dim stateErr As String
    stateErr = AssertBatchState(db, dni, root, idA, idB, "alpha-h.txt", "beta-h.txt", False, True)
    If Len(stateErr) > 0 Then
        Test_ASTC_H2C_BatchMovesTwoAnexos = JsonFail("Estado post-commit inconsistente: " & stateErr, logs)
        GoTo CleanUp
    End If

    logs.Add "H2C batch de 2 anexos: ambos ficheros en Actual, ambos EsHistorico='No', ninguno en HISTORICO."
    Test_ASTC_H2C_BatchMovesTwoAnexos = JsonOk("h2c-batch-success", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, TEST_USER_H2C_BASE, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    Test_ASTC_H2C_BatchMovesTwoAnexos = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, TEST_USER_H2C_BASE, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

' ============================================================
' PUBLIC TEST ATOMS — rollback at each seam
' ============================================================

Public Function Test_ASTC_C2H_RollsBackAfterCopyFails() As String
    Test_ASTC_C2H_RollsBackAfterCopyFails = RunRollbackAtSeamTest("AfterCopy", "10")
End Function

Public Function Test_ASTC_C2H_RollsBackAfterDbUpdateFails() As String
    Test_ASTC_C2H_RollsBackAfterDbUpdateFails = RunRollbackAtSeamTest("AfterDbUpdate", "11")
End Function

Public Function Test_ASTC_C2H_RollsBackAfterSourceDeleteFails() As String
    Test_ASTC_C2H_RollsBackAfterSourceDeleteFails = RunRollbackAtSeamTest("AfterSourceDelete", "12")
End Function

Public Function Test_ASTC_C2H_RollsBackBeforeCommitFails() As String
    Test_ASTC_C2H_RollsBackBeforeCommitFails = RunRollbackAtSeamTest("BeforeCommit", "13")
End Function

' ============================================================
' PUBLIC TEST ATOMS — preflight rejections
' ============================================================

Public Function Test_ASTC_C2H_RejectsMissingSource() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Dim sourceA As String
    Dim sourceB As String
    Dim histDir As String
    Dim actualDir As String
    Dim stateErr As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    dni = TEST_DNI_PREFIX & "20"
    root = TempTestRoot("c2h-missing-source")
    DeleteFolderIfExists root
    CreateFolderTree root

    SetupEntornoForTempRoot db, root, logs

    SeedC2HFixture db, TEST_USER_REJECTS_BASE, dni, root, "alpha.pdf", "beta.pdf", idA, idB

    ' Wipe the source file for idA — the coordinator must refuse BEFORE any mutation.
    sourceA = EnsureSlash(root) & dni & "\Actual\alpha.pdf"
    If fso.FileExists(sourceA) Then fso.DeleteFile sourceA, True

    Set coordinator = New AnexoSelectionTransactionCoordinator
    Set ids = BuildIDsCollection(idA, idB)
    If coordinator.MoverLote(ids, True, db, "", errMsg) Then
        Test_ASTC_C2H_RejectsMissingSource = JsonFail("MoverLote debería rechazar source ausente pero devolvió éxito.", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "origen no encontrado", vbTextCompare) = 0 Then
        Test_ASTC_C2H_RejectsMissingSource = JsonFail("Mensaje debería mencionar 'origen no encontrado', got: " & errMsg, logs)
        GoTo CleanUp
    End If

    ' Preflight semantics: NO mutation, NO snapshot. The test deliberately
    ' deleted source A pre-run, so the post-rejection state is:
    '   - source A: still absent (preflight did not recreate it)
    '   - source B: still present (untouched)
    '   - no destination files created
    '   - DB unchanged (both rows still EsHistorico='No')
    '   - no snapshot folder under %TEMP%\HPS\AnexoSelectionSnapshots\
    actualDir = EnsureSlash(root) & dni & "\Actual\"
    histDir = EnsureSlash(root) & dni & "\HISTORICO\"
    sourceB = actualDir & "beta.pdf"

    If fso.FileExists(sourceA) Then
        stateErr = "source A se recreó tras rechazo de preflight (no debería): " & sourceA
    ElseIf Not fso.FileExists(sourceB) Then
        stateErr = "source B desapareció tras rechazo de preflight (no debería): " & sourceB
    ElseIf fso.FileExists(histDir & "alpha.pdf") Then
        stateErr = "destino alpha.pdf apareció tras rechazo de preflight (no debería)"
    ElseIf fso.FileExists(histDir & "beta.pdf") Then
        stateErr = "destino beta.pdf apareció tras rechazo de preflight (no debería)"
    ElseIf EsHistoricoInDb(db, idA) <> "No" Then
        stateErr = "IDAnexo " & idA & " EsHistorico cambió tras rechazo: '" & EsHistoricoInDb(db, idA) & "'"
    ElseIf EsHistoricoInDb(db, idB) <> "No" Then
        stateErr = "IDAnexo " & idB & " EsHistorico cambió tras rechazo: '" & EsHistoricoInDb(db, idB) & "'"
    End If

    If Len(stateErr) > 0 Then
        Test_ASTC_C2H_RejectsMissingSource = JsonFail("Estado post-rechazo inconsistente: " & stateErr, logs)
        GoTo CleanUp
    End If

    logs.Add "Source ausente rechazado en preflight: source A sigue ausente (no se recreó), source B intacto, ningún destino, DB sin cambios."
    Test_ASTC_C2H_RejectsMissingSource = JsonOk("rejects-missing-source", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    Test_ASTC_C2H_RejectsMissingSource = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

Public Function Test_ASTC_C2H_RejectsExistingDestination() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Dim sourceA As String
    Dim sourceB As String
    Dim destA As String
    Dim destB As String
    Dim actualDir As String
    Dim histDir As String
    Dim stateErr As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    dni = TEST_DNI_PREFIX & "21"
    root = TempTestRoot("c2h-existing-dest")
    DeleteFolderIfExists root
    CreateFolderTree root

    SetupEntornoForTempRoot db, root, logs

    SeedC2HFixture db, TEST_USER_REJECTS_BASE + 1, dni, root, "alpha.pdf", "beta.pdf", idA, idB

    ' Pre-create a destination file — coordinator MUST refuse (no overwrite).
    actualDir = EnsureSlash(root) & dni & "\Actual\"
    histDir = EnsureSlash(root) & dni & "\HISTORICO\"
    sourceA = actualDir & "alpha.pdf"
    sourceB = actualDir & "beta.pdf"
    destA = histDir & "alpha.pdf"
    destB = histDir & "beta.pdf"
    CreateTextFile destA, "destino pre-existente"

    Set coordinator = New AnexoSelectionTransactionCoordinator
    Set ids = BuildIDsCollection(idA, idB)
    If coordinator.MoverLote(ids, True, db, "", errMsg) Then
        Test_ASTC_C2H_RejectsExistingDestination = JsonFail("MoverLote debería rechazar destino existente pero devolvió éxito.", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "ya existe el archivo de destino", vbTextCompare) = 0 Then
        Test_ASTC_C2H_RejectsExistingDestination = JsonFail("Mensaje debería mencionar 'ya existe el archivo de destino', got: " & errMsg, logs)
        GoTo CleanUp
    End If

    ' Preflight semantics: NO mutation. The pre-existing destination MUST
    ' remain untouched (content preserved), sources MUST be intact, DB
    ' unchanged. The shared AssertBatchState helper assumes "no destinations
    ' after rollback", so we assert manually here to keep the intentional
    ' pre-existing destination.
    stateErr = ""
    If Not fso.FileExists(destA) Then
        stateErr = "Destino pre-existente desapareció tras rechazo."
    ElseIf Not fso.FileExists(sourceA) Then
        stateErr = "source A desapareció tras rechazo: " & sourceA
    ElseIf Not fso.FileExists(sourceB) Then
        stateErr = "source B desapareció tras rechazo: " & sourceB
    ElseIf fso.FileExists(destB) Then
        stateErr = "destino B apareció inesperadamente: " & destB
    ElseIf EsHistoricoInDb(db, idA) <> "No" Then
        stateErr = "IDAnexo " & idA & " EsHistorico cambió tras rechazo: '" & EsHistoricoInDb(db, idA) & "'"
    ElseIf EsHistoricoInDb(db, idB) <> "No" Then
        stateErr = "IDAnexo " & idB & " EsHistorico cambió tras rechazo: '" & EsHistoricoInDb(db, idB) & "'"
    End If

    If Len(stateErr) > 0 Then
        Test_ASTC_C2H_RejectsExistingDestination = JsonFail("Estado post-rechazo inconsistente: " & stateErr, logs)
        GoTo CleanUp
    End If

    logs.Add "Destino existente rechazado en preflight: contenido original preservado, sources intactos, DB sin cambios."
    Test_ASTC_C2H_RejectsExistingDestination = JsonOk("rejects-existing-dest", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 1, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    Test_ASTC_C2H_RejectsExistingDestination = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 1, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

Public Function Test_ASTC_C2H_RejectsDuplicateIDAnexo() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    dni = TEST_DNI_PREFIX & "22"
    root = TempTestRoot("c2h-duplicate-id")
    DeleteFolderIfExists root
    CreateFolderTree root

    SetupEntornoForTempRoot db, root, logs

    SeedC2HFixture db, TEST_USER_REJECTS_BASE + 2, dni, root, "alpha.pdf", "beta.pdf", idA, idB

    Set coordinator = New AnexoSelectionTransactionCoordinator
    Set ids = BuildIDsCollection(idA, idA)   ' same ID twice
    If coordinator.MoverLote(ids, True, db, "", errMsg) Then
        Test_ASTC_C2H_RejectsDuplicateIDAnexo = JsonFail("MoverLote debería rechazar IDAnexo duplicado pero devolvió éxito.", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "duplicado", vbTextCompare) = 0 Then
        Test_ASTC_C2H_RejectsDuplicateIDAnexo = JsonFail("Mensaje debería mencionar 'duplicado', got: " & errMsg, logs)
        GoTo CleanUp
    End If

    Dim stateErr As String
    stateErr = AssertBatchState(db, dni, root, idA, idB, "alpha.pdf", "beta.pdf", True, False)
    If Len(stateErr) > 0 Then
        Test_ASTC_C2H_RejectsDuplicateIDAnexo = JsonFail("Estado post-rechazo inconsistente: " & stateErr, logs)
        GoTo CleanUp
    End If

    logs.Add "IDAnexo duplicado rechazado en preflight antes de cualquier mutación."
    Test_ASTC_C2H_RejectsDuplicateIDAnexo = JsonOk("rejects-duplicate-id", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 2, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    Test_ASTC_C2H_RejectsDuplicateIDAnexo = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 2, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

Public Function Test_ASTC_C2H_RejectsMismatchedEsHistorico() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    dni = TEST_DNI_PREFIX & "23"
    root = TempTestRoot("c2h-mismatch-estado")
    DeleteFolderIfExists root
    CreateFolderTree root

    SetupEntornoForTempRoot db, root, logs

    ' Seed H2C fixtures (EsHistorico='Sí') but try to C2H them. Coordinator must refuse.
    SeedH2CFixture db, TEST_USER_REJECTS_BASE + 3, dni, root, "alpha-h.pdf", "beta-h.pdf", idA, idB

    Set coordinator = New AnexoSelectionTransactionCoordinator
    Set ids = BuildIDsCollection(idA, idB)
    If coordinator.MoverLote(ids, True, db, "", errMsg) Then
        Test_ASTC_C2H_RejectsMismatchedEsHistorico = JsonFail("MoverLote debería rechazar EsHistorico inconsistente pero devolvió éxito.", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "EsHistorico", vbTextCompare) = 0 Then
        Test_ASTC_C2H_RejectsMismatchedEsHistorico = JsonFail("Mensaje debería mencionar 'EsHistorico', got: " & errMsg, logs)
        GoTo CleanUp
    End If

    ' Estado HISTORICO preservado.
    If EsHistoricoInDb(db, idA) <> "Sí" Or EsHistoricoInDb(db, idB) <> "Sí" Then
        Test_ASTC_C2H_RejectsMismatchedEsHistorico = JsonFail("EsHistorico cambió tras rechazo: A='" & _
            EsHistoricoInDb(db, idA) & "', B='" & EsHistoricoInDb(db, idB) & "'.", logs)
        GoTo CleanUp
    End If

    logs.Add "EsHistorico='Sí' rechazado al pedir C2H — la fila no estaba en el estado origen esperado."
    Test_ASTC_C2H_RejectsMismatchedEsHistorico = JsonOk("rejects-mismatch-estado", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 3, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    Test_ASTC_C2H_RejectsMismatchedEsHistorico = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 3, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

Public Function Test_ASTC_C2H_RejectsDuplicateDestinationPath() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim root As String
    Dim dni As String
    Dim idA As Long, idB As Long
    Dim ids As Collection
    Dim coordinator As AnexoSelectionTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    dni = TEST_DNI_PREFIX & "24"
    root = TempTestRoot("c2h-dup-destpath")
    DeleteFolderIfExists root
    CreateFolderTree root

    SetupEntornoForTempRoot db, root, logs

    ' Two anexos with the SAME filename in DB, distinct IDAnexos. Both
    ' resolve to the same destination path. SeedC2HFixture writes one
    ' physical file (second CreateTextFile overwrites the first because
    ' Windows disallows two files with the same name in one folder), so
    ' BOTH URLAnexo calls resolve to the same physical source — preflight
    ' must NOT trip the source-existence check, and must reach the
    ' duplicate-destination rejection at the end of the validation chain.
    SeedC2HFixture db, TEST_USER_REJECTS_BASE + 4, dni, root, "same.pdf", "same.pdf", idA, idB

    Set coordinator = New AnexoSelectionTransactionCoordinator
    Set ids = BuildIDsCollection(idA, idB)
    If coordinator.MoverLote(ids, True, db, "", errMsg) Then
        Test_ASTC_C2H_RejectsDuplicateDestinationPath = JsonFail("MoverLote debería rechazar dos IDAnexos con mismo destino pero devolvió éxito.", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "mismo destino", vbTextCompare) = 0 Then
        Test_ASTC_C2H_RejectsDuplicateDestinationPath = JsonFail("Mensaje debería mencionar 'mismo destino', got: " & errMsg, logs)
        GoTo CleanUp
    End If

    Dim stateErr As String
    stateErr = AssertBatchState(db, dni, root, idA, idB, "same.pdf", "same.pdf", True, False)
    ' Note: A and B share the same NombreAnexo "same.pdf" so both URLAnexo
    ' paths point at <root>\<DNI>\Actual\same.pdf (one physical file on
    ' disk) and both destinations resolve to <root>\<DNI>\HISTORICO\same.pdf.
    ' The preflight rejection happens BEFORE any copy/UPDATE/delete, so
    ' both source files are still present (same file actually) and neither
    ' destination exists.
    If Len(stateErr) > 0 Then
        Test_ASTC_C2H_RejectsDuplicateDestinationPath = JsonFail("Estado post-rechazo inconsistente: " & stateErr, logs)
        GoTo CleanUp
    End If

    logs.Add "Dos IDAnexos resolviendo al mismo path destino rechazados: preflight defensivo funcionando."
    Test_ASTC_C2H_RejectsDuplicateDestinationPath = JsonOk("rejects-dup-destpath", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 4, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Exit Function
EH:
    Test_ASTC_C2H_RejectsDuplicateDestinationPath = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    CleanupFixture db, TEST_USER_REJECTS_BASE + 4, idA, idB
    DeleteFolderIfExists root
    RestoreConfigValues db
    TeardownEntorno
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
End Function

' NOTE: Test_ASTC_C2H_RejectsRecordsAffectedNotOne was removed. It only
' triggered the BeforeDbUpdate seam and never actually exercised the
' coordinator's "RecordsAffected <> 1" branch — the seam raised BEFORE
' the UPDATE ran, so db.RecordsAffected was never read. A genuine
' deterministic test would need to delete the row between preflight and
' the UPDATE, which requires either an extra in-transaction seam or a
' second open DAO.Database, and is not worth the cost relative to the
' value of the success-path cardinality checks in
' Test_ASTC_C2H_BatchMovesTwoAnexos / Test_ASTC_H2C_BatchMovesTwoAnexos
' plus the AfterDbUpdate rollback atom.
