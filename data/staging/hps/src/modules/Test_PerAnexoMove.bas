Attribute VB_Name = "Test_PerAnexoMove"
Option Compare Database
Option Explicit

' ============================================================
' ANEXOS-HISTORICOS hotfix tests for the PER-ANEXO move path
' (AnexoUsuarioHPS.PasarAnexo / PasarAnexoAHistorico).
'
' These tests cover the bugs reported on 2026-07-10:
'   1. UsuarioHPS.URLCarpetaAnexosHistoricos must derive from
'      URLCarpetaAnexosActuales (adds \HISTORICO\ subfolder inside <DNI>\),
'      NOT from m_ObjEntorno.URLCarpetaAnexosHistoricos (which is the
'      separate network root for users that have been moved to
'      TbUsuariosHistoricos).
'   2. AnexoUsuarioHPS.PasarAnexo must fail LOUD when the source file is
'      missing — previously the If fso.FileExists(...) guard made the
'      CopyFile/DeleteFile a no-op, and the DB UPDATE at the end still
'      ran, leaving DB and filesystem out of sync ("no localizado" symptom).
'   3. Same defect in PasarAnexoAHistorico: DeleteFile ran unconditionally
'      even when the source was missing.
'
' Test ID ranges (>= 900000 to avoid colliding with production data):
'   TEST_ID_USER    = 900201
'   TEST_ID_ANEXO   = 900202
' ============================================================

Private Const TEST_ID_USER As Long = 900201
Private Const TEST_ID_ANEXO As Long = 900202
Private Const TEST_DNI As String = "DNI900201"

' ============================================================
' Helpers (module-local — declaration order: consts, JSON wrappers,
' other private helpers, public atoms).
' ============================================================

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

Private Function EnsureSlash(ByVal path As String) As String
    If Len(path) = 0 Then Exit Function
    If Right$(path, 1) = "\" Then
        EnsureSlash = path
    Else
        EnsureSlash = path & "\"
    End If
End Function

Private Function TempPerAnexoRoot(ByVal suffix As String) As String
    TempPerAnexoRoot = EnsureSlash(Environ$("TEMP")) & "HPS\PerAnexoMoveTests\" & suffix & "\"
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

Private Function EsHistoricoInDb(ByVal p_IDAnexo As Long) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim connectString As String
    Dim backendPath As String

    backendPath = EnsureSlash(CurrentProject.path) & "HPST.accdb"
    connectString = BackendConnectString()
    Set db = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, False, connectString)
    Set rs = db.OpenRecordset("SELECT EsHistorico FROM TbAnexosUsuariosHPS WHERE IDAnexo=" & p_IDAnexo, dbOpenSnapshot)
    If rs.EOF Then
        EsHistoricoInDb = "<missing>"
    Else
        EsHistoricoInDb = Nz(rs!EsHistorico, "")
    End If
    rs.Close
    db.Close
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
    If Len(password) = 0 Then Err.Raise 1000, , "BackendConnectString: DYSFLOW_BACKEND_PASSWORD, HPS_BACKEND_PASSWORD, ACCESS_VBA_PASSWORD, or .dysflow/backend.pwd is required for PerAnexoMove tests."
    BackendConnectString = "MS Access;PWD=" & password
End Function

Private Sub SetDatosEnLocalForTest(ByVal p_Value As String)
    On Error Resume Next
    Application.TempVars.Remove "DatosEnLocal"
    Err.Clear
    Application.TempVars.Add "DatosEnLocal", p_Value
    On Error GoTo 0
End Sub

Private Sub EnsureUserAndAnexoInDb()
    ' Seed: TbUsuarios row (so UsuarioHPS.Usuario can resolve) + TbAnexosUsuariosHPS row
    ' with EsHistorico='No'. Both use TEST_ID_* ranges so other tests don't collide.
    Dim db As DAO.Database
    Dim connectString As String
    Dim backendPath As String
    Dim fso As Object

    backendPath = EnsureSlash(CurrentProject.path) & "HPST.accdb"
    connectString = BackendConnectString()
    Set db = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, False, connectString)
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDAnexo=" & TEST_ID_ANEXO, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & TEST_ID_USER, dbFailOnError
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & TEST_ID_USER & ", '" & TEST_DNI & "', 'PerAnexoTest', 'Fixture')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, IDUsuario, NombreAnexo, EsHistorico) VALUES (" & TEST_ID_ANEXO & ", " & TEST_ID_USER & ", 'missing-source.pdf', 'No')", dbFailOnError
    If Err.Number <> 0 Then
        Debug.Print "EnsureUserAndAnexoInDb err=" & Err.Number & " " & Err.Description
    End If
    On Error GoTo 0
    db.Close
End Sub

Private Sub CleanupUserAndAnexoInDb()
    Dim db As DAO.Database
    Dim connectString As String
    Dim backendPath As String

    backendPath = EnsureSlash(CurrentProject.path) & "HPST.accdb"
    connectString = BackendConnectString()
    Set db = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, False, connectString)

    On Error Resume Next
    db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDAnexo=" & TEST_ID_ANEXO, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & TEST_ID_USER, dbFailOnError
    On Error GoTo 0
    db.Close
End Sub

' ============================================================
' PUBLIC TEST ATOMS
' ============================================================

Public Function Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI() As String
    ' [anexos subcarpetas] Locks the path convention for ACTIVE users.
    ' Both URLs derive from the SAME configured root (m_ObjEntorno.URLCarpetaAnexos)
    ' and are SIBLINGS under the DNI subfolder:
    '   URLCarpetaAnexosActuales  = <root>\<DNI>\Actual\
    '   URLCarpetaAnexosHistoricos = <root>\<DNI>\HISTORICO\
    ' NOT nested (the earlier nested form <root>\<DNI>\Actual\HISTORICO\
    ' was a bug that made historical anexos unreachable from PasarAnexo).
    '
    ' Top-level convention recap (los paths vienen de APP_ROOT_* en TbConfiguracionHPS
    ' resueltos via Entorno.ResolveURLCarpetaAnexosCandidate):
    '   <APP_ROOT_*>\HPS\ANEXOS\HPS\<DNI>\Actual\<file>     <- current anexos of active user
    '   <APP_ROOT_*>\HPS\ANEXOS\HPS\<DNI>\HISTORICO\<file>   <- historical anexos of active user
    '   <APP_ROOT_*>\HPS\ANEXOS\HISTORICO\<DNI>\<file>       <- ALL anexos of HISTORICAL user (flat)
    Dim logs As Collection
    Dim db As DAO.Database
    Dim testEntorno As entorno
    Dim u As UsuarioHPS
    Dim sActuales As String
    Dim sHistoricos As String
    Dim eOldEnOficina As EnumSiNo
    Dim eOldAccesoDatosTE As EnumSiNo
    Dim sOldDatosEnLocal As String
    Dim bHadDatosEnLocal As Boolean
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    eOldEnOficina = m_EnOficina
    eOldAccesoDatosTE = m_AccesoADatosTE
    On Error Resume Next
    sOldDatosEnLocal = CStr(Application.TempVars("DatosEnLocal"))
    bHadDatosEnLocal = (Err.Number = 0)
    Err.Clear
    On Error GoTo EH

    ' Save the live m_ObjEntorno to restore on teardown. Then build a
    ' deterministic entorno by populating TbConfiguracionHPS with the Local
    ' profile, mimicking Test_HPSConfig_LocalModeUsesLocalAnexosConfig.
    Dim oldEntorno As entorno
    Set oldEntorno = m_ObjEntorno

    On Error Resume Next
    db.Execute "DELETE FROM TbConfiguracionHPS", dbFailOnError
    Err.Clear
    On Error GoTo EH

    SetDatosEnLocalForTest "S" & Chr$(237)
    ConfigurarTablaConfiguracion_PerfilLocal
    m_EnOficina = EnumSiNo.No
    m_AccesoADatosTE = EnumSiNo.Sí

    Set testEntorno = New entorno
    testEntorno.CargarConfiguracion
    Set m_ObjEntorno = testEntorno

    Set u = New UsuarioHPS
    u.DNI = TEST_DNI
    u.ID = CStr(TEST_ID_USER)
    ' Force the property GET to recompute (the cache lives on the instance).
    sActuales = u.URLCarpetaAnexosActuales
    sHistoricos = u.URLCarpetaAnexosHistoricos

    If Len(sActuales) = 0 Then
        Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail("URLCarpetaAnexosActuales returned empty; cannot verify sibling.", logs)
        GoTo CleanUp
    End If
    If Len(sHistoricos) = 0 Then
        Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail("URLCarpetaAnexosHistoricos returned empty; the new derivation must populate it.", logs)
        GoTo CleanUp
    End If
    ' [anexos subcarpetas] The contract: BOTH URLs share the same <root>\<DNI>\
    ' prefix (without the trailing per-user suffix); Actuales appends "\Actual\"
    ' and Historicos appends "\HISTORICO\". Strip each suffix and assert equal.
    Dim expectedActualesBase As String
    Dim expectedHistoricosBase As String
    If Right$(sActuales, Len("\Actual\")) = "\Actual\" Then
        expectedActualesBase = Left$(sActuales, Len(sActuales) - Len("\Actual\"))
    Else
        Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail("URLCarpetaAnexosActuales MUST end with '\Actual\' but got: '" & sActuales & "'.", logs)
        GoTo CleanUp
    End If
    If Right$(sHistoricos, Len("\HISTORICO\")) = "\HISTORICO\" Then
        expectedHistoricosBase = Left$(sHistoricos, Len(sHistoricos) - Len("\HISTORICO\"))
    Else
        Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail("URLCarpetaAnexosHistoricos MUST end with '\HISTORICO\' but got: '" & sHistoricos & "'.", logs)
        GoTo CleanUp
    End If
    If expectedActualesBase <> expectedHistoricosBase Then
        Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail( _
            "URLCarpetaAnexosActuales and URLCarpetaAnexosHistoricos MUST share the same <root>\<DNI>\ prefix. " & _
            "Actuales='" & sActuales & "', Historicos='" & sHistoricos & "'.", logs)
        GoTo CleanUp
    End If
    ' Negative: the OLD nested form had URLs that differed only in suffix.
    ' The new convention gives them the same prefix and distinct suffixes.
    ' Specifically, Historicos MUST NOT equal Actuales + "HISTORICO\" (nested).
    If sHistoricos = sActuales & "HISTORICO\" Or sHistoricos = sActuales & "\HISTORICO\" Then
        Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail( _
            "URLCarpetaAnexosHistoricos is NESTED inside URLCarpetaAnexosActuales — this is the old bug. " & _
            "Actuales='" & sActuales & "', Historicos='" & sHistoricos & "'.", logs)
        GoTo CleanUp
    End If
    ' Negative: Historicos MUST NOT derive from the legacy ANEXOS_HISTORICO_PATH
    ' (a DIFFERENT network root). It must share Actuales's root.
    Dim sOldWrongPath As String
    sOldWrongPath = "C:\Users\adm1\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\HPS\ANEXOS\HISTORICO\" & TEST_DNI & "\"
    If sHistoricos = sOldWrongPath Then
        Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail( _
            "URLCarpetaAnexosHistoricos still derives from the old ANEXOS_HISTORICO_PATH root, not from URLCarpetaAnexosActuales. " & _
            "Got='" & sHistoricos & "'.", logs)
        GoTo CleanUp
    End If

    logs.Add "URLCarpetaAnexosActuales  = " & sActuales
    logs.Add "URLCarpetaAnexosHistoricos = " & sHistoricos
    logs.Add "Both share <root>\<DNI>\ prefix; actual anexos land in '\Actual\', historical in '\HISTORICO\' (siblings, not nested)."
    Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonOk("anexos-urls-siblings", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    Set u = Nothing
    Set testEntorno = Nothing
    Set m_ObjEntorno = oldEntorno
    m_EnOficina = eOldEnOficina
    m_AccesoADatosTE = eOldAccesoDatosTE
    Application.TempVars.Remove "DatosEnLocal"
    If bHadDatosEnLocal Then
        Application.TempVars.Add "DatosEnLocal", sOldDatosEnLocal
    End If
    Set db = Nothing
    On Error GoTo 0
    Exit Function
EH:
    Test_AnexoUsuarioHPS_AnexosUrls_AreSiblingsUnderDNI = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    Set m_ObjEntorno = oldEntorno
    On Error GoTo 0
End Function

Public Function Test_PasarAnexo_FailsLoudlyWhenSourceMissing() As String
    ' [anexos subcarpetas] Before the fix, the guard
    '   If fso.FileExists(m_URLArchivoInicial) Then ... End If
    ' silently skipped CopyFile/DeleteFile, and the DB UPDATE at the end of
    ' PasarAnexo still ran — leaving DB and filesystem out of sync
    ' ("no localizado" symptom). The fix must raise p_Error and NOT touch
    ' EsHistorico on the row.
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oldEntorno As entorno
    Dim testEntorno As entorno
    Dim testRoot As String
    Dim anexo As AnexoUsuarioHPS
    Dim preEsHistorico As String
    Dim postEsHistorico As String
    Dim sErr As String
    Dim eOldEnOficina As EnumSiNo
    Dim eOldAccesoDatosTE As EnumSiNo
    Dim sOldDatosEnLocal As String
    Dim bHadDatosEnLocal As Boolean
    Set logs = New Collection
    On Error GoTo EH

    ' 0) Save globals (mirror the pattern of Test_HPSConfig_LocalModeUsesLocalAnexosConfig)
    Set db = CurrentDb
    eOldEnOficina = m_EnOficina
    eOldAccesoDatosTE = m_AccesoADatosTE
    On Error Resume Next
    sOldDatosEnLocal = CStr(Application.TempVars("DatosEnLocal"))
    bHadDatosEnLocal = (Err.Number = 0)
    Err.Clear
    On Error GoTo EH
    Set oldEntorno = m_ObjEntorno

    ' 1) Seed the DB row + a known filesystem root (we will NOT create the
    '    source file — the missing file IS what we are testing).
    EnsureUserAndAnexoInDb
    testRoot = TempPerAnexoRoot("pasaranexo-missing-source")
    If Dir$(testRoot, vbDirectory) = "" Then CreateFolderTree testRoot

    ' 2) Capture pre-state EsHistorico from the DB.
    preEsHistorico = EsHistoricoInDb(TEST_ID_ANEXO)
    If preEsHistorico <> "No" Then
        Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonFail("Pre-state EsHistorico should be 'No' (seeded), got: '" & preEsHistorico & "'.", logs)
        GoTo CleanUp
    End If

    ' 3) Wire a deterministic entorno pointing at testRoot so the derived URL
    '    resolves to <testRoot>\<DNI>\missing-source.pdf (which we never
    '    create on disk). Same pattern as the URLCarpetaAnexosHistoricos
    '    derivation test above.
    On Error Resume Next
    db.Execute "DELETE FROM TbConfiguracionHPS", dbFailOnError
    Err.Clear
    On Error GoTo EH

    SetDatosEnLocalForTest "S" & Chr$(237)
    ConfigurarTablaConfiguracion_PerfilLocal
    m_EnOficina = EnumSiNo.No
    m_AccesoADatosTE = EnumSiNo.Sí

    Set testEntorno = New entorno
    testEntorno.CargarConfiguracion
    Set m_ObjEntorno = testEntorno

    ' 4) Build the AnexoUsuarioHPS instance the same way the form does:
    '    ID + a stub Usuario + EsHistorico. We do NOT call
    '    getAnexoUsuarioHPS() because that would re-read EsHistorico from
    '    the row (already 'No') and we'd lose the intent. Instead we
    '    construct directly so the per-anexo path runs against OUR state.
    Set anexo = New AnexoUsuarioHPS
    anexo.IDAnexo = CStr(TEST_ID_ANEXO)
    anexo.NombreAnexo = "missing-source.pdf"
    anexo.idUsuario = CStr(TEST_ID_USER)
    anexo.EsHistorico = "No"

    ' 5) Touch URLAnexo to force the property GET to compose
    '    m_sURLAnexo = <Entorno.URLCarpetaAnexos>\<DNI>\missing-source.pdf
    '    (we ignore the result — the property GET has the side effect of
    '    populating m_sURLAnexo, which is what PasarAnexo reads).
    Dim sUrl As String
    sUrl = anexo.URLAnexo
    logs.Add "Anexo.URLAnexo resolved to: " & sUrl
    If Len(sUrl) = 0 Then
        Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonFail("Anexo.URLAnexo returned empty; cannot set up missing-source scenario.", logs)
        GoTo CleanUp
    End If
    ' Sanity: the derived URL must NOT exist on disk (we never created it).
    If Dir$(sUrl) <> "" Then
        Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonFail("Derived URL unexpectedly exists on disk: " & sUrl & ". Cannot set up missing-source scenario.", logs)
        GoTo CleanUp
    End If

    ' 6) Call PasarAnexo Sí. The source file does not exist, so the FIXED
    '    implementation must raise p_Error and NOT update EsHistorico.
    Call anexo.PasarAnexo(EnumSiNo.Sí, sErr)
    If Len(sErr) = 0 Then
        Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonFail("PasarAnexo returned no p_Error when source file was missing. " & _
            "Before the fix this would silently update EsHistorico and leave the file in the old location.", logs)
        GoTo CleanUp
    End If
    If InStr(1, sErr, "origen no encontrado", vbTextCompare) = 0 Then
        Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonFail("Expected 'archivo origen no encontrado' in p_Error, got: " & sErr, logs)
        GoTo CleanUp
    End If

    ' 7) Critical assertion: the DB row's EsHistorico must be UNCHANGED.
    '    Before the fix it would be flipped to 'Sí' even though the file
    '    was never moved.
    postEsHistorico = EsHistoricoInDb(TEST_ID_ANEXO)
    If postEsHistorico <> preEsHistorico Then
        Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonFail("EsHistorico changed from '" & preEsHistorico & "' to '" & postEsHistorico & _
            "' despite the source file being missing. DB and filesystem are now out of sync.", logs)
        GoTo CleanUp
    End If

    logs.Add "PasarAnexo raised p_Error and DB EsHistorico stayed '" & postEsHistorico & "' — no DB/filesystem divergence."
    Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonOk("pasaranexo-fails-loudly", logs)
    GoTo CleanUp
CleanUp:
    On Error Resume Next
    Set anexo = Nothing
    Set testEntorno = Nothing
    Set m_ObjEntorno = oldEntorno
    m_EnOficina = eOldEnOficina
    m_AccesoADatosTE = eOldAccesoDatosTE
    Application.TempVars.Remove "DatosEnLocal"
    If bHadDatosEnLocal Then
        Application.TempVars.Add "DatosEnLocal", sOldDatosEnLocal
    End If
    Set db = Nothing
    CleanupUserAndAnexoInDb
    DeleteFolderIfExists testRoot
    On Error GoTo 0
    Exit Function
EH:
    Test_PasarAnexo_FailsLoudlyWhenSourceMissing = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    Set m_ObjEntorno = oldEntorno
    CleanupUserAndAnexoInDb
    DeleteFolderIfExists testRoot
    On Error GoTo 0
End Function
