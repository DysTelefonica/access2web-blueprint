Attribute VB_Name = "Test_LocalReadAccessAuthorization"
Option Compare Database
Option Explicit

Private Const TEST_TABLE As String = "TbConfiguracionUsuariosAccesoLocal"

Public Function Test_LRA_EntornoAllowlist_LoadsConfiguredUsersFromBackendTable() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim tempPath As String
    Dim entorno As entorno
    Dim colUsuarios As Scripting.Dictionary
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH
    
    tempPath = TempDbPath("loads")
    Set db = CrearBackendTemporalConTabla(tempPath, True, logs)
    db.Execute "INSERT INTO " & TEST_TABLE & " (CorreoUsuario, Activo) VALUES ('usuario.configurado@telefonica.com', True)", dbFailOnError
    db.Execute "INSERT INTO " & TEST_TABLE & " (CorreoUsuario, Activo) VALUES ('usuario.inactivo@telefonica.com', False)", dbFailOnError
    
    Set entorno = New entorno
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(db, errMsg)
    
    If Not colUsuarios.Exists("usuario.configurado@telefonica.com") Then
        Test_LRA_EntornoAllowlist_LoadsConfiguredUsersFromBackendTable = JsonFail("Expected configured active user to be loaded from the injected backend.", logs)
        GoTo SALIR
    End If
    If colUsuarios.Exists("usuario.inactivo@telefonica.com") Then
        Test_LRA_EntornoAllowlist_LoadsConfiguredUsersFromBackendTable = JsonFail("Inactive user must not be authorized.", logs)
        GoTo SALIR
    End If
    If errMsg <> vbNullString Then
        Test_LRA_EntornoAllowlist_LoadsConfiguredUsersFromBackendTable = JsonFail("Loader returned unexpected error: " & errMsg, logs)
        GoTo SALIR
    End If
    
    logs.Add "Loaded active users from injected backend table and ignored inactive rows."
    Test_LRA_EntornoAllowlist_LoadsConfiguredUsersFromBackendTable = JsonOk("local-read-configured-users", logs)
SALIR:
    DisposeTempDb tempPath, db
    Exit Function
EH:
    Test_LRA_EntornoAllowlist_LoadsConfiguredUsersFromBackendTable = JsonFail("Unexpected error: " & Err.Description, logs)
    DisposeTempDb tempPath, db
End Function

Public Function Test_LRA_EntornoAllowlist_IsCaseInsensitive() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim tempPath As String
    Dim entorno As entorno
    Dim colUsuarios As Scripting.Dictionary
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH
    
    tempPath = TempDbPath("case")
    Set db = CrearBackendTemporalConTabla(tempPath, True, logs)
    db.Execute "INSERT INTO " & TEST_TABLE & " (CorreoUsuario, Activo) VALUES ('Usuario.Mixto@Telefonica.Com', True)", dbFailOnError
    
    Set entorno = New entorno
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(db, errMsg)
    
    If Not colUsuarios.Exists("usuario.mixto@telefonica.com") Then
        Test_LRA_EntornoAllowlist_IsCaseInsensitive = JsonFail("Allowlist dictionary must be case-insensitive.", logs)
        GoTo SALIR
    End If
    If colUsuarios.CompareMode <> TextCompare Then
        Test_LRA_EntornoAllowlist_IsCaseInsensitive = JsonFail("Allowlist dictionary CompareMode must be TextCompare.", logs)
        GoTo SALIR
    End If
    If errMsg <> vbNullString Then
        Test_LRA_EntornoAllowlist_IsCaseInsensitive = JsonFail("Loader returned unexpected error: " & errMsg, logs)
        GoTo SALIR
    End If
    
    logs.Add "Email lookup is case-insensitive through TextCompare dictionary."
    Test_LRA_EntornoAllowlist_IsCaseInsensitive = JsonOk("local-read-case-insensitive", logs)
SALIR:
    DisposeTempDb tempPath, db
    Exit Function
EH:
    Test_LRA_EntornoAllowlist_IsCaseInsensitive = JsonFail("Unexpected error: " & Err.Description, logs)
    DisposeTempDb tempPath, db
End Function

Public Function Test_LRA_EntornoAllowlist_CreatesAndSeedsBackendTableWhenMissing() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim tempPath As String
    Dim entorno As entorno
    Dim colUsuarios As Scripting.Dictionary
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH
    
    Set entorno = New entorno
    tempPath = TempDbPath("missing")
    Set db = CrearBackendTemporalSinTabla(tempPath, logs)
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(db, errMsg)
    If errMsg <> vbNullString Then
        Test_LRA_EntornoAllowlist_CreatesAndSeedsBackendTableWhenMissing = JsonFail("Bootstrap should not return an error: " & errMsg, logs)
        GoTo SALIR
    End If
    If Not TablaExiste(db, TEST_TABLE) Then
        Test_LRA_EntornoAllowlist_CreatesAndSeedsBackendTableWhenMissing = JsonFail("Missing backend table should be created automatically.", logs)
        GoTo SALIR
    End If
    If CountRows(db, TEST_TABLE, vbNullString) <> 4 Then
        Test_LRA_EntornoAllowlist_CreatesAndSeedsBackendTableWhenMissing = JsonFail("Bootstrap should insert exactly the four legacy seed rows.", logs)
        GoTo SALIR
    End If
    If Not HasLegacySeeds(colUsuarios) Then
        Test_LRA_EntornoAllowlist_CreatesAndSeedsBackendTableWhenMissing = JsonFail("Bootstrap should load the four legacy seed emails into the allowlist.", logs)
        GoTo SALIR
    End If
    
    logs.Add "Missing backend table was created, seeded, and loaded."
    Test_LRA_EntornoAllowlist_CreatesAndSeedsBackendTableWhenMissing = JsonOk("local-read-bootstrap-created", logs)
SALIR:
    DisposeTempDb tempPath, db
    Exit Function
EH:
    Test_LRA_EntornoAllowlist_CreatesAndSeedsBackendTableWhenMissing = JsonFail("Unexpected error: " & Err.Description, logs)
    DisposeTempDb tempPath, db
End Function

Public Function Test_LRA_EntornoAllowlist_BootstrapIsIdempotent() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim tempPath As String
    Dim entorno As entorno
    Dim colUsuarios As Scripting.Dictionary
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH
    
    Set entorno = New entorno
    tempPath = TempDbPath("idempotent")
    Set db = CrearBackendTemporalSinTabla(tempPath, logs)
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(db, errMsg)
    If errMsg <> vbNullString Then
        Test_LRA_EntornoAllowlist_BootstrapIsIdempotent = JsonFail("First bootstrap returned unexpected error: " & errMsg, logs)
        GoTo SALIR
    End If
    errMsg = vbNullString
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(db, errMsg)
    If errMsg <> vbNullString Then
        Test_LRA_EntornoAllowlist_BootstrapIsIdempotent = JsonFail("Second load returned unexpected error: " & errMsg, logs)
        GoTo SALIR
    End If
    If CountRows(db, TEST_TABLE, vbNullString) <> 4 Then
        Test_LRA_EntornoAllowlist_BootstrapIsIdempotent = JsonFail("Calling loader twice must not duplicate bootstrap seed rows.", logs)
        GoTo SALIR
    End If
    If Not HasLegacySeeds(colUsuarios) Then
        Test_LRA_EntornoAllowlist_BootstrapIsIdempotent = JsonFail("Second load should still return the seeded allowlist.", logs)
        GoTo SALIR
    End If
    
    logs.Add "Repeated loader calls preserved exactly four seed rows."
    Test_LRA_EntornoAllowlist_BootstrapIsIdempotent = JsonOk("local-read-bootstrap-idempotent", logs)
SALIR:
    DisposeTempDb tempPath, db
    Exit Function
EH:
    Test_LRA_EntornoAllowlist_BootstrapIsIdempotent = JsonFail("Unexpected error: " & Err.Description, logs)
    DisposeTempDb tempPath, db
End Function

Public Function Test_LRA_EntornoAllowlist_PreservesExistingEmptyBackendTable() As String
    Dim logs As Collection
    Dim dbEmpty As DAO.Database
    Dim emptyPath As String
    Dim entorno As entorno
    Dim colUsuarios As Scripting.Dictionary
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH
    
    Set entorno = New entorno
    emptyPath = TempDbPath("empty")
    Set dbEmpty = CrearBackendTemporalConTabla(emptyPath, True, logs)
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(dbEmpty, errMsg)
    If colUsuarios.Count <> 0 Then
        Test_LRA_EntornoAllowlist_PreservesExistingEmptyBackendTable = JsonFail("Existing empty table must preserve admin intent and return an empty allowlist.", logs)
        GoTo SALIR
    End If
    If CountRows(dbEmpty, TEST_TABLE, vbNullString) <> 0 Then
        Test_LRA_EntornoAllowlist_PreservesExistingEmptyBackendTable = JsonFail("Existing empty table must not be seeded by runtime bootstrap.", logs)
        GoTo SALIR
    End If
    If InStr(1, errMsg, "no contiene usuarios activos", vbTextCompare) = 0 Then
        Test_LRA_EntornoAllowlist_PreservesExistingEmptyBackendTable = JsonFail("Empty table should report a useful deny-by-default error.", logs)
        GoTo SALIR
    End If
    
    logs.Add "Existing empty configuration table is preserved and denies local read access."
    Test_LRA_EntornoAllowlist_PreservesExistingEmptyBackendTable = JsonOk("local-read-existing-empty-preserved", logs)
SALIR:
    DisposeTempDb emptyPath, dbEmpty
    Exit Function
EH:
    Test_LRA_EntornoAllowlist_PreservesExistingEmptyBackendTable = JsonFail("Unexpected error: " & Err.Description, logs)
    DisposeTempDb emptyPath, dbEmpty
End Function

Public Function Test_LRA_EntornoAllowlist_DoesNotReactivateDisabledLegacySeed() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim tempPath As String
    Dim entorno As entorno
    Dim colUsuarios As Scripting.Dictionary
    Dim errMsg As String
    Dim legacyEmail As String
    Set logs = New Collection
    On Error GoTo EH
    
    legacyEmail = "esperanza.delalamoarriba@telefonica.com"
    Set entorno = New entorno
    tempPath = TempDbPath("disabledlegacy")
    Set db = CrearBackendTemporalConTabla(tempPath, True, logs)
    db.Execute "INSERT INTO " & TEST_TABLE & " (CorreoUsuario, Activo) VALUES ('" & legacyEmail & "', False)", dbFailOnError
    
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(db, errMsg)
    If colUsuarios.Exists(legacyEmail) Then
        Test_LRA_EntornoAllowlist_DoesNotReactivateDisabledLegacySeed = JsonFail("Disabled legacy seed email must not be included in the allowlist.", logs)
        GoTo SALIR
    End If
    If CountRows(db, TEST_TABLE, "CorreoUsuario='" & legacyEmail & "' AND Activo=False") <> 1 Then
        Test_LRA_EntornoAllowlist_DoesNotReactivateDisabledLegacySeed = JsonFail("Existing disabled legacy seed row must remain Activo=False.", logs)
        GoTo SALIR
    End If
    If CountRows(db, TEST_TABLE, "CorreoUsuario='" & legacyEmail & "'") <> 1 Then
        Test_LRA_EntornoAllowlist_DoesNotReactivateDisabledLegacySeed = JsonFail("Loader must not duplicate an existing legacy seed email.", logs)
        GoTo SALIR
    End If
    If InStr(1, errMsg, "no contiene usuarios activos", vbTextCompare) = 0 Then
        Test_LRA_EntornoAllowlist_DoesNotReactivateDisabledLegacySeed = JsonFail("Disabled-only existing table should report deny-by-default error.", logs)
        GoTo SALIR
    End If
    
    logs.Add "Existing disabled legacy seed remained inactive, absent from allowlist, and unique."
    Test_LRA_EntornoAllowlist_DoesNotReactivateDisabledLegacySeed = JsonOk("local-read-disabled-legacy-seed-preserved", logs)
SALIR:
    DisposeTempDb tempPath, db
    Exit Function
EH:
    Test_LRA_EntornoAllowlist_DoesNotReactivateDisabledLegacySeed = JsonFail("Unexpected error: " & Err.Description, logs)
    DisposeTempDb tempPath, db
End Function

Public Function Test_LRA_EntornoAllowlist_TreatsRowsActiveWhenActivoColumnMissing() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim tempPath As String
    Dim entorno As entorno
    Dim colUsuarios As Scripting.Dictionary
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH
    
    tempPath = TempDbPath("noactivo")
    Set db = CrearBackendTemporalConTabla(tempPath, False, logs)
    db.Execute "INSERT INTO " & TEST_TABLE & " (CorreoUsuario) VALUES ('sin.activo@telefonica.com')", dbFailOnError
    
    Set entorno = New entorno
    Set colUsuarios = entorno.CargarUsuariosPermitidosEnLocal(db, errMsg)
    If errMsg <> vbNullString Then
        Test_LRA_EntornoAllowlist_TreatsRowsActiveWhenActivoColumnMissing = JsonFail("Loader returned unexpected error without Activo column: " & errMsg, logs)
        GoTo SALIR
    End If
    If Not colUsuarios.Exists("sin.activo@telefonica.com") Then
        Test_LRA_EntornoAllowlist_TreatsRowsActiveWhenActivoColumnMissing = JsonFail("Rows should be treated as active when optional Activo column is absent.", logs)
        GoTo SALIR
    End If
    
    logs.Add "Existing table without Activo column loaded rows as active."
    Test_LRA_EntornoAllowlist_TreatsRowsActiveWhenActivoColumnMissing = JsonOk("local-read-no-activo-column", logs)
SALIR:
    DisposeTempDb tempPath, db
    Exit Function
EH:
    Test_LRA_EntornoAllowlist_TreatsRowsActiveWhenActivoColumnMissing = JsonFail("Unexpected error: " & Err.Description, logs)
    DisposeTempDb tempPath, db
End Function

Private Function CrearBackendTemporalConTabla(ByVal p_TempPath As String, ByVal p_ConActivo As Boolean, ByRef logs As Collection) As DAO.Database
    Dim db As DAO.Database
    
    Set db = CrearBackendTemporalSinTabla(p_TempPath, logs)
    If p_ConActivo Then
        db.Execute "CREATE TABLE " & TEST_TABLE & " (ID COUNTER PRIMARY KEY, CorreoUsuario TEXT(255), Activo YESNO, FechaAlta DATETIME, Observaciones MEMO)", dbFailOnError
    Else
        db.Execute "CREATE TABLE " & TEST_TABLE & " (ID COUNTER PRIMARY KEY, CorreoUsuario TEXT(255))", dbFailOnError
    End If
    logs.Add "Created temp backend table " & TEST_TABLE & "."
    Set CrearBackendTemporalConTabla = db
End Function

Private Function CrearBackendTemporalSinTabla(ByVal p_TempPath As String, ByRef logs As Collection) As DAO.Database
    If Len(Dir$(p_TempPath)) > 0 Then
        Kill p_TempPath
    End If
    Set CrearBackendTemporalSinTabla = DBEngine.Workspaces(0).CreateDatabase(p_TempPath, dbLangGeneral, dbVersion120)
    logs.Add "Created isolated temp backend DB: " & p_TempPath
End Function

Private Sub DisposeTempDb(ByVal p_TempPath As String, ByRef p_Db As DAO.Database)
    On Error Resume Next
    If Not p_Db Is Nothing Then
        p_Db.Close
        Set p_Db = Nothing
    End If
    If Len(p_TempPath) > 0 Then
        If Len(Dir$(p_TempPath)) > 0 Then
            Kill p_TempPath
        End If
    End If
    On Error GoTo 0
End Sub

Private Function TempDbPath(ByVal p_Suffix As String) As String
    TempDbPath = Environ$("TEMP") & "\hps_lra_" & p_Suffix & "_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 100000)) & ".accdb"
End Function

Private Function TablaExiste(ByRef p_Db As DAO.Database, ByVal p_TableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In p_Db.TableDefs
        If StrComp(tdf.Name, p_TableName, vbTextCompare) = 0 Then
            TablaExiste = True
            Exit Function
        End If
    Next tdf
End Function

Private Function CountRows(ByRef p_Db As DAO.Database, ByVal p_TableName As String, ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    sql = "SELECT Count(*) AS TotalRows FROM " & p_TableName
    If Len(p_Where) > 0 Then
        sql = sql & " WHERE " & p_Where
    End If
    Set rs = p_Db.OpenRecordset(sql, dbOpenSnapshot)
    CountRows = CLng(Nz(rs.Fields("TotalRows").value, 0))
    rs.Close
    Set rs = Nothing
End Function

Private Function HasLegacySeeds(ByRef p_Usuarios As Scripting.Dictionary) As Boolean
    If p_Usuarios Is Nothing Then
        HasLegacySeeds = False
        Exit Function
    End If
    If Not p_Usuarios.Exists("esperanza.delalamoarriba@telefonica.com") Then Exit Function
    If Not p_Usuarios.Exists("andres.romandelperal@telefonica.com") Then Exit Function
    If Not p_Usuarios.Exists("martina.torralbarodriguez@telefonica.com") Then Exit Function
    If Not p_Usuarios.Exists("almudena.cardenasvelloso.ext@telefonica.com") Then Exit Function
    HasLegacySeeds = True
End Function

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
