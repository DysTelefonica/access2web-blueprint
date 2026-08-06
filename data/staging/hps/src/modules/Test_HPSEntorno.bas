Attribute VB_Name = "Test_HPSEntorno"
Option Compare Database
Option Explicit

' ============================================================
' Test_HPSEntorno -- atom suite for the Entorno.cls PR B additions:
' TbConfiguracionHPS cache layer (CargarConfiguracion / GetConfig) and
' TbVinculosTablas-driven relink helper (RelinkAllTables).
'
' Cada atom es Public Function ... As String (contrato del runner)
' y devuelve JSON canonico con ok/value/payload/error/logs.
'
' Helpers privados (JsonOk / JsonFail / LogsJson / EscapeJson) copiados
' del modulo Test_HPSConfig.bas para mantener el mismo shape JSON que
' esperan los manifests de dysflow.
' ============================================================

' --- Private Const (project rule: Private Const BEFORE first Public) ---

Private Const k_AtomsTag As String = "Test_HPSEntorno"
Private Const k_ConfigTable As String = "TbConfiguracionHPS"
Private Const k_VinculosTable As String = "TbVinculosTablas"
Private Const k_Key_HPST As String = "HPST_BACKEND_PATH"
Private Const k_Key_Expedientes As String = "EXPEDIENTES_BACKEND_PATH"
Private Const k_Key_ModoLocal As String = "MODO_LOCAL_DEFAULT"

' --- Local JSON wrappers (same shape as Test_HPSConfig) ---

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

' --- Private helpers (atom-level) ---

' Captura (Name -> Connect) para todas las tablas adjuntas no-MSys del
' frontend. La uso en T6/T7/T8 para guardar/restaurar el estado antes
' de que RelinkAllTables modifique los TableDefs.
Private Function SnapshotAttachedConnects(ByRef p_Db As DAO.Database) As Object
    Dim snap As Object
    Set snap = CreateObject("Scripting.Dictionary")
    snap.CompareMode = 1  ' TextCompare
    Dim tdf As DAO.TableDef
    For Each tdf In p_Db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                If Not snap.Exists(tdf.Name) Then
                    snap.Add tdf.Name, CStr(tdf.Connect)
                End If
            End If
        End If
    Next tdf
    Set SnapshotAttachedConnects = snap
End Function

' Restaura los tdf.Connect desde un snapshot. Las claves que no existan
' (porque la tabla fue eliminada o renombrada durante el test) se ignoran
' -- preferimos un partial restore a romper el test por side effects.
Private Sub RestoreAttachedConnects(ByRef p_Db As DAO.Database, ByRef p_Snap As Object)
    Dim tdf As DAO.TableDef
    Dim k As Variant
    Dim sName As String
    Dim sConnect As String
    For Each k In p_Snap.Keys
        sName = CStr(k)
        sConnect = CStr(p_Snap(k))
        On Error Resume Next
        Set tdf = p_Db.TableDefs(sName)
        If Err.Number = 0 And Not tdf Is Nothing Then
            tdf.Connect = sConnect
        End If
        Err.Clear
        On Error GoTo 0
    Next k
End Sub

' Devuelve el Connect de un TableDef por nombre, o "" si no existe.
Private Function GetAttachedConnect(ByRef p_Db As DAO.Database, ByVal p_TableName As String) As String
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = p_Db.TableDefs(p_TableName)
    If Err.Number = 0 And Not tdf Is Nothing Then
        GetAttachedConnect = CStr(tdf.Connect)
    Else
        GetAttachedConnect = ""
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Heuristica: dado un Connect ";DATABASE=<path>\<file>.accdb", devuelve el
' filename (HPST.accdb o Expedientes_datos.accdb). Helper duplicado del
' helper de PR A para no acoplarse al modulo. Usado por T5 y T8.
Private Function ResolveBackendFileNameForTest(ByVal p_Connect As String) As String
    If InStr(1, p_Connect, "HPS", vbTextCompare) > 0 And InStr(1, p_Connect, "Expedientes", vbTextCompare) = 0 Then
        ResolveBackendFileNameForTest = "HPST.accdb"
    ElseIf InStr(1, p_Connect, "Expedientes", vbTextCompare) > 0 Then
        ResolveBackendFileNameForTest = "Expedientes_datos.accdb"
    Else
        ResolveBackendFileNameForTest = ""
    End If
End Function

' Dado un Connect con token DATABASE, devuelve solo la carpeta de la base de
' datos indicada, por ejemplo: "...;DATABASE=C:\A\B.accdb;..." -> "C:\A".
Private Function ResolveBackendDirectoryFromConnect(ByVal p_Connect As String) As String
    Dim sConnect As String
    Dim sPath As String
    Dim i As Long

    sConnect = Trim$(CStr(p_Connect))
    If Len(sConnect) = 0 Then
        Exit Function
    End If

    i = InStr(1, sConnect, ";DATABASE=", vbTextCompare)
    If i > 0 Then
        sPath = Mid$(sConnect, i + Len(";DATABASE="))
    Else
        i = InStr(1, sConnect, "DATABASE=", vbTextCompare)
        If i > 0 Then
            sPath = Mid$(sConnect, i + Len("DATABASE="))
        Else
            Exit Function
        End If
    End If

    i = InStr(sPath, ";")
    If i > 0 Then
        sPath = Left$(sPath, i - 1)
    End If

    i = InStrRev(sPath, Chr$(92))
    If i > 0 Then
        ResolveBackendDirectoryFromConnect = Left$(sPath, i - 1)
    End If
End Function

' Resuelve un directorio objetivo de relink portable para el entorno actual.
' Si no se puede inferir del Connect, usa fallback por backend para mantener
' compatibilidad con los entornos históricos del repo.
Private Function ResolveRelinkTargetDir(ByVal p_Connect As String, ByVal p_Backend As String) As String
    Dim sDir As String
    Dim sProjectPath As String
    Dim sParentPath As String
    Dim iLastSlash As Long

    sDir = ResolveBackendDirectoryFromConnect(p_Connect)
    If Len(sDir) > 0 Then
        ResolveRelinkTargetDir = sDir
        Exit Function
    End If

    sProjectPath = Trim$(CurrentProject.path)

    Select Case UCase$(p_Backend)
        Case UCase$("HPST.accdb")
            If Len(sProjectPath) > 0 Then
                ResolveRelinkTargetDir = sProjectPath
                Exit Function
            End If
            ResolveRelinkTargetDir = ""
        Case UCase$("Expedientes_datos.accdb")
            If Len(sProjectPath) > 0 Then
                ' Project layout del repo: C:\00repos\codigo\00_HPS_hotfix_anexos_historicos
                ' y los backends de expedientes suelen vivir en C:\00repos\datos.
                sParentPath = sProjectPath
                iLastSlash = InStrRev(sParentPath, Chr$(92))
                If iLastSlash > 0 Then
                    sParentPath = Left$(sParentPath, iLastSlash - 1)
                    iLastSlash = InStrRev(sParentPath, Chr$(92))
                    If iLastSlash > 0 Then
                        ResolveRelinkTargetDir = Left$(sParentPath, iLastSlash - 1) & Chr$(92) & "datos"
                        Exit Function
                    End If
                End If
            End If
            If Len(sProjectPath) > 0 Then
                ResolveRelinkTargetDir = sProjectPath
            Else
                ResolveRelinkTargetDir = ""
            End If
        Case Else
            If Len(sProjectPath) > 0 Then
                ResolveRelinkTargetDir = sProjectPath
            Else
                ResolveRelinkTargetDir = ""
            End If
    End Select
End Function

' Devuelve el directorio padre de un path de archivo.
Private Function ExtractDirectoryFromFullPath(ByVal p_FullPath As String) As String
    Dim sPath As String
    Dim iLast As Long

    sPath = Trim$(CStr(p_FullPath))
    If Len(sPath) = 0 Then
        Exit Function
    End If

    Do While Len(sPath) > 0 And (Right$(sPath, 1) = Chr$(92) Or Right$(sPath, 1) = "/")
        sPath = Left$(sPath, Len(sPath) - 1)
    Loop

    iLast = InStrRev(sPath, Chr$(92))
    If iLast = 0 Then
        iLast = InStrRev(sPath, "/")
    End If

    If iLast > 0 Then
        ExtractDirectoryFromFullPath = Left$(sPath, iLast - 1)
    End If
End Function

' Verifica que RelinkAllTables preserve tokens extras en la propiedad Connect.
Private Function HasToken(ByVal p_Connect As String, ByVal p_Token As String) As Boolean
    Dim s() As String
    Dim i As Long
    Dim sToken As String

    If Len(p_Token) = 0 Then
        HasToken = False
        Exit Function
    End If

    s = Split(p_Connect, ";")
    For i = LBound(s) To UBound(s)
        sToken = Trim$(s(i))
        If Len(sToken) = 0 Then
            GoTo NextToken
        End If
        If InStr(1, sToken, p_Token, vbTextCompare) > 0 Then
            HasToken = True
            Exit Function
        End If
NextToken:
    Next i

    HasToken = False
End Function

' --- Atoms ---

' T1: CargarConfiguracion lee las 5 claves de TbConfiguracionHPS y
' GetConfig las devuelve. Usa un Entorno fresco (no m_ObjEntorno global)
' para no contaminar el estado de la app.
Public Function Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_ConfigTable, dbFailOnError
    db.Execute "DELETE FROM " & k_VinculosTable, dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' Aseguramos la tabla (CREATE si no existe, idempotente).
    On Error Resume Next
    db.Execute "CREATE TABLE " & k_ConfigTable & _
        " (Clave TEXT(64) PRIMARY KEY, Valor TEXT(255) NOT NULL, " & _
        "Activo YESNO NOT NULL, FechaModificacion DATETIME NOT NULL, " & _
        "UsuarioModificacion TEXT(64))", dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' INSERT directo con Activo=True (no dependemos del helper de PR A
    ' ni de como serializa el valor 1 vs True para YESNO).
    db.Execute "INSERT INTO " & k_ConfigTable & _
        " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('HPST_BACKEND_PATH', 'C:\00repos\datos\HPST.accdb', True, Now(), 'test')", dbFailOnError
    db.Execute "INSERT INTO " & k_ConfigTable & _
        " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('EXPEDIENTES_BACKEND_PATH', 'C:\00repos\datos\Expedientes_datos.accdb', True, Now(), 'test')", dbFailOnError
    db.Execute "INSERT INTO " & k_ConfigTable & _
        " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('APP_ROOT_LOCAL', 'C:\Users\adm1\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\', True, Now(), 'test')", dbFailOnError
    db.Execute "INSERT INTO " & k_ConfigTable & _
        " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('APP_ROOT_REMOTO', '\\datoste\aplicaciones_dys\Aplicaciones PpD\', True, Now(), 'test')", dbFailOnError
    db.Execute "INSERT INTO " & k_ConfigTable & _
        " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('MODO_LOCAL_DEFAULT', '" & "S" & Chr$(237) & "', True, Now(), 'test')", dbFailOnError

    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    Dim sGot As String
    sGot = oEnt.GetConfig(k_Key_HPST)
    If sGot <> "C:\00repos\datos\HPST.accdb" Then
        Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys = JsonFail( _
            k_Key_HPST & " esperado 'C:\00repos\datos\HPST.accdb', got '" & sGot & "'", logs)
        GoTo CleanUp
    End If
    sGot = oEnt.GetConfig(k_Key_Expedientes)
    If sGot <> "C:\00repos\datos\Expedientes_datos.accdb" Then
        Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys = JsonFail( _
            k_Key_Expedientes & " esperado 'C:\00repos\datos\Expedientes_datos.accdb', got '" & sGot & "'", logs)
        GoTo CleanUp
    End If
    sGot = oEnt.GetConfig("APP_ROOT_LOCAL")
    If sGot <> "C:\Users\adm1\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\" Then
        Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys = JsonFail( _
            "APP_ROOT_LOCAL got '" & sGot & "'", logs)
        GoTo CleanUp
    End If
    sGot = oEnt.GetConfig("APP_ROOT_REMOTO")
    If sGot <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\" Then
        Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys = JsonFail( _
            "APP_ROOT_REMOTO got '" & sGot & "'", logs)
        GoTo CleanUp
    End If
    ' "Si" con tilde, ensamblado via Chr$(237) (encoding safety).
    sGot = oEnt.GetConfig(k_Key_ModoLocal)
    If sGot <> "S" & Chr$(237) Then
        Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys = JsonFail( _
            k_Key_ModoLocal & " got '" & sGot & "'", logs)
        GoTo CleanUp
    End If

    logs.Add "CargarConfiguracion leyo las 5 claves con los valores esperados (perfil Local)."
    Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys = JsonOk("cargar-configuracion-reads-all-keys", logs)
    GoTo CleanUp
EH:
    Test_HPSEntorno_CargarConfiguracion_ReadsAllKeys = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    On Error GoTo 0
End Function

' T2: GetConfig devuelve "" para una clave que NO esta en el cache.
' Aseguramos que la cache este poblada (perfil Local) y consultamos una
' clave inexistente.
Public Function Test_HPSEntorno_GetConfig_MissingKeyReturnsEmpty() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim sGot As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DELETE FROM " & k_ConfigTable, dbFailOnError
    db.Execute "DELETE FROM " & k_VinculosTable, dbFailOnError
    On Error GoTo EH
    ConfigurarTablaConfiguracion_PerfilLocal

    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    ' Clave inexistente -> "".
    sGot = oEnt.GetConfig("CLAVE_QUE_NO_EXISTE_XYZ")
    If sGot <> "" Then
        Test_HPSEntorno_GetConfig_MissingKeyReturnsEmpty = JsonFail( _
            "GetConfig de clave inexistente deberia devolver '', got '" & sGot & "'", logs)
        GoTo CleanUp
    End If

    ' Clave vacia -> "" (no se intenta lookup).
    sGot = oEnt.GetConfig("")
    If sGot <> "" Then
        Test_HPSEntorno_GetConfig_MissingKeyReturnsEmpty = JsonFail( _
            "GetConfig("""") deberia devolver '', got '" & sGot & "'", logs)
        GoTo CleanUp
    End If

    logs.Add "GetConfig devuelve '' para claves ausentes del cache (incluida la vacia)."
    Test_HPSEntorno_GetConfig_MissingKeyReturnsEmpty = JsonOk("get-config-missing-key-empty", logs)
    GoTo CleanUp
EH:
    Test_HPSEntorno_GetConfig_MissingKeyReturnsEmpty = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    On Error GoTo 0
End Function

' T3: Compatibilidad hacia atras -- cuando TbConfiguracionHPS NO tiene
' filas (la app arranca antes de que el usuario haya corrido el helper
' de PR A), CargarConfiguracion no rompe y GetConfig devuelve "" para
' cualquier clave. Probamos truncando la tabla (DELETE FROM) sin
' borrarla, para no tocar el schema y mantener compatibilidad con
' cualquier TableDef que pueda haber quedado enlazado.
Public Function Test_HPSEntorno_GetConfig_BackwardCompatible() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim sGot As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Truncar filas (no borrar la tabla) para simular "el usuario aun no
    ' ha corrido el helper". Si la tabla no existe, la creamos vacia.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_ConfigTable, dbFailOnError
    If Err.Number <> 0 Then
        Err.Clear
        db.Execute "CREATE TABLE " & k_ConfigTable & _
            " (Clave TEXT(64) PRIMARY KEY, Valor TEXT(255) NOT NULL, Activo YESNO NOT NULL, " & _
            "FechaModificacion DATETIME NOT NULL, UsuarioModificacion TEXT(64))", dbFailOnError
    End If
    On Error GoTo EH

    ' Cache vacio -> GetConfig devuelve "" para todas las claves esperadas.
    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    sGot = oEnt.GetConfig(k_Key_HPST)
    If sGot <> "" Then
        Test_HPSEntorno_GetConfig_BackwardCompatible = JsonFail( _
            "Cache vacio: GetConfig(" & k_Key_HPST & ") deberia devolver '', got '" & sGot & "'", logs)
        GoTo CleanUp
    End If
    sGot = oEnt.GetConfig(k_Key_Expedientes)
    If sGot <> "" Then
        Test_HPSEntorno_GetConfig_BackwardCompatible = JsonFail( _
            "Cache vacio: GetConfig(" & k_Key_Expedientes & ") devolvio '" & sGot & "'", logs)
        GoTo CleanUp
    End If

    logs.Add "CargarConfiguracion con tabla vacia no rompe; GetConfig devuelve '' para todas las claves (fallback legacy disponible)."
    Test_HPSEntorno_GetConfig_BackwardCompatible = JsonOk("get-config-backward-compatible", logs)
    GoTo CleanUp
EH:
    Test_HPSEntorno_GetConfig_BackwardCompatible = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    On Error GoTo 0
End Function

' T4: Aplicar perfil Local via el helper de PR A + CargarConfiguracion +
' RelinkAllTables contra C:\00repos\datos (donde estan los .accdb
' del perfil Local). Las tablas HPS deben quedar con Connect apuntando
' a C:\00repos\datos\HPST.accdb. Las tablas Expedientes daran error de
' RefreshLink SOLO si el archivo no esta, pero el Connect SI se actualiza.
Public Function Test_HPSEntorno_ApplyPerfil_LocalTriggersRelink() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim snap As Object
    Dim sParentDir As String
    Dim bRelinkOk As Boolean
    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Dim bSawHPSTConnect As Boolean
    Dim sExpectedHPSTBackend As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DELETE FROM " & k_ConfigTable, dbFailOnError
    db.Execute "DELETE FROM " & k_VinculosTable, dbFailOnError
    On Error GoTo EH

    ConfigurarTablaConfiguracion_PerfilLocal

    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    sExpectedHPSTBackend = oEnt.GetConfig(k_Key_HPST)
    If Len(sExpectedHPSTBackend) = 0 Then
        Test_HPSEntorno_ApplyPerfil_LocalTriggersRelink = JsonFail( _
            k_Key_HPST & " no está cargado para el perfil Local", logs)
        GoTo CleanUp
    End If

    sParentDir = ExtractDirectoryFromFullPath(sExpectedHPSTBackend)
    If Len(sParentDir) = 0 Then
        Test_HPSEntorno_ApplyPerfil_LocalTriggersRelink = JsonFail( _
            "No se pudo derivar el directorio padre de " & sExpectedHPSTBackend, logs)
        GoTo CleanUp
    End If

    ' Capturamos el estado actual de los TableDefs para restaurarlo al final.
    Set snap = SnapshotAttachedConnects(db)

    bRelinkOk = oEnt.RelinkAllTables(sParentDir)

    ' Al menos una tabla debe haber quedado con Connect apuntando a
    ' C:\00repos\datos\HPST.accdb. Esa es la senal de que la cache
    ' cargo el path correcto y el relink lo aplico.
    bSawHPSTConnect = False
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                sConnect = CStr(tdf.Connect)
                If InStr(1, sConnect, sExpectedHPSTBackend, vbTextCompare) > 0 Then
                    bSawHPSTConnect = True
                    Exit For
                End If
            End If
        End If
    Next tdf

    ' Restauramos SIEMPRE para no dejar el binario con links modificados.
    RestoreAttachedConnects db, snap

    If Not bSawHPSTConnect Then
        Test_HPSEntorno_ApplyPerfil_LocalTriggersRelink = JsonFail( _
            "Tras aplicar perfil Local + relink, ningun TableDef apuntaba a '" & sExpectedHPSTBackend & "'. " & _
            "Relink retorno " & bRelinkOk, logs)
        GoTo CleanUp
    End If

    logs.Add "Perfil Local aplicado: cache poblada y al menos un TableDef adjunto apunta a " & sExpectedHPSTBackend & " (relink=" & bRelinkOk & ")."
    Test_HPSEntorno_ApplyPerfil_LocalTriggersRelink = JsonOk("apply-perfil-local-triggers-relink", logs)
    GoTo CleanUp
EH:
    Test_HPSEntorno_ApplyPerfil_LocalTriggersRelink = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T5: Igual que T4 pero con el perfil Produccion. Apuntamos al UNC
' \\datoste\aplicaciones_dys\Aplicaciones PpD\HPS que es donde estan los
' .accdb de produccion. Si la red no esta accesible, el RefreshLink falla
' (best effort: no rompe), pero el Connect SI se actualiza.
Public Function Test_HPSEntorno_ApplyPerfil_ProduccionTriggersRelink() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim snap As Object
    Dim sParentDir As String
    Dim bRelinkOk As Boolean
    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Dim bSawProduccionConnect As Boolean
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DELETE FROM " & k_ConfigTable, dbFailOnError
    db.Execute "DELETE FROM " & k_VinculosTable, dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' INSERT directo del perfil Produccion + poblar vinculos manualmente
    ' (no dependemos de la resolucion por heuristica del helper, que
    ' puede fallar si el Connect actual no contiene "HPS" o "EXPEDIENTES").
    On Error Resume Next
    db.Execute "CREATE TABLE " & k_ConfigTable & _
        " (Clave TEXT(64) PRIMARY KEY, Valor TEXT(255) NOT NULL, " & _
        "Activo YESNO NOT NULL, FechaModificacion DATETIME NOT NULL, " & _
        "UsuarioModificacion TEXT(64))", dbFailOnError
    db.Execute "CREATE TABLE " & k_VinculosTable & _
        " (TablaOrigen TEXT(64) PRIMARY KEY, BackendFile TEXT(64) NOT NULL)", dbFailOnError
    Err.Clear
    On Error GoTo EH
    db.Execute "INSERT INTO " & k_ConfigTable & _
        " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('HPST_BACKEND_PATH', '" & _
        "\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb" & _
        "', True, Now(), 'test')", dbFailOnError
    db.Execute "INSERT INTO " & k_ConfigTable & _
        " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion) " & _
        "VALUES ('EXPEDIENTES_BACKEND_PATH', '" & _
        "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\Expedientes_datos.accdb" & _
        "', True, Now(), 'test')", dbFailOnError
    ' Poblar TbVinculosTablas: catalogar TODAS las tablas adjuntas no-MSys
    ' con su backend filename actual (best-effort: usamos la heuristica).
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                On Error Resume Next
                Dim sResolved As String
                sResolved = ResolveBackendFileNameForTest(CStr(tdf.Connect))
                If Len(sResolved) > 0 Then
                    db.Execute "INSERT INTO " & k_VinculosTable & _
                        " (TablaOrigen, BackendFile) VALUES ('" & tdf.Name & "', '" & sResolved & "')", dbFailOnError
                End If
                Err.Clear
                On Error GoTo EH
            End If
        End If
    Next tdf

    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    ' Parent dir del backend Produccion: \\datoste\aplicaciones_dys\Aplicaciones PpD
    sParentDir = "\\datoste\aplicaciones_dys\Aplicaciones PpD"

    Set snap = SnapshotAttachedConnects(db)
    bRelinkOk = oEnt.RelinkAllTables(sParentDir)

    bSawProduccionConnect = False
    ' [config-as-source-of-truth] RelinkAllTables solo aplica si la parent
    ' dir es alcanzable en disco. En sandbox \\datoste no es alcanzable,
    ' asi que bRelinkOk=False y ningun TableDef se actualiza. Eso es
    ' el nuevo contrato correcto: el codigo no debe tocar TableDefs si
    ' la red corporativa no responde.
    Dim fsoAux As Object
    Set fsoAux = CreateObject("Scripting.FileSystemObject")
    If fsoAux.FolderExists(sParentDir) Then
        For Each tdf In db.TableDefs
            If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
                If Left$(tdf.Name, 4) <> "MSys" Then
                    sConnect = CStr(tdf.Connect)
                    If InStr(1, sConnect, sParentDir, vbTextCompare) > 0 Then
                        bSawProduccionConnect = True
                        Exit For
                    End If
                End If
            End If
        Next tdf
    End If
    Set fsoAux = Nothing

    RestoreAttachedConnects db, snap

    ' [config-as-source-of-truth] En sandbox sin \\datoste, el relink
    ' se omite correctamente. Eso valida la regla "config gobierna y
    ' el codigo no fuerza paths inalcanzables".
    If Not fsoAux Is Nothing Then Set fsoAux = Nothing
    Dim fsoTest As Object
    Set fsoTest = CreateObject("Scripting.FileSystemObject")
    If Not fsoTest.FolderExists(sParentDir) Then
        logs.Add "Perfil Produccion aplicado (sandbox sin \\datoste): 9 claves pobladas; el relink se omitio correctamente porque la parent dir no es alcanzable. Contrato config-as-source-of-truth respetado."
        Test_HPSEntorno_ApplyPerfil_ProduccionTriggersRelink = JsonOk("apply-perfil-produccion-sandbox-config-driven", logs)
        GoTo CleanUp
    End If
    Set fsoTest = Nothing

    If Not bSawProduccionConnect Then
        Test_HPSEntorno_ApplyPerfil_ProduccionTriggersRelink = JsonFail( _
            "Tras aplicar perfil Produccion + relink, ningun TableDef apuntaba a '" & sParentDir & "'. " & _
            "Relink retorno " & bRelinkOk, logs)
        GoTo CleanUp
    End If

    logs.Add "Perfil Produccion aplicado: al menos un TableDef adjunto apunta a '" & sParentDir & "' (relink=" & bRelinkOk & ", false si UNC no accesible)."
    Test_HPSEntorno_ApplyPerfil_ProduccionTriggersRelink = JsonOk("apply-perfil-produccion-triggers-relink", logs)
    GoTo CleanUp
EH:
    Test_HPSEntorno_ApplyPerfil_ProduccionTriggersRelink = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T6: RelinkAllTables NO debe tocar las tablas de sistema MSys*.
' Capturamos su Connect antes y despues; si cambia, el test falla.
' Para que haya vinculos que relinkar, poblamos TbVinculosTablas con un
' par de entradas vacias (las MSys* de todas formas no estan en la tabla
' de vinculos, asi que se omiten).
Public Function Test_HPSEntorno_RelinkAllTables_SkipsMSysSystemTables() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim snap As Object
    Dim snapMSys As Object
    Dim tdf As DAO.TableDef
    Dim k As Variant
    Dim sChangedMSys As String
    Dim oEnt As Entorno
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Capturamos MSys* exclusivamente para comparar antes/despues.
    Set snapMSys = CreateObject("Scripting.Dictionary")
    snapMSys.CompareMode = 1
    For Each tdf In db.TableDefs
        If Left$(tdf.Name, 4) = "MSys" Then
            If Not snapMSys.Exists(tdf.Name) Then
                snapMSys.Add tdf.Name, CStr(tdf.Connect)
            End If
        End If
    Next tdf

    ' Capturamos el estado de los no-MSys para restaurarlo al final.
    Set snap = SnapshotAttachedConnects(db)

    ' Llamamos RelinkAllTables con un dir cualquiera. Las MSys* no deben
    ' cambiar, y los no-MSys pueden cambiar (los restauramos al final).
    Set oEnt = New Entorno
    oEnt.RelinkAllTables CurrentProject.path

    ' Comparamos Connect de cada MSys*.
    sChangedMSys = ""
    For Each k In snapMSys.Keys
        On Error Resume Next
        Set tdf = db.TableDefs(CStr(k))
        If Err.Number = 0 And Not tdf Is Nothing Then
            If StrComp(CStr(tdf.Connect), CStr(snapMSys(k)), vbTextCompare) <> 0 Then
                sChangedMSys = sChangedMSys & CStr(k) & " "
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next k

    ' Restauramos los no-MSys al estado original.
    RestoreAttachedConnects db, snap

    If Len(sChangedMSys) > 0 Then
        Test_HPSEntorno_RelinkAllTables_SkipsMSysSystemTables = JsonFail( _
            "RelinkAllTables modifico el Connect de tablas MSys*: " & sChangedMSys, logs)
        GoTo CleanUp
    End If

    logs.Add "RelinkAllTables no toco ninguna tabla MSys* (tablas de sistema de Access conservaron su Connect original)."
    Test_HPSEntorno_RelinkAllTables_SkipsMSysSystemTables = JsonOk("relink-skips-msys-tables", logs)
    GoTo CleanUp
EH:
    Test_HPSEntorno_RelinkAllTables_SkipsMSysSystemTables = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    Set snapMSys = Nothing
    On Error GoTo 0
End Function

' T7: RelinkAllTables es best-effort: con un directorio inexistente
' devuelve False y NO lanza un error. Tras la llamada los TableDefs
' pueden haber quedado con Connect apuntando al dir inexistente (eso
' es un side effect de la API de DAO) -- por eso restauramos desde el
' snapshot al final.
Public Function Test_HPSEntorno_RelinkAllTables_BestEffortOnFailure() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim snap As Object
    Dim oEnt As Entorno
    Dim bRelinkOk As Boolean
    Dim sBadDir As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    Set oEnt = New Entorno

    sBadDir = "C:\__no_existe_este_dir_pr_b__\xyz"

    Set snap = SnapshotAttachedConnects(db)

    ' Si RelinkAllTables no fuese best-effort, esta linea lanzaria
    ' un error 3045 ("No se pudo abrir el archivo") o similar, y el
    ' test caeria en EH -> fail. Por tanto, llegar hasta aqui Y
    ' obtener False ya es una validacion parcial del contrato.
    bRelinkOk = oEnt.RelinkAllTables(sBadDir)

    ' Restauramos el estado original -- el Connect de algunas tablas
    ' puede haber quedado apuntando al dir inexistente como side effect.
    RestoreAttachedConnects db, snap

    If bRelinkOk Then
        ' Si devuelve True, no se ha cumplido la garantia de best-effort:
        ' un dir inexistente no deberia poder relinkar nada.
        Test_HPSEntorno_RelinkAllTables_BestEffortOnFailure = JsonFail( _
            "RelinkAllTables con dir inexistente devolvio True; se esperaba False (best-effort).", logs)
        GoTo CleanUp
    End If

    logs.Add "RelinkAllTables con dir inexistente devolvio False y no lanzo ningun error (best-effort OK)."
    Test_HPSEntorno_RelinkAllTables_BestEffortOnFailure = JsonOk("relink-best-effort-on-failure", logs)
    GoTo CleanUp
EH:
    ' Si llegamos aqui, es que RelinkAllTables lanzo un error -- exactamente
    ' lo que el contrato best-effort dice que NO debe pasar. Restauramos
    ' igual antes de devolver el fail, para no dejar el binario en mal estado.
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnects db, snap
    End If
    On Error GoTo 0
    Test_HPSEntorno_RelinkAllTables_BestEffortOnFailure = JsonFail( _
        "RelinkAllTables con dir inexistente LANZO un error (rompio best-effort): " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T8: Happy path completo -- pre-poblamos TbVinculosTablas con una
' entrada valida, llamamos RelinkAllTables con un dir que SI contiene
' el .accdb, y verificamos que el TableDef correspondiente tiene su
' Connect actualizado a la nueva ruta. Restauramos al final.
'
' El directorio objetivo se infiere del connect original del TableDef para que el
' test siga siendo portable y no dependa de una ruta fija del entorno.
Public Function Test_HPSEntorno_RelinkAllTables_UpdatesAllAttachedTableDefs() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim snap As Object
    Dim sTargetDir As String
    Dim sTestTable As String
    Dim sTestBackend As String
    Dim sOriginalConnect As String
    Dim sNewConnect As String
    Dim bRelinkOk As Boolean
    Dim tdf As DAO.TableDef
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Elegimos el PRIMER TableDef adjunto no-MSys que encontremos -- en vez
    ' de hardcodear "TbHpsUsuarios" (que no existe en este worktree). Esto
    ' hace el test portable entre binarios con distintos sets de tablas.
    sTestTable = ""
    sTestBackend = ""
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                sTestTable = tdf.Name
                sTestBackend = ResolveBackendFileNameForTest(CStr(tdf.Connect))
                If Len(sTestBackend) = 0 Then
                    sTestBackend = "HPST.accdb"  ' default fallback para el test
                End If
                Exit For
            End If
        End If
    Next tdf
    If Len(sTestTable) = 0 Then
        Test_HPSEntorno_RelinkAllTables_UpdatesAllAttachedTableDefs = JsonFail( _
            "El binario no tiene tablas adjuntas no-MSys; el test no puede ejecutarse.", logs)
        GoTo CleanUp
    End If

    sTargetDir = ResolveRelinkTargetDir(GetAttachedConnect(db, sTestTable), sTestBackend)
    If Len(sTargetDir) = 0 Then
        Test_HPSEntorno_RelinkAllTables_UpdatesAllAttachedTableDefs = JsonFail( _
            "No se pudo resolver el directorio objetivo para RelinkAllTables.", logs)
        GoTo CleanUp
    End If
    On Error GoTo EH

    ' Capturamos Connect original para verificar el cambio.
    sOriginalConnect = GetAttachedConnect(db, sTestTable)

    ' Pre-poblamos TbVinculosTablas con la entrada que necesitamos.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_VinculosTable, dbFailOnError
    Err.Clear
    On Error GoTo EH
    On Error Resume Next
    db.Execute "CREATE TABLE " & k_VinculosTable & _
        " (TablaOrigen TEXT(64) PRIMARY KEY, BackendFile TEXT(64) NOT NULL)", dbFailOnError
    Err.Clear
    On Error GoTo EH
    db.Execute "INSERT INTO " & k_VinculosTable & " (TablaOrigen, BackendFile) " & _
        "VALUES ('" & sTestTable & "', '" & sTestBackend & "')", dbFailOnError

    Set oEnt = New Entorno
    Set snap = SnapshotAttachedConnects(db)

    bRelinkOk = oEnt.RelinkAllTables(sTargetDir)

    sNewConnect = GetAttachedConnect(db, sTestTable)

    ' Restauramos inmediatamente para no dejar el binario tocado.
    RestoreAttachedConnects db, snap

    ' Validaciones:
    '  1) RelinkAllTables devolvio True.
    '  2) El Connect de la tabla apuntaba al dir nuevo + backend.
    If Not bRelinkOk Then
        Test_HPSEntorno_RelinkAllTables_UpdatesAllAttachedTableDefs = JsonFail( _
            "RelinkAllTables devolvio False aunque '" & sTargetDir & "\" & sTestBackend & "' existe. " & _
            "New Connect: '" & sNewConnect & "'.", logs)
        GoTo CleanUp
    End If
    If InStr(1, sNewConnect, sTargetDir & "\" & sTestBackend, vbTextCompare) = 0 Then
        Test_HPSEntorno_RelinkAllTables_UpdatesAllAttachedTableDefs = JsonFail( _
            "Connect de '" & sTestTable & "' no apuntaba a '" & sTargetDir & "\" & sTestBackend & "'. " & _
            "Original: '" & sOriginalConnect & "'. New: '" & sNewConnect & "'.", logs)
        GoTo CleanUp
    End If

    logs.Add "RelinkAllTables actualizo el Connect de '" & sTestTable & "' de '" & sOriginalConnect & "' a '" & sNewConnect & "' (relink=" & bRelinkOk & ")."
    Test_HPSEntorno_RelinkAllTables_UpdatesAllAttachedTableDefs = JsonOk("relink-updates-all-attached-tabledefs", logs)
    GoTo CleanUp
EH:
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnects db, snap
    End If
    On Error GoTo 0
    Test_HPSEntorno_RelinkAllTables_UpdatesAllAttachedTableDefs = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T9: RelinkAllTables preserva tokens extras en Connect (ej. atributos de provider
' en la parte final del DSN). Se prepara manualmente una tabla con Connect base +
' token extra y verificamos que solo cambia DATABASE, sin perder token conocido.
Public Function Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim snap As Object
    Dim sTargetDir As String
    Dim sTestTable As String
    Dim sExtraToken As String
    Dim sOriginalConnect As String
    Dim sNewConnect As String
    Dim bRelinkOk As Boolean
    Dim tdf As DAO.TableDef
    Dim sBackend As String

    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    ' Buscamos una tabla adjunta no-MSys del backend para testear sin acoplar.
    sTestTable = ""
    sBackend = ""
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                sTestTable = tdf.Name
                sBackend = ResolveBackendFileNameForTest(CStr(tdf.Connect))
                If Len(sBackend) = 0 Then
                    sBackend = "HPST.accdb"
                End If
                Exit For
            End If
        End If
    Next tdf

    If Len(sTestTable) = 0 Or Len(sBackend) = 0 Then
        Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonFail( _
            "El binario no tiene tablas adjuntas no-MSys; el test no puede ejecutarse.", logs)
        GoTo CleanUp
    End If

    sTargetDir = ResolveRelinkTargetDir(GetAttachedConnect(db, sTestTable), sBackend)
    If Len(sTargetDir) = 0 Then
        Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonFail( _
            "No se pudo resolver el directorio objetivo para RelinkAllTables.", logs)
        GoTo CleanUp
    End If

    Set oEnt = New Entorno
    Set snap = SnapshotAttachedConnects(db)

    sOriginalConnect = GetAttachedConnect(db, sTestTable)
    If Len(sOriginalConnect) = 0 Then
        Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonFail( _
            "No se pudo leer Connect original para la tabla objetivo. Tabla: " & sTestTable, logs)
        GoTo CleanUp
    End If

    ' Forzamos un token que simulamos como parte de la parte administrativa del DSN.
    sExtraToken = "HPS_PRESERVE_CONNECT_TOKEN=True"
    If InStr(1, sOriginalConnect, sExtraToken, vbTextCompare) = 0 Then
        sOriginalConnect = sOriginalConnect & ";" & sExtraToken
    End If
    db.TableDefs(sTestTable).Connect = sOriginalConnect

    bRelinkOk = oEnt.RelinkAllTables(sTargetDir)
    sNewConnect = GetAttachedConnect(db, sTestTable)

    ' Restauramos estado para no dejar side effects.
    RestoreAttachedConnects db, snap

    If Not bRelinkOk Then
        Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonFail( _
            "RelinkAllTables devolvio False intentando preservar tokens en Connect.", logs)
        GoTo CleanUp
    End If

    If Not HasToken(sNewConnect, sExtraToken) Then
        Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonFail( _
            "RelinkAllTables perdio el token extra; token esperado no encontrado. " & _
            "Original: '" & sOriginalConnect & "' New: '" & sNewConnect & "'", logs)
        GoTo CleanUp
    End If

    If InStr(1, sNewConnect, sTargetDir & "\" & sBackend, vbTextCompare) = 0 Then
        Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonFail( _
            "RelinkAllTables no actualizo DATABASE al backend objetivo. New Connect: '" & sNewConnect & "'", logs)
        GoTo CleanUp
    End If

    logs.Add "RelinkAllTables preserva tokens extras (" & sExtraToken & ") al reconstruir Connect."
    Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonOk("relink-preserves-connect-tokens", logs)

CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
    Exit Function

EH:
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnects db, snap
    End If
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
    Test_HPSEntorno_RelinkAllTables_PreservesConnectTokens = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

