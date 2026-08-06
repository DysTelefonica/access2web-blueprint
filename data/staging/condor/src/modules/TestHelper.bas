Attribute VB_Name = "TestHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: TestHelper.bas
' RESPONSABILIDAD: Helpers canónicos para baterías de test access-vba-tdd v2.4
'
' PROVEER:
'   - BuildJsonOk(value, logs()) As String
'   - BuildJsonFail(errorMsg, logs()) As String
'   - EscapeJsonString(s) As String
'   - JsonStringArray(logs()) As String
'   - SqlStr(s) As String          — v1.9 escapa apostrophes en SQL
'   - GetTestDb()/CloseTestDb()    — v1.9 canonical test DB lifecycle
'   - AssertSandboxBackend(logs(), sandboxPath, error) As Boolean — guard read-only
'   - AssertLocalBackend() As Boolean  — alias legacy read-only
'   - BeginTestSession(logs(), error) As Boolean — lifecycle canónico
'   - EndTestSession(logs(), error) As Boolean — lifecycle canónico
'   - ResetTestSession(error) — cleanup idempotente
'   - SetupProdGlobalsForTest(error, tempRoot, suiteTag, userName) — v2.4.3 prod env
'   - CleanupProdTempRoot(tempRoot) — v2.4.3 borra tempRoot bajo %TEMP%
'   - SuiteSetup / SuiteTeardown — aliases legacy (compatibilidad)
' ==========================================================================

' --- Estado original del backend (para SuiteTeardown) ---
Private m_OriginalBackendActivo As String

' --- Módulo privado para caching del DB de test ---
Private m_TestDb As DAO.Database

' ==========================================================================
' BuildJsonOk: Construye JSON de test PASSED
'   logs() es un array de strings con traza de ejecución
' ==========================================================================
Public Function BuildJsonOk(ByVal value As String, ByRef logs() As String) As String
    Dim json As String
    json = "{""ok"":true,""value"":""" & EscapeJsonString(value) & """,""payload"":null,""error"":null,""logs"":" & JsonStringArray(logs()) & "}"
    BuildJsonOk = json
End Function

' --------------------------------------------------------------------------
' BuildJsonFail: Construye JSON de test FAILED
'   logs() es un array de strings con traza de ejecución
' --------------------------------------------------------------------------
Public Function BuildJsonFail(ByVal errorMsg As String, ByRef logs() As String) As String
    Dim json As String
    json = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJsonString(errorMsg) & """,""logs"":" & JsonStringArray(logs()) & "}"
    BuildJsonFail = json
End Function

' --------------------------------------------------------------------------
' EscapeJsonString: Escapa caracteres especiales para JSON
' --------------------------------------------------------------------------
Public Function EscapeJsonString(ByVal s As String) As String
    Dim result As String
    result = s
    ' Backslash — must be first to avoid double-escaping
    result = Replace(result, "\", "\\")
    ' Double quotes
    result = Replace(result, """", "\""")
    ' Newlines
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbCr, "\n")
    ' Tabs
    result = Replace(result, vbTab, "\t")
    EscapeJsonString = result
End Function

' --------------------------------------------------------------------------
' JsonStringArray: Convierte VBA String array a JSON array string
' --------------------------------------------------------------------------
Public Function JsonStringArray(ByRef logs() As String) As String
    Dim result As String
    Dim i As Long
    Dim ub As Long
    Dim lb As Long

    On Error Resume Next
    lb = LBound(logs)
    ub = UBound(logs)
    If Err.Number <> 0 Then
        ' Empty array
        JsonStringArray = "[]"
        Exit Function
    End If

    result = "["
    For i = lb To ub
        If i > lb Then result = result & ","
        result = result & """" & EscapeJsonString(logs(i)) & """"
    Next i
    result = result & "]"

    JsonStringArray = result
End Function

' ==========================================================================
' v1.8 CANONICAL ADDITIONS — AssertLocalBackend + Lifecycle
' ==========================================================================

' --------------------------------------------------------------------------
' AssertLocalBackend: Guard para ejecución COM aislada
'   Retorna True si el routing de test apunta a un sandbox local validado
'   Uso: al inicio de cada test individual que puede correr vía COM standalone
' --------------------------------------------------------------------------
Public Function AssertLocalBackend() As Boolean
    Dim bak As String
    Dim sandboxPath As String

    On Error Resume Next
    If m_TestingMode And Len(Trim$(m_BackendSandboxURL)) > 0 Then
        sandboxPath = Trim$(m_BackendSandboxURL)
        AssertLocalBackend = IsSafeCondorSandboxPath(sandboxPath)
        Exit Function
    End If
    bak = Nz(Application.TempVars("BackendActivo"), "")
    AssertLocalBackend = (bak = "TEST" Or bak = "SANDBOX")
End Function



' --------------------------------------------------------------------------
' RestoreBackend: Restaura el backend activo original (para SuiteTeardown)
'   Lee desde TbConfiguracionBackends y actualiza TempVars
' --------------------------------------------------------------------------
Public Sub RestoreBackend()
    Dim cfgError As String
    On Error Resume Next
    Call LeeConfiguracionLocal(cfgError)
End Sub

' ==========================================================================
' SUITELIFECYCLE — SuiteSetup / SuiteTeardown
'   Se llaman UNA SOLA VEZ por batería (no por test individual)
' ==========================================================================

' --------------------------------------------------------------------------
' BeginTestSession: Configura el entorno de test ANTES del primer test
'   - Valida sandbox local sin mutar TbConfiguracionBackends
'   - Activa m_TestingMode y ejecuta Test_EVE(True)
' --------------------------------------------------------------------------
Public Function BeginTestSession(ByRef logs() As String, ByRef p_Error As String) As Boolean
    Dim eveError As String
    Dim sandboxPath As String

    BeginTestSession = False
    p_Error = ""
    ' Guardar backend original
    m_OriginalBackendActivo = Nz(Application.TempVars("BackendActivo"), "")

    ' Validate and force sandbox/local routing before any test write.
    Call ForceLocalBackend(p_Error)
    If p_Error <> "" Then Exit Function

    If Not AssertSandboxBackend(logs, sandboxPath, p_Error) Then
        If Len(p_Error) = 0 Then
            p_Error = "TESTS BLOCKED: sandbox backend assertion failed"
        End If
        Call ResetTestSession
        Exit Function
    End If

    If Not Test_EVE(True, eveError) Then
        p_Error = "TESTS BLOCKED: " & eveError
        Call ResetTestSession
        Exit Function
    End If

    ' Verificar que el guard pasa
    If Not AssertLocalBackend() Then
        p_Error = "TESTS BLOCKED: ForceLocalBackend did not enable safe backend routing"
        Call ResetTestSession
        Exit Function
    End If

    BeginTestSession = True
End Function

' --------------------------------------------------------------------------
' EndTestSession: Limpia el entorno de test DESPUÉS del último test
'   - Ejecuta Test_EVE(False) y resetea routing/cache de test
' --------------------------------------------------------------------------
Public Function EndTestSession(ByRef logs() As String, Optional ByRef p_Error As String = "") As Boolean
    Dim eveError As String

    EndTestSession = False
    p_Error = ""
    If Not Test_EVE(False, eveError) Then
        p_Error = eveError
    End If

    Call ResetTestSession(p_Error)
    Call RestoreBackend
    EndTestSession = (Len(p_Error) = 0)
End Function

' --------------------------------------------------------------------------
' SuiteSetup: alias legacy; usar BeginTestSession en código nuevo.
' --------------------------------------------------------------------------
Public Sub SuiteSetup(ByRef p_Error As String)
    Dim logs() As String
    logs = NewLogsArray(3)
    If Not BeginTestSession(logs, p_Error) Then
        If Len(p_Error) = 0 Then
            p_Error = "TESTS BLOCKED: BeginTestSession returned False"
        End If
    End If
End Sub

' --------------------------------------------------------------------------
' SuiteTeardown: alias legacy; usar EndTestSession en código nuevo.
' --------------------------------------------------------------------------
Public Sub SuiteTeardown(ByRef p_Error As String)
    Dim logs() As String
    logs = NewLogsArray(3)
    If Not EndTestSession(logs, p_Error) Then
        If Len(p_Error) = 0 Then
            p_Error = "EndTestSession returned False"
        End If
    End If
End Sub

' ==========================================================================
' HELPERS DE LOG PARA BATERÍA — traza acumulable
' ==========================================================================

' --------------------------------------------------------------------------
' NewLogsArray: Helper para crear arrays de logs???
'   Usage: Dim logs() As String: logs = NewLogsArray(3)
' --------------------------------------------------------------------------
Public Function NewLogsArray(ByVal size As Long) As String()
    Dim arr() As String
    ReDim arr(0 To size) As String
    NewLogsArray = arr
End Function

' ==========================================================================
' v1.9 CANONICAL ADDITIONS — SqlStr + GetTestDb/CloseTestDb
' ==========================================================================

' --------------------------------------------------------------------------
' SqlStr: Escapa apostrophes para SQL con Replace simple
'   Returns a complete quoted SQL text literal.
'   Use instead of direct concatenation: "'" & value & "'".
' --------------------------------------------------------------------------
Public Function SqlStr(ByVal s As String) As String
    SqlStr = "'" & Replace(s, "'", "''") & "'"
End Function

' --------------------------------------------------------------------------
' ForceLocalBackend: validate sandbox before any fixture write.
'   Reads TbConfiguracionBackends, validates filesystem + DAO open, then enables
'   m_TestingMode so normal getdb() calls route to the sandbox. On any failure,
'   leaves testing mode disabled and returns TESTS BLOCKED through p_Error.
' --------------------------------------------------------------------------
Public Sub ForceLocalBackend(ByRef p_Error As String)
    Dim dbFrontend As DAO.Database
    Dim rs As DAO.Recordset
    Dim sandboxPath As String
    Dim backendPassword As String
    Dim fsoLocal As Object
    Dim dbProbe As DAO.Database

    On Error GoTo EH
    p_Error = ""
    m_TestingMode = False
    m_BackendSandboxURL = ""

    Set dbFrontend = CurrentDb
    Set rs = dbFrontend.OpenRecordset( _
        "SELECT TOP 1 BackendSandbox, BackendTest, PasswordBackend " & _
        "FROM TbConfiguracionBackends WHERE Habilitado=True", _
        dbOpenSnapshot)

    If rs.EOF Then
        p_Error = "TESTS BLOCKED: TbConfiguracionBackends has no enabled row"
        GoTo CleanExit
    End If

    sandboxPath = Trim$(Nz(rs!BackendSandbox, ""))
    If Len(sandboxPath) = 0 Then
        sandboxPath = Trim$(Nz(rs!BackendTest, ""))
    End If
    backendPassword = Nz(rs!PasswordBackend, GetPasswordDB())

    If Len(sandboxPath) = 0 Then
        p_Error = "TESTS BLOCKED: BackendSandbox/BackendTest path is empty"
        GoTo CleanExit
    End If

    If Not IsSafeCondorSandboxPath(sandboxPath) Then
        p_Error = "TESTS BLOCKED: unsafe CONDOR sandbox path: " & sandboxPath
        GoTo CleanExit
    End If

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If Not fsoLocal.FileExists(sandboxPath) Then
        p_Error = "TESTS BLOCKED: sandbox backend not found: " & sandboxPath
        GoTo CleanExit
    End If

    Set dbProbe = DBEngine.Workspaces(0).OpenDatabase(sandboxPath, False, True, ";PWD=" & backendPassword)
    dbProbe.Close
    Set dbProbe = Nothing

    m_BackendSandboxURL = sandboxPath
    m_PasswordBackend = backendPassword
    m_BackendSandboxPassword = backendPassword
    m_TestingMode = True

    ResetGlobals
    m_BackendSandboxURL = sandboxPath
    m_PasswordBackend = backendPassword
    m_BackendSandboxPassword = backendPassword
    m_TestingMode = True

CleanExit:
    On Error Resume Next
    If Not dbProbe Is Nothing Then dbProbe.Close
    Set dbProbe = Nothing
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set dbFrontend = Nothing
    Exit Sub

EH:
    p_Error = "TESTS BLOCKED: ForceLocalBackend failed: " & Err.description
    m_TestingMode = False
    m_BackendSandboxURL = ""
    Resume CleanExit
End Sub

' --------------------------------------------------------------------------
' AssertSandboxBackend: read-only v2.4 guard before enabling fixture writes.
'   Reads the already-resolved sandbox session path when present, otherwise reads
'   TbConfiguracionBackends without mutating configuration rows.
' --------------------------------------------------------------------------
Public Function AssertSandboxBackend(ByRef logs() As String, ByRef p_SandboxPath As String, ByRef p_Error As String) As Boolean
    Dim dbFrontend As DAO.Database
    Dim rs As DAO.Recordset
    Dim sandboxPath As String
    Dim backendPassword As String
    Dim fsoLocal As Object
    Dim dbProbe As DAO.Database

    On Error GoTo EH
    AssertSandboxBackend = False
    p_Error = ""
    p_SandboxPath = ""

    sandboxPath = Trim$(m_BackendSandboxURL)
    backendPassword = m_PasswordBackend

    If Len(sandboxPath) = 0 Then
        Set dbFrontend = CurrentDb
        Set rs = dbFrontend.OpenRecordset( _
            "SELECT TOP 1 BackendSandbox, BackendTest, PasswordBackend " & _
            "FROM TbConfiguracionBackends WHERE Habilitado=True", _
            dbOpenSnapshot)

        If rs.EOF Then
            p_Error = "TESTS BLOCKED: TbConfiguracionBackends has no enabled row"
            GoTo CleanExit
        End If

        sandboxPath = Trim$(Nz(rs!BackendSandbox, ""))
        If Len(sandboxPath) = 0 Then
            sandboxPath = Trim$(Nz(rs!BackendTest, ""))
        End If
        backendPassword = Nz(rs!PasswordBackend, GetPasswordDB())

        ' Cache the resolved password in the canonical test-side name so
        ' CheckAndReconnect can read it without depending on the production
        ' m_PasswordBackend (which LeeConfiguracionLocal may overwrite).
        m_PasswordBackend = backendPassword
        m_BackendSandboxPassword = backendPassword
    End If

    If Len(sandboxPath) = 0 Then
        p_Error = "TESTS BLOCKED: BackendSandbox/BackendTest path is empty"
        GoTo CleanExit
    End If

    If Not IsSafeCondorSandboxPath(sandboxPath) Then
        p_Error = "TESTS BLOCKED: unsafe CONDOR sandbox path: " & sandboxPath
        GoTo CleanExit
    End If

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If fsoLocal Is Nothing Then
        p_Error = "TESTS BLOCKED: FileSystemObject is unavailable"
        GoTo CleanExit
    End If

    If Not fsoLocal.FileExists(sandboxPath) Then
        p_Error = "TESTS BLOCKED: sandbox backend not found: " & sandboxPath
        GoTo CleanExit
    End If

    Set dbProbe = DBEngine.Workspaces(0).OpenDatabase(sandboxPath, False, True, ";PWD=" & backendPassword)
    If dbProbe Is Nothing Then
        p_Error = "TESTS BLOCKED: sandbox backend did not open: " & sandboxPath
        GoTo CleanExit
    End If

    p_SandboxPath = sandboxPath
    AssertSandboxBackend = True

CleanExit:
    On Error Resume Next
    If Not dbProbe Is Nothing Then
        dbProbe.Close
    End If
    Set dbProbe = Nothing
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    Set dbFrontend = Nothing
    Set fsoLocal = Nothing
    Exit Function

EH:
    p_Error = "TESTS BLOCKED: AssertSandboxBackend failed: " & Err.Number & " - " & Err.description
    Resume CleanExit
End Function

' --------------------------------------------------------------------------
' Test_EVE: wrapper canónico para el ciclo de vida de tests v2.4.2.
'   True ejecuta EVE con m_TestingMode ya activo; False no reinicia EVE.
' --------------------------------------------------------------------------
Public Function Test_EVE(ByVal p_Enable As Boolean, ByRef p_Error As String) As Boolean
    On Error GoTo EH
    p_Error = ""

    If p_Enable Then
        Call EVE
    End If

    Test_EVE = True
    Exit Function

EH:
    p_Error = "Test_EVE(" & CStr(p_Enable) & ") failed: " & Err.Number & " - " & Err.description
    Test_EVE = False
End Function

' --------------------------------------------------------------------------
' IsSafeCondorSandboxPath: guard mínimo v2.4.2 para no tocar backends remotos.
' --------------------------------------------------------------------------
Private Function IsSafeCondorSandboxPath(ByVal p_Path As String) As Boolean
    Dim normalizedPath As String
    Dim isUnc As Boolean
    Dim hasCondorFingerprint As Boolean
    Dim hasBlockedFingerprint As Boolean

    normalizedPath = LCase$(Trim$(p_Path))
    If Len(normalizedPath) = 0 Then Exit Function

    isUnc = (Left$(normalizedPath, 2) = "\\")
    hasCondorFingerprint = (InStr(1, normalizedPath, "condor", vbTextCompare) > 0 And _
                            InStr(1, normalizedPath, "_datos.accdb", vbTextCompare) > 0)
    hasBlockedFingerprint = (InStr(1, normalizedPath, "\datoste\", vbTextCompare) > 0 Or _
                             InStr(1, normalizedPath, "aplicaciones_dys", vbTextCompare) > 0 Or _
                             InStr(1, normalizedPath, "aplicaciones ppd", vbTextCompare) > 0)

    If isUnc Then Exit Function
    If hasBlockedFingerprint Then Exit Function
    If Not hasCondorFingerprint Then Exit Function

    IsSafeCondorSandboxPath = True
End Function

' --------------------------------------------------------------------------
' RemoveTempVarIfExists: elimina TempVars de test sin depender de estado previo.
' --------------------------------------------------------------------------
Private Sub RemoveTempVarIfExists(ByVal p_Name As String)
    On Error Resume Next
    Application.TempVars.Remove p_Name
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' ResetTestSession: defensively exits test routing and closes cached test DBs.
' --------------------------------------------------------------------------
Public Sub ResetTestSession(Optional ByRef p_Error As String = "")
    p_Error = ""
    On Error Resume Next
    CloseTestDb
    If Not g_dbCondor Is Nothing Then
        g_dbCondor.Close
        Set g_dbCondor = Nothing
    End If
    m_TestingMode = False
    m_BackendSandboxURL = ""
    m_BackendSandboxPassword = ""
    Call RemoveTempVarIfExists("BackendPathSandbox")
    Call RemoveTempVarIfExists("BackendPathConfigurado")
    Call RemoveTempVarIfExists("DatosEnLocal")
    If Err.Number <> 0 Then
        p_Error = "ResetTestSession cleanup warning: " & Err.Number & " - " & Err.description
    End If
End Sub

' --------------------------------------------------------------------------
' GetTestDb: Cached test database connection (v1.9 canonical)
'   Crea y caching una conexión al backend TEST
'   Uso: llamar una vez en SuiteSetup, usar m_TestDb en fixtures
'   CloseTestDb se llama en SuiteTeardown
' --------------------------------------------------------------------------
Public Function GetTestDb() As DAO.Database
    If m_TestDb Is Nothing Then
        Set m_TestDb = getdb()
    End If
    Set GetTestDb = m_TestDb
End Function

' --------------------------------------------------------------------------
' CloseTestDb: Cierra la conexión cached de test (v1.9 canonical)
'   Llamar en SuiteTeardown SIN EXCEPCION
' --------------------------------------------------------------------------
Public Sub CloseTestDb()
    On Error Resume Next
    If Not m_TestDb Is Nothing Then
        m_TestDb.Close
        Set m_TestDb = Nothing
    End If
End Sub

' ==========================================================================
' v2.4.3 ADDITIONS — Production-environment globals harness
'   EVE() does not set m_ObjEntorno / m_ObjUsuarioActivo /
'   m_URLRutaAplicacionLocal / TempVars("DatosEnLocal") for the production
'   guard path, and the production services depend on those globals. Each
'   _Strict suite was duplicating the same ~15-line block to re-apply them
'   after BeginTestSession; SetupProdGlobalsForTest centralizes that
'   pattern. Lives in TestHelper so every suite shares ONE definition.
' ==========================================================================

' --------------------------------------------------------------------------
' SetupProdGlobalsForTest: Aplica las 4 globales de producción que los
'   code paths de producción necesitan pero que BeginTestSession no aplica.
'   Llamado por las suites _Strict después de BeginTestSession.
'
'   Parametros:
'     p_Error     (ByRef, OUT) — "" si OK; descripción si falló.
'     p_TempRoot  (ByRef, IN/OUT) — Si vacío, crea uno bajo TEMP con el
'                  prefijo p_SuiteTag + timestamp + random. Si viene con
'                  valor, se respeta (la suite lo está reutilizando).
'     p_SuiteTag  (String) — prefijo del folder temporal. Ej: "cap005_cdcasub".
'     p_UserName  (String) — nombre del usuario activo de prueba. Ej: "CAP005 Test User".
'
'   Aplica:
'     1) Application.TempVars("DatosEnLocal") = "SÍ"
'     2) m_URLRutaAplicacionLocal = p_TempRoot & "app\"
'     3) Set m_ObjEntorno = New Entorno
'     4) Set m_ObjUsuarioActivo = New Usuario ; .nombre = p_UserName
'
'   Idempotente: safe to call multiple times in the same session (los
'   globals se sobreescriben con los valores provistos).
'   NO muta TbConfiguracionBackends ni el sandbox URL (eso es BeginTestSession).
' --------------------------------------------------------------------------
Public Sub SetupProdGlobalsForTest(ByRef p_Error As String, _
                                    ByRef p_TempRoot As String, _
                                    ByVal p_SuiteTag As String, _
                                    ByVal p_UserName As String)
    On Error GoTo EH
    Dim tempBase As String

    p_Error = ""

    ' 1. Crear tempRoot si no fue provisto. Mismo formato que las 4 suites
    '    _Strict que ya tenían este patrón inlined.
    If Len(Trim$(p_TempRoot)) = 0 Then
        tempBase = Environ$("TEMP")
        If Right$(tempBase, 1) <> "\" Then tempBase = tempBase & "\"
        p_TempRoot = tempBase & "condor_" & p_SuiteTag & "_" & _
                     Format$(Now, "yyyymmddhhnnss") & "_" & _
                     CStr(Int(Rnd() * 100000)) & "\"
        Call EnsureFolderExists(p_TempRoot)
    End If

    ' 2. Aplicar las 4 globales de producción
    Application.TempVars("DatosEnLocal") = "SÍ"
    m_URLRutaAplicacionLocal = p_TempRoot & "app\"
    Set m_ObjEntorno = New Entorno
    Set m_ObjUsuarioActivo = New Usuario
    m_ObjUsuarioActivo.nombre = p_UserName

    Exit Sub

EH:
    p_Error = "SetupProdGlobalsForTest: " & Err.Description
End Sub

' --------------------------------------------------------------------------
' CleanupProdTempRoot: Borra el folder temporal creado por
'   SetupProdGlobalsForTest. Solo borra si está bajo %TEMP% (defensive
'   guard contra paths productivos). Reemplaza el CleanupTempRoot que
'   estaba duplicado en 5 suites _Strict.
' --------------------------------------------------------------------------
Public Sub CleanupProdTempRoot(ByVal p_TempRoot As String)
    Dim fsoLocal As Object
    Dim tempRoot As String

    On Error Resume Next
    If Len(Trim$(p_TempRoot)) = 0 Then Exit Sub

    tempRoot = LCase$(Trim$(Environ$("TEMP")))
    If Len(tempRoot) > 0 Then
        If Right$(tempRoot, 1) <> "\" Then tempRoot = tempRoot & "\"
        If InStr(1, LCase$(p_TempRoot), tempRoot, vbTextCompare) <> 1 Then Exit Sub
    End If

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If fsoLocal Is Nothing Then Exit Sub

    If fsoLocal.FolderExists(p_TempRoot) Then
        fsoLocal.DeleteFolder p_TempRoot, True
    End If
    Set fsoLocal = Nothing
End Sub

' --------------------------------------------------------------------------
' EnsureFolderExists: Crea la carpeta y todos sus padres si no existe.
'   Reemplaza el EnsureFolder privado duplicado en 5 suites _Strict.
'   Recursivo: si el padre no existe, lo crea primero.
' --------------------------------------------------------------------------
Private Sub EnsureFolderExists(ByVal p_Path As String)
    Dim fsoLocal As Object
    Dim parentPath As String

    On Error GoTo CleanExit
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If fsoLocal Is Nothing Then Exit Sub

    If fsoLocal.FolderExists(p_Path) Then GoTo CleanExit

    parentPath = fsoLocal.GetParentFolderName(Left$(p_Path, Len(p_Path) - 1))
    If Len(parentPath) > 0 Then
        If Not fsoLocal.FolderExists(parentPath) Then
            Call EnsureFolderExists(parentPath & "\")
        End If
    End If
    fsoLocal.CreateFolder p_Path

CleanExit:
    Set fsoLocal = Nothing
End Sub

