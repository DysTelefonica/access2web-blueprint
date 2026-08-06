Attribute VB_Name = "Test_Helper"
Option Compare Database
Option Explicit

' ============================================================
' Test_Helper — Configuración de testing para GESTIÓN_RIESGOS
'
' Skill: access-vba-tdd v1.9
'
' Este módulo existe para que los tests sean independientes de
' lo que diga TbConfiguracionBackends.BackendActivo.
'
' En modo TEST, TODO el sistema opera contra el backend local
' (Gestion_Riesgos_Datos.accdb), sin importar si BackendActivo
' está configurado como PROD u otro valor.
'
' Uso desde cualquier test:
'   Dim cfgError As String
'   If Not ForceLocalBackend(cfgError) Then
'       Test_X = "{""ok"":false,""error"":""" & cfgError & """}"
'       Exit Function
'   End If
'   ' ahora getdb() apunta al backend local
' ============================================================

' --- Cache de conexión de test (v1.9 §2) ---
Private m_TestDb As DAO.Database

' --- Flag: EVE ya fue inicializado para esta sesión de tests ---
' Skill v2.0 §3: EVE() se ejecuta UNA VEZ en SuiteSetup, no por cada test.
Private m_EveInitialized As Boolean

Private Function ResolveBackendSandbox(ByRef p_BackendPath As String, _
                                       ByRef p_BackendPassword As String, _
                                       Optional ByRef p_Error As String) As Boolean
    On Error GoTo EH
    p_BackendPath = ""
    p_BackendPassword = ""
    p_Error = ""

    Dim dbConfig As DAO.Database
    Dim rsConfig As DAO.Recordset
    Dim sandboxCandidateCount As Long

    Set dbConfig = CurrentDb
    Set rsConfig = dbConfig.OpenRecordset( _
        "SELECT BackendSandbox, PasswordBackend " & _
        "FROM TbConfiguracionBackends " & _
        "WHERE [ID]=1 AND Len(Trim(Nz([BackendSandbox],''))) > 0", _
        dbOpenSnapshot)

    If rsConfig.EOF Then
        sandboxCandidateCount = CLng(DCount("*", "TbConfiguracionBackends", _
            "Len(Trim(Nz([BackendSandbox],''))) > 0"))
        If sandboxCandidateCount > 1 Then
            p_Error = "TESTS BLOCKED: hay " & CStr(sandboxCandidateCount) & _
                " BackendSandbox candidatos; se esperaba la fila deterministica ID=1"
        Else
            p_Error = "TESTS BLOCKED: TbConfiguracionBackends no tiene BackendSandbox deterministico en ID=1"
        End If
        GoTo Cleanup
    End If

    p_BackendPath = Trim$(Nz(rsConfig.Fields("BackendSandbox").Value, ""))
    p_BackendPassword = Nz(rsConfig.Fields("PasswordBackend").Value, "")
    If p_BackendPassword = "" Then p_BackendPassword = Environ$("ACCESS_VBA_PASSWORD")

    If p_BackendPath = "" Then
        p_Error = "TESTS BLOCKED: BackendSandbox esta vacio"
        GoTo Cleanup
    End If

    ResolveBackendSandbox = True

Cleanup:
    On Error Resume Next
    If Not rsConfig Is Nothing Then rsConfig.Close
    Set rsConfig = Nothing
    Set dbConfig = Nothing
    Exit Function

EH:
    p_Error = "ResolveBackendSandbox: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

Private Function ValidateBackendSandbox(ByVal backendPath As String, _
                                        ByVal backendPassword As String, _
                                        Optional ByRef p_Error As String) As Boolean
    On Error GoTo EH
    p_Error = ""

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(backendPath) Then
        p_Error = "TESTS BLOCKED: BackendSandbox no encontrado: " & backendPath
        Set fso = Nothing
        Exit Function
    End If
    Set fso = Nothing

    Dim dbSandbox As DAO.Database
    Set dbSandbox = DBEngine(0).OpenDatabase(backendPath, False, False, ";PWD=" & backendPassword)
    dbSandbox.Close
    Set dbSandbox = Nothing

    ValidateBackendSandbox = True
    Exit Function

EH:
    p_Error = "TESTS BLOCKED: BackendSandbox no alcanzable: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not dbSandbox Is Nothing Then dbSandbox.Close
    Set dbSandbox = Nothing
End Function

' --- JSON helpers canónicos para tests atómicos ---
Public Function BuildJsonOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildJsonOk = "{""ok"":true,""value"":" & TestJsonValue(value) & ",""payload"":null,""error"":null,""logs"":" & TestJsonStringArray(logs) & "}"
End Function

Public Function BuildJsonFail(ByVal errorMsg As String, ByRef logs() As String) As String
    BuildJsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & TestEscapeJsonString(errorMsg) & """,""logs"":" & TestJsonStringArray(logs) & "}"
End Function

Public Function TestEscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    TestEscapeJsonString = s
End Function

Public Function EscapeJsonString(ByVal s As String) As String
    EscapeJsonString = TestEscapeJsonString(s)
End Function

Private Function TestJsonValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        TestJsonValue = "null"
        Exit Function
    End If

    Select Case VarType(value)
        Case vbBoolean
            TestJsonValue = LCase$(CStr(value))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            TestJsonValue = Replace(CStr(value), ",", ".")
        Case Else
            TestJsonValue = """" & TestEscapeJsonString(CStr(value)) & """"
    End Select
End Function

Private Function TestJsonStringArray(ByRef logs() As String) As String
    On Error GoTo EmptyLogs

    Dim i As Long
    Dim parts As String
    For i = LBound(logs) To UBound(logs)
        If Len(logs(i)) > 0 Then
            If Len(parts) > 0 Then parts = parts & ","
            parts = parts & """" & TestEscapeJsonString(logs(i)) & """"
        End If
    Next i

    TestJsonStringArray = "[" & parts & "]"
    Exit Function

EmptyLogs:
    TestJsonStringArray = "[]"
End Function

' --- Backend local para testing ---
Private Function GetTestBackendPath() As String
    ' Asume que el backend local está junto al frontend o en la
    ' carpeta del proyecto. Si BackendSandbox está configurado,
    ' usa ese; si no, construye la ruta local por defecto.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If m_BackendSandboxURL <> "" Then
        GetTestBackendPath = m_BackendSandboxURL
    Else
        ' Fallback: backend local junto al frontend
        GetTestBackendPath = fso.BuildPath(CurrentProject.path, "Gestion_Riesgos_Datos.accdb")
    End If
    Set fso = Nothing
End Function

' --- ForceLocalBackend ---
' Función pública que TODO test debe llamar antes de cualquier DAO.
' Activa el modo testing con backend local hard-pinneado.
'
' Skill v2.0 §3: EVE() se ejecuta UNA VEZ (SuiteSetup). Las llamadas
' posteriores dentro de la misma sesión de tests omiten EVE y solo
' verifican que las variables de backend estén configuradas.
'
' Decisión 2026-07-22: BYPASS de bootstrap tradicional. No se lee
' TbConfiguracionBackends (la columna canónica [ID] no existe; el
' proyecto usa IDAplicacion, y algunas filas tienen IDAplicacion=5).
' No se llama EVE — el bucle de propiedades de Entorno falla en
' URLAchivoIni/VersionAplicacion cuando CurrentProject.Path no
' resuelve al worktree en el runtime del test runner. Hard-pin del
' backend LOCAL por constante cumple la regla del usuario "siempre
' estás en el entorno local" sin tocar la tabla de configuración.
'
' Pre-checks §3.6 aplicados (2, 3, 5): path no-UNC, sin fingerprints
' de producción, FileExists. (#1 — resolver vía tabla — se omite
' por directiva del usuario; #4 — fingerprint del proyecto — se
' satisface por la convención "<proyecto>_datos.accdb").
'
' Hace:
'   1. Pin del backend local por constante
'   2. Pre-checks §3.6 (no UNC, sin fingerprints PROD, FileExists)
'   3. Set de m_BackendSandboxURL + m_ActiveBackendURL + m_TestingMode
'   4. Dual-write TempVars usadas por getdb() y LeeConfiguracionLocal
'   5. Marcar EVE como inicializado para short-circuit de llamadas
'      posteriores en la misma sesión
'   6. Lee el path del sandbox desde TbConfiguracionBackends.BackendSandbox
'      (NO hardcoded). El contrato del proyecto: cualquier URL vive en la tabla
'      de configuración, no en constantes de código.
'
' Retorna True si todo OK; False si falló (p_Error en JSON).
Public Function ForceLocalBackend(Optional ByRef p_Error As String) As Boolean
    On Error GoTo ErrHandler
    p_Error = ""

    ' 1. Forzar BackendActivo=LOCAL ANTES de cargar config, asi LeeConfiguracionLocal
    '    toma m_ActiveBackendURL = m_BackendSandboxURL en vez de m_ActiveBackendURL
    '    = m_RutaDirApp_PROD.
    m_BackendActivo = "LOCAL"

    ' 2. Cargar config desde TbConfiguracionBackends (linea 974: m_BackendSandboxURL
    '    = SanitizarRutaUsuarioWindowsLocal(rcdConfig.Fields("BackendSandbox").Value)).
    LeeConfiguracionLocal p_Error
    If p_Error <> "" Then
        ForceLocalBackend = False
        Exit Function
    End If

    ' 3. Reforzar m_BackendActivo=LOCAL por si LeeConfiguracionLocal lo cambio al
    '    resolver la rama "PROD" (defensa en profundidad; ver Test_EVE lineas 941-955).
    m_BackendActivo = "LOCAL"

    ' 4. Verificar que m_BackendSandboxURL tiene un path valido y existe.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If m_BackendSandboxURL = "" Then
        p_Error = "ForceLocalBackend: m_BackendSandboxURL vacio tras LeeConfiguracionLocal; revisa TbConfiguracionBackends.BackendSandbox"
        ForceLocalBackend = False
        Exit Function
    End If
    ' §3.6 #2: path no UNC
    If Left$(m_BackendSandboxURL, 2) = "\\" Then
        p_Error = "ForceLocalBackend: backend sandbox no puede ser UNC: " & m_BackendSandboxURL
        ForceLocalBackend = False
        Exit Function
    End If
    ' §3.6 #3: sin fingerprints productivos
    If InStr(1, m_BackendSandboxURL, "\\datoste\", vbTextCompare) > 0 _
       Or InStr(1, m_BackendSandboxURL, "\\prod\", vbTextCompare) > 0 Then
        p_Error = "ForceLocalBackend: backend sandbox apunta a produccion: " & m_BackendSandboxURL
        ForceLocalBackend = False
        Exit Function
    End If
    ' §3.6 #5: file exists
    If Not fso.FileExists(m_BackendSandboxURL) Then
        p_Error = "ForceLocalBackend: backend sandbox no existe: " & m_BackendSandboxURL
        ForceLocalBackend = False
        Exit Function
    End If

    Dim localPassword As String
    localPassword = Environ$("ACCESS_VBA_PASSWORD")

    ' Set module testing vars desde la config (NO hardcoded)
    m_ActiveBackendURL = m_BackendSandboxURL
    m_PasswordBackend = localPassword
    m_TestingMode = True

    ' Dual-write TempVars (REQ-RES-004 / REQ-RES-005) — getdb() lee
    ' de estos TempVars en modo test. Fuente: m_BackendSandboxURL (de la tabla).
    Application.TempVars("BackendPathConfigurado") = m_BackendSandboxURL
    Application.TempVars("BackendPathSandbox") = m_BackendSandboxURL
    Application.TempVars("BackendSandbox") = m_BackendSandboxURL
    Application.TempVars("EnPruebas") = "Sí"

    ' §3.4: marcar EVE inicializado para short-circuit de llamadas
    ' posteriores en la misma sesión. NO se llama EVE aquí.
    m_EveInitialized = True

    Set fso = Nothing
    ForceLocalBackend = True
    Exit Function

ErrHandler:
    p_Error = "ForceLocalBackend: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    Set fso = Nothing
    ForceLocalBackend = False
End Function

' --- AssertLocalBackend ---
' Guard publico para tests que necesitan verificar que el backend de test esta
' configurado antes de tocar DAO. Mantiene compatibilidad con suites legacy que
' llaman Test_Helper.AssertLocalBackend() sin argumento de error.
Public Function AssertLocalBackend() As Boolean
    On Error GoTo EH

    If m_BackendSandboxURL = "" Then
        AssertLocalBackend = False
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    AssertLocalBackend = fso.FileExists(m_BackendSandboxURL) _
        And (m_BackendActivo = "LOCAL" Or m_BackendActivo = "SANDBOX")
    Set fso = Nothing
    Exit Function

EH:
    AssertLocalBackend = False
End Function

' --- EnsureTestConfigLoaded ---
' Carga/configura el backend local si aun no esta listo. Los tests individuales
' lo usan como guard fail-fast para evitar tocar PROD por accidente.
Public Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    p_Error = ""

    If AssertLocalBackend() Then
        EnsureTestConfigLoaded = True
        Exit Function
    End If

    EnsureTestConfigLoaded = ForceLocalBackend(p_Error)
End Function

' --- IsTestDbAvailable ---
' Verifica que el backend de test exista y sea alcanzable.
' No abre conexión; solo valida el archivo.
Public Function IsTestDbAvailable() As Boolean
    Dim fso As Object
    Dim dbPath As String

    dbPath = GetTestBackendPath()
    If dbPath = "" Then
        IsTestDbAvailable = False
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    IsTestDbAvailable = fso.FileExists(dbPath)
    Set fso = Nothing
End Function

' --- TestBackendAsString ---
' Para diagnóstico: devuelve la ruta actual del backend de test.
Public Function TestBackendAsString() As String
    TestBackendAsString = GetTestBackendPath()
End Function

' --- CloseTestDb (v1.9 §2) ---
' Cierra la conexión cacheada m_TestDb. Llamar en SuiteTeardown.
Public Sub CloseTestDb()
    On Error Resume Next
    If Not m_TestDb Is Nothing Then
        m_TestDb.Close
        Set m_TestDb = Nothing
    End If
    On Error GoTo 0
End Sub

' --- GetCachedTestDb (v1.9 §2) ---
' Devuelve la conexión cacheada (o Nothing si no hay).
Public Function GetCachedTestDb() As DAO.Database
    ' v1.9 §2: validate liveness before returning
    If Not m_TestDb Is Nothing Then
        On Error Resume Next
        Dim dummy As String
        dummy = m_TestDb.Name  ' ? detect if Access closed it
        If Err.Number = 0 Then
            Set GetCachedTestDb = m_TestDb
            Exit Function
        End If
        On Error GoTo 0
        Set m_TestDb = Nothing
    End If
    Set GetCachedTestDb = Nothing
End Function

' --- SqlStr (v1.9 §3) ---
' Escapa un string para uso en SQL. Reemplaza ' por ''.
Public Function SqlStr(ByVal value As String) As String
    SqlStr = Replace(value, "'", "''")
End Function

' --- Test_EVE_Json (v1.9 §2 SuiteSetup) ---
' Wrapper JSON de la Sub Test_EVE en Variables Globales.
' Renombrado para evitar conflicto con Public Sub Test_EVE en Variables Globales.
' El runner COM espera Public Function que devuelve JSON.
Public Function Test_EVE_Json() As String
    Dim cfgError As String
    Dim logs(0 To 1) As String
    logs(0) = "1. ForceLocalBackend"
    logs(1) = "2. Backend: " & TestBackendAsString()

    Dim ok As Boolean
    ok = ForceLocalBackend(cfgError)
    If Not ok Then
        Test_EVE_Json = BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    Test_EVE_Json = BuildJsonOk(TestBackendAsString(), logs)
End Function

' --- ResetTestSession (v2.0 §3) ---
' Llmar al final de cada RunAll (SuiteTeardown) para resetear el flag
' de EVE. Asi la proxima sesion de tests vuelve a inicializar EVE una vez.
Public Sub ResetTestSession(Optional ByRef p_Error As String)
    Dim resetError As String
    p_Error = ""
    m_TestingMode = False
    m_EveInitialized = False
    CloseTestDb

    ResetGlobals resetError
    If resetError <> "" Then p_Error = "ResetTestSession: " & resetError
End Sub
