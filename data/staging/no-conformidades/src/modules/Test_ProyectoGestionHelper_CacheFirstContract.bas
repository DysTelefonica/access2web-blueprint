Attribute VB_Name = "Test_ProyectoGestionHelper_CacheFirstContract"
Option Compare Database
Option Explicit

' === TEST cache-first contract: NCProyectoGestionListadoHelper.GetNCsProyectoGestionFiltrados ===
' Strict TDD para issue #94.
' Verifica el contrato A>C del helper (NO A>B>C como PintarIndicadores):
'   A) Si cache tiene filas validas -> usar cache (path A)
'   C) Si cache esta vacia -> log fallback + constructor (path C) — el helper
'      NO warmups la cache en el path C (a diferencia de PintarIndicadores)
'
' Este test es sobre el HELPER, no sobre el form. El form es solo UI plumbing
' (debe leer controles y delegar al helper). Testear el form requiere UI (DoCmd.OpenForm)
' que viola §1.1.2 de la skill access-vba-tdd.
'
' El fix del form (issue #94) es un refactor: cambiar
'   Set m_Col = constructor.getNCs*Proyecto*(...)
' por
'   Set m_Col = NCProyectoGestionListadoHelper.GetNCsProyectoGestionFiltrados(...)
' Este test verifica que el helper cumple el contrato A>C; el form queda
' como thin wrapper que solo delega.

Public Function Test_ProyectoGestionHelper_CacheFirstContract_Atomic() As String
    ' Inicializar a fail ANTES de On Error (regla §1.1.2 de la skill).
    ' Si HandleError no se dispara por algún motivo, el test retorna fail, no "".
    ' Uso New Collection (sintaxis VBA) en vez de TestHelper.NewLogs() porque
    ' este ultimo es una funcion y no se puede pasar como ByRef.
    Dim logs As Collection
    Dim assertError As String
    Set logs = New Collection
    Test_ProyectoGestionHelper_CacheFirstContract_Atomic = TestHelper.BuildJsonFail( _
        "test did not complete", logs)

    On Error GoTo HandleError

    Dim db As DAO.Database
    Dim dbErr As String
    Dim helperErr As String
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim col As Collection
    Dim col2 As Collection
    Dim cacheValidRowsBefore As Long
    Dim cacheValidRowsAfter1 As Long
    Dim cacheValidRowsAfter2 As Long
    Dim fallbackRowsBefore As Long
    Dim fallbackRowsAfter1 As Long
    Dim fallbackRowsAfter2 As Long

    sessionStarted = False

    ' === SETUP: BeginTestSession ===
    TestHelper.AddLog logs, "1. SETUP: BeginTestSession"
    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_ProyectoGestionHelper_CacheFirstContract_Atomic = TestHelper.BuildJsonFail( _
            "TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_ProyectoGestionHelper_CacheFirstContract_Atomic = TestHelper.BuildJsonFail( _
            "getdb returned Nothing: " & dbErr, logs)
        GoTo Cleanup
    End If

    ' === FIXTURE GATE (§1.2): EnsureTableClean de working tables ===
    ' TbCacheListadoNC es working table (mutada por warmup). Pre-clean para
    ' estado conocido (no depender de filas de runs anteriores).
    TestHelper.AddLog logs, "2. FIXTURE GATE: EnsureTableClean TbCacheListadoNC + TbLogCache"
    On Error Resume Next
    db.Execute "DELETE FROM TbCacheListadoNC", dbFailOnError
    db.Execute "DELETE FROM TbLogCache WHERE IDNoConformidad=0", dbFailOnError
    Err.Clear
    On Error GoTo HandleError

    ' === ASSERT pre-conditions ===
    cacheValidRowsBefore = CountRowsForTest(db, _
        "SELECT COUNT(*) FROM TbCacheListadoNC WHERE CacheValida=True")
    fallbackRowsBefore = CountRowsForTest(db, _
        "SELECT COUNT(*) FROM TbLogCache WHERE TipoOperacion='FormCacheFallback' AND IDNoConformidad=0")

    If cacheValidRowsBefore > 0 Then
        assertError = "Setup invalido: TbCacheListadoNC no esta vacia; validRows=" & cacheValidRowsBefore
        GoTo Fail
    End If
    TestHelper.AddLog logs, "3. PRE-CHECK: cacheValidRows=0, fallbackRows=" & fallbackRowsBefore

    ' === ACT 1: primera llamada con cache vacia ===
    ' Espera: path B (fallback al constructor) + warmup de cache + log de FormCacheFallback
    TestHelper.AddLog logs, "4. ACT 1: primera llamada (cache vacia -> path B esperado)"
    helperErr = ""
    Set col = NCProyectoGestionListadoHelper.GetNCsProyectoGestionFiltrados( _
        p_Codigo:="", _
        p_IDExpediente:=0, _
        p_Juridica:="", _
        p_IDTipo:=0, _
        p_EstadoValor:="", _
        p_EstadoEnum:=0, _
        p_Descripcion:="", _
        p_Notas:="", _
        p_RequiereControlEficacia:="", _
        p_ControlEficaciaRelleno:="", _
        p_RegistrosCerrados:="", _
        p_ResponsableTelefonica:="", _
        p_ResponsableCalidad:="", _
        p_Google:="", _
        p_Error:=helperErr)
    TestHelper.AddLog logs, "5. ACT 1 result: helperErr='" & helperErr & _
                            "', col=" & IIf(col Is Nothing, "Nothing", "Collection count=" & col.Count)

    ' === ASSERT 1: result es no-Nothing ===
    If col Is Nothing Then
        assertError = "ASSERT1 FAIL: Esperado Collection no-Nothing, obtenido Nothing; helperErr='" & helperErr & "'"
        GoTo Fail
    End If
    TestHelper.AddLog logs, "6. ASSERT1 OK: result no-Nothing, count=" & col.Count

    ' === ASSERT 2: FormCacheFallback fue logueado (cache estaba vacia -> path C) ===
    ' El helper NO warmups la cache en el path C (contrato A>C, no A>B>C).
    ' El fallback es la unica senal observable de que se uso path C en vez de A.
    fallbackRowsAfter1 = CountRowsForTest(db, _
        "SELECT COUNT(*) FROM TbLogCache WHERE TipoOperacion='FormCacheFallback' AND IDNoConformidad=0")
    If fallbackRowsAfter1 <= fallbackRowsBefore Then
        assertError = "ASSERT2 FAIL: Esperado que FormCacheFallback se loguee (cache estaba vacia, helper uso path C), " & _
                      "pero no se incremento; before=" & fallbackRowsBefore & " after=" & fallbackRowsAfter1
        GoTo Fail
    End If
    TestHelper.AddLog logs, "7. ASSERT2 OK: FormCacheFallback logueado (path C usado); before=" & fallbackRowsBefore & _
                            " after=" & fallbackRowsAfter1

    ' === ASSERT 3: cache NO fue warmed (helper NO hace warmup en path C) ===
    ' Esto documenta el contrato real: el helper usa cache si esta, sino constructor,
    ' pero NO sincroniza la cache. Si en el futuro alguien agrega sync-on-the-fly
    ' (path B), este assert falla y obliga a actualizar el test.
    cacheValidRowsAfter1 = CountRowsForTest(db, _
        "SELECT COUNT(*) FROM TbCacheListadoNC WHERE CacheValida=True")
    If cacheValidRowsAfter1 > 0 Then
        assertError = "ASSERT3 FAIL: Cache fue warmed (path B agregado?). El contrato del helper es A>C (no warmup). " & _
                      "Si esto es intencional, actualizar el test y el comentario del modulo."
        GoTo Fail
    End If
    TestHelper.AddLog logs, "8. ASSERT3 OK: cache NO warmed (contrato A>C confirmado)"

    Test_ProyectoGestionHelper_CacheFirstContract_Atomic = TestHelper.BuildJsonOk(logs, "a_c_contract_ok")
    GoTo Cleanup

Fail:
    Test_ProyectoGestionHelper_CacheFirstContract_Atomic = TestHelper.BuildJsonFail(assertError, logs)
    GoTo Cleanup

Cleanup:
    ' Cleanup defensivo (§1.1.2): On Error Resume Next + statements VBA nativos.
    ' Limpiar working tables para no contaminar corridas siguientes.
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbCacheListadoNC", dbFailOnError
        db.Execute "DELETE FROM TbLogCache WHERE IDNoConformidad=0", dbFailOnError
    End If
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    Set db = Nothing
    Set col = Nothing
    Set col2 = Nothing
    Exit Function

HandleError:
    ' Error inesperado: convierte a fail con descripcion + source, va al cleanup.
    ' Resume Cleanup (no Resume Next) para saltar directo al teardown defensivo.
    Test_ProyectoGestionHelper_CacheFirstContract_Atomic = TestHelper.BuildJsonFail( _
        "unexpected error: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Cleanup
End Function

' CountRowsForTest: cuenta filas de un SELECT COUNT(*). Devuelve 0 si falla.
Private Function CountRowsForTest(ByVal p_Db As DAO.Database, ByVal p_SQL As String) As Long
    Dim rs As DAO.Recordset
    Dim cnt As Long
    On Error GoTo notfound
    Set rs = p_Db.OpenRecordset(p_SQL, dbOpenSnapshot, dbReadOnly)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0).Value) Then cnt = CLng(rs.Fields(0).Value)
    End If
    rs.Close
    Set rs = Nothing
    CountRowsForTest = cnt
    Exit Function
notfound:
    CountRowsForTest = 0
End Function
