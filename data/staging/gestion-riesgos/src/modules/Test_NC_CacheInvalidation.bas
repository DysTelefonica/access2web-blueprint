Attribute VB_Name = "Test_NC_CacheInvalidation"
Option Compare Database
Option Explicit

' ============================================================
' Test_NC_CacheInvalidation
'
' Tests de la invalidacion de SearchComboCache desde NC.Registrar y
' NC.Eliminar (issue #35 / cache judgment 24/07/2026).
'
' Schema evidence: Test_Cache.bas ya cubrio el contrato de SearchComboCache
' (cold miss, warm hit, reset, invalidacion quirurgica). Este modulo verifica
' que las mutaciones reales de TbNoConformidades en produccion invalidan
' el cache, no solo los tests del propio modulo SearchComboCache.
'
' Fixture rule: deterministic NC ID 911400 dentro de BeginTrans/Rollback,
' nunca persiste en backend compartido. Test_Helper.ResetTestSession
' garantiza sandbox limpio al final.
'
' Estructura (regla TDD §1.8 — todas las Private antes de las Public):
'   1. Attribute VB_Name + Options
'   2. Header comment
'   3. Private Const (fixtures)
'   4. Private Function helpers
'   5. Public Function atoms
' ============================================================

' --- 3. Private Const (fixtures) ---
Private Const NC_CACHE_FIX_ID As Long = 911400
Private Const NC_CACHE_FIX_EXP As String = "EXP_NC_CACHE_FIXTURE"
Private Const NC_CACHE_FIX_PROY As String = "PROY_NC_CACHE_FIXTURE"
Private Const NC_CACHE_FIX_VEH As String = "VEH_NC_CACHE_FIXTURE"

' --- 4. Private Function helpers ---

' Helper local: equivalente al SearchComboCachePrepareDb de Test_Cache.bas.
' Reusamos ForceLocalBackend + GetTestDb; no necesitamos sentinel check porque
' usamos IDs disjuntos (911400 vs 911350-911352).
Private Function SearchComboCachePrepareDb(ByRef p_db As DAO.Database, ByRef p_Error As String) As Boolean
    p_Error = ""
    If Not Test_Helper.ForceLocalBackend(p_Error) Then
        p_Error = "TESTS BLOCKED: ForceLocalBackend failed: " & p_Error
        Exit Function
    End If

    Set p_db = Test_Fixtures.GetTestDb(p_Error)
    If p_db Is Nothing Then
        p_Error = "TESTS BLOCKED: GetTestDb returned Nothing: " & p_Error
        Exit Function
    End If

    SearchComboCachePrepareDb = True
End Function

' --- 5. Public Function atoms ---

' ============================================================
' ATOMO 1 — NC.Eliminar invalida SearchComboCache
'
' Escenario: Tras un INSERT de NC + warm-cache, una llamada a NC.Eliminar
' debe limpiar el cache para que la siguiente lectura vea el cambio.
'
' Pre: cache HitCount=N (warm).
' Act: NC.Eliminar.
' Post: cache HitCount=0 (Reset fue llamado desde NC.Eliminar).
' ============================================================
Public Function Test_NC_Eliminar_InvalidaSearchComboCache() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim warmValues As Scripting.Dictionary
    Dim postDeleteValues As Scripting.Dictionary
    Dim m_NC As nc
    Dim preDeleteHitCount As Long

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: deterministic NC row (ID=" & NC_CACHE_FIX_ID & ") inside BeginTrans"
    logs(2) = "3. Act: reset + cold load (MissCount=1, HitCount=0)"
    logs(3) = "4. Act: warm load (HitCount=1)"
    logs(4) = "5. Assert pre: HitCount=1 confirms cache is warm"
    logs(5) = "6. Act: NC.Eliminar on the fixture ID"
    logs(6) = "7. Assert post: HitCount must be 0 (Reset was called from Eliminar)"
    logs(7) = "8. Teardown: rollback transaction"

    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True

    db.Execute "INSERT INTO TbNoConformidades " & _
        "(IDNoConformidad, CodigoNoConformidad, EXPEDIENTE, PROYECTO, VEHICULO, Juridica, TIPO, ESTADO, FECHAAPERTURA) " & _
        "VALUES (" & NC_CACHE_FIX_ID & ", 'NC_CACHE_FIXTURE', '" & NC_CACHE_FIX_EXP & "', '" & _
        NC_CACHE_FIX_PROY & "', '" & NC_CACHE_FIX_VEH & "', 'TDE', 'NC', 'REGISTRADA', Date())", _
        dbFailOnError

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail

    ' Cold load: MissCount=1
    Set warmValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If SearchComboCache_MissCount() <> 1 Then
        errMsg = "expected MissCount=1 after cold load, got " & SearchComboCache_MissCount()
        GoTo Fail
    End If
    If SearchComboCache_HitCount() <> 0 Then
        errMsg = "expected HitCount=0 on cold access, got " & SearchComboCache_HitCount()
        GoTo Fail
    End If

    ' Warm load: HitCount=1
    Set warmValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If SearchComboCache_HitCount() <> 1 Then
        errMsg = "expected HitCount=1 on warm access, got " & SearchComboCache_HitCount()
        GoTo Fail
    End If

    preDeleteHitCount = SearchComboCache_HitCount()
    logs(5) = "5. Assert pre: HitCount=" & preDeleteHitCount & " confirms cache is warm"

    ' Act: NC.Eliminar
    Set m_NC = New nc
    m_NC.IDNoConformidad = CStr(NC_CACHE_FIX_ID)
    m_NC.Eliminar errMsg
    If errMsg <> "" Then GoTo Fail

    ' Assert post: NC.Eliminar must have called SearchComboCache_Reset.
    ' After Reset, HitCount and MissCount are both 0.
    If SearchComboCache_HitCount() <> 0 Then
        errMsg = "NC.Eliminar did NOT invalidate SearchComboCache: HitCount=" & _
            SearchComboCache_HitCount() & " (expected 0 after Reset)"
        GoTo Fail
    End If

    ' Sanity: the next GetDistinctValues should be a fresh miss (counter goes 0 -> 1)
    Set postDeleteValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If SearchComboCache_MissCount() <> 1 Then
        errMsg = "expected MissCount=1 after Reset+reload, got " & SearchComboCache_MissCount()
        GoTo Fail
    End If

    Test_NC_Eliminar_InvalidaSearchComboCache = Test_Helper.BuildJsonOk("nc_eliminar_invalida_cache_pass", logs)

Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set warmValues = Nothing
    Set postDeleteValues = Nothing
    Set m_NC = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

Fail:
    Test_NC_Eliminar_InvalidaSearchComboCache = Test_Helper.BuildJsonFail(errMsg, logs)
    GoTo Teardown

EH:
    errMsg = "Test_NC_Eliminar_InvalidaSearchComboCache: " & Err.Description
    GoTo Fail
End Function
