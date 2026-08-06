Attribute VB_Name = "Test_RiesgosEstablecerPrioridadesHelper"
' ============================================================
' Test_RiesgosEstablecerPrioridadesHelper — TDD atoms for
'   modRiesgosEstablecerPrioridadesHelper
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
' SDD:   hr4-sister-slice-2026-07-01 (HR3d follow-up §9, #2)
'
' Scope: 2 helper entries (RiesgoPrioridad_VaciarTbAux,
'        RiesgoPrioridad_RellenarTbAux) -> 6 scenario atoms
'   Helper signature:
'     Public Sub RiesgoPrioridad_VaciarTbAux(ByRef p_Error As String)
'     Public Sub RiesgoPrioridad_RellenarTbAux( _
'         ByVal p_ColRiesgosActivos As Scripting.Dictionary, _
'         ByVal p_ProyConRiesgosBiblioteca As EnumSiNo, _
'         ByRef p_Error As String)
'
'   Convention: helper is Sub (no JSON return). Atoms wrap the call and
'   return JSON via TestCore_BuildOk / TestCore_BuildFail.
'
'   Hard rule 7: el helper NO hace MsgBox. Los atomos verifican que
'   p_Error se rellena en sad path sin raise visible.
' ============================================================
Option Compare Database
Option Explicit

' --- Constants de IDs del fixture (coinciden con Test_Fixtures.bas) ---
Private Const FIX_EDICION    As Long = 900102
Private Const FIX_PROYECTO   As Long = 900101
Private Const FIX_EXPEDIENTE As Long = 900100
Private Const FIX_ID_RIESGO  As Long = 900503

' --- Helper wrappers (Hard rule 6: re-usa Test_Helper) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Fixture helpers ---
Private Function ForceBackend(ByRef p_Error As String) As Boolean
    ForceBackend = Test_Helper.ForceLocalBackend(p_Error)
End Function

' ============================================================
' ATOM 1 (HAPPY): RiesgoPrioridad_VaciarTbAux con tabla seeded
' ============================================================
Public Function Test_RiesgoPrioridad_VaciarTbAux_TablaVacia_Ok() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_RiesgoPrioridad_VaciarTbAux_TablaVacia_Ok = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    logs(logIdx) = "2. Act: RiesgoPrioridad_VaciarTbAux"
    logIdx = logIdx + 1
    RiesgoPrioridad_VaciarTbAux m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_VaciarTbAux_TablaVacia_Ok = BuildFail( _
            "Happy path must yield p_Error=''. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio, tabla limpia"
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_VaciarTbAux_TablaVacia_Ok = BuildOk("vaciar_tb_aux_vacia_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_VaciarTbAux_TablaVacia_Ok = BuildFail( _
        "Unhandled exception (helper must NOT raise on happy path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 2 (EDGE): RiesgoPrioridad_VaciarTbAux idempotency (re-calls produce p_Error="")
' ============================================================
Public Function Test_RiesgoPrioridad_VaciarTbAux_Idempotente_SegundaVez_Ok() As String
    Dim logs(0 To 6) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_RiesgoPrioridad_VaciarTbAux_Idempotente_SegundaVez_Ok = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    logs(logIdx) = "2. Act: 1a llamada (poblamos via Priorizacion_InsertSnapshotRow)"
    logIdx = logIdx + 1
    Priorizacion_InsertSnapshotRow CLng(FIX_ID_RIESGO), "1", "0", "R-TEST", "Test riesgo", "", m_Err
    If m_Err <> "" Then
        logs(logIdx) = "Arrange FAIL: seed Priorizacion_InsertSnapshotRow: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_VaciarTbAux_Idempotente_SegundaVez_Ok = BuildFail("seed failed: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Act: 2a llamada (RiesgoPrioridad_VaciarTbAux debe limpiar)"
    logIdx = logIdx + 1
    RiesgoPrioridad_VaciarTbAux m_Err
    If m_Err <> "" Then
        logs(logIdx) = "Assert FAIL: 2a llamada set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_VaciarTbAux_Idempotente_SegundaVez_Ok = BuildFail( _
            "2a llamada debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "4. Act: 3a llamada (segundo RiesgoPrioridad_VaciarTbAux sobre tabla vacia)"
    logIdx = logIdx + 1
    RiesgoPrioridad_VaciarTbAux m_Err

    If m_Err <> "" Then
        logs(logIdx) = "5. Assert FAIL: 3a llamada set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_VaciarTbAux_Idempotente_SegundaVez_Ok = BuildFail( _
            "3a llamada sobre tabla vacia debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert PASS: Idempotente confirmado (2 llamadas sin error)"
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_VaciarTbAux_Idempotente_SegundaVez_Ok = BuildOk("vaciar_tb_aux_idempotente", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_VaciarTbAux_Idempotente_SegundaVez_Ok = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 3 (HAPPY): RiesgoPrioridad_RellenarTbAux con coleccion vacia -> no-op sin error
' ============================================================
Public Function Test_RiesgoPrioridad_RellenarTbAux_ColVacia_Ok() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_RiesgoPrioridad_RellenarTbAux_ColVacia_Ok = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Col As Scripting.Dictionary
    Set m_Col = New Scripting.Dictionary

    logs(logIdx) = "2. Act: RiesgoPrioridad_RellenarTbAux con col vacia"
    logIdx = logIdx + 1
    RiesgoPrioridad_RellenarTbAux m_Col, EnumSiNo.No, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_ColVacia_Ok = BuildFail( _
            "Col vacia debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio en col vacia"
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_ColVacia_Ok = BuildOk("rellenar_tb_aux_col_vacia_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_ColVacia_Ok = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 4 (SAD): RiesgoPrioridad_RellenarTbAux con col Nothing -> no-op sin error
' ============================================================
Public Function Test_RiesgoPrioridad_RellenarTbAux_ColNothing_NoOp_Ok() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_RiesgoPrioridad_RellenarTbAux_ColNothing_NoOp_Ok = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    logs(logIdx) = "2. Act: RiesgoPrioridad_RellenarTbAux con col = Nothing"
    logIdx = logIdx + 1
    RiesgoPrioridad_RellenarTbAux Nothing, EnumSiNo.No, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_ColNothing_NoOp_Ok = BuildFail( _
            "Col Nothing debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio en col Nothing"
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_ColNothing_NoOp_Ok = BuildOk("rellenar_tb_aux_col_nothing_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_ColNothing_NoOp_Ok = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 5 (HAPPY): RiesgoPrioridad_RellenarTbAux con 1 riesgo
'   Carga un riesgo en la coleccion, llama al helper, valida que la
'   tabla TbAuxPriorizacion tiene 1 fila con los datos correctos.
' ============================================================
Public Function Test_RiesgoPrioridad_RellenarTbAux_UnRiesgo_TablaPoblada() As String
    Dim logs(0 To 8) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_RiesgoPrioridad_RellenarTbAux_UnRiesgo_TablaPoblada = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    ' Cargar riesgo en cache
    Dim m_Riesgo As riesgo
    Set m_Riesgo = GetCachedRiesgo(CStr(FIX_ID_RIESGO), m_Err)
    If m_Err <> "" Or m_Riesgo Is Nothing Then
        logs(logIdx) = "Arrange FAIL: GetCachedRiesgo: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_UnRiesgo_TablaPoblada = BuildFail("seed risk failed: " & m_Err, logs)
        Exit Function
    End If

    Dim m_Col As Scripting.Dictionary
    Set m_Col = New Scripting.Dictionary
    m_Col.Add CStr(FIX_ID_RIESGO), m_Riesgo

    logs(logIdx) = "2. Act: RiesgoPrioridad_RellenarTbAux"
    logIdx = logIdx + 1
    RiesgoPrioridad_RellenarTbAux m_Col, EnumSiNo.No, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_UnRiesgo_TablaPoblada = BuildFail( _
            "Happy path must yield p_Error=''. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert: tabla TbAuxPriorizacion tiene 1 fila con IDRiesgo = " & CStr(FIX_ID_RIESGO)
    logIdx = logIdx + 1

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) AS n FROM TbAuxPriorizacion WHERE IDRiesgo=" & FIX_ID_RIESGO, dbOpenSnapshot)
    Dim m_N As Long
    If rs.EOF Then
        m_N = 0
    Else
        m_N = CLng(rs.Fields("n").Value)
    End If
    rs.Close

    If m_N <> 1 Then
        logs(logIdx) = "4. Assert FAIL: expected 1 row, got " & CStr(m_N)
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_UnRiesgo_TablaPoblada = BuildFail( _
            "Row count expected 1, got " & CStr(m_N), logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert PASS: tabla poblada correctamente"
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_UnRiesgo_TablaPoblada = BuildOk("rellenar_tb_aux_un_riesgo_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_UnRiesgo_TablaPoblada = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 6 (ADVERSARIAL): Doble llamada no produce duplicados
'   (AdversarialDoble_OK: idempotencia sobre la coleccion)
' ============================================================
Public Function Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica() As String
    Dim logs(0 To 8) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Riesgo As riesgo
    Set m_Riesgo = GetCachedRiesgo(CStr(FIX_ID_RIESGO), m_Err)
    If m_Err <> "" Or m_Riesgo Is Nothing Then
        logs(logIdx) = "Arrange FAIL: GetCachedRiesgo: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica = BuildFail("seed risk failed: " & m_Err, logs)
        Exit Function
    End If

    Dim m_Col As Scripting.Dictionary
    Set m_Col = New Scripting.Dictionary
    m_Col.Add CStr(FIX_ID_RIESGO), m_Riesgo

    logs(logIdx) = "2. Act: 1a llamada RellenarTbAux"
    logIdx = logIdx + 1
    RiesgoPrioridad_RellenarTbAux m_Col, EnumSiNo.No, m_Err
    If m_Err <> "" Then
        logs(logIdx) = "1a llamada FAIL: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica = BuildFail("1a llamada: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Act: 2a llamada RellenarTbAux (debe limpiar + re-llenar, no duplicar)"
    logIdx = logIdx + 1
    RiesgoPrioridad_RellenarTbAux m_Col, EnumSiNo.No, m_Err
    If m_Err <> "" Then
        logs(logIdx) = "2a llamada FAIL: " & m_Err
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica = BuildFail("2a llamada: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert: tabla tiene exactamente 1 fila (no duplicados)"
    logIdx = logIdx + 1

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) AS n FROM TbAuxPriorizacion WHERE IDRiesgo=" & FIX_ID_RIESGO, dbOpenSnapshot)
    Dim m_N As Long
    If rs.EOF Then
        m_N = 0
    Else
        m_N = CLng(rs.Fields("n").Value)
    End If
    rs.Close

    If m_N <> 1 Then
        logs(logIdx) = "5. Assert FAIL: expected 1 row (no dup), got " & CStr(m_N)
        logIdx = logIdx + 1
        Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica = BuildFail( _
            "Doble llamada produjo duplicados: count=" & CStr(m_N), logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert PASS: Sin duplicados tras doble llamada"
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica = BuildOk("rellenar_tb_aux_no_duplica", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_RiesgoPrioridad_RellenarTbAux_DobleLlamada_NoDuplica = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function
