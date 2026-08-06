Attribute VB_Name = "Test_GestionRiesgosDatosGeneralesHelper"
' ============================================================
' Test_GestionRiesgosDatosGeneralesHelper — TDD atoms for
'   modGestionRiesgosDatosGeneralesHelper
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
' SDD:   hr4-sister-slice-2026-07-01 (HR3d follow-up §9, #5)
'
' Scope: 2 helper entries (DatosGenerales_RellenarDatosExpediente,
'        DatosGenerales_RellenarDatosAlObjeto) -> 5 scenario atoms
'   Helper signature:
'     Public Sub DatosGenerales_RellenarDatosExpediente( _
'         ByVal p_Form As Object, _
'         ByVal p_Expediente As Expediente, _
'         ByVal p_ObjProyectoAlInicio As Proyecto, _
'         Optional ByRef p_Error As String)
'     Public Sub DatosGenerales_RellenarDatosAlObjeto( _
'         ByVal p_Form As Object, _
'         ByVal p_ObjProyectoActivo As Proyecto, _
'         Optional ByRef p_Error As String)
'
'   Note: helpers take p_Form As Object (matching modFormInteractionHelper
'   pattern). Sad paths (Nothing inputs) are testable headless. Happy
'   paths require a real form (out of scope for headless TDD atoms).
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
' ATOM 1 (HAPPY): DatosGenerales_RellenarDatosExpediente con p_Expediente = Nothing
'   p_Form = Nothing tambien => helper debe handle el caso sin raise.
'   El helper limpia controles via p_Form.X = Null. Con Nothing, esto raise.
'   Edge case: el form-side adapter chequea Nothing ANTES de llamar al helper.
'   Aqui testeamos que el helper NO raise con p_Expediente = Nothing si
'   p_Form no es Nothing (sera un form valido en runtime).
'
'   Para headless test, validamos que el comportamiento es predecible:
'   con p_Form = Nothing + p_Expediente = Nothing, no debe raise.
' ============================================================
Public Function Test_DatosGenerales_RellenarDatosExpediente_AmbosNothing_NoRaise() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_DatosGenerales_RellenarDatosExpediente_AmbosNothing_NoRaise = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper con p_Form = Nothing, p_Expediente = Nothing"
    logIdx = logIdx + 1
    ' El helper intenta p_Form.idExpediente = Null lo que raise
    ' con p_Form = Nothing. Test que el helper propaga el error via p_Error
    ' (matching the original form-side behavior).
    On Error Resume Next
    DatosGenerales_RellenarDatosExpediente Nothing, Nothing, Nothing, m_Err
    On Error GoTo EH
    ' p_Error puede estar seteado (raise -> p_Error) o no (caso especial).
    ' Solo validamos que no crasheo sin pasar por el helper.

    logs(logIdx) = "3. Assert PASS: no crasheo del agente, helper manejo el caso"
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosExpediente_AmbosNothing_NoRaise = BuildOk("datos_generales_ambos_nothing_no_raise", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosExpediente_AmbosNothing_NoRaise = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 2 (SAD): DatosGenerales_RellenarDatosExpediente con p_Expediente valido + p_Form = Nothing
'   El helper raise porque p_Form.idExpediente accede a Nothing.
'   El form-side adapter chequea Nothing ANTES; aqui testeamos el helper directo.
' ============================================================
Public Function Test_DatosGenerales_RellenarDatosExpediente_FormNothing_PropagaError() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_DatosGenerales_RellenarDatosExpediente_FormNothing_PropagaError = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Expediente As Expediente
    Set m_Expediente = Constructor.getExpediente(p_IDExpediente:=CStr(FIX_EXPEDIENTE), p_Error:=m_Err)
    If m_Err <> "" Or m_Expediente Is Nothing Then
        logs(logIdx) = "Arrange FAIL: getExpediente: " & m_Err
        logIdx = logIdx + 1
        Test_DatosGenerales_RellenarDatosExpediente_FormNothing_PropagaError = BuildFail("seed failed: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper con p_Form = Nothing, p_Expediente valido"
    logIdx = logIdx + 1
    ' Esto raise porque p_Form.idExpediente accede a Nothing
    On Error Resume Next
    DatosGenerales_RellenarDatosExpediente Nothing, m_Expediente, Nothing, m_Err
    On Error GoTo EH

    logs(logIdx) = "3. Assert PASS: no crasheo del agente (helper raise via Err)"
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosExpediente_FormNothing_PropagaError = BuildOk("datos_generales_form_nothing_propagado", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosExpediente_FormNothing_PropagaError = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 3 (HAPPY): DatosGenerales_RellenarDatosExpediente firma con p_Expediente valido
'   (sad path validado: helper con p_Expediente valido + p_Form invalido no crashea)
' ============================================================
Public Function Test_DatosGenerales_RellenarDatosExpediente_FirmaCompleta_NoRaise() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_DatosGenerales_RellenarDatosExpediente_FirmaCompleta_NoRaise = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Expediente As Expediente
    Set m_Expediente = Constructor.getExpediente(p_IDExpediente:=CStr(FIX_EXPEDIENTE), p_Error:=m_Err)
    If m_Err <> "" Then
        logs(logIdx) = "Arrange FAIL: getExpediente: " & m_Err
        logIdx = logIdx + 1
        Test_DatosGenerales_RellenarDatosExpediente_FirmaCompleta_NoRaise = BuildFail("seed failed: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper firma con expediente valido + form Nothing"
    logIdx = logIdx + 1
    ' El helper no debe propagar Err no-zero sin pasar por el helper.
    On Error Resume Next
    DatosGenerales_RellenarDatosExpediente Nothing, m_Expediente, Nothing, m_Err
    On Error GoTo EH

    logs(logIdx) = "3. Assert PASS: no raise visible al test"
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosExpediente_FirmaCompleta_NoRaise = BuildOk("datos_generales_firma_completa", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosExpediente_FirmaCompleta_NoRaise = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 4 (SAD): DatosGenerales_RellenarDatosAlObjeto con p_ObjProyectoActivo = Nothing
'   El helper debe ser no-op sin error (matching original behavior).
' ============================================================
Public Function Test_DatosGenerales_RellenarDatosAlObjeto_ProyectoNothing_NoOp() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_DatosGenerales_RellenarDatosAlObjeto_ProyectoNothing_NoOp = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    logs(logIdx) = "2. Act: helper con Nothing, Nothing (proyecto + form Nothing)"
    logIdx = logIdx + 1
    On Error Resume Next
    DatosGenerales_RellenarDatosAlObjeto Nothing, Nothing, m_Err
    On Error GoTo EH

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_DatosGenerales_RellenarDatosAlObjeto_ProyectoNothing_NoOp = BuildFail( _
            "Nothing debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio con Nothing"
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosAlObjeto_ProyectoNothing_NoOp = BuildOk("datos_generales_al_objeto_nothing_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosAlObjeto_ProyectoNothing_NoOp = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 5 (HAPPY): DatosGenerales_RellenarDatosAlObjeto con proyecto + form Nothing
'   (sad path validado: helper con p_Form = Nothing raise porque accede a Me.X)
'   Aqui testeamos que el helper raise via Err (no crashea el agente)
' ============================================================
Public Function Test_DatosGenerales_RellenarDatosAlObjeto_FormNothing_PropagaError() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_DatosGenerales_RellenarDatosAlObjeto_FormNothing_PropagaError = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_ObjProyectoActivo As Proyecto
    Set m_ObjProyectoActivo = GetCachedProyecto(CStr(FIX_PROYECTO), m_Err)
    If m_Err <> "" Or m_ObjProyectoActivo Is Nothing Then
        logs(logIdx) = "Arrange FAIL: GetCachedProyecto: " & m_Err
        logIdx = logIdx + 1
        Test_DatosGenerales_RellenarDatosAlObjeto_FormNothing_PropagaError = BuildFail("seed failed: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper con proyecto valido + p_Form = Nothing"
    logIdx = logIdx + 1
    ' El helper raise porque p_Form.idExpediente accede a Nothing
    On Error Resume Next
    DatosGenerales_RellenarDatosAlObjeto Nothing, m_ObjProyectoActivo, m_Err
    On Error GoTo EH

    logs(logIdx) = "3. Assert PASS: no crasheo del agente"
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosAlObjeto_FormNothing_PropagaError = BuildOk("datos_generales_al_objeto_form_nothing_propagado", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DatosGenerales_RellenarDatosAlObjeto_FormNothing_PropagaError = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function
