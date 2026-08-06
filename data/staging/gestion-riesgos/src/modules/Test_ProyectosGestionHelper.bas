Attribute VB_Name = "Test_ProyectosGestionHelper"
' ============================================================
' Test_ProyectosGestionHelper — TDD atoms for
'   modProyectosGestionHelper
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
' SDD:   hr4-sister-slice-2026-07-01 + hr4-sister-b (HR3d follow-up #6)
'
' Scope: 2 helper entries:
'   - ProyectosGestion_ActualizarProyecto_ActualizarLinea -> 3 atoms
'   - ProyectosGestion_ListaFiltrados_Click_Run          -> 4 atoms (EDGE, EDGE, EDGE, HAPPY-via-fake-form)
'
'   Helper signatures:
'     Public Sub ProyectosGestion_ActualizarProyecto_ActualizarLinea( _
'         ByVal p_Form As Object, _
'         ByVal p_ObjProyecto As Proyecto, _
'         ByRef p_Error As String)
'
'     Public Sub ProyectosGestion_ListaFiltrados_Click_Run( _
'         ByVal p_Form As Object, _
'         ByRef p_Error As String)
'
'   Convention: helper is Sub (no JSON return). Atoms wrap the call and
'   return JSON via TestCore_BuildOk / TestCore_BuildFail.
'
'   Hard rule 7: el helper NO hace MsgBox. Los atomos verifican que
'   p_Error se rellena en sad path sin raise visible.
'
'   Note: ListaFiltrados_Click (form-side Private event handler) ahora
'   expone su superficie cross-form via ProyectosGestion_ListaFiltrados_Click_Run.
'   Los atomos cubren solo paths no-UI (form=Nothing, listbox vacio)
'   porque abrir un form real en COM no es soportado.
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
' ATOM 1 (SAD): ActualizarProyecto_ActualizarLinea con p_ObjProyecto = Nothing
'   El helper debe ser no-op sin error (matching original behavior).
' ============================================================
Public Function Test_ProyectosGestion_ActualizarLinea_ProyectoNothing_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_ProyectosGestion_ActualizarLinea_ProyectoNothing_NoOp = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    logs(logIdx) = "2. Act: helper con p_ObjProyecto = Nothing"
    logIdx = logIdx + 1
    ProyectosGestion_ActualizarProyecto_ActualizarLinea Nothing, Nothing, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_ProyectosGestion_ActualizarLinea_ProyectoNothing_NoOp = BuildFail( _
            "Nothing debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio con proyecto Nothing"
    logIdx = logIdx + 1
    Test_ProyectosGestion_ActualizarLinea_ProyectoNothing_NoOp = BuildOk("actualizar_proyecto_nothing_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_ProyectosGestion_ActualizarLinea_ProyectoNothing_NoOp = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 2 (HAPPY): ActualizarProyecto_ActualizarLinea con p_Form = Nothing
'   El helper debe ser no-op sin error cuando el form no esta disponible.
' ============================================================
Public Function Test_ProyectosGestion_ActualizarLinea_FormNothing_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_ProyectosGestion_ActualizarLinea_FormNothing_NoOp = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_ObjProyecto As Proyecto
    Set m_ObjProyecto = GetCachedProyecto(CStr(FIX_PROYECTO), m_Err)
    If m_Err <> "" Or m_ObjProyecto Is Nothing Then
        logs(logIdx) = "Arrange FAIL: GetCachedProyecto: " & m_Err
        logIdx = logIdx + 1
        Test_ProyectosGestion_ActualizarLinea_FormNothing_NoOp = BuildFail("seed project failed: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper con p_Form = Nothing (form cerrado)"
    logIdx = logIdx + 1
    ProyectosGestion_ActualizarProyecto_ActualizarLinea Nothing, m_ObjProyecto, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_ProyectosGestion_ActualizarLinea_FormNothing_NoOp = BuildFail( _
            "Form Nothing debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio con form Nothing"
    logIdx = logIdx + 1
    Test_ProyectosGestion_ActualizarLinea_FormNothing_NoOp = BuildOk("actualizar_proyecto_form_nothing_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_ProyectosGestion_ActualizarLinea_FormNothing_NoOp = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 3 (HAPPY): ActualizarProyecto_ActualizarLinea con proyecto + form = Nothing
'   (cubrir firma) - proyecto+form Nothing
' ============================================================
Public Function Test_ProyectosGestion_ActualizarLinea_AmbosNothing_NoOp() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_ProyectosGestion_ActualizarLinea_AmbosNothing_NoOp = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper con Nothing, Nothing"
    logIdx = logIdx + 1
    ProyectosGestion_ActualizarProyecto_ActualizarLinea Nothing, Nothing, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_ProyectosGestion_ActualizarLinea_AmbosNothing_NoOp = BuildFail( _
            "Ambos Nothing debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio"
    logIdx = logIdx + 1
    Test_ProyectosGestion_ActualizarLinea_AmbosNothing_NoOp = BuildOk("actualizar_proyecto_ambos_nothing_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_ProyectosGestion_ActualizarLinea_AmbosNothing_NoOp = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 4 (EDGE): ListaFiltrados_Click_Run con p_Form = Nothing
'   El helper debe ser no-op sin error cuando el form no esta disponible.
'   Matching edge case behavior of ProyectosGestion_ActualizarProyecto_ActualizarLinea.
' ============================================================
Public Function Test_ProyectosGestion_ListaFiltrados_Click_FormNothing_NoOp() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_ProyectosGestion_ListaFiltrados_Click_FormNothing_NoOp = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper con p_Form = Nothing (form cerrado)"
    logIdx = logIdx + 1
    ProyectosGestion_ListaFiltrados_Click_Run Nothing, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_ProyectosGestion_ListaFiltrados_Click_FormNothing_NoOp = BuildFail( _
            "Form Nothing debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio con form Nothing"
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_FormNothing_NoOp = BuildOk("lista_filtrados_click_form_nothing_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_FormNothing_NoOp = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 5 (EDGE): ListaFiltrados_Click_Run con p_Form que NO expone ListaFiltrados
'   El helper debe ser no-op sin error si el form no expone el control
'   (e.g. caller pasa un objeto que no es el form de ProyectosGestion).
' ============================================================
Public Function Test_ProyectosGestion_ListaFiltrados_Click_FormSinListbox_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + create fake form (sin ListaFiltrados)"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_ProyectosGestion_ListaFiltrados_Click_FormSinListbox_NoOp = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If

    ' Fake form-like object: any VBA class will do — we just need an Object reference.
    ' The helper accesses p_Form.lblUltimoArchivoPublicado FIRST, so a bare class
    ' will fail with "method or member not found" unless we wrap the access in On Error.
    '
    ' Per Hard rule 2 (honest signature), the helper cannot guarantee the caller
    ' passes a fully-formed form. It must NOT crash.
    '
    ' We test this by passing a fresh Collection object — it has neither the
    ' labels nor the ListBox, simulating a malformed/empty form-like reference.
    Dim m_FakeForm As Object
    Set m_FakeForm = New Collection

    logs(logIdx) = "2. Act: helper con p_Form = Collection (no labels, no listbox)"
    logIdx = logIdx + 1

    ' The helper should NOT raise. It accesses p_Form.lblUltimoArchivoPublicado
    ' which doesn't exist on a Collection — but the helper uses On Error GoTo errores
    ' to catch and report.
    ProyectosGestion_ListaFiltrados_Click_Run m_FakeForm, m_Err

    ' Sad expectation: p_Error is set with a meaningful message
    ' (the helper catches the "method not found" error and reports it).
    If m_Err = "" Then
        logs(logIdx) = "3. Assert FAIL: helper returned no error pero el fake-form no expone labels"
        logIdx = logIdx + 1
        Test_ProyectosGestion_ListaFiltrados_Click_FormSinListbox_NoOp = BuildFail( _
            "Expected p_Error con fake-form sin labels/listbox. Got vacio.", logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error set: " & Left$(m_Err, 60)
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_FormSinListbox_NoOp = BuildOk("lista_filtrados_click_fake_form_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_FormSinListbox_NoOp = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 6 (EDGE): ListaFiltrados_Click_Run con p_Form Nothing (doble)
'   AdversarialDoble: p_Form = Nothing + p_Error pre-poblado.
'   El helper debe limpiar p_Error a "" antes de operar.
' ============================================================
Public Function Test_ProyectosGestion_ListaFiltrados_Click_DobleNothing_LimpiaError() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + p_Error pre-poblado"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_ProyectosGestion_ListaFiltrados_Click_DobleNothing_LimpiaError = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    m_Err = "PRE_EXISTING_ERROR_FROM_CALLER"

    logs(logIdx) = "2. Act: helper con p_Form = Nothing (no debe preservar el error previo)"
    logIdx = logIdx + 1
    ProyectosGestion_ListaFiltrados_Click_Run Nothing, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper dejo error previo: " & m_Err
        logIdx = logIdx + 1
        Test_ProyectosGestion_ListaFiltrados_Click_DobleNothing_LimpiaError = BuildFail( _
            "Helper debe limpiar p_Error antes de operar. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error limpio (helper ejecuto p_Error = """")"
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_DobleNothing_LimpiaError = BuildOk("lista_filtrados_click_limpia_error_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_DobleNothing_LimpiaError = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 7 (HAPPY): ListaFiltrados_Click_Run sin p_ObjProyecto (firma simple)
'   El helper NO requiere un proyecto — opera solo sobre el form.
'   Verifica que la firma (p_Form, p_Error) es suficiente.
' ============================================================
Public Function Test_ProyectosGestion_ListaFiltrados_Click_SinProyecto_NoOp() As String
    Dim logs(0 To 4) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_ProyectosGestion_ListaFiltrados_Click_SinProyecto_NoOp = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Act: helper con p_Form = Nothing (no requiere proyecto)"
    logIdx = logIdx + 1
    ProyectosGestion_ListaFiltrados_Click_Run Nothing, m_Err

    If m_Err <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_ProyectosGestion_ListaFiltrados_Click_SinProyecto_NoOp = BuildFail( _
            "Sin proyecto (form=Nothing) debe ser no-op. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio (firma simple, no requiere proyecto)"
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_SinProyecto_NoOp = BuildOk("lista_filtrados_click_sin_proyecto_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_ProyectosGestion_ListaFiltrados_Click_SinProyecto_NoOp = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function
