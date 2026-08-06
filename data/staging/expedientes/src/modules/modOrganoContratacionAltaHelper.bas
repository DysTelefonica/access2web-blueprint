Attribute VB_Name = "modOrganoContratacionAltaHelper"
Option Compare Database
Option Explicit

' modOrganoContratacionAltaHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormOrganoContratacion (Alta/Edición form).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: OrganoContratacionAlta_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #30 commit 97bdb21):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_OrganoContratacionAltaHelper_OpenForm` opened a real Access form, causing
'     VBE interruption in headless test runs (user-reported 2026-06-26)
'
' IMPORTANT: This module covers the Alta/Edición form (Form_FormOrganoContratacion),
' which is structurally different from the Gestion list (Form_FormOrganoContratacionGestion,
' covered by modOrganoContratacionHelper). Prefix disambiguation:
'   - `OrganoContratacionAlta_*` — Alta/Edición form helpers (this module)
'   - `OrganoContratacion_*` (no Alta suffix) — Gestion list helpers (modOrganoContratacionHelper)
'
' New design (4 helpers — different from Gestion list's 5):
'   1. OrganoContratacionAlta_Abrir_Inicializar(p_Modo, p_IDEntidad, ByRef p_Entidad, p_Error)
'      — Form_Open
'      Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}
'   2. OrganoContratacionAlta_VerificarCambios(p_ValoresActuales, p_ValoresOriginales, p_Error)
'      Returns JSON: {ok, payload:{hayCambios, diffs}, error, logs}
'   3. OrganoContratacionAlta_Registrar(p_Valores, p_Error) — ComandoRegistrar_Click
'      Returns JSON: {ok, payload:{ok, id, modo}, error, logs}
'      The helper decides alta vs edicion based on the m_ObjOrganoContratacionActivo
'      project-conventional global (set by the parent form's Alta/Editar handler).
'   4. OrganoContratacionAlta_Cerrar(p_Error) — form cleanup
'      Returns JSON: {ok, payload:{ok}, error, logs}
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.
'
' vba-access §10.1 declaration ordering: Private Const / Private Function at top,
' Public Function atoms after. No mid-module consts.
'
' vba-access §1.6.1: split guards (no IIf/And short-circuit on the same object).
'
' vba-access §1.4: On Error GoTo EH with single exit label; Err.Raise 1000 for
' functional errors (preserves context for the caller).

' === Module-level constants (all at top per vba-access §10.1) ============================

' Caption text for each mode.
Private Const ORGANO_CONTRATACION_ALTA_TITULO_ALTA As String = "ALTA DE ÓRGANO DE CONTRATACIÓN"
Private Const ORGANO_CONTRATACION_ALTA_TITULO_EDICION As String = "EDICIÓN DE ÓRGANO DE CONTRATACIÓN"

' Supported modo values (callers pass these in p_Modo).
Private Const ORGANO_CONTRATACION_ALTA_MODO_ALTA As String = "alta"
Private Const ORGANO_CONTRATACION_ALTA_MODO_EDICION As String = "edicion"

' Fields that the Alta/Edición form manages (used by VerificarCambios + Registrar).
Private Const ORGANO_CONTRATACION_ALTA_FIELD_NAME As String = "OrganoContratacion"
Private Const ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION As String = "DESCRIPCION"

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all rework code agrees.
Private Const ORGANO_CONTRATACION_ALTA_TEST_ID_BASE As Long = 900750


' === Local helpers (all at top per vba-access §10.1) ====================================

' --- BuildJsonPayload ----------------------------------------------------------------
' Wraps a Dictionary payload in the canonical JSON envelope:
'   {"ok":true,"value":null,"payload":<payloadJson>,"error":null,"logs":[...]}
' Logs that are empty strings are stripped (consistent with TestHelper.JsonStringArray).
' Uses JsonConverter.ConvertToJson for safe payload serialization.
Private Function BuildJsonPayload( _
    ByVal p_Ok As Boolean, _
    ByVal p_Payload As Object, _
    ByVal p_ErrorMsg As String, _
    ByRef p_Logs() As String _
) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then
        payloadJson = "null"
    Else
        payloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If

    Dim logsJson As String
    logsJson = TestHelper.JsonStringArray(p_Logs)

    Dim errorJson As String
    If p_Ok Then
        errorJson = "null"
    Else
        errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    End If

    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & _
                       ",""value"":null" & _
                       ",""payload"":" & payloadJson & _
                       ",""error"":" & errorJson & _
                       ",""logs"":" & logsJson & "}"
End Function

' --- BuildEmptyEntidad --------------------------------------------------------------
' Returns a fresh empty Dictionary {OrganoContratacion, DESCRIPCION} with empty strings.
' Used by Abrir_Inicializar in alta mode and as the default for ByRef p_Entidad
' before the helper populates it.
Private Function BuildEmptyEntidad() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(ORGANO_CONTRATACION_ALTA_FIELD_NAME) = ""
    d(ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION) = ""
    Set BuildEmptyEntidad = d
End Function

' --- NormalizarModo --------------------------------------------------------------
' Validates and normalizes the p_Modo parameter. Returns "" if invalid.
Private Function NormalizarModo(ByVal p_Modo As String) As String
    Dim m As String
    m = LCase$("" & p_Modo)
    If m = ORGANO_CONTRATACION_ALTA_MODO_ALTA Then
        NormalizarModo = ORGANO_CONTRATACION_ALTA_MODO_ALTA
    ElseIf m = ORGANO_CONTRATACION_ALTA_MODO_EDICION Then
        NormalizarModo = ORGANO_CONTRATACION_ALTA_MODO_EDICION
    Else
        NormalizarModo = ""
    End If
End Function

' --- ExtraerValor ----------------------------------------------------------------
' Split guard: extracts a string value from a Dictionary (returns "" if missing).
Private Function ExtraerValor(ByVal p_Dict As Object, ByVal p_Key As String) As String
    If p_Dict Is Nothing Then
        ExtraerValor = ""
        Exit Function
    End If
    If p_Dict.Exists(p_Key) Then
        ExtraerValor = CStr(p_Dict(p_Key))
    Else
        ExtraerValor = ""
    End If
End Function


' === Public API =========================================================================

' --- OrganoContratacionAlta_Abrir_Inicializar -----------------------------------------
' Pure-data init: given the form's mode (alta/edicion) and the entity ID, returns
' the caption to display + the precargados fields for the form to render.
'
' p_Modo: "alta" | "edicion"
' p_IDEntidad: ID of the OrganoContratacion to load (used only when p_Modo="edicion")
' ByRef p_Entidad: helper populates this with a Dictionary {OrganoContratacion, DESCRIPCION}
'                  of the loaded entity, OR an empty Dictionary for alta mode.
'
' Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}.
Public Function OrganoContratacionAlta_Abrir_Inicializar( _
    ByVal p_Modo As String, _
    ByVal p_IDEntidad As String, _
    ByRef p_Entidad As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim modo As String
    modo = NormalizarModo(p_Modo)
    If Len(modo) = 0 Then
        p_Error = "OrganoContratacionAlta_Abrir_Inicializar: p_Modo must be 'alta' or 'edicion', got '" & p_Modo & "'"
        Set p_Entidad = Nothing
        OrganoContratacionAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim titulo As String
    If modo = ORGANO_CONTRATACION_ALTA_MODO_ALTA Then
        titulo = ORGANO_CONTRATACION_ALTA_TITULO_ALTA
        ' Initialize p_Entidad with empty values.
        Set p_Entidad = BuildEmptyEntidad()
        payload("titulo") = titulo
        payload("modo") = modo
        payload("hasEntidad") = False
        payload("id") = ""
        Set payload("entidad") = Nothing

        logs(0) = "Abrir_Inicializar: mode=alta, p_Entidad=empty"
        OrganoContratacionAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Edicion mode — load the entity by ID.
    Dim idEntidad As String
    idEntidad = Trim$("" & p_IDEntidad)
    If Len(idEntidad) = 0 Then
        p_Error = "OrganoContratacionAlta_Abrir_Inicializar: p_IDEntidad is required in edicion mode"
        Set p_Entidad = Nothing
        OrganoContratacionAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Load the OrganoContratacion class instance, then build the Dictionary shape for the form.
    Dim loaded As OrganoContratacion
    Dim loadErr As String
    Set loaded = constructor.getOrganoContratacion(p_IDOrganoContratacion:=idEntidad, p_Error:=loadErr)
    If Len(loadErr) > 0 Then
        p_Error = "OrganoContratacionAlta_Abrir_Inicializar: " & loadErr
        Set p_Entidad = Nothing
        OrganoContratacionAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If loaded Is Nothing Then
        p_Error = "OrganoContratacionAlta_Abrir_Inicializar: No se ha podido encontrar el Órgano de Contratación registrado (id=" & idEntidad & ")"
        Set p_Entidad = Nothing
        OrganoContratacionAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Populate p_Entidad as Dictionary for the form to render.
    Set p_Entidad = BuildEmptyEntidad()
    If loaded.OrganoContratacion <> "" Then
        p_Entidad(ORGANO_CONTRATACION_ALTA_FIELD_NAME) = CStr(loaded.OrganoContratacion)
    End If
    If loaded.DESCRIPCION <> "" Then
        p_Entidad(ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION) = CStr(loaded.DESCRIPCION)
    End If

    ' Also update the project-conventional global so Registrar picks it up later.
    Set m_ObjOrganoContratacionActivo = loaded

    titulo = ORGANO_CONTRATACION_ALTA_TITULO_EDICION
    payload("titulo") = titulo
    payload("modo") = modo
    payload("hasEntidad") = True
    payload("id") = idEntidad
    Set payload("entidad") = p_Entidad

    logs(0) = "Abrir_Inicializar: mode=edicion, id=" & idEntidad
    logs(1) = "Abrir_Inicializar: precargados populated"

    OrganoContratacionAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OrganoContratacionAlta_Abrir_Inicializar: " & Err.Description
    End If
    Set p_Entidad = Nothing
    OrganoContratacionAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OrganoContratacionAlta_VerificarCambios -------------------------------------------
' Pure-data diff: compares current form values against the original snapshot.
' Returns whether anything changed and a per-field diff Dictionary.
'
' p_ValoresActuales: Dictionary {OrganoContratacion, DESCRIPCION} — current form values
' p_ValoresOriginales: Dictionary {OrganoContratacion, DESCRIPCION} — snapshot at open
'                     (or Nothing for alta mode)
'
' For alta mode (p_ValoresOriginales is Nothing), hayCambios is True iff any
' field is non-empty. For edicion mode, hayCambios is True iff any field differs.
'
' Returns JSON: {ok, payload:{hayCambios, diffs:{OrganoContratacion, DESCRIPCION}}, error, logs}.
Public Function OrganoContratacionAlta_VerificarCambios( _
    ByVal p_ValoresActuales As Object, _
    ByVal p_ValoresOriginales As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    If p_ValoresActuales Is Nothing Then
        p_Error = "OrganoContratacionAlta_VerificarCambios: p_ValoresActuales is Nothing"
        OrganoContratacionAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim diffs As Object
    Set diffs = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_ValoresActuales, ORGANO_CONTRATACION_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_ValoresActuales, ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION)

    Dim hayCambios As Boolean
    hayCambios = False

    If p_ValoresOriginales Is Nothing Then
        ' Alta mode — any non-empty value counts as a change.
        If Len(currentName) > 0 Then
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_NAME) = False
        End If
        If Len(currentDesc) > 0 Then
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION) = False
        End If
        logs(0) = "VerificarCambios: alta mode, hayCambios=" & CStr(hayCambios)
    Else
        ' Edicion mode — diff per field.
        Dim originalName As String
        Dim originalDesc As String

        originalName = ExtraerValor(p_ValoresOriginales, ORGANO_CONTRATACION_ALTA_FIELD_NAME)
        originalDesc = ExtraerValor(p_ValoresOriginales, ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION)

        If StrComp(currentName, originalName, vbTextCompare) <> 0 Then
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_NAME) = False
        End If

        If StrComp(currentDesc, originalDesc, vbTextCompare) <> 0 Then
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION) = False
        End If

        logs(0) = "VerificarCambios: edicion mode, hayCambios=" & CStr(hayCambios)
    End If

    payload("hayCambios") = hayCambios
    Set payload("diffs") = diffs

    OrganoContratacionAlta_VerificarCambios = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OrganoContratacionAlta_VerificarCambios: " & Err.Description
    End If
    OrganoContratacionAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OrganoContratacionAlta_Registrar -----------------------------------------------
' Pure-data persist: given the final form values, persist via OrganoContratacionOperaciones.
' The helper decides alta vs edicion based on the project-conventional global
' m_ObjOrganoContratacionActivo:
'   - If m_ObjOrganoContratacionActivo is Nothing  -> alta mode (insert new row)
'   - If m_ObjOrganoContratacionActivo is set      -> edicion mode (update existing row)
'
' p_Valores: Dictionary {OrganoContratacion, DESCRIPCION} — final values to persist
'
' Returns JSON: {ok, payload:{ok, id, modo}, error, logs}.
Public Function OrganoContratacionAlta_Registrar( _
    ByVal p_Valores As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Valores Is Nothing Then
        p_Error = "OrganoContratacionAlta_Registrar: p_Valores is Nothing"
        OrganoContratacionAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_Valores, ORGANO_CONTRATACION_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_Valores, ORGANO_CONTRATACION_ALTA_FIELD_DESCRIPCION)

    ' Determine mode and capture id.
    Dim modo As String
    Dim idEntidad As String
    If m_ObjOrganoContratacionActivo Is Nothing Then
        modo = ORGANO_CONTRATACION_ALTA_MODO_ALTA
        idEntidad = ""
    Else
        modo = ORGANO_CONTRATACION_ALTA_MODO_EDICION
        If m_ObjOrganoContratacionActivo.IDOrganoContratacion <> "" Then
            idEntidad = CStr(m_ObjOrganoContratacionActivo.IDOrganoContratacion)
        Else
            idEntidad = ""
        End If
    End If

    logs(0) = "Registrar: mode=" & modo & ", id=" & idEntidad

    ' Build the OrganoContratacion class instance from p_Valores (and existing m_ObjOrganoContratacionActivo).
    Dim oc As OrganoContratacion
    If m_ObjOrganoContratacionActivo Is Nothing Then
        Set oc = New OrganoContratacion
    Else
        Set oc = m_ObjOrganoContratacionActivo
    End If

    oc.OrganoContratacion = currentName
    oc.DESCRIPCION = currentDesc

    ' Update the project-conventional global so subsequent operations see the entity.
    Set m_ObjOrganoContratacionActivo = oc

    ' Delegate to OrganoContratacionOperaciones for the actual DB write.
    Dim m_OrganoContratacionOp As New OrganoContratacionOperaciones
    Dim m_AlInicio As OrganoContratacion
    If modo = ORGANO_CONTRATACION_ALTA_MODO_EDICION Then
        Dim loadErr As String
        Set m_AlInicio = constructor.getOrganoContratacion(p_IDOrganoContratacion:=idEntidad, p_Error:=loadErr)
    Else
        Set m_AlInicio = Nothing
    End If

    With m_OrganoContratacionOp
        Set .OrganoContratacion = oc
        Dim registrarErr As String
        .Registrar m_AlInicio, registrarErr
        If Len(registrarErr) > 0 Then
            p_Error = "OrganoContratacionAlta_Registrar: " & registrarErr
            OrganoContratacionAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
        End If
    End With

    ' Capture the resulting ID after Registrar has stamped it.
    Dim resultingId As String
    If oc.IDOrganoContratacion <> "" Then
        resultingId = CStr(oc.IDOrganoContratacion)
    Else
        resultingId = idEntidad
    End If

    payload("ok") = True
    payload("id") = resultingId
    payload("modo") = modo

    logs(1) = "Registrar: persisted, resultingId=" & resultingId

    OrganoContratacionAlta_Registrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OrganoContratacionAlta_Registrar: " & Err.Description
    End If
    OrganoContratacionAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OrganoContratacionAlta_Cerrar --------------------------------------------------
' Form cleanup. Releases the project-conventional m_ObjOrganoContratacionActivo global so the
' next form open starts clean. Returns ok=true on success.
'
' Returns JSON: {ok, payload:{ok}, error, logs}.
Public Function OrganoContratacionAlta_Cerrar( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    ' Release the project-conventional global.
    Set m_ObjOrganoContratacionActivo = Nothing

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ok") = True

    logs(0) = "Cerrar: m_ObjOrganoContratacionActivo cleared"

    OrganoContratacionAlta_Cerrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OrganoContratacionAlta_Cerrar: " & Err.Description
    End If
    OrganoContratacionAlta_Cerrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function