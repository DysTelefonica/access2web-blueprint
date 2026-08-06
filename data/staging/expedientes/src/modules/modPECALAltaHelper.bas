Attribute VB_Name = "modPECALAltaHelper"
Option Compare Database
Option Explicit

' modPECALAltaHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormPECAL (Alta/Edición form).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: PECALAlta_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #30 commit 97bdb21):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_PECALAltaHelper_OpenForm` opened a real Access form, causing
'     VBE interruption in headless test runs (user-reported 2026-06-26)
'
' IMPORTANT: This module covers the Alta/Edición form (Form_FormPECAL),
' which is structurally different from the Gestion list (Form_FormPECALESGestion,
' covered by modPECALHelper). Prefix disambiguation:
'   - `PECALAlta_*` — Alta/Edición form helpers (this module)
'   - `PECAL_*` (no Alta suffix) — Gestion list helpers (modPECALHelper)
'
' New design (4 helpers — different from Gestion list's 5):
'   1. PECALAlta_Abrir_Inicializar(p_Modo, p_IDEntidad, ByRef p_Entidad, p_Error)
'      — Form_Open
'      Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}
'   2. PECALAlta_VerificarCambios(p_ValoresActuales, p_ValoresOriginales, p_Error)
'      Returns JSON: {ok, payload:{hayCambios, diffs}, error, logs}
'   3. PECALAlta_Registrar(p_Valores, p_Error) — ComandoRegistrar_Click
'      Returns JSON: {ok, payload:{ok, id, modo}, error, logs}
'      The helper decides alta vs edicion based on the m_ObjPECALActiva
'      project-conventional global (set by the parent form's Alta/Editar handler).
'   4. PECALAlta_Cerrar(p_Error) — form cleanup
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
Private Const PECAL_ALTA_TITULO_ALTA As String = "ALTA DE PECAL"
Private Const PECAL_ALTA_TITULO_EDICION As String = "EDICIÓN DE PECAL"

' Supported modo values (callers pass these in p_Modo).
Private Const PECAL_ALTA_MODO_ALTA As String = "alta"
Private Const PECAL_ALTA_MODO_EDICION As String = "edicion"

' Fields that the Alta/Edición form manages (used by VerificarCambios + Registrar).
Private Const PECAL_ALTA_FIELD_NAME As String = "PECAL"
Private Const PECAL_ALTA_FIELD_DESCRIPCION As String = "DESCRIPCION"

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all rework code agrees.
Private Const PECAL_ALTA_TEST_ID_BASE As Long = 900850


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
' Returns a fresh empty Dictionary {PECAL, DESCRIPCION} with empty strings.
' Used by Abrir_Inicializar in alta mode and as the default for ByRef p_Entidad
' before the helper populates it.
Private Function BuildEmptyEntidad() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(PECAL_ALTA_FIELD_NAME) = ""
    d(PECAL_ALTA_FIELD_DESCRIPCION) = ""
    Set BuildEmptyEntidad = d
End Function

' --- NormalizarModo --------------------------------------------------------------
' Validates and normalizes the p_Modo parameter. Returns "" if invalid.
Private Function NormalizarModo(ByVal p_Modo As String) As String
    Dim m As String
    m = LCase$("" & p_Modo)
    If m = PECAL_ALTA_MODO_ALTA Then
        NormalizarModo = PECAL_ALTA_MODO_ALTA
    ElseIf m = PECAL_ALTA_MODO_EDICION Then
        NormalizarModo = PECAL_ALTA_MODO_EDICION
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

' --- PECALAlta_Abrir_Inicializar ---------------------------------------------
' Pure-data init: given the form's mode (alta/edicion) and the entity ID, returns
' the caption to display + the precargados fields for the form to render.
'
' p_Modo: "alta" | "edicion"
' p_IDEntidad: ID of the PECAL to load (used only when p_Modo="edicion")
' ByRef p_Entidad: helper populates this with a Dictionary {PECAL, DESCRIPCION}
'                  of the loaded entity, OR an empty Dictionary for alta mode.
'
' Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}.
Public Function PECALAlta_Abrir_Inicializar( _
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
        p_Error = "PECALAlta_Abrir_Inicializar: p_Modo must be 'alta' or 'edicion', got '" & p_Modo & "'"
        Set p_Entidad = Nothing
        PECALAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim titulo As String
    If modo = PECAL_ALTA_MODO_ALTA Then
        titulo = PECAL_ALTA_TITULO_ALTA
        ' Initialize p_Entidad with empty values.
        Set p_Entidad = BuildEmptyEntidad()
        payload("titulo") = titulo
        payload("modo") = modo
        payload("hasEntidad") = False
        payload("id") = ""
        Set payload("entidad") = Nothing

        logs(0) = "Abrir_Inicializar: mode=alta, p_Entidad=empty"
        PECALAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Edicion mode — load the entity by ID.
    Dim idEntidad As String
    idEntidad = Trim$("" & p_IDEntidad)
    If Len(idEntidad) = 0 Then
        p_Error = "PECALAlta_Abrir_Inicializar: p_IDEntidad is required in edicion mode"
        Set p_Entidad = Nothing
        PECALAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Load the PECAL class instance, then build the Dictionary shape for the form.
    Dim loaded As PECAL
    Dim loadErr As String
    Set loaded = constructor.getPecal(p_IDPEcal:=idEntidad, p_Error:=loadErr)
    If Len(loadErr) > 0 Then
        p_Error = "PECALAlta_Abrir_Inicializar: " & loadErr
        Set p_Entidad = Nothing
        PECALAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If loaded Is Nothing Then
        p_Error = "PECALAlta_Abrir_Inicializar: No se ha podido encontrar la PECAL registrada (id=" & idEntidad & ")"
        Set p_Entidad = Nothing
        PECALAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Populate p_Entidad as Dictionary for the form to render.
    Set p_Entidad = BuildEmptyEntidad()
    If loaded.PECAL <> "" Then
        p_Entidad(PECAL_ALTA_FIELD_NAME) = CStr(loaded.PECAL)
    End If
    If loaded.DESCRIPCION <> "" Then
        p_Entidad(PECAL_ALTA_FIELD_DESCRIPCION) = CStr(loaded.DESCRIPCION)
    End If

    ' Also update the project-conventional global so Registrar picks it up later.
    Set m_ObjPECALActiva = loaded

    titulo = PECAL_ALTA_TITULO_EDICION
    payload("titulo") = titulo
    payload("modo") = modo
    payload("hasEntidad") = True
    payload("id") = idEntidad
    Set payload("entidad") = p_Entidad

    logs(0) = "Abrir_Inicializar: mode=edicion, id=" & idEntidad
    logs(1) = "Abrir_Inicializar: precargados populated"

    PECALAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "PECALAlta_Abrir_Inicializar: " & Err.Description
    End If
    Set p_Entidad = Nothing
    PECALAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- PECALAlta_VerificarCambios ---------------------------------------------
' Pure-data diff: compares current form values against the original snapshot.
' Returns whether anything changed and a per-field diff Dictionary.
'
' p_ValoresActuales: Dictionary {PECAL, DESCRIPCION} — current form values
' p_ValoresOriginales: Dictionary {PECAL, DESCRIPCION} — snapshot at open
'                     (or Nothing for alta mode)
'
' For alta mode (p_ValoresOriginales is Nothing), hayCambios is True iff any
' field is non-empty. For edicion mode, hayCambios is True iff any field differs.
'
' Returns JSON: {ok, payload:{hayCambios, diffs:{PECAL, DESCRIPCION}}, error, logs}.
Public Function PECALAlta_VerificarCambios( _
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
        p_Error = "PECALAlta_VerificarCambios: p_ValoresActuales is Nothing"
        PECALAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim diffs As Object
    Set diffs = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_ValoresActuales, PECAL_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_ValoresActuales, PECAL_ALTA_FIELD_DESCRIPCION)

    Dim hayCambios As Boolean
    hayCambios = False

    If p_ValoresOriginales Is Nothing Then
        ' Alta mode — any non-empty value counts as a change.
        If Len(currentName) > 0 Then
            diffs(PECAL_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(PECAL_ALTA_FIELD_NAME) = False
        End If
        If Len(currentDesc) > 0 Then
            diffs(PECAL_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(PECAL_ALTA_FIELD_DESCRIPCION) = False
        End If
        logs(0) = "VerificarCambios: alta mode, hayCambios=" & CStr(hayCambios)
    Else
        ' Edicion mode — diff per field.
        Dim originalName As String
        Dim originalDesc As String

        originalName = ExtraerValor(p_ValoresOriginales, PECAL_ALTA_FIELD_NAME)
        originalDesc = ExtraerValor(p_ValoresOriginales, PECAL_ALTA_FIELD_DESCRIPCION)

        If StrComp(currentName, originalName, vbTextCompare) <> 0 Then
            diffs(PECAL_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(PECAL_ALTA_FIELD_NAME) = False
        End If

        If StrComp(currentDesc, originalDesc, vbTextCompare) <> 0 Then
            diffs(PECAL_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(PECAL_ALTA_FIELD_DESCRIPCION) = False
        End If

        logs(0) = "VerificarCambios: edicion mode, hayCambios=" & CStr(hayCambios)
    End If

    payload("hayCambios") = hayCambios
    Set payload("diffs") = diffs

    PECALAlta_VerificarCambios = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "PECALAlta_VerificarCambios: " & Err.Description
    End If
    PECALAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- PECALAlta_Registrar --------------------------------------------------
' Pure-data persist: given the final form values, persist via PECALOperaciones.
' The helper decides alta vs edicion based on the project-conventional global
' m_ObjPECALActiva:
'   - If m_ObjPECALActiva is Nothing  -> alta mode (insert new row)
'   - If m_ObjPECALActiva is set      -> edicion mode (update existing row)
'
' p_Valores: Dictionary {PECAL, DESCRIPCION} — final values to persist
'
' Returns JSON: {ok, payload:{ok, id, modo}, error, logs}.
Public Function PECALAlta_Registrar( _
    ByVal p_Valores As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Valores Is Nothing Then
        p_Error = "PECALAlta_Registrar: p_Valores is Nothing"
        PECALAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_Valores, PECAL_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_Valores, PECAL_ALTA_FIELD_DESCRIPCION)

    ' Determine mode and capture id.
    Dim modo As String
    Dim idEntidad As String
    If m_ObjPECALActiva Is Nothing Then
        modo = PECAL_ALTA_MODO_ALTA
        idEntidad = ""
    Else
        modo = PECAL_ALTA_MODO_EDICION
        If m_ObjPECALActiva.IDPECAL <> "" Then
            idEntidad = CStr(m_ObjPECALActiva.IDPECAL)
        Else
            idEntidad = ""
        End If
    End If

    logs(0) = "Registrar: mode=" & modo & ", id=" & idEntidad

    ' Build the PECAL class instance from p_Valores (and existing m_ObjPECALActiva).
    Dim pe As PECAL
    If m_ObjPECALActiva Is Nothing Then
        Set pe = New PECAL
    Else
        Set pe = m_ObjPECALActiva
    End If

    pe.PECAL = currentName
    pe.DESCRIPCION = currentDesc

    ' Update the project-conventional global so subsequent operations see the entity.
    Set m_ObjPECALActiva = pe

    ' Delegate to PECALOperaciones for the actual DB write.
    Dim m_PECALOp As New PECALOperaciones
    Dim m_AlInicio As PECAL
    If modo = PECAL_ALTA_MODO_EDICION Then
        Dim loadErr As String
        Set m_AlInicio = constructor.getPecal(p_IDPEcal:=idEntidad, p_Error:=loadErr)
    Else
        Set m_AlInicio = Nothing
    End If

    With m_PECALOp
        Set .PECAL = pe
        Dim registrarErr As String
        .Registrar m_AlInicio, registrarErr
        If Len(registrarErr) > 0 Then
            p_Error = "PECALAlta_Registrar: " & registrarErr
            PECALAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
        End If
    End With

    ' Capture the resulting ID after Registrar has stamped it.
    Dim resultingId As String
    If pe.IDPECAL <> "" Then
        resultingId = CStr(pe.IDPECAL)
    Else
        resultingId = idEntidad
    End If

    payload("ok") = True
    payload("id") = resultingId
    payload("modo") = modo

    logs(1) = "Registrar: persisted, resultingId=" & resultingId

    PECALAlta_Registrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "PECALAlta_Registrar: " & Err.Description
    End If
    PECALAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- PECALAlta_Cerrar -----------------------------------------------------
' Form cleanup. Releases the project-conventional m_ObjPECALActiva global so the
' next form open starts clean. Returns ok=true on success.
'
' Returns JSON: {ok, payload:{ok}, error, logs}.
Public Function PECALAlta_Cerrar( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    ' Release the project-conventional global.
    Set m_ObjPECALActiva = Nothing

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ok") = True

    logs(0) = "Cerrar: m_ObjPECALActiva cleared"

    PECALAlta_Cerrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "PECALAlta_Cerrar: " & Err.Description
    End If
    PECALAlta_Cerrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function