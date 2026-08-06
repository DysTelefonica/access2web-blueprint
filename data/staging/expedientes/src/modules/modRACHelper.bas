Attribute VB_Name = "modRACHelper"
Option Compare Database
Option Explicit

' modRACHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormRAC (Alta/Edición form).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: RAC_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_RACHelper_OpenForm` opened a real Access form, causing VBE interruption
'
' IMPORTANT: This module covers the Alta/Edición form (Form_FormRAC), which is
' structurally different from the Gestion list (Form_FormRACSGestion, covered
' by modRACSHelper). Prefix disambiguation:
'   - `RAC_*` (singular) — Alta/Edición form helpers (this module)
'   - `RACS_*` (plural) — Gestion list helpers (modRACSHelper)
'
' New design (4 helpers — different from Gestion list's 5):
'   1. RAC_Abrir_Inicializar(p_Modo, p_IDEntidad, ByRef p_Entidad, p_Error) — Form_Open
'      Returns JSON: {ok, payload:{titulo, modo, hasEntidad, entidad}, error, logs}
'   2. RAC_Verificar_Cambios(p_ValoresActuales, p_ValoresOriginales, p_Error)
'      Returns JSON: {ok, payload:{hayCambios, diffs}, error, logs}
'   3. RAC_Registrar(p_Valores, p_Error) — ComandoRegistrar_Click
'      Returns JSON: {ok, payload:{ok, id, modo}, error, logs}
'      The helper decides alta vs edicion based on the m_ObjRACActivo global
'      (set by the form's Alta/Editar handler before opening this form).
'   4. RAC_Cerrar(p_Error) — form cleanup
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
Private Const RAC_TITULO_ALTA As String = "ALTA DE RAC"
Private Const RAC_TITULO_EDICION As String = "EDICIÓN DE RAC"

' Supported modo values (callers pass these in p_Modo).
Private Const RAC_MODO_ALTA As String = "alta"
Private Const RAC_MODO_EDICION As String = "edicion"

' Fields that the Alta/Edición form manages (used by Verificar_Cambios + Registrar).
Private Const RAC_FIELD_RAC As String = "RAC"
Private Const RAC_FIELD_CORREO As String = "CORREO"
Private Const RAC_FIELD_DESCRIPCION As String = "DESCRIPCION"

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all rework code agrees.
Private Const RAC_TEST_ID_BASE As Long = 900900


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

' --- BuildEmptyRACEntity --------------------------------------------------------------
' Returns a fresh empty Dictionary {RAC, CORREO, DESCRIPCION} with empty strings.
' Used by Abrir_Inicializar in alta mode and as the default for ByRef p_Entidad
' before the helper populates it.
Private Function BuildEmptyRACEntity() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(RAC_FIELD_RAC) = ""
    d(RAC_FIELD_CORREO) = ""
    d(RAC_FIELD_DESCRIPCION) = ""
    Set BuildEmptyRACEntity = d
End Function

' --- RAC_NormalizarModo --------------------------------------------------------------
' Validates and normalizes the p_Modo parameter. Returns "" if invalid.
Private Function RAC_NormalizarModo(ByVal p_Modo As String) As String
    Dim m As String
    m = LCase$("" & p_Modo)
    If m = RAC_MODO_ALTA Then
        RAC_NormalizarModo = RAC_MODO_ALTA
    ElseIf m = RAC_MODO_EDICION Then
        RAC_NormalizarModo = RAC_MODO_EDICION
    Else
        RAC_NormalizarModo = ""
    End If
End Function

' --- RAC_ExtraerValor ----------------------------------------------------------------
' Split guard: extracts a string value from a Dictionary (returns "" if missing).
Private Function RAC_ExtraerValor(ByVal p_Dict As Object, ByVal p_Key As String) As String
    If p_Dict Is Nothing Then
        RAC_ExtraerValor = ""
        Exit Function
    End If
    If p_Dict.Exists(p_Key) Then
        RAC_ExtraerValor = CStr(p_Dict(p_Key))
    Else
        RAC_ExtraerValor = ""
    End If
End Function


' === Public API =========================================================================

' --- RAC_Abrir_Inicializar -----------------------------------------------------------
' Pure-data init: given the form's mode (alta/edicion) and the entity ID, returns
' the caption to display + the precargados fields for the form to render.
'
' p_Modo: "alta" | "edicion"
' p_IDEntidad: ID of the RAC to load (used only when p_Modo="edicion")
' ByRef p_Entidad: helper populates this with a Dictionary {RAC, CORREO, DESCRIPCION}
'                  of the loaded entity, OR an empty Dictionary for alta mode.
'
' Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}.
Public Function RAC_Abrir_Inicializar( _
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
    modo = RAC_NormalizarModo(p_Modo)
    If Len(modo) = 0 Then
        p_Error = "RAC_Abrir_Inicializar: p_Modo must be 'alta' or 'edicion', got '" & p_Modo & "'"
        Set p_Entidad = Nothing
        RAC_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim titulo As String
    If modo = RAC_MODO_ALTA Then
        titulo = RAC_TITULO_ALTA
        ' Initialize p_Entidad with empty values.
        Set p_Entidad = BuildEmptyRACEntity()
        payload("titulo") = titulo
        payload("modo") = modo
        payload("hasEntidad") = False
        payload("id") = ""
        Set payload("entidad") = Nothing

        logs(0) = "Abrir_Inicializar: mode=alta, p_Entidad=empty"
        RAC_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Edicion mode — load the entity by ID.
    Dim idEntidad As String
    idEntidad = Trim$("" & p_IDEntidad)
    If Len(idEntidad) = 0 Then
        p_Error = "RAC_Abrir_Inicializar: p_IDEntidad is required in edicion mode"
        Set p_Entidad = Nothing
        RAC_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Load the RAC class instance, then build the Dictionary shape for the form.
    Dim loaded As RAC
    Dim loadErr As String
    Set loaded = constructor.getRAC(p_IDRAC:=idEntidad, p_Error:=loadErr)
    If Len(loadErr) > 0 Then
        p_Error = "RAC_Abrir_Inicializar: " & loadErr
        Set p_Entidad = Nothing
        RAC_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If loaded Is Nothing Then
        p_Error = "RAC_Abrir_Inicializar: No se ha podido encontrar la RAC registrada (id=" & idEntidad & ")"
        Set p_Entidad = Nothing
        RAC_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Populate p_Entidad as Dictionary for the form to render.
    Set p_Entidad = BuildEmptyRACEntity()
    If loaded.RAC <> "" Then
        p_Entidad(RAC_FIELD_RAC) = CStr(loaded.RAC)
    End If
    If loaded.CORREO <> "" Then
        p_Entidad(RAC_FIELD_CORREO) = CStr(loaded.CORREO)
    End If
    If loaded.DESCRIPCION <> "" Then
        p_Entidad(RAC_FIELD_DESCRIPCION) = CStr(loaded.DESCRIPCION)
    End If

    ' Also update the project-conventional global so Registrar picks it up later.
    Set m_ObjRACActivo = loaded

    titulo = RAC_TITULO_EDICION
    payload("titulo") = titulo
    payload("modo") = modo
    payload("hasEntidad") = True
    payload("id") = idEntidad
    Set payload("entidad") = p_Entidad

    logs(0) = "Abrir_Inicializar: mode=edicion, id=" & idEntidad
    logs(1) = "Abrir_Inicializar: precargados populated"

    RAC_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RAC_Abrir_Inicializar: " & Err.Description
    End If
    Set p_Entidad = Nothing
    RAC_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- RAC_Verificar_Cambios -----------------------------------------------------------
' Pure-data diff: compares current form values against the original snapshot.
' Returns whether anything changed and a per-field diff Dictionary.
'
' p_ValoresActuales: Dictionary {RAC, CORREO, DESCRIPCION} — current form values
' p_ValoresOriginales: Dictionary {RAC, CORREO, DESCRIPCION} — snapshot at open
'                     (or Nothing for alta mode)
'
' For alta mode (p_ValoresOriginales is Nothing), hayCambios is True iff any
' field is non-empty. For edicion mode, hayCambios is True iff any field differs.
'
' Returns JSON: {ok, payload:{hayCambios, diffs:{RAC, CORREO, DESCRIPCION}}, error, logs}.
Public Function RAC_Verificar_Cambios( _
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
        p_Error = "RAC_Verificar_Cambios: p_ValoresActuales is Nothing"
        RAC_Verificar_Cambios = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim diffs As Object
    Set diffs = CreateObject("Scripting.Dictionary")

    Dim currentRAC As String
    Dim currentCorreo As String
    Dim currentDesc As String

    currentRAC = RAC_ExtraerValor(p_ValoresActuales, RAC_FIELD_RAC)
    currentCorreo = RAC_ExtraerValor(p_ValoresActuales, RAC_FIELD_CORREO)
    currentDesc = RAC_ExtraerValor(p_ValoresActuales, RAC_FIELD_DESCRIPCION)

    Dim hayCambios As Boolean
    hayCambios = False

    If p_ValoresOriginales Is Nothing Then
        ' Alta mode — any non-empty value counts as a change.
        If Len(currentRAC) > 0 Then
            diffs(RAC_FIELD_RAC) = True
            hayCambios = True
        Else
            diffs(RAC_FIELD_RAC) = False
        End If
        If Len(currentCorreo) > 0 Then
            diffs(RAC_FIELD_CORREO) = True
            hayCambios = True
        Else
            diffs(RAC_FIELD_CORREO) = False
        End If
        If Len(currentDesc) > 0 Then
            diffs(RAC_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(RAC_FIELD_DESCRIPCION) = False
        End If
        logs(0) = "Verificar_Cambios: alta mode, hayCambios=" & CStr(hayCambios)
    Else
        ' Edicion mode — diff per field.
        Dim originalRAC As String
        Dim originalCorreo As String
        Dim originalDesc As String

        originalRAC = RAC_ExtraerValor(p_ValoresOriginales, RAC_FIELD_RAC)
        originalCorreo = RAC_ExtraerValor(p_ValoresOriginales, RAC_FIELD_CORREO)
        originalDesc = RAC_ExtraerValor(p_ValoresOriginales, RAC_FIELD_DESCRIPCION)

        If StrComp(currentRAC, originalRAC, vbTextCompare) <> 0 Then
            diffs(RAC_FIELD_RAC) = True
            hayCambios = True
        Else
            diffs(RAC_FIELD_RAC) = False
        End If

        If StrComp(currentCorreo, originalCorreo, vbTextCompare) <> 0 Then
            diffs(RAC_FIELD_CORREO) = True
            hayCambios = True
        Else
            diffs(RAC_FIELD_CORREO) = False
        End If

        If StrComp(currentDesc, originalDesc, vbTextCompare) <> 0 Then
            diffs(RAC_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(RAC_FIELD_DESCRIPCION) = False
        End If

        logs(0) = "Verificar_Cambios: edicion mode, hayCambios=" & CStr(hayCambios)
    End If

    payload("hayCambios") = hayCambios
    Set payload("diffs") = diffs

    RAC_Verificar_Cambios = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RAC_Verificar_Cambios: " & Err.Description
    End If
    RAC_Verificar_Cambios = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- RAC_Registrar -------------------------------------------------------------------
' Pure-data persist: given the final form values, persist via RACOperaciones.
' The helper decides alta vs edicion based on the project-conventional global
' m_ObjRACActivo:
'   - If m_ObjRACActivo is Nothing  -> alta mode (insert new row)
'   - If m_ObjRACActivo is set      -> edicion mode (update existing row)
'
' p_Valores: Dictionary {RAC, CORREO, DESCRIPCION} — final values to persist
'
' Returns JSON: {ok, payload:{ok, id, modo}, error, logs}.
Public Function RAC_Registrar( _
    ByVal p_Valores As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Valores Is Nothing Then
        p_Error = "RAC_Registrar: p_Valores is Nothing"
        RAC_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim currentRAC As String
    Dim currentCorreo As String
    Dim currentDesc As String

    currentRAC = RAC_ExtraerValor(p_Valores, RAC_FIELD_RAC)
    currentCorreo = RAC_ExtraerValor(p_Valores, RAC_FIELD_CORREO)
    currentDesc = RAC_ExtraerValor(p_Valores, RAC_FIELD_DESCRIPCION)

    ' Determine mode and capture id.
    Dim modo As String
    Dim idEntidad As String
    If m_ObjRACActivo Is Nothing Then
        modo = RAC_MODO_ALTA
        idEntidad = ""
    Else
        modo = RAC_MODO_EDICION
        If m_ObjRACActivo.IDRAC <> "" Then
            idEntidad = CStr(m_ObjRACActivo.IDRAC)
        Else
            idEntidad = ""
        End If
    End If

    logs(0) = "Registrar: mode=" & modo & ", id=" & idEntidad

    ' Build the RAC class instance from p_Valores (and existing m_ObjRACActivo).
    Dim rac As RAC
    If m_ObjRACActivo Is Nothing Then
        Set rac = New RAC
    Else
        Set rac = m_ObjRACActivo
    End If

    rac.RAC = currentRAC
    rac.CORREO = currentCorreo
    rac.DESCRIPCION = currentDesc

    ' Update the project-conventional global so subsequent operations see the entity.
    Set m_ObjRACActivo = rac

    ' Delegate to RACOperaciones for the actual DB write.
    Dim m_RACOp As New RACOperaciones
    With m_RACOp
        Set .RAC = rac
        Dim registrarErr As String
        .Registrar Nothing, registrarErr
        If Len(registrarErr) > 0 Then
            p_Error = "RAC_Registrar: " & registrarErr
            RAC_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
        End If
    End With

    ' Capture the resulting ID after Registrar has stamped it.
    Dim resultingId As String
    If rac.IDRAC <> "" Then
        resultingId = CStr(rac.IDRAC)
    Else
        resultingId = idEntidad
    End If

    payload("ok") = True
    payload("id") = resultingId
    payload("modo") = modo

    logs(1) = "Registrar: persisted, resultingId=" & resultingId

    RAC_Registrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RAC_Registrar: " & Err.Description
    End If
    RAC_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- RAC_Cerrar ----------------------------------------------------------------------
' Form cleanup. Releases the project-conventional m_ObjRACActivo global so the
' next form open starts clean. Returns ok=true on success.
'
' Returns JSON: {ok, payload:{ok}, error, logs}.
Public Function RAC_Cerrar( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    ' Release the project-conventional global.
    Set m_ObjRACActivo = Nothing

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ok") = True

    logs(0) = "Cerrar: m_ObjRACActivo cleared"

    RAC_Cerrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RAC_Cerrar: " & Err.Description
    End If
    RAC_Cerrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
