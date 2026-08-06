Attribute VB_Name = "modOficinaProgramaAltaHelper"
Option Compare Database
Option Explicit

' modOficinaProgramaAltaHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormOficinaPrograma (Alta/Edición form).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: OficinaProgramaAlta_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_OficinaProgramaAltaHelper_OpenForm` opened a real Access form, causing
'     VBE interruption in headless test runs (user-reported 2026-06-26)
'
' IMPORTANT: This module covers the Alta/Edición form (Form_FormOficinaPrograma),
' which is structurally different from the Gestion list (Form_FormOficinasProgramaGestion,
' covered by modOficinaProgramaHelper, commit 6838014 / PR #33 pilot).
' Prefix disambiguation:
'   - `OficinaProgramaAlta_*` — Alta/Edición form helpers (this module)
'   - `OficinaPrograma_*` (no Alta suffix) — Gestion list helpers (modOficinaProgramaHelper)
'
' New design (4 helpers — matches the OrganoContratacionAlta pattern):
'   1. OficinaProgramaAlta_Abrir_Inicializar(p_Modo, p_IDEntidad, ByRef p_Entidad, p_Error)
'      — Form_Open
'      Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}
'   2. OficinaProgramaAlta_VerificarCambios(p_ValoresActuales, p_ValoresOriginales, p_Error)
'      Returns JSON: {ok, payload:{hayCambios, diffs}, error, logs}
'   3. OficinaProgramaAlta_Registrar(p_Valores, p_Error) — ComandoRegistrar_Click
'      Returns JSON: {ok, payload:{ok, id, modo}, error, logs}
'   4. OficinaProgramaAlta_Cerrar(p_Error) — form cleanup
'      Returns JSON: {ok, payload:{ok}, error, logs}
'
' 2 fields per form (OficinaPrograma, DESCRIPCION).
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
Private Const OFICINA_PROGRAMA_ALTA_TITULO_ALTA As String = "ALTA DE OFICINA DE PROGRAMA"
Private Const OFICINA_PROGRAMA_ALTA_TITULO_EDICION As String = "EDICIÓN DE OFICINA DE PROGRAMA"

' Supported modo values (callers pass these in p_Modo).
Private Const OFICINA_PROGRAMA_ALTA_MODO_ALTA As String = "alta"
Private Const OFICINA_PROGRAMA_ALTA_MODO_EDICION As String = "edicion"

' Fields that the Alta/Edición form manages.
Private Const OFICINA_PROGRAMA_ALTA_FIELD_NAME As String = "OficinaPrograma"
Private Const OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION As String = "DESCRIPCION"

' Sentinel name that blocks deletion in the Gestion list (preserved here as a
' runtime invariant: any entity named "N/A" is structural and must not be edited
' via this form's Registrar path either, per the user-facing error).
Private Const OFICINA_PROGRAMA_ALTA_NA_SENTINEL As String = "N/A"

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all rework code agrees.
Private Const OFICINA_PROGRAMA_ALTA_TEST_ID_BASE As Long = 900730


' === Local helpers (all at top per vba-access §10.1) ====================================

' --- BuildJsonPayload ----------------------------------------------------------------
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
Private Function BuildEmptyEntidad() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(OFICINA_PROGRAMA_ALTA_FIELD_NAME) = ""
    d(OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION) = ""
    Set BuildEmptyEntidad = d
End Function

' --- NormalizarModo --------------------------------------------------------------
Private Function NormalizarModo(ByVal p_Modo As String) As String
    Dim m As String
    m = LCase$("" & p_Modo)
    If m = OFICINA_PROGRAMA_ALTA_MODO_ALTA Then
        NormalizarModo = OFICINA_PROGRAMA_ALTA_MODO_ALTA
    ElseIf m = OFICINA_PROGRAMA_ALTA_MODO_EDICION Then
        NormalizarModo = OFICINA_PROGRAMA_ALTA_MODO_EDICION
    Else
        NormalizarModo = ""
    End If
End Function

' --- ExtraerValor ----------------------------------------------------------------
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

' --- OficinaProgramaAlta_Abrir_Inicializar -----------------------------------------
' Pure-data init: given the form's mode (alta/edicion) and the entity ID, returns
' the caption to display + the precargados fields for the form to render.
'
' p_Modo: "alta" | "edicion"
' p_IDEntidad: ID of the OficinaPrograma to load (used only when p_Modo="edicion")
' ByRef p_Entidad: helper populates this with a Dictionary {OficinaPrograma, DESCRIPCION}
'                  of the loaded entity, OR an empty Dictionary for alta mode.
'
' Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}.
Public Function OficinaProgramaAlta_Abrir_Inicializar( _
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
        p_Error = "OficinaProgramaAlta_Abrir_Inicializar: p_Modo must be 'alta' or 'edicion', got '" & p_Modo & "'"
        Set p_Entidad = Nothing
        OficinaProgramaAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim titulo As String
    If modo = OFICINA_PROGRAMA_ALTA_MODO_ALTA Then
        titulo = OFICINA_PROGRAMA_ALTA_TITULO_ALTA
        Set p_Entidad = BuildEmptyEntidad()
        payload("titulo") = titulo
        payload("modo") = modo
        payload("hasEntidad") = False
        payload("id") = ""
        Set payload("entidad") = Nothing

        logs(0) = "Abrir_Inicializar: mode=alta, p_Entidad=empty"
        OficinaProgramaAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Edicion mode — load the entity by ID.
    Dim idEntidad As String
    idEntidad = Trim$("" & p_IDEntidad)
    If Len(idEntidad) = 0 Then
        p_Error = "OficinaProgramaAlta_Abrir_Inicializar: p_IDEntidad is required in edicion mode"
        Set p_Entidad = Nothing
        OficinaProgramaAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Load the OficinaPrograma class instance, then build the Dictionary shape for the form.
    Dim loaded As OficinaPrograma
    Dim loadErr As String
    Set loaded = constructor.getOficinaPrograma(p_IDOficinaPrograma:=idEntidad, p_Error:=loadErr)
    If Len(loadErr) > 0 Then
        p_Error = "OficinaProgramaAlta_Abrir_Inicializar: " & loadErr
        Set p_Entidad = Nothing
        OficinaProgramaAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If loaded Is Nothing Then
        p_Error = "OficinaProgramaAlta_Abrir_Inicializar: No se ha podido encontrar la Oficina de Programa registrada (id=" & idEntidad & ")"
        Set p_Entidad = Nothing
        OficinaProgramaAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Populate p_Entidad as Dictionary for the form to render.
    Set p_Entidad = BuildEmptyEntidad()
    If loaded.OficinaPrograma <> "" Then
        p_Entidad(OFICINA_PROGRAMA_ALTA_FIELD_NAME) = CStr(loaded.OficinaPrograma)
    End If
    If loaded.DESCRIPCION <> "" Then
        p_Entidad(OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION) = CStr(loaded.DESCRIPCION)
    End If

    ' Also update the project-conventional global so Registrar picks it up later.
    Set m_ObjOficinaProgramaActiva = loaded

    titulo = OFICINA_PROGRAMA_ALTA_TITULO_EDICION
    payload("titulo") = titulo
    payload("modo") = modo
    payload("hasEntidad") = True
    payload("id") = idEntidad
    Set payload("entidad") = p_Entidad

    logs(0) = "Abrir_Inicializar: mode=edicion, id=" & idEntidad
    logs(1) = "Abrir_Inicializar: precargados populated"

    OficinaProgramaAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaProgramaAlta_Abrir_Inicializar: " & Err.Description
    End If
    Set p_Entidad = Nothing
    OficinaProgramaAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OficinaProgramaAlta_VerificarCambios -------------------------------------------
' Pure-data diff: compares current form values against the original snapshot.
'
' p_ValoresActuales: Dictionary {OficinaPrograma, DESCRIPCION} — current form values
' p_ValoresOriginales: Dictionary {OficinaPrograma, DESCRIPCION} — snapshot at open
'                     (or Nothing for alta mode)
'
' For alta mode (p_ValoresOriginales is Nothing), hayCambios is True iff any
' field is non-empty. For edicion mode, hayCambios is True iff any field differs.
'
' Returns JSON: {ok, payload:{hayCambios, diffs:{OficinaPrograma, DESCRIPCION}}, error, logs}.
Public Function OficinaProgramaAlta_VerificarCambios( _
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
        p_Error = "OficinaProgramaAlta_VerificarCambios: p_ValoresActuales is Nothing"
        OficinaProgramaAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim diffs As Object
    Set diffs = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_ValoresActuales, OFICINA_PROGRAMA_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_ValoresActuales, OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION)

    Dim hayCambios As Boolean
    hayCambios = False

    If p_ValoresOriginales Is Nothing Then
        ' Alta mode — any non-empty value counts as a change.
        If Len(currentName) > 0 Then
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_NAME) = False
        End If
        If Len(currentDesc) > 0 Then
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION) = False
        End If
        logs(0) = "VerificarCambios: alta mode, hayCambios=" & CStr(hayCambios)
    Else
        ' Edicion mode — diff per field.
        Dim originalName As String
        Dim originalDesc As String

        originalName = ExtraerValor(p_ValoresOriginales, OFICINA_PROGRAMA_ALTA_FIELD_NAME)
        originalDesc = ExtraerValor(p_ValoresOriginales, OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION)

        If StrComp(currentName, originalName, vbTextCompare) <> 0 Then
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_NAME) = False
        End If

        If StrComp(currentDesc, originalDesc, vbTextCompare) <> 0 Then
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION) = False
        End If

        logs(0) = "VerificarCambios: edicion mode, hayCambios=" & CStr(hayCambios)
    End If

    payload("hayCambios") = hayCambios
    Set payload("diffs") = diffs

    OficinaProgramaAlta_VerificarCambios = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaProgramaAlta_VerificarCambios: " & Err.Description
    End If
    OficinaProgramaAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OficinaProgramaAlta_Registrar -----------------------------------------------
' Pure-data persist: given the final form values, persist via OficinaProgramaOperaciones.
' The helper decides alta vs edicion based on the project-conventional global
' m_ObjOficinaProgramaActivo:
'   - If m_ObjOficinaProgramaActivo is Nothing  -> alta mode (insert new row)
'   - If m_ObjOficinaProgramaActivo is set      -> edicion mode (update existing row)
'
' p_Valores: Dictionary {OficinaPrograma, DESCRIPCION} — final values to persist
'
' Returns JSON: {ok, payload:{ok, id, modo}, error, logs}.
Public Function OficinaProgramaAlta_Registrar( _
    ByVal p_Valores As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Valores Is Nothing Then
        p_Error = "OficinaProgramaAlta_Registrar: p_Valores is Nothing"
        OficinaProgramaAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_Valores, OFICINA_PROGRAMA_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_Valores, OFICINA_PROGRAMA_ALTA_FIELD_DESCRIPCION)

    ' Determine mode and capture id.
    Dim modo As String
    Dim idEntidad As String
    If m_ObjOficinaProgramaActiva Is Nothing Then
        modo = OFICINA_PROGRAMA_ALTA_MODO_ALTA
        idEntidad = ""
    Else
        modo = OFICINA_PROGRAMA_ALTA_MODO_EDICION
        If m_ObjOficinaProgramaActiva.IDOficinaPrograma <> "" Then
            idEntidad = CStr(m_ObjOficinaProgramaActiva.IDOficinaPrograma)
        Else
            idEntidad = ""
        End If
    End If

    logs(0) = "Registrar: mode=" & modo & ", id=" & idEntidad

    ' Build the OficinaPrograma class instance from p_Valores (and existing m_ObjOficinaProgramaActiva).
    Dim op As OficinaPrograma
    If m_ObjOficinaProgramaActiva Is Nothing Then
        Set op = New OficinaPrograma
    Else
        Set op = m_ObjOficinaProgramaActiva
    End If

    op.OficinaPrograma = currentName
    op.DESCRIPCION = currentDesc

    ' Update the project-conventional global so subsequent operations see the entity.
    Set m_ObjOficinaProgramaActiva = op

    ' Delegate to OficinaProgramaOperaciones for the actual DB write.
    Dim m_OficinaProgramaOp As New OficinaProgramaOperaciones
    Dim m_AlInicio As OficinaPrograma
    If modo = OFICINA_PROGRAMA_ALTA_MODO_EDICION Then
        Dim loadErr As String
        Set m_AlInicio = constructor.getOficinaPrograma(p_IDOficinaPrograma:=idEntidad, p_Error:=loadErr)
    Else
        Set m_AlInicio = Nothing
    End If

    With m_OficinaProgramaOp
        Set .OficinaPrograma = op
        Dim registrarErr As String
        .Registrar m_AlInicio, registrarErr
        If Len(registrarErr) > 0 Then
            p_Error = "OficinaProgramaAlta_Registrar: " & registrarErr
            OficinaProgramaAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
        End If
    End With

    ' Capture the resulting ID after Registrar has stamped it.
    Dim resultingId As String
    If op.IDOficinaPrograma <> "" Then
        resultingId = CStr(op.IDOficinaPrograma)
    Else
        resultingId = idEntidad
    End If

    payload("ok") = True
    payload("id") = resultingId
    payload("modo") = modo

    logs(1) = "Registrar: persisted, resultingId=" & resultingId

    OficinaProgramaAlta_Registrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaProgramaAlta_Registrar: " & Err.Description
    End If
    OficinaProgramaAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OficinaProgramaAlta_Cerrar --------------------------------------------------
' Form cleanup. Releases the project-conventional m_ObjOficinaProgramaActiva global
' so the next form open starts clean. Returns ok=true on success.
'
' Returns JSON: {ok, payload:{ok}, error, logs}.
Public Function OficinaProgramaAlta_Cerrar( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    ' Release the project-conventional global.
    Set m_ObjOficinaProgramaActiva = Nothing

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ok") = True

    logs(0) = "Cerrar: m_ObjOficinaProgramaActiva cleared"

    OficinaProgramaAlta_Cerrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaProgramaAlta_Cerrar: " & Err.Description
    End If
    OficinaProgramaAlta_Cerrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
