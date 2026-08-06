Attribute VB_Name = "modLugarEjecucionHelper"
Option Compare Database
Option Explicit

' modLugarEjecucionHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormLugarEjecucionGestion.
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: LugarEjecucion_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - helpers did `DoCmd.OpenForm` and `Application.Echo` directly
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_LugarEjecucionHelper_OpenForm` opened a real Access form, causing VBE interruption
'     in headless test runs (user-reported 2026-06-26, mirrored from pilot)
'
' New design (5 helpers, NOT 8 — the other 3 are pure UI orchestration, kept in form):
'   1. LugarEjecucion_Abrir_Inicializar(p_EsAdministrador, p_HasOpenArgs, p_Error) — Form_Open
'      Returns JSON: {ok, payload:{showAlta, focusAlta, showElegir}, error, logs}
'   2. LugarEjecucion_Buscar_Listar(p_Lugares, p_Filter, p_Error) — ComandoBuscar_Click
'      Returns JSON: {ok, payload:{rowSource, count}, error, logs}
'   3. LugarEjecucion_Seleccionar_Cargar(p_IDSeleccionado, p_Lugares, p_EsAdministrador, p_Error)
'      — ListaFiltrados_Click
'      Returns JSON: {ok, payload:{entity:{...}, enableEditar, enableEliminar}, error, logs}
'   4. LugarEjecucion_Eliminar_Borrar(p_LugarEjecucion, p_PromptResult, p_Error)
'      — ComandoEliminar_Click
'      Returns JSON: {ok, value:cancelled|deleted, error, logs}
'   5. LugarEjecucion_DobleClick_AbrirEdicion(p_HasElegir, p_EditarEnabled, p_Error)
'      — ListaFiltrados_DblClick
'      Returns JSON: {ok, payload:{action:choose|edit|none}, error, logs}
'
' UI orchestration that stays in the form (rule #1 — never testable business logic):
'   - Alta button: form does DoCmd.OpenForm / FormInteraction_FormularioAbierto / Forms(...)
'   - Edición button: form does DoCmd.OpenForm / FormInteraction_FormularioAbierto / Forms(...)
'   - Limpiar button: form does Me.LUGAREJECUCION = Null (pure UI action, no helper needed)
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter. On failure: set p_Error,
' return fail JSON. On success: return canonical JSON envelope.
'
' vba-access §10.1 declaration ordering: all Private Const / Private Function at top,
' Public Function atoms after. No mid-module consts.
'
' vba-access §1.6.1: split guards (no IIf/And short-circuit on the same object).
'
' vba-access §1.4: On Error GoTo EH with single exit label; Err.Raise 1000 for
' functional errors (preserves context for the caller).

' === Module-level constants (all at top per vba-access §10.1) ============================

' Column header for the filtered list (matches Access listbox RowSource convention).
Private Const LUGAR_EJECUCION_HEADERS As String = "IDLugarEjecucion;LUGAR;DESCRIPCION"

' Field separator inside rowSource (matches Access listbox RowSource convention).
Private Const LUGAR_EJECUCION_FIELDS_SEP As String = ";"

' Maximum rows emitted in rowSource. Defensive cap so a runaway test does not
' build a multi-MB CSV in a single string.
Private Const LUGAR_EJECUCION_MAX_ROWS As Long = 5000

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all rework code agrees.
Private Const LUGAR_EJECUCION_TEST_ID_BASE As Long = 900600


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

' --- LugarEjecucion_SerializarEntidad -------------------------------------------------
' Serializes a single entity Dictionary {IDLugarEjecucion, LugarEjecucion, DESCRIPCION}
' into a plain Dictionary suitable for JSON. Returns Nothing if p_Entidad is Nothing
' or missing required keys (split guards per vba-access §1.6.1).
Private Function LugarEjecucion_SerializarEntidad(ByVal p_Entidad As Object) As Object
    If p_Entidad Is Nothing Then
        Set LugarEjecucion_SerializarEntidad = Nothing
        Exit Function
    End If

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")

    If p_Entidad.Exists("IDLugarEjecucion") Then
        out("IDLugarEjecucion") = CStr(p_Entidad("IDLugarEjecucion"))
    Else
        out("IDLugarEjecucion") = ""
    End If

    If p_Entidad.Exists("LugarEjecucion") Then
        out("LugarEjecucion") = CStr(p_Entidad("LugarEjecucion"))
    Else
        out("LugarEjecucion") = ""
    End If

    If p_Entidad.Exists("DESCRIPCION") Then
        out("DESCRIPCION") = CStr(p_Entidad("DESCRIPCION"))
    Else
        out("DESCRIPCION") = ""
    End If

    Set LugarEjecucion_SerializarEntidad = out
End Function

' --- LugarEjecucion_BuildRowLine ------------------------------------------------------
' Builds a single rowSource line (semicolon-separated) for the listbox.
' Sanitizes embedded semicolons by replacing them with colons (Access listbox
' RowSource cannot contain field separators inside a value).
Private Function LugarEjecucion_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_Lugar As String, _
    ByVal p_Desc As String _
) As String
    Dim safeLugar As String
    safeLugar = Replace(p_Lugar, LUGAR_EJECUCION_FIELDS_SEP, ":")

    Dim safeDesc As String
    safeDesc = Replace(p_Desc, LUGAR_EJECUCION_FIELDS_SEP, ":")

    LugarEjecucion_BuildRowLine = p_ID & LUGAR_EJECUCION_FIELDS_SEP & _
                                  safeLugar & LUGAR_EJECUCION_FIELDS_SEP & _
                                  safeDesc
End Function


' === Public API =========================================================================

' --- LugarEjecucion_Abrir_Inicializar -------------------------------------------------
' Pure-data init: given admin flag + OpenArgs presence, decide which buttons to show
' and where to put focus. The form renders the JSON into its own controls.
' Returns JSON: {ok, payload:{showAlta, focusAlta, showElegir}, error, logs}.
Public Function LugarEjecucion_Abrir_Inicializar( _
    ByVal p_EsAdministrador As Boolean, _
    ByVal p_HasOpenArgs As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim showAlta As Boolean
    Dim focusAlta As Boolean
    Dim showElegir As Boolean

    If p_EsAdministrador Then
        showAlta = True
        focusAlta = True
    Else
        showAlta = False
        focusAlta = False
    End If

    If p_HasOpenArgs Then
        showElegir = True
    Else
        showElegir = False
    End If

    payload("showAlta") = showAlta
    payload("focusAlta") = focusAlta
    payload("showElegir") = showElegir

    logs(0) = "Abrir_Inicializar: admin=" & CStr(p_EsAdministrador) & _
              ", hasOpenArgs=" & CStr(p_HasOpenArgs)
    logs(1) = "Abrir_Inicializar: showAlta=" & CStr(showAlta) & _
              ", focusAlta=" & CStr(focusAlta) & _
              ", showElegir=" & CStr(showElegir)

    LugarEjecucion_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "LugarEjecucion_Abrir_Inicializar: " & Err.Description
    End If
    LugarEjecucion_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- LugarEjecucion_Buscar_Listar -----------------------------------------------------
' Pure-data filter: given a Dictionary of entities {ID -> stubEntity} and a filter
' text, returns a semicolon-separated rowSource CSV + count. The form assigns
' rowSource to Me.ListaFiltrados.RowSource (Access handles the header row separately).
'
' p_Lugares is a Scripting.Dictionary (or Nothing). Each value must expose
' "IDLugarEjecucion", "LugarEjecucion", "DESCRIPCION" keys.
'
' Returns JSON: {ok, payload:{rowSource, count}, error, logs}.
Public Function LugarEjecucion_Buscar_Listar( _
    ByVal p_Lugares As Object, _
    ByVal p_Filter As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim rowSource As String
    rowSource = ""

    Dim rowCount As Long
    rowCount = 0

    ' Defensive: empty/Nothing collection -> 0 rows, no filter applied.
    If p_Lugares Is Nothing Then
        logs(0) = "Buscar_Listar: p_Lugares is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        LugarEjecucion_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim filterText As String
    filterText = "" & p_Filter

    Dim k As Variant
    Dim entity As Object
    Dim currentId As String
    Dim currentLugar As String
    Dim currentDesc As String

    For Each k In p_Lugares.Keys
        If rowCount >= LUGAR_EJECUCION_MAX_ROWS Then
            logs(3) = "Buscar_Listar: hit MAX_ROWS=" & LUGAR_EJECUCION_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set entity = p_Lugares(k)
        If entity Is Nothing Then
            ' Skip malformed entries — split guard before touching properties.
        Else
            If entity.Exists("IDLugarEjecucion") Then
                currentId = CStr(entity("IDLugarEjecucion"))
            Else
                currentId = ""
            End If

            If entity.Exists("LugarEjecucion") Then
                currentLugar = CStr(entity("LugarEjecucion"))
            Else
                currentLugar = ""
            End If

            If entity.Exists("DESCRIPCION") Then
                currentDesc = CStr(entity("DESCRIPCION"))
            Else
                currentDesc = ""
            End If

            If Len(filterText) > 0 Then
                ' Case-insensitive substring match on the LugarEjecucion name field.
                If InStr(1, currentLugar, filterText, vbTextCompare) = 0 Then
                    ' No match — skip without counting.
                Else
                    If Len(rowSource) > 0 Then
                        rowSource = rowSource & vbCrLf
                    End If
                    rowSource = rowSource & LugarEjecucion_BuildRowLine(currentId, currentLugar, currentDesc)
                    rowCount = rowCount + 1
                End If
            Else
                ' No filter — include everything.
                If Len(rowSource) > 0 Then
                    rowSource = rowSource & vbCrLf
                End If
                rowSource = rowSource & LugarEjecucion_BuildRowLine(currentId, currentLugar, currentDesc)
                rowCount = rowCount + 1
            End If
        End If
        Set entity = Nothing
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount

    logs(0) = "Buscar_Listar: filterText=" & filterText & ", rows=" & rowCount
    logs(1) = "Buscar_Listar: rowSource length=" & Len(rowSource)

    LugarEjecucion_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "LugarEjecucion_Buscar_Listar: " & Err.Description
    End If
    LugarEjecucion_Buscar_Listar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- LugarEjecucion_Seleccionar_Cargar ------------------------------------------------
' Pure-data selection: given the selected ID (extracted by the form from the listbox
' Column(0)) and the entities Dictionary, returns the selected entity plus the
' enable flags for ComandoEditar/Eliminar.
'
' p_IDSeleccionado is "" or whitespace when nothing is selected — returns ok with
' enableEditar=false, enableEliminar=false, entity=null.
'
' Returns JSON: {ok, payload:{entity, enableEditar, enableEliminar}, error, logs}.
Public Function LugarEjecucion_Seleccionar_Cargar( _
    ByVal p_IDSeleccionado As String, _
    ByVal p_Lugares As Object, _
    ByVal p_EsAdministrador As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim selectedId As String
    selectedId = Trim$("" & p_IDSeleccionado)

    ' No selection — both buttons disabled, entity is null.
    If Len(selectedId) = 0 Then
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        logs(0) = "Seleccionar_Cargar: empty selection"
        LugarEjecucion_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Look up the entity by ID in the Dictionary.
    Dim entity As Object
    If p_Lugares Is Nothing Then
        Set entity = Nothing
    Else
        If p_Lugares.Exists(selectedId) Then
            Set entity = p_Lugares(selectedId)
        Else
            Set entity = Nothing
        End If
    End If

    If entity Is Nothing Then
        logs(0) = "Seleccionar_Cargar: id=" & selectedId & " not found in collection"
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        LugarEjecucion_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Found the entity — always enable Editar; Eliminar only for admins.
    Dim serializedEntity As Object
    Set serializedEntity = LugarEjecucion_SerializarEntidad(entity)

    Set payload("entity") = serializedEntity
    payload("enableEditar") = True
    If p_EsAdministrador Then
        payload("enableEliminar") = True
    Else
        payload("enableEliminar") = False
    End If

    logs(0) = "Seleccionar_Cargar: id=" & selectedId
    logs(1) = "Seleccionar_Cargar: isAdmin=" & CStr(p_EsAdministrador)

    LugarEjecucion_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "LugarEjecucion_Seleccionar_Cargar: " & Err.Description
    End If
    LugarEjecucion_Seleccionar_Cargar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- LugarEjecucion_Eliminar_Borrar --------------------------------------------------
' Pure-data delete: given a stub entity {IDLugarEjecucion, LugarEjecucion, DESCRIPCION}
' and an optional injected prompt result, decide whether to delegate to DAO.
'
' p_PromptResult conventions (per e2e rule #5):
'   0   -> show a real MsgBox (production path; form never calls this — only direct
'          test fixtures may pass 0 to exercise the live prompt)
'   vbYes (6) -> user said yes, proceed with delete
'   vbNo (7)  -> user cancelled, return value="cancelled"
'
' For tests, callers MUST pass vbYes/vbNo explicitly — never 0 — so the test is
' never blocked by a real modal in headless COM.
'
' Returns JSON: {ok, value:cancelled|deleted, error, logs}.
Public Function LugarEjecucion_Eliminar_Borrar( _
    ByVal p_LugarEjecucion As Object, _
    Optional ByRef p_PromptResult As Long = 0, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    ' Validate stub entity.
    If p_LugarEjecucion Is Nothing Then
        p_Error = "LugarEjecucion_Eliminar_Borrar: p_LugarEjecucion is Nothing"
        LugarEjecucion_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idLugar As String
    Dim lugarNombre As String

    If p_LugarEjecucion.Exists("IDLugarEjecucion") Then
        idLugar = CStr(p_LugarEjecucion("IDLugarEjecucion"))
    Else
        idLugar = ""
    End If

    If p_LugarEjecucion.Exists("LugarEjecucion") Then
        lugarNombre = CStr(p_LugarEjecucion("LugarEjecucion"))
    Else
        lugarNombre = ""
    End If

    If Len(idLugar) = 0 Then
        p_Error = "LugarEjecucion_Eliminar_Borrar: missing IDLugarEjecucion"
        LugarEjecucion_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(0) = "Eliminar_Borrar: id=" & idLugar & ", name=" & lugarNombre

    ' Prompt handling.
    Dim promptResult As Long
    If p_PromptResult <> 0 Then
        promptResult = p_PromptResult
    Else
        ' Production path — real MsgBox. Tests MUST NOT pass 0 here.
        promptResult = MsgBox("¿Desea realmente borrar el lugar de ejecución seleccionado?", _
                              vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar")
    End If

    logs(2) = "Eliminar_Borrar: promptResult=" & promptResult

    If promptResult <> vbYes Then
        logs(3) = "Eliminar_Borrar: cancelled by user"
        LugarEjecucion_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    ' Delegate to DAO via Helper_EntidadCRUD + LugarEjecucionOperaciones.
    ' We pass empty confirmation string so the generic helper does NOT prompt again
    ' — the user already confirmed above.
    Dim m_LugarEjecucionOp As New LugarEjecucionOperaciones

    ' The Operaciones class expects a real LugarEjecucion class object — wrap the
    ' stub Dictionary values into a temporary LugarEjecucion instance.
    Dim wrapped As New LugarEjecucion
    If p_LugarEjecucion.Exists("IDLugarEjecucion") Then
        wrapped.IDLugarEjecucion = CStr(p_LugarEjecucion("IDLugarEjecucion"))
    End If
    If p_LugarEjecucion.Exists("LugarEjecucion") Then
        wrapped.LugarEjecucion = CStr(p_LugarEjecucion("LugarEjecucion"))
    End If
    If p_LugarEjecucion.Exists("DESCRIPCION") Then
        wrapped.DESCRIPCION = CStr(p_LugarEjecucion("DESCRIPCION"))
    End If

    Set m_LugarEjecucionOp.LugarEjecucion = wrapped

    Dim daoErr As String
    Dim daoResult As String
    daoResult = Helper_EntidadCRUD.EliminarEntidadGenerico( _
        m_LugarEjecucionOp, "LugarEjecucion", wrapped, "", daoErr)
    If daoErr <> "" Then
        p_Error = daoErr
        logs(3) = "Eliminar_Borrar: DAO error=" & daoErr
        LugarEjecucion_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' EliminarEntidadGenerico returns "" on cancel (we don't pass confirm so it skips),
    ' or the DAO result string on success/error.
    If Len(daoResult) > 0 And daoResult <> "OK" Then
        ' DAO returned a non-empty result other than OK — treat as cancellation/failure.
        logs(3) = "Eliminar_Borrar: DAO returned '" & daoResult & "' — treating as cancelled"
        LugarEjecucion_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    logs(3) = "Eliminar_Borrar: deleted"
    LugarEjecucion_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "LugarEjecucion_Eliminar_Borrar: " & Err.Description
    End If
    LugarEjecucion_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- LugarEjecucion_DobleClick_AbrirEdicion ------------------------------------------
' Pure-data dispatch: given the form's current state (HasElegir, EditarEnabled),
' decide which action the double-click should trigger.
'
' Dispatch rules:
'   - If p_HasElegir is True  -> action="choose"  (parent form's cmdElegir will fire)
'   - If p_EditarEnabled is True -> action="edit"  (form opens the alta form in edit mode)
'   - Else -> action="none"   (no-op; form just ignores the double-click)
'
' Returns JSON: {ok, payload:{action:choose|edit|none}, error, logs}.
Public Function LugarEjecucion_DobleClick_AbrirEdicion( _
    ByVal p_HasElegir As Boolean, _
    ByVal p_EditarEnabled As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim action As String

    If p_HasElegir Then
        action = "choose"
        logs(0) = "DobleClick_AbrirEdicion: choose dispatched (cmdElegir visible)"
    Else
        If p_EditarEnabled Then
            action = "edit"
            logs(0) = "DobleClick_AbrirEdicion: edit dispatched (ComandoEditar enabled)"
        Else
            action = "none"
            logs(0) = "DobleClick_AbrirEdicion: no action (cmdElegir hidden, ComandoEditar disabled)"
        End If
    End If

    payload("action") = action

    LugarEjecucion_DobleClick_AbrirEdicion = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "LugarEjecucion_DobleClick_AbrirEdicion: " & Err.Description
    End If
    LugarEjecucion_DobleClick_AbrirEdicion = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
