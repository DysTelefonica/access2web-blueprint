Attribute VB_Name = "modOficinaProgramaHelper"
Option Compare Database
Option Explicit

' modOficinaProgramaHelper — REWORK-PILOT (2026-06-26)
' Pure-data helpers for Form_FormOficinasProgramaGestion.
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: OficinaPrograma_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in the previous pilot, see docs/audit/pilot-oficina-programa.md):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_OficinaProgramaHelper_OpenForm` opened a real Access form, causing VBE interruption
'     in headless test runs (user-reported 2026-06-26)
'
' New design (5 helpers, NOT 8 — the other 3 are pure UI orchestration, kept in form):
'   1. OficinaPrograma_Abrir_Inicializar(p_EsAdministrador, p_HasOpenArgs, p_Error) — Form_Open
'      Returns JSON: {ok, payload:{showAlta, focusAlta, showElegir}, error, logs}
'   2. OficinaPrograma_Buscar_Listar(p_Oficinas, p_Filter, p_Error) — ComandoBuscar_Click
'      Returns JSON: {ok, payload:{rowSource, count}, error, logs}
'      (rowSource is a semicolon-separated CSV; the form assigns it to ListaFiltrados.RowSource)
'   3. OficinaPrograma_Seleccionar_Cargar(p_IDSeleccionado, p_Oficinas, p_EsAdministrador, p_Error)
'      — ListaFiltrados_Click
'      Returns JSON: {ok, payload:{entity:{...}, enableEditar, enableEliminar}, error, logs}
'   4. OficinaPrograma_Eliminar_Borrar(p_OficinaPrograma, p_PromptResult, p_Error)
'      — ComandoEliminar_Click
'      Returns JSON: {ok, value:cancelled|deleted, error, logs}
'      (p_PromptResult=0 means "ask the user via MsgBox"; non-zero is the injected answer for tests)
'   5. OficinaPrograma_DobleClick_AbrirEdicion(p_HasElegir, p_EditarEnabled, p_Error)
'      — ListaFiltrados_DblClick
'      Returns JSON: {ok, payload:{action:choose|edit|none}, error, logs}
'
' UI orchestration that stays in the form (rule #1 — never testable business logic):
'   - Alta button: form does DoCmd.OpenForm / FormularioAbierto / Forms(...)
'   - Edición button: form does DoCmd.OpenForm / FormularioAbierto / Forms(...)
'   - Limpiar button: form does Me.OficinaPrograma = Null (pure UI action, no helper needed)
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

' Column header for the filtered list (matches the legacy pilot's RowSource seed).
Private Const OFICINA_PROGRAMA_HEADERS As String = "IDOficinaPrograma;OFICINA;DESCRIPCIÓN"

' Field separator inside rowSource (matches Access listbox RowSource convention).
Private Const OFICINA_PROGRAMA_FIELDS_SEP As String = ";"

' Maximum rows emitted in rowSource. Defensive cap so a runaway test does not
' build a multi-MB CSV in a single string.
Private Const OFICINA_PROGRAMA_MAX_ROWS As Long = 5000

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all pilot code agrees.
Private Const OFICINA_PROGRAMA_TEST_ID_BASE As Long = 900500


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

    BuildJsonPayload = "{""ok"":" & IIf(p_Ok, "true", "false") & _
                       ",""value"":null" & _
                       ",""payload"":" & payloadJson & _
                       ",""error"":" & errorJson & _
                       ",""logs"":" & logsJson & "}"
End Function

' --- OficinaPrograma_SerializarEntidad -----------------------------------------------
' Serializes a single entity Dictionary {IDOficinaPrograma, OficinaPrograma, DESCRIPCION}
' into a plain Dictionary suitable for JSON. Returns Nothing if p_Entidad is Nothing
' or missing required keys (split guards per vba-access §1.6.1).
Private Function OficinaPrograma_SerializarEntidad(ByVal p_Entidad As Object) As Object
    If p_Entidad Is Nothing Then
        Set OficinaPrograma_SerializarEntidad = Nothing
        Exit Function
    End If

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")

    If p_Entidad.Exists("IDOficinaPrograma") Then
        out("IDOficinaPrograma") = CStr(p_Entidad("IDOficinaPrograma"))
    Else
        out("IDOficinaPrograma") = ""
    End If

    If p_Entidad.Exists("OficinaPrograma") Then
        out("OficinaPrograma") = CStr(p_Entidad("OficinaPrograma"))
    Else
        out("OficinaPrograma") = ""
    End If

    If p_Entidad.Exists("DESCRIPCION") Then
        out("DESCRIPCION") = CStr(p_Entidad("DESCRIPCION"))
    Else
        out("DESCRIPCION") = ""
    End If

    Set OficinaPrograma_SerializarEntidad = out
End Function

' --- OficinaPrograma_BuildRowLine ----------------------------------------------------
' Builds a single rowSource line (semicolon-separated) for the listbox.
' Sanitizes embedded semicolons by replacing them with colons (Access listbox
' RowSource cannot contain field separators inside a value).
Private Function OficinaPrograma_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_Name As String, _
    ByVal p_Desc As String _
) As String
    Dim safeName As String
    safeName = Replace(p_Name, OFICINA_PROGRAMA_FIELDS_SEP, ":")

    Dim safeDesc As String
    safeDesc = Replace(p_Desc, OFICINA_PROGRAMA_FIELDS_SEP, ":")

    OficinaPrograma_BuildRowLine = p_ID & OFICINA_PROGRAMA_FIELDS_SEP & _
                                    safeName & OFICINA_PROGRAMA_FIELDS_SEP & _
                                    safeDesc
End Function


' === Public API =========================================================================

' --- OficinaPrograma_Abrir_Inicializar -----------------------------------------------
' Pure-data init: given admin flag + OpenArgs presence, decide which buttons to show
' and where to put focus. The form renders the JSON into its own controls.
' Returns JSON: {ok, payload:{showAlta, focusAlta, showElegir}, error, logs}.
Public Function OficinaPrograma_Abrir_Inicializar( _
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

    OficinaPrograma_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaPrograma_Abrir_Inicializar: " & Err.Description
    End If
    OficinaPrograma_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OficinaPrograma_Buscar_Listar ----------------------------------------------------
' Pure-data filter: given a Dictionary of entities {ID -> stubEntity} and a filter
' text, returns a semicolon-separated rowSource CSV + count. The form assigns
' rowSource to Me.ListaFiltrados.RowSource (Access handles the header row separately).
'
' p_Oficinas is a Scripting.Dictionary (or Nothing). Each value must expose
' "IDOficinaPrograma", "OficinaPrograma", "DESCRIPCION" keys.
'
' Returns JSON: {ok, payload:{rowSource, count}, error, logs}.
Public Function OficinaPrograma_Buscar_Listar( _
    ByVal p_Oficinas As Object, _
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
    If p_Oficinas Is Nothing Then
        logs(0) = "Buscar_Listar: p_Oficinas is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        OficinaPrograma_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim filterText As String
    filterText = "" & p_Filter

    Dim k As Variant
    Dim entity As Object
    Dim currentName As String
    Dim currentId As String
    Dim currentDesc As String

    For Each k In p_Oficinas.Keys
        If rowCount >= OFICINA_PROGRAMA_MAX_ROWS Then
            logs(3) = "Buscar_Listar: hit MAX_ROWS=" & OFICINA_PROGRAMA_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set entity = p_Oficinas(k)
        If entity Is Nothing Then
            ' Skip malformed entries — split guard before touching properties.
        Else
            If entity.Exists("IDOficinaPrograma") Then
                currentId = CStr(entity("IDOficinaPrograma"))
            Else
                currentId = ""
            End If

            If entity.Exists("OficinaPrograma") Then
                currentName = CStr(entity("OficinaPrograma"))
            Else
                currentName = ""
            End If

            If entity.Exists("DESCRIPCION") Then
                currentDesc = CStr(entity("DESCRIPCION"))
            Else
                currentDesc = ""
            End If

            If Len(filterText) > 0 Then
                ' Case-insensitive substring match on the name field.
                If InStr(1, currentName, filterText, vbTextCompare) = 0 Then
                    ' No match — skip without counting.
                Else
                    If Len(rowSource) > 0 Then
                        rowSource = rowSource & vbCrLf
                    End If
                    rowSource = rowSource & OficinaPrograma_BuildRowLine(currentId, currentName, currentDesc)
                    rowCount = rowCount + 1
                End If
            Else
                ' No filter — include everything.
                If Len(rowSource) > 0 Then
                    rowSource = rowSource & vbCrLf
                End If
                rowSource = rowSource & OficinaPrograma_BuildRowLine(currentId, currentName, currentDesc)
                rowCount = rowCount + 1
            End If
        End If
        Set entity = Nothing
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount

    logs(0) = "Buscar_Listar: filterText=" & filterText & ", rows=" & rowCount
    logs(1) = "Buscar_Listar: rowSource length=" & Len(rowSource)

    OficinaPrograma_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaPrograma_Buscar_Listar: " & Err.Description
    End If
    OficinaPrograma_Buscar_Listar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OficinaPrograma_Seleccionar_Cargar -----------------------------------------------
' Pure-data selection: given the selected ID (extracted by the form from the listbox
' Column(0)) and the entities Dictionary, returns the selected entity plus the
' enable flags for ComandoEditar/Eliminar.
'
' p_IDSeleccionado is "" or whitespace when nothing is selected — returns ok with
' enableEditar=false, enableEliminar=false, entity=null.
'
' Returns JSON: {ok, payload:{entity, enableEditar, enableEliminar}, error, logs}.
Public Function OficinaPrograma_Seleccionar_Cargar( _
    ByVal p_IDSeleccionado As String, _
    ByVal p_Oficinas As Object, _
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
        OficinaPrograma_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Look up the entity by ID in the Dictionary.
    Dim entity As Object
    If p_Oficinas Is Nothing Then
        Set entity = Nothing
    Else
        If p_Oficinas.Exists(selectedId) Then
            Set entity = p_Oficinas(selectedId)
        Else
            Set entity = Nothing
        End If
    End If

    If entity Is Nothing Then
        logs(0) = "Seleccionar_Cargar: id=" & selectedId & " not found in collection"
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        OficinaPrograma_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Found the entity — always enable Editar; Eliminar only for admins.
    Dim serializedEntity As Object
    Set serializedEntity = OficinaPrograma_SerializarEntidad(entity)

    Set payload("entity") = serializedEntity
    payload("enableEditar") = True
    If p_EsAdministrador Then
        payload("enableEliminar") = True
    Else
        payload("enableEliminar") = False
    End If

    logs(0) = "Seleccionar_Cargar: id=" & selectedId
    logs(1) = "Seleccionar_Cargar: isAdmin=" & CStr(p_EsAdministrador)

    OficinaPrograma_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaPrograma_Seleccionar_Cargar: " & Err.Description
    End If
    OficinaPrograma_Seleccionar_Cargar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OficinaPrograma_Eliminar_Borrar -------------------------------------------------
' Pure-data delete: given a stub entity {IDOficinaPrograma, OficinaPrograma, DESCRIPCION}
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
' Returns JSON: {ok, value:cancelled|deleted|skipped-n-a, error, logs}.
Public Function OficinaPrograma_Eliminar_Borrar( _
    ByVal p_OficinaPrograma As Object, _
    Optional ByRef p_PromptResult As Long = 0, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    ' Validate stub entity.
    If p_OficinaPrograma Is Nothing Then
        p_Error = "OficinaPrograma_Eliminar_Borrar: p_OficinaPrograma is Nothing"
        OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idOficina As String
    Dim oficinaProgramaName As String

    If p_OficinaPrograma.Exists("IDOficinaPrograma") Then
        idOficina = CStr(p_OficinaPrograma("IDOficinaPrograma"))
    Else
        idOficina = ""
    End If

    If p_OficinaPrograma.Exists("OficinaPrograma") Then
        oficinaProgramaName = CStr(p_OficinaPrograma("OficinaPrograma"))
    Else
        oficinaProgramaName = ""
    End If

    If Len(idOficina) = 0 Then
        p_Error = "OficinaPrograma_Eliminar_Borrar: missing IDOficinaPrograma"
        OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(0) = "Eliminar_Borrar: id=" & idOficina & ", name=" & oficinaProgramaName

    ' Guard against deleting the structural "N/A" sentinel — this matches the legacy
    ' behaviour in OficinaProgramaOperaciones.Eliminar and the user-facing message.
    If oficinaProgramaName = "N/A" Then
        p_Error = "Es una opción Estructural que otras herramientas necesitan. No se puede borrar"
        logs(1) = "Eliminar_Borrar: blocked N/A sentinel"
        OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Prompt handling.
    Dim promptResult As Long
    If p_PromptResult <> 0 Then
        promptResult = p_PromptResult
    Else
        ' Production path — real MsgBox. Tests MUST NOT pass 0 here.
        promptResult = MsgBox("¿Desea realmente borrar la Oficina de Programa seleccionada?", _
                              vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar")
    End If

    logs(2) = "Eliminar_Borrar: promptResult=" & promptResult

    If promptResult <> vbYes Then
        logs(3) = "Eliminar_Borrar: cancelled by user"
        OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    ' Delegate to DAO via Helper_EntidadCRUD + OficinaProgramaOperaciones.
    ' We pass empty confirmation string so the generic helper does NOT prompt again
    ' — the user already confirmed above.
    Dim m_OficinaProgramaOp As New OficinaProgramaOperaciones

    ' The Operaciones class expects a real OficinaPrograma class object — wrap the
    ' stub Dictionary values into a temporary OficinaPrograma instance.
    Dim wrapped As New OficinaPrograma
    If p_OficinaPrograma.Exists("IDOficinaPrograma") Then
        wrapped.IDOficinaPrograma = CStr(p_OficinaPrograma("IDOficinaPrograma"))
    End If
    If p_OficinaPrograma.Exists("OficinaPrograma") Then
        wrapped.OficinaPrograma = CStr(p_OficinaPrograma("OficinaPrograma"))
    End If
    If p_OficinaPrograma.Exists("DESCRIPCION") Then
        wrapped.DESCRIPCION = CStr(p_OficinaPrograma("DESCRIPCION"))
    End If

    Set m_OficinaProgramaOp.OficinaPrograma = wrapped

    Dim daoErr As String
    Dim daoResult As String
    daoResult = Helper_EntidadCRUD.EliminarEntidadGenerico( _
        m_OficinaProgramaOp, "OficinaPrograma", wrapped, "", daoErr)
    If daoErr <> "" Then
        p_Error = daoErr
        logs(3) = "Eliminar_Borrar: DAO error=" & daoErr
        OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' EliminarEntidadGenerico returns "" on cancel (we don't pass confirm so it skips),
    ' or the DAO result string on success/error.
    If Len(daoResult) > 0 And daoResult <> "OK" Then
        ' DAO returned a non-empty result other than OK — treat as cancellation/failure.
        logs(3) = "Eliminar_Borrar: DAO returned '" & daoResult & "' — treating as cancelled"
        OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    logs(3) = "Eliminar_Borrar: deleted"
    OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaPrograma_Eliminar_Borrar: " & Err.Description
    End If
    OficinaPrograma_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- OficinaPrograma_DobleClick_AbrirEdicion ------------------------------------------
' Pure-data dispatch: given the form's current state (HasElegir, EditarEnabled),
' decide which action the double-click should trigger.
'
' Dispatch rules:
'   - If p_HasElegir is True  -> action="choose"  (parent form's cmdElegir will fire)
'   - If p_EditarEnabled is True -> action="edit"  (form opens the alta form in edit mode)
'   - Else -> action="none"   (no-op; form just ignores the double-click)
'
' Returns JSON: {ok, payload:{action:choose|edit|none}, error, logs}.
Public Function OficinaPrograma_DobleClick_AbrirEdicion( _
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

    OficinaPrograma_DobleClick_AbrirEdicion = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "OficinaPrograma_DobleClick_AbrirEdicion: " & Err.Description
    End If
    OficinaPrograma_DobleClick_AbrirEdicion = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function