Attribute VB_Name = "modUsuariosHelper"
Option Compare Database
Option Explicit

' modUsuariosHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormUsuariosGestion (SELECT-ONLY Gestion form —
' no Alta/Editar/Eliminar buttons).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: Usuarios_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - helpers did `Application.Echo` directly
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_UsuariosHelper_OpenForm` opened a real Access form, causing VBE interruption
'     in headless test runs (user-reported 2026-06-26)
'
' New design (5 SELECT-only helpers — no Alta/Editar/Eliminar):
'   1. Usuarios_Abrir_Inicializar(p_EsAdministrador, p_Error) — Form_Open
'      Returns JSON: {ok, payload:{hasElegir}, error, logs}
'   2. Usuarios_Buscar_Listar(p_Usuarios, p_Filter, p_EstadoFiltro, p_Error)
'      — ComandoBuscar_Click (composite filter: text + activos)
'      Returns JSON: {ok, payload:{rowSource, count}, error, logs}
'   3. Usuarios_Seleccionar_Cargar(p_IDUsuario, p_Usuarios, p_EsAdministrador, p_Error)
'      — ListaFiltrados_Click (no buttons enabled — SELECT-only)
'      Returns JSON: {ok, payload:{entity, hasSelection}, error, logs}
'   4. Usuarios_Limpiar_Reset(p_Error) — ComandoLimpiar_Click
'      Returns JSON: {ok, payload:{ok}, error, logs}
'   5. Usuarios_DobleClick_AbrirDetalle(p_DetalleEnabled, p_Error)
'      — ListaFiltrados_DblClick (only "choose" action — no edit)
'      Returns JSON: {ok, payload:{action:choose|none}, error, logs}
'
' Note: only 5 names (no Alta/Editar/Eliminar because this is a SELECT-only form).
'
' Composite filter semantics (Activos):
'   - "Sí"   -> only active (FechaBaja is empty/null)
'   - "No"   -> only inactive (FechaBaja is a date)
'   - "Todos" -> no activos filter
'
' Default Activos differs between Abrir_Inicializar ("Sí") and Limpiar_Reset
' ("Todos") — pre-existing legacy behavior preserved. The form sets the default
' Activos itself (via FormInteraction_EstablecerValorControl in Form_Open);
' Limpiar_Reset tells the form what value to set via the JSON payload.
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.
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
Private Const USUARIOS_HEADERS As String = "Id;Nombre"

' Field separator inside rowSource (matches Access listbox RowSource convention).
Private Const USUARIOS_FIELDS_SEP As String = ";"

' Maximum rows emitted in rowSource. Defensive cap so a runaway test does not
' build a multi-MB CSV in a single string.
Private Const USUARIOS_MAX_ROWS As Long = 5000

' Activos filter values.
Private Const USUARIOS_ACTIVOS_SI As String = "Sí"
Private Const USUARIOS_ACTIVOS_NO As String = "No"
Private Const USUARIOS_ACTIVOS_TODOS As String = "Todos"

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all pilot code agrees.
Private Const USUARIOS_TEST_ID_BASE As Long = 900720


' === Local helpers (all at top per vba-access §10.1) ====================================

' --- BuildJsonPayload ----------------------------------------------------------------
' Wraps a Dictionary payload in the canonical JSON envelope:
'   {"ok":true,"value":null,"payload":<payloadJson>,"error":null,"logs":[...]}
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

' --- Usuarios_SerializarEntidad -----------------------------------------------
' Serializes a single entity Dictionary {ID, Nombre, FechaBaja}
' into a plain Dictionary suitable for JSON. Returns Nothing if p_Entidad is Nothing.
Private Function Usuarios_SerializarEntidad(ByVal p_Entidad As Object) As Object
    If p_Entidad Is Nothing Then
        Set Usuarios_SerializarEntidad = Nothing
        Exit Function
    End If

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")

    If p_Entidad.Exists("ID") Then
        out("ID") = CStr(p_Entidad("ID"))
    Else
        out("ID") = ""
    End If

    If p_Entidad.Exists("Nombre") Then
        out("Nombre") = CStr(p_Entidad("Nombre"))
    Else
        out("Nombre") = ""
    End If

    Set Usuarios_SerializarEntidad = out
End Function

' --- Usuarios_BuildRowLine ----------------------------------------------------
' Builds a single rowSource line (semicolon-separated) for the listbox.
' Sanitizes embedded semicolons by replacing them with colons (Access listbox
' RowSource cannot contain field separators inside a value).
Private Function Usuarios_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String _
) As String
    Dim safeNombre As String
    safeNombre = Replace(p_Nombre, USUARIOS_FIELDS_SEP, ":")

    Usuarios_BuildRowLine = p_ID & USUARIOS_FIELDS_SEP & safeNombre
End Function

' --- Usuarios_EsActivo --------------------------------------------------------
' Returns True when p_FechaBaja indicates an active user (empty/null/non-date).
' Split guards per vba-access §1.6.1.
Private Function Usuarios_EsActivo(ByVal p_FechaBaja As Variant) As Boolean
    If IsNull(p_FechaBaja) Then
        Usuarios_EsActivo = True
        Exit Function
    End If

    If Len("" & p_FechaBaja) = 0 Then
        Usuarios_EsActivo = True
        Exit Function
    End If

    If IsDate(p_FechaBaja) Then
        Usuarios_EsActivo = False
    Else
        Usuarios_EsActivo = True
    End If
End Function


' === Public API =========================================================================

' --- Usuarios_Abrir_Inicializar -----------------------------------------------
' Pure-data init: given admin flag + OpenArgs presence, decide whether to show
' cmdElegir (only shown when OpenArgs are present — parent form is using this
' form as a picker).
'
' Note: this helper does NOT default Activos — the form does it via
' FormInteraction_EstablecerValorControl in Form_Open (per legacy behavior:
' Abrir_Inicializar defaults "Sí", Limpiar_Reset defaults "Todos").
'
' Returns JSON: {ok, payload:{hasElegir, hasOpenArgs}, error, logs}.
Public Function Usuarios_Abrir_Inicializar( _
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

    Dim hasElegir As Boolean
    If p_HasOpenArgs Then
        hasElegir = True
    Else
        hasElegir = False
    End If

    payload("hasElegir") = hasElegir
    payload("hasOpenArgs") = p_HasOpenArgs

    logs(0) = "Abrir_Inicializar: admin=" & CStr(p_EsAdministrador) & _
              ", hasOpenArgs=" & CStr(p_HasOpenArgs) & _
              ", hasElegir=" & CStr(hasElegir)

    Usuarios_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Usuarios_Abrir_Inicializar: " & Err.Description
    End If
    Usuarios_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Usuarios_Buscar_Listar ----------------------------------------------------
' Pure-data filter: given a Dictionary of entities {ID -> stubEntity} and
' composite filter (text + Activos), returns a semicolon-separated rowSource
' CSV + count.
'
' p_Usuarios is a Scripting.Dictionary (or Nothing). Each value must expose
' "ID", "Nombre", "FechaBaja" keys. FechaBaja is empty/null/non-date for active
' users, a date for inactive.
'
' p_EstadoFiltro: "Sí" | "No" | "Todos" (any other value treated as "Todos").
'
' Text filter: case-insensitive substring match on Nombre.
'
' Returns JSON: {ok, payload:{rowSource, count}, error, logs}.
Public Function Usuarios_Buscar_Listar( _
    ByVal p_Usuarios As Object, _
    ByVal p_Filter As String, _
    ByVal p_EstadoFiltro As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(7)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim rowSource As String
    rowSource = ""

    Dim rowCount As Long
    rowCount = 0

    ' Defensive: empty/Nothing collection -> 0 rows, no filter applied.
    If p_Usuarios Is Nothing Then
        logs(0) = "Buscar_Listar: p_Usuarios is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        payload("activos") = ""
        Usuarios_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim filterText As String
    filterText = "" & p_Filter

    ' Activos filter: normalize to canonical value.
    Dim estadoFiltro As String
    Dim estadoIn As String
    estadoIn = "" & p_EstadoFiltro
    If estadoIn = USUARIOS_ACTIVOS_SI Then
        estadoFiltro = USUARIOS_ACTIVOS_SI
    ElseIf estadoIn = USUARIOS_ACTIVOS_NO Then
        estadoFiltro = USUARIOS_ACTIVOS_NO
    Else
        estadoFiltro = USUARIOS_ACTIVOS_TODOS
    End If

    Dim k As Variant
    Dim entity As Object
    Dim currentId As String
    Dim currentNombre As String
    Dim currentFechaBaja As Variant

    For Each k In p_Usuarios.Keys
        If rowCount >= USUARIOS_MAX_ROWS Then
            logs(3) = "Buscar_Listar: hit MAX_ROWS=" & USUARIOS_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set entity = p_Usuarios(k)
        If entity Is Nothing Then
            ' Skip malformed entries — split guard before touching properties.
        Else
            If entity.Exists("ID") Then
                currentId = CStr(entity("ID"))
            Else
                currentId = ""
            End If

            If entity.Exists("Nombre") Then
                currentNombre = CStr(entity("Nombre"))
            Else
                currentNombre = ""
            End If

            If entity.Exists("FechaBaja") Then
                currentFechaBaja = entity("FechaBaja")
            Else
                currentFechaBaja = Null
            End If

            ' Text filter: substring on Nombre.
            Dim passText As Boolean
            passText = True
            If Len(filterText) > 0 Then
                If InStr(1, currentNombre, filterText, vbTextCompare) = 0 Then
                    passText = False
                End If
            End If

            ' Activos filter.
            Dim passActivos As Boolean
            passActivos = True
            If estadoFiltro = USUARIOS_ACTIVOS_SI Then
                ' Only active — FechaBaja must be empty/null/non-date.
                If Not Usuarios_EsActivo(currentFechaBaja) Then
                    passActivos = False
                End If
            ElseIf estadoFiltro = USUARIOS_ACTIVOS_NO Then
                ' Only inactive — FechaBaja must be a date.
                If Usuarios_EsActivo(currentFechaBaja) Then
                    passActivos = False
                End If
            End If
            ' "Todos": no activos filter — passActivos stays True.

            If passText And passActivos Then
                If Len(rowSource) > 0 Then
                    rowSource = rowSource & vbCrLf
                End If
                rowSource = rowSource & Usuarios_BuildRowLine(currentId, currentNombre)
                rowCount = rowCount + 1
            End If
        End If
        Set entity = Nothing
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    payload("activos") = estadoFiltro

    logs(0) = "Buscar_Listar: filterText=" & filterText & _
              ", activos=" & estadoFiltro & ", rows=" & rowCount
    logs(1) = "Buscar_Listar: rowSource length=" & Len(rowSource)

    Usuarios_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Usuarios_Buscar_Listar: " & Err.Description
    End If
    Usuarios_Buscar_Listar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Usuarios_Seleccionar_Cargar -----------------------------------------------
' Pure-data selection: given the selected ID and the entities Dictionary, returns
' the selected entity. SELECT-only form does NOT enable ComandoEditar/Eliminar
' (no such buttons) — only loads the selection state for the parent form's
' event dispatch.
'
' p_IDSeleccionado is "" or whitespace when nothing is selected — returns ok with
' entity=null, hasSelection=false.
'
' Returns JSON: {ok, payload:{entity, hasSelection}, error, logs}.
Public Function Usuarios_Seleccionar_Cargar( _
    ByVal p_IDUsuario As String, _
    ByVal p_Usuarios As Object, _
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
    selectedId = Trim$("" & p_IDUsuario)

    ' No selection — entity null.
    If Len(selectedId) = 0 Then
        Set payload("entity") = Nothing
        payload("hasSelection") = False
        logs(0) = "Seleccionar_Cargar: empty selection"
        Usuarios_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Look up the entity by ID in the Dictionary.
    Dim entity As Object
    If p_Usuarios Is Nothing Then
        Set entity = Nothing
    Else
        If p_Usuarios.Exists(selectedId) Then
            Set entity = p_Usuarios(selectedId)
        Else
            Set entity = Nothing
        End If
    End If

    If entity Is Nothing Then
        logs(0) = "Seleccionar_Cargar: id=" & selectedId & " not found in collection"
        Set payload("entity") = Nothing
        payload("hasSelection") = False
        Usuarios_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Found the entity — serialize for the form.
    Dim serializedEntity As Object
    Set serializedEntity = Usuarios_SerializarEntidad(entity)

    Set payload("entity") = serializedEntity
    payload("hasSelection") = True

    logs(0) = "Seleccionar_Cargar: id=" & selectedId
    logs(1) = "Seleccionar_Cargar: isAdmin=" & CStr(p_EsAdministrador)

    Usuarios_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Usuarios_Seleccionar_Cargar: " & Err.Description
    End If
    Usuarios_Seleccionar_Cargar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Usuarios_Limpiar_Reset ----------------------------------------------------
' Pure-data reset: returns the new default Activos value the form should set
' ("Todos"). The form clears its own text filter control.
'
' This helper does NOT touch any control — it just returns the canonical
' default so atoms can assert the decision without opening the form.
'
' Returns JSON: {ok, payload:{activosDefault}, error, logs}.
Public Function Usuarios_Limpiar_Reset( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    payload("activosDefault") = USUARIOS_ACTIVOS_TODOS

    logs(0) = "Limpiar_Reset: text filter clear, Activos=" & USUARIOS_ACTIVOS_TODOS

    Usuarios_Limpiar_Reset = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Usuarios_Limpiar_Reset: " & Err.Description
    End If
    Usuarios_Limpiar_Reset = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Usuarios_DobleClick_AbrirDetalle ------------------------------------------
' Pure-data dispatch: given the form's current state (DetalleEnabled = cmdElegir
' is visible), decide which action the double-click should trigger.
'
' Dispatch rules (SELECT-only form has NO edit button):
'   - If p_DetalleEnabled is True  -> action="choose"  (parent form's cmdElegir will fire)
'   - Else -> action="none"   (no-op; form just ignores the double-click)
'
' Returns JSON: {ok, payload:{action:choose|none}, error, logs}.
Public Function Usuarios_DobleClick_AbrirDetalle( _
    ByVal p_DetalleEnabled As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim action As String

    If p_DetalleEnabled Then
        action = "choose"
        logs(0) = "DobleClick_AbrirDetalle: choose dispatched (cmdElegir visible)"
    Else
        action = "none"
        logs(0) = "DobleClick_AbrirDetalle: no action (cmdElegir hidden — SELECT-only form has no edit)"
    End If

    payload("action") = action

    Usuarios_DobleClick_AbrirDetalle = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Usuarios_DobleClick_AbrirDetalle: " & Err.Description
    End If
    Usuarios_DobleClick_AbrirDetalle = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
