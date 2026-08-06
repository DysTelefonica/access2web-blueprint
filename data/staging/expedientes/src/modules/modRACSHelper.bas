Attribute VB_Name = "modRACSHelper"
Option Compare Database
Option Explicit

' modRACSHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormRACSGestion.
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: RACS_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - helpers did `DoCmd.OpenForm` and `Application.Echo` directly
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_RACSHelper_OpenForm` opened a real Access form
'
' Naming convention: prefix is `RACS_` (plural) to disambiguate from `RAC_*`
' helpers extracted for the Alta/Edición form (Form_FormRAC) in modRACHelper.
' The two modules have different roles: this one is the Gestion list, the
' other is the Alta/Edición single-record form.
'
' New design (5 helpers, NOT 8 — the other 3 are pure UI orchestration, kept in form):
'   1. RACS_Abrir_Inicializar(p_EsAdministrador, p_HasOpenArgs, p_Error)
'   2. RACS_Buscar_Listar(p_RACs, p_Filter, p_Error)
'   3. RACS_Seleccionar_Cargar(p_IDSeleccionado, p_RACs, p_EsAdministrador, p_Error)
'   4. RACS_Eliminar_Borrar(p_RAC, p_PromptResult, p_Error)
'   5. RACS_DobleClick_AbrirEdicion(p_HasElegir, p_EditarEnabled, p_Error)
'
' UI orchestration that stays in the form (rule #1 — never testable business logic):
'   - Alta button: form does DoCmd.OpenForm / FormInteraction_FormularioAbierto / Forms(...)
'   - Edición button: form does DoCmd.OpenForm / FormInteraction_FormularioAbierto / Forms(...)
'   - Limpiar button: form does Me.RAC = Null (pure UI action, no helper needed)
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

' Column header for the filtered list (matches Access listbox RowSource convention).
Private Const RACS_HEADERS As String = "IDRAC;RAC"

' Field separator inside rowSource (matches Access listbox RowSource convention).
Private Const RACS_FIELDS_SEP As String = ";"

' Maximum rows emitted in rowSource. Defensive cap so a runaway test does not
' build a multi-MB CSV in a single string.
Private Const RACS_MAX_ROWS As Long = 5000

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all rework code agrees.
Private Const RACS_TEST_ID_BASE As Long = 900800


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

' --- RACS_SerializarEntidad -----------------------------------------------------------
' Serializes a single entity Dictionary {IDRAC, RAC, CORREO, DESCRIPCION} into a
' plain Dictionary suitable for JSON. Returns Nothing if p_Entidad is Nothing
' or missing required keys (split guards per vba-access §1.6.1).
Private Function RACS_SerializarEntidad(ByVal p_Entidad As Object) As Object
    If p_Entidad Is Nothing Then
        Set RACS_SerializarEntidad = Nothing
        Exit Function
    End If

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")

    If p_Entidad.Exists("IDRAC") Then
        out("IDRAC") = CStr(p_Entidad("IDRAC"))
    Else
        out("IDRAC") = ""
    End If

    If p_Entidad.Exists("RAC") Then
        out("RAC") = CStr(p_Entidad("RAC"))
    Else
        out("RAC") = ""
    End If

    If p_Entidad.Exists("CORREO") Then
        out("CORREO") = CStr(p_Entidad("CORREO"))
    Else
        out("CORREO") = ""
    End If

    If p_Entidad.Exists("DESCRIPCION") Then
        out("DESCRIPCION") = CStr(p_Entidad("DESCRIPCION"))
    Else
        out("DESCRIPCION") = ""
    End If

    Set RACS_SerializarEntidad = out
End Function

' --- RACS_BuildRowLine ---------------------------------------------------------------
' Builds a single rowSource line (semicolon-separated) for the listbox.
' Sanitizes embedded semicolons by replacing them with colons (Access listbox
' RowSource cannot contain field separators inside a value).
Private Function RACS_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_RAC As String _
) As String
    Dim safeRAC As String
    safeRAC = Replace(p_RAC, RACS_FIELDS_SEP, ":")

    RACS_BuildRowLine = p_ID & RACS_FIELDS_SEP & safeRAC
End Function


' === Public API =========================================================================

' --- RACS_Abrir_Inicializar -----------------------------------------------------------
Public Function RACS_Abrir_Inicializar( _
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

    RACS_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RACS_Abrir_Inicializar: " & Err.Description
    End If
    RACS_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- RACS_Buscar_Listar --------------------------------------------------------------
Public Function RACS_Buscar_Listar( _
    ByVal p_RACs As Object, _
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
    If p_RACs Is Nothing Then
        logs(0) = "Buscar_Listar: p_RACs is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        RACS_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim filterText As String
    filterText = "" & p_Filter

    Dim k As Variant
    Dim entity As Object
    Dim currentId As String
    Dim currentRAC As String

    For Each k In p_RACs.Keys
        If rowCount >= RACS_MAX_ROWS Then
            logs(3) = "Buscar_Listar: hit MAX_ROWS=" & RACS_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set entity = p_RACs(k)
        If entity Is Nothing Then
            ' Skip malformed entries.
        Else
            If entity.Exists("IDRAC") Then
                currentId = CStr(entity("IDRAC"))
            Else
                currentId = ""
            End If

            If entity.Exists("RAC") Then
                currentRAC = CStr(entity("RAC"))
            Else
                currentRAC = ""
            End If

            If Len(filterText) > 0 Then
                If InStr(1, currentRAC, filterText, vbTextCompare) = 0 Then
                    ' No match — skip without counting.
                Else
                    If Len(rowSource) > 0 Then
                        rowSource = rowSource & vbCrLf
                    End If
                    rowSource = rowSource & RACS_BuildRowLine(currentId, currentRAC)
                    rowCount = rowCount + 1
                End If
            Else
                ' No filter — include everything.
                If Len(rowSource) > 0 Then
                    rowSource = rowSource & vbCrLf
                End If
                rowSource = rowSource & RACS_BuildRowLine(currentId, currentRAC)
                rowCount = rowCount + 1
            End If
        End If
        Set entity = Nothing
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount

    logs(0) = "Buscar_Listar: filterText=" & filterText & ", rows=" & rowCount
    logs(1) = "Buscar_Listar: rowSource length=" & Len(rowSource)

    RACS_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RACS_Buscar_Listar: " & Err.Description
    End If
    RACS_Buscar_Listar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- RACS_Seleccionar_Cargar ---------------------------------------------------------
Public Function RACS_Seleccionar_Cargar( _
    ByVal p_IDSeleccionado As String, _
    ByVal p_RACs As Object, _
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
        RACS_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Look up the entity by ID in the Dictionary.
    Dim entity As Object
    If p_RACs Is Nothing Then
        Set entity = Nothing
    Else
        If p_RACs.Exists(selectedId) Then
            Set entity = p_RACs(selectedId)
        Else
            Set entity = Nothing
        End If
    End If

    If entity Is Nothing Then
        logs(0) = "Seleccionar_Cargar: id=" & selectedId & " not found in collection"
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        RACS_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Found the entity — always enable Editar; Eliminar only for admins.
    Dim serializedEntity As Object
    Set serializedEntity = RACS_SerializarEntidad(entity)

    Set payload("entity") = serializedEntity
    payload("enableEditar") = True
    If p_EsAdministrador Then
        payload("enableEliminar") = True
    Else
        payload("enableEliminar") = False
    End If

    logs(0) = "Seleccionar_Cargar: id=" & selectedId
    logs(1) = "Seleccionar_Cargar: isAdmin=" & CStr(p_EsAdministrador)

    RACS_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RACS_Seleccionar_Cargar: " & Err.Description
    End If
    RACS_Seleccionar_Cargar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- RACS_Eliminar_Borrar ------------------------------------------------------------
Public Function RACS_Eliminar_Borrar( _
    ByVal p_RAC As Object, _
    Optional ByRef p_PromptResult As Long = 0, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    ' Validate stub entity.
    If p_RAC Is Nothing Then
        p_Error = "RACS_Eliminar_Borrar: p_RAC is Nothing"
        RACS_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idRAC As String
    Dim racNombre As String

    If p_RAC.Exists("IDRAC") Then
        idRAC = CStr(p_RAC("IDRAC"))
    Else
        idRAC = ""
    End If

    If p_RAC.Exists("RAC") Then
        racNombre = CStr(p_RAC("RAC"))
    Else
        racNombre = ""
    End If

    If Len(idRAC) = 0 Then
        p_Error = "RACS_Eliminar_Borrar: missing IDRAC"
        RACS_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(0) = "Eliminar_Borrar: id=" & idRAC & ", name=" & racNombre

    ' Prompt handling.
    Dim promptResult As Long
    If p_PromptResult <> 0 Then
        promptResult = p_PromptResult
    Else
        ' Production path — real MsgBox. Tests MUST NOT pass 0 here.
        promptResult = MsgBox("¿Desea realmente borrar el RAC seleccionado?", _
                              vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar")
    End If

    logs(2) = "Eliminar_Borrar: promptResult=" & promptResult

    If promptResult <> vbYes Then
        logs(3) = "Eliminar_Borrar: cancelled by user"
        RACS_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    ' Delegate to DAO via Helper_EntidadCRUD + RACOperaciones.
    Dim m_RACOp As New RACOperaciones

    Dim wrapped As New RAC
    If p_RAC.Exists("IDRAC") Then
        wrapped.IDRAC = CStr(p_RAC("IDRAC"))
    End If
    If p_RAC.Exists("RAC") Then
        wrapped.RAC = CStr(p_RAC("RAC"))
    End If
    If p_RAC.Exists("CORREO") Then
        wrapped.CORREO = CStr(p_RAC("CORREO"))
    End If
    If p_RAC.Exists("DESCRIPCION") Then
        wrapped.DESCRIPCION = CStr(p_RAC("DESCRIPCION"))
    End If

    Set m_RACOp.RAC = wrapped

    Dim daoErr As String
    Dim daoResult As String
    daoResult = Helper_EntidadCRUD.EliminarEntidadGenerico( _
        m_RACOp, "RAC", wrapped, "", daoErr)
    If daoErr <> "" Then
        p_Error = daoErr
        logs(3) = "Eliminar_Borrar: DAO error=" & daoErr
        RACS_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If Len(daoResult) > 0 And daoResult <> "OK" Then
        logs(3) = "Eliminar_Borrar: DAO returned '" & daoResult & "' — treating as cancelled"
        RACS_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    logs(3) = "Eliminar_Borrar: deleted"
    RACS_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RACS_Eliminar_Borrar: " & Err.Description
    End If
    RACS_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- RACS_DobleClick_AbrirEdicion ----------------------------------------------------
Public Function RACS_DobleClick_AbrirEdicion( _
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

    RACS_DobleClick_AbrirEdicion = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RACS_DobleClick_AbrirEdicion: " & Err.Description
    End If
    RACS_DobleClick_AbrirEdicion = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
