Attribute VB_Name = "modComercialHelper"
Option Compare Database
Option Explicit

' modComercialHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormComercialesGestion.
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: Comercial_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed:
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - helpers did `DoCmd.OpenForm` and `Application.Echo` directly
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'
' New design (5 helpers):
'   1. Comercial_Abrir_Inicializar(p_EsAdministrador, p_HasOpenArgs, p_Error)
'      Returns JSON: {ok, payload:{showAlta, focusAlta, showElegir}, error, logs}
'   2. Comercial_Buscar_Listar(p_Comerciales, p_Filter, p_Error)
'      Returns JSON: {ok, payload:{rowSource, count}, error, logs}
'      (rowSource is semicolon-separated; form assigns to ListaFiltrados.RowSource)
'      Filter field: Comercial (text name only, NOT DESCRIPCION — legacy listbox is 2-col)
'   3. Comercial_Seleccionar_Cargar(p_IDSeleccionado, p_Comerciales, p_EsAdministrador, p_Error)
'      Returns JSON: {ok, payload:{entity, enableEditar, enableEliminar}, error, logs}
'   4. Comercial_Eliminar_Borrar(p_Comercial, p_PromptResult, p_Error)
'      Returns JSON: {ok, value:cancelled|deleted, error, logs}
'      (p_PromptResult=0 means "ask the user via MsgBox"; non-zero is the injected answer)
'   5. Comercial_DobleClick_AbrirEdicion(p_HasElegir, p_EditarEnabled, p_Error)
'      Returns JSON: {ok, payload:{action:choose|edit|none}, error, logs}
'
' UI orchestration that stays in the form (rule #1):
'   - Alta button: form does DoCmd.OpenForm "FormComercial"
'   - Edición button: form does DoCmd.OpenForm "FormComercial" with m_ObjComercialActivo set
'   - Limpiar button: form does Me.Comercial = Null (pure UI action, no helper needed)
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.

' === Module-level constants (all at top per vba-access §10.1) ============================

' Column header for the filtered list. Legacy listbox shows ID + Comercial only
' (DESCRIPCION is NOT in the listbox for Comercial — keep the 2-col shape).
Private Const COMERCIAL_HEADERS As String = "IDComercial;Comercial"

' Field separator inside rowSource.
Private Const COMERCIAL_FIELDS_SEP As String = ";"

' Maximum rows emitted in rowSource. Defensive cap so a runaway test does not
' build a multi-MB CSV in a single string.
Private Const COMERCIAL_MAX_ROWS As Long = 5000

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs.
Private Const COMERCIAL_TEST_ID_BASE As Long = 900500


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

' --- Comercial_SerializarEntidad -----------------------------------------------------
' Serializes a single entity Dictionary {IDComercial, Comercial, DESCRIPCION} into a
' plain Dictionary suitable for JSON. Returns Nothing if p_Entidad is Nothing or
' missing required keys (split guards per vba-access §1.6.1).
Private Function Comercial_SerializarEntidad(ByVal p_Entidad As Object) As Object
    If p_Entidad Is Nothing Then
        Set Comercial_SerializarEntidad = Nothing
        Exit Function
    End If

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")

    If p_Entidad.Exists("IDComercial") Then
        out("IDComercial") = CStr(p_Entidad("IDComercial"))
    Else
        out("IDComercial") = ""
    End If

    If p_Entidad.Exists("Comercial") Then
        out("Comercial") = CStr(p_Entidad("Comercial"))
    Else
        out("Comercial") = ""
    End If

    If p_Entidad.Exists("DESCRIPCION") Then
        out("DESCRIPCION") = CStr(p_Entidad("DESCRIPCION"))
    Else
        out("DESCRIPCION") = ""
    End If

    Set Comercial_SerializarEntidad = out
End Function

' --- Comercial_BuildRowLine ---------------------------------------------------------
' Builds a single rowSource line for the listbox (2-col: ID + Comercial).
' Sanitizes embedded semicolons by replacing them with colons.
Private Function Comercial_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String _
) As String
    Dim safeNombre As String
    safeNombre = Replace(p_Nombre, COMERCIAL_FIELDS_SEP, ":")

    Comercial_BuildRowLine = p_ID & COMERCIAL_FIELDS_SEP & safeNombre
End Function


' === Public API =========================================================================

' --- Comercial_Abrir_Inicializar ---------------------------------------------------
' Pure-data init: given admin flag + OpenArgs presence, decide which buttons to show
' and where to put focus. The form renders the JSON into its own controls.
' Returns JSON: {ok, payload:{showAlta, focusAlta, showElegir}, error, logs}.
Public Function Comercial_Abrir_Inicializar( _
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

    Comercial_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Comercial_Abrir_Inicializar: " & Err.Description
    End If
    Comercial_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Comercial_Buscar_Listar --------------------------------------------------------
' Pure-data filter: given a Dictionary of entities {ID -> stubEntity} and a filter
' text, returns a semicolon-separated rowSource CSV + count.
'
' p_Comerciales is a Scripting.Dictionary (or Nothing). Each value must expose
' "IDComercial", "Comercial", "DESCRIPCION" keys.
'
' Filter logic: substring match on the Comercial name field. Empty filter -> include all.
'
' Returns JSON: {ok, payload:{rowSource, count}, error, logs}.
Public Function Comercial_Buscar_Listar( _
    ByVal p_Comerciales As Object, _
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
    If p_Comerciales Is Nothing Then
        logs(0) = "Buscar_Listar: p_Comerciales is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        Comercial_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim filterText As String
    filterText = "" & p_Filter

    Dim k As Variant
    Dim entity As Object
    Dim currentId As String
    Dim currentNombre As String

    For Each k In p_Comerciales.Keys
        If rowCount >= COMERCIAL_MAX_ROWS Then
            logs(3) = "Buscar_Listar: hit MAX_ROWS=" & COMERCIAL_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set entity = p_Comerciales(k)
        If entity Is Nothing Then
            ' Skip malformed entries.
        Else
            If entity.Exists("IDComercial") Then
                currentId = CStr(entity("IDComercial"))
            Else
                currentId = ""
            End If

            If entity.Exists("Comercial") Then
                currentNombre = CStr(entity("Comercial"))
            Else
                currentNombre = ""
            End If

            ' Filter on the Comercial name field (matches legacy behavior).
            If Len(filterText) > 0 Then
                If InStr(1, currentNombre, filterText, vbTextCompare) = 0 Then
                    ' No match — skip without counting.
                Else
                    If Len(rowSource) > 0 Then
                        rowSource = rowSource & vbCrLf
                    End If
                    rowSource = rowSource & Comercial_BuildRowLine(currentId, currentNombre)
                    rowCount = rowCount + 1
                End If
            Else
                ' No filter — include everything.
                If Len(rowSource) > 0 Then
                    rowSource = rowSource & vbCrLf
                End If
                rowSource = rowSource & Comercial_BuildRowLine(currentId, currentNombre)
                rowCount = rowCount + 1
            End If
        End If
        Set entity = Nothing
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount

    logs(0) = "Buscar_Listar: filterText=" & filterText & ", rows=" & rowCount
    logs(1) = "Buscar_Listar: rowSource length=" & Len(rowSource)

    Comercial_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Comercial_Buscar_Listar: " & Err.Description
    End If
    Comercial_Buscar_Listar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Comercial_Seleccionar_Cargar ---------------------------------------------------
' Pure-data selection: given the selected ID (extracted by the form from the listbox
' Column(0)) and the entities Dictionary, returns the selected entity plus the
' enable flags for ComandoEditar/Eliminar.
'
' Returns JSON: {ok, payload:{entity, enableEditar, enableEliminar}, error, logs}.
Public Function Comercial_Seleccionar_Cargar( _
    ByVal p_IDSeleccionado As String, _
    ByVal p_Comerciales As Object, _
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
        Comercial_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim entity As Object
    If p_Comerciales Is Nothing Then
        Set entity = Nothing
    Else
        If p_Comerciales.Exists(selectedId) Then
            Set entity = p_Comerciales(selectedId)
        Else
            Set entity = Nothing
        End If
    End If

    If entity Is Nothing Then
        logs(0) = "Seleccionar_Cargar: id=" & selectedId & " not found in collection"
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        Comercial_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim serializedEntity As Object
    Set serializedEntity = Comercial_SerializarEntidad(entity)

    Set payload("entity") = serializedEntity
    payload("enableEditar") = True
    If p_EsAdministrador Then
        payload("enableEliminar") = True
    Else
        payload("enableEliminar") = False
    End If

    logs(0) = "Seleccionar_Cargar: id=" & selectedId
    logs(1) = "Seleccionar_Cargar: isAdmin=" & CStr(p_EsAdministrador)

    Comercial_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Comercial_Seleccionar_Cargar: " & Err.Description
    End If
    Comercial_Seleccionar_Cargar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Comercial_Eliminar_Borrar -----------------------------------------------------
' Pure-data delete: given a stub entity {IDComercial, Comercial, DESCRIPCION} and an
' optional injected prompt result, decide whether to delegate to DAO.
'
' For tests, callers MUST pass vbYes/vbNo explicitly — never 0 — so the test is
' never blocked by a real modal in headless COM.
'
' Returns JSON: {ok, value:cancelled|deleted, error, logs}.
Public Function Comercial_Eliminar_Borrar( _
    ByVal p_Comercial As Object, _
    Optional ByRef p_PromptResult As Long = 0, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Comercial Is Nothing Then
        p_Error = "Comercial_Eliminar_Borrar: p_Comercial is Nothing"
        Comercial_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idComercial As String
    Dim comercialNombre As String

    If p_Comercial.Exists("IDComercial") Then
        idComercial = CStr(p_Comercial("IDComercial"))
    Else
        idComercial = ""
    End If

    If p_Comercial.Exists("Comercial") Then
        comercialNombre = CStr(p_Comercial("Comercial"))
    Else
        comercialNombre = ""
    End If

    If Len(idComercial) = 0 Then
        p_Error = "Comercial_Eliminar_Borrar: missing IDComercial"
        Comercial_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(0) = "Eliminar_Borrar: id=" & idComercial & ", name=" & comercialNombre

    Dim promptResult As Long
    If p_PromptResult <> 0 Then
        promptResult = p_PromptResult
    Else
        promptResult = MsgBox("¿Desea realmente borrar al comercial seleccionado?", _
                              vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar")
    End If

    logs(2) = "Eliminar_Borrar: promptResult=" & promptResult

    If promptResult <> vbYes Then
        logs(3) = "Eliminar_Borrar: cancelled by user"
        Comercial_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    Dim m_ComercialOP As New ComercialOperaciones

    Dim wrapped As New Comercial
    If p_Comercial.Exists("IDComercial") Then
        wrapped.IDComercial = CStr(p_Comercial("IDComercial"))
    End If
    If p_Comercial.Exists("Comercial") Then
        wrapped.Comercial = CStr(p_Comercial("Comercial"))
    End If
    If p_Comercial.Exists("DESCRIPCION") Then
        wrapped.DESCRIPCION = CStr(p_Comercial("DESCRIPCION"))
    End If

    Set m_ComercialOP.Comercial = wrapped

    Dim daoErr As String
    Dim daoResult As String
    daoResult = Helper_EntidadCRUD.EliminarEntidadGenerico( _
        m_ComercialOP, "Comercial", wrapped, "", daoErr)
    If daoErr <> "" Then
        p_Error = daoErr
        logs(3) = "Eliminar_Borrar: DAO error=" & daoErr
        Comercial_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If Len(daoResult) > 0 And daoResult <> "OK" Then
        logs(3) = "Eliminar_Borrar: DAO returned '" & daoResult & "' — treating as cancelled"
        Comercial_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    logs(3) = "Eliminar_Borrar: deleted"
    Comercial_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Comercial_Eliminar_Borrar: " & Err.Description
    End If
    Comercial_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Comercial_DobleClick_AbrirEdicion ----------------------------------------------
' Pure-data dispatch: given the form's current state (HasElegir, EditarEnabled),
' decide which action the double-click should trigger.
'
' Returns JSON: {ok, payload:{action:choose|edit|none}, error, logs}.
Public Function Comercial_DobleClick_AbrirEdicion( _
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

    Comercial_DobleClick_AbrirEdicion = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Comercial_DobleClick_AbrirEdicion: " & Err.Description
    End If
    Comercial_DobleClick_AbrirEdicion = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
