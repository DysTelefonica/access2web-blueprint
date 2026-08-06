Attribute VB_Name = "modSuministradorAltaHelper"
Option Compare Database
Option Explicit

' modSuministradorAltaHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormSuministrador (Alta/Edición form).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: SuministradorAlta_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - helpers accepted `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests called `DoCmd.OpenForm TEST_FORM_NAME` and passed Forms(TEST_FORM_NAME) as p_Form
'   - `Test_SuministradorAltaHelper_OpenForm` opened a real Access form, causing
'     VBE interruption in headless test runs (user-reported 2026-06-26)
'
' IMPORTANT: This module covers the Alta/Edición form (Form_FormSuministrador),
' which is structurally different from the Gestion list (Form_FormSuministradoresGestion,
' covered by modSuministradorHelper). Prefix disambiguation:
'   - `SuministradorAlta_*` — Alta/Edición form helpers (this module)
'   - `Suministrador_*` (no Alta suffix) — Gestion list helpers (modSuministradorHelper)
'
' New design (4 helpers — different from Gestion list's 5):
'   1. SuministradorAlta_Abrir_Inicializar(p_Modo, p_IDEntidad, ByRef p_Entidad, p_Error)
'      — Form_Open
'      Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}
'   2. SuministradorAlta_VerificarCambios(p_ValoresActuales, p_ValoresOriginales, p_Error)
'      Returns JSON: {ok, payload:{hayCambios, diffs}, error, logs}
'   3. SuministradorAlta_Registrar(p_Valores, p_Error) — ComandoRegistrar_Click
'      Returns JSON: {ok, payload:{ok, id, modo}, error, logs}
'      The helper decides alta vs edicion based on the m_ObjSuministradorActivo
'      project-conventional global (set by the parent form's Alta/Editar handler).
'   4. SuministradorAlta_Cerrar(p_Error) — form cleanup
'      Returns JSON: {ok, payload:{ok}, error, logs}
'
' 7 fields per form (Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP, TramitadoraHPS).
' Special event signature: `Alta(p_ID As String)` includes the new ID as parameter.
' The helper returns "alta" in JSON; the form passes m_ObjSuministradorActivo.IDSuministrador
' to the event.
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
Private Const SUMINISTRADOR_ALTA_TITULO_ALTA As String = "ALTA DE SUMINISTRADOR"
Private Const SUMINISTRADOR_ALTA_TITULO_EDICION As String = "EDICIÓN DE SUMINISTRADOR"

' Supported modo values (callers pass these in p_Modo).
Private Const SUMINISTRADOR_ALTA_MODO_ALTA As String = "alta"
Private Const SUMINISTRADOR_ALTA_MODO_EDICION As String = "edicion"

' Fields that the Alta/Edición form manages (used by VerificarCambios + Registrar).
Private Const SUMINISTRADOR_ALTA_FIELD_NOMBRE As String = "Nombre"
Private Const SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO As String = "Nemotecnico"
Private Const SUMINISTRADOR_ALTA_FIELD_CIF As String = "CIF"
Private Const SUMINISTRADOR_ALTA_FIELD_DIRECCION As String = "Direccion"
Private Const SUMINISTRADOR_ALTA_FIELD_CIUDAD As String = "Ciudad"
Private Const SUMINISTRADOR_ALTA_FIELD_CP As String = "CP"
Private Const SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS As String = "TramitadoraHPS"

' Test fixture ID base. NOT used by the helper itself — tests use this to build
' stub entities with predictable IDs. Kept here so all rework code agrees.
Private Const SUMINISTRADOR_ALTA_TEST_ID_BASE As Long = 900750


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
' Returns a fresh empty Dictionary with the 7 Suministrador fields, all set to "".
' Used by Abrir_Inicializar in alta mode and as the default for ByRef p_Entidad
' before the helper populates it.
Private Function BuildEmptyEntidad() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(SUMINISTRADOR_ALTA_FIELD_NOMBRE) = ""
    d(SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO) = ""
    d(SUMINISTRADOR_ALTA_FIELD_CIF) = ""
    d(SUMINISTRADOR_ALTA_FIELD_DIRECCION) = ""
    d(SUMINISTRADOR_ALTA_FIELD_CIUDAD) = ""
    d(SUMINISTRADOR_ALTA_FIELD_CP) = ""
    d(SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS) = ""
    Set BuildEmptyEntidad = d
End Function

' --- NormalizarModo --------------------------------------------------------------
' Validates and normalizes the p_Modo parameter. Returns "" if invalid.
Private Function NormalizarModo(ByVal p_Modo As String) As String
    Dim m As String
    m = LCase$("" & p_Modo)
    If m = SUMINISTRADOR_ALTA_MODO_ALTA Then
        NormalizarModo = SUMINISTRADOR_ALTA_MODO_ALTA
    ElseIf m = SUMINISTRADOR_ALTA_MODO_EDICION Then
        NormalizarModo = SUMINISTRADOR_ALTA_MODO_EDICION
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

' --- SuministradorAlta_Abrir_Inicializar -----------------------------------------
' Pure-data init: given the form's mode (alta/edicion) and the entity ID, returns
' the caption to display + the precargados fields for the form to render.
'
' p_Modo: "alta" | "edicion"
' p_IDEntidad: ID of the Suministrador to load (used only when p_Modo="edicion")
' ByRef p_Entidad: helper populates this with a Dictionary {Nombre, Nemotecnico,
'                  CIF, Direccion, Ciudad, CP, TramitadoraHPS} of the loaded entity,
'                  OR an empty Dictionary for alta mode.
'
' Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}.
Public Function SuministradorAlta_Abrir_Inicializar( _
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
        p_Error = "SuministradorAlta_Abrir_Inicializar: p_Modo must be 'alta' or 'edicion', got '" & p_Modo & "'"
        Set p_Entidad = Nothing
        SuministradorAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim titulo As String
    If modo = SUMINISTRADOR_ALTA_MODO_ALTA Then
        titulo = SUMINISTRADOR_ALTA_TITULO_ALTA
        ' Initialize p_Entidad with empty values.
        Set p_Entidad = BuildEmptyEntidad()
        payload("titulo") = titulo
        payload("modo") = modo
        payload("hasEntidad") = False
        payload("id") = ""
        Set payload("entidad") = Nothing

        logs(0) = "Abrir_Inicializar: mode=alta, p_Entidad=empty"
        SuministradorAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Edicion mode — load the entity by ID.
    Dim idEntidad As String
    idEntidad = Trim$("" & p_IDEntidad)
    If Len(idEntidad) = 0 Then
        p_Error = "SuministradorAlta_Abrir_Inicializar: p_IDEntidad is required in edicion mode"
        Set p_Entidad = Nothing
        SuministradorAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Load the Suministrador class instance, then build the Dictionary shape for the form.
    Dim loaded As Suministrador
    Dim loadErr As String
    Set loaded = constructor.getSuministrador(p_IDSuministrador:=idEntidad, p_Error:=loadErr)
    If Len(loadErr) > 0 Then
        p_Error = "SuministradorAlta_Abrir_Inicializar: " & loadErr
        Set p_Entidad = Nothing
        SuministradorAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If loaded Is Nothing Then
        p_Error = "SuministradorAlta_Abrir_Inicializar: No se ha podido encontrar el Suministrador registrado (id=" & idEntidad & ")"
        Set p_Entidad = Nothing
        SuministradorAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Populate p_Entidad as Dictionary for the form to render.
    Set p_Entidad = BuildEmptyEntidad()
    If loaded.Nombre <> "" Then p_Entidad(SUMINISTRADOR_ALTA_FIELD_NOMBRE) = CStr(loaded.Nombre)
    If loaded.Nemotecnico <> "" Then p_Entidad(SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO) = CStr(loaded.Nemotecnico)
    If loaded.CIF <> "" Then p_Entidad(SUMINISTRADOR_ALTA_FIELD_CIF) = CStr(loaded.CIF)
    If loaded.Direccion <> "" Then p_Entidad(SUMINISTRADOR_ALTA_FIELD_DIRECCION) = CStr(loaded.Direccion)
    If loaded.Ciudad <> "" Then p_Entidad(SUMINISTRADOR_ALTA_FIELD_CIUDAD) = CStr(loaded.Ciudad)
    If loaded.CP <> "" Then p_Entidad(SUMINISTRADOR_ALTA_FIELD_CP) = CStr(loaded.CP)
    If loaded.TramitadoraHPS = "Sí" Or loaded.TramitadoraHPS = "No" Then
        p_Entidad(SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS) = CStr(loaded.TramitadoraHPS)
    End If

    ' Also update the project-conventional global so Registrar picks it up later.
    Set m_ObjSuministradorActivo = loaded

    titulo = SUMINISTRADOR_ALTA_TITULO_EDICION
    payload("titulo") = titulo
    payload("modo") = modo
    payload("hasEntidad") = True
    payload("id") = idEntidad
    Set payload("entidad") = p_Entidad

    logs(0) = "Abrir_Inicializar: mode=edicion, id=" & idEntidad
    logs(1) = "Abrir_Inicializar: precargados populated"

    SuministradorAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "SuministradorAlta_Abrir_Inicializar: " & Err.Description
    End If
    Set p_Entidad = Nothing
    SuministradorAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- SuministradorAlta_VerificarCambios -------------------------------------------
' Pure-data diff: compares current form values against the original snapshot.
' Returns whether anything changed and a per-field diff Dictionary.
'
' p_ValoresActuales: Dictionary {Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP,
'                            TramitadoraHPS} — current form values
' p_ValoresOriginales: Dictionary {Nombre, ...} — snapshot at open
'                     (or Nothing for alta mode)
'
' For alta mode (p_ValoresOriginales is Nothing), hayCambios is True iff any
' field is non-empty. For edicion mode, hayCambios is True iff any field differs.
'
' Returns JSON: {ok, payload:{hayCambios, diffs:{Nombre, ...}}, error, logs}.
Public Function SuministradorAlta_VerificarCambios( _
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
        p_Error = "SuministradorAlta_VerificarCambios: p_ValoresActuales is Nothing"
        SuministradorAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim diffs As Object
    Set diffs = CreateObject("Scripting.Dictionary")

    Dim currentNombre As String
    Dim currentNemo As String
    Dim currentCIF As String
    Dim currentDir As String
    Dim currentCiudad As String
    Dim currentCP As String
    Dim currentTramitadora As String

    currentNombre = ExtraerValor(p_ValoresActuales, SUMINISTRADOR_ALTA_FIELD_NOMBRE)
    currentNemo = ExtraerValor(p_ValoresActuales, SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO)
    currentCIF = ExtraerValor(p_ValoresActuales, SUMINISTRADOR_ALTA_FIELD_CIF)
    currentDir = ExtraerValor(p_ValoresActuales, SUMINISTRADOR_ALTA_FIELD_DIRECCION)
    currentCiudad = ExtraerValor(p_ValoresActuales, SUMINISTRADOR_ALTA_FIELD_CIUDAD)
    currentCP = ExtraerValor(p_ValoresActuales, SUMINISTRADOR_ALTA_FIELD_CP)
    currentTramitadora = ExtraerValor(p_ValoresActuales, SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS)

    Dim hayCambios As Boolean
    hayCambios = False

    If p_ValoresOriginales Is Nothing Then
        ' Alta mode — any non-empty value counts as a change.
        If Len(currentNombre) > 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_NOMBRE) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_NOMBRE) = False
        End If
        If Len(currentNemo) > 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO) = False
        End If
        If Len(currentCIF) > 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_CIF) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_CIF) = False
        End If
        If Len(currentDir) > 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_DIRECCION) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_DIRECCION) = False
        End If
        If Len(currentCiudad) > 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_CIUDAD) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_CIUDAD) = False
        End If
        If Len(currentCP) > 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_CP) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_CP) = False
        End If
        If Len(currentTramitadora) > 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS) = False
        End If
        logs(0) = "VerificarCambios: alta mode, hayCambios=" & CStr(hayCambios)
    Else
        ' Edicion mode — diff per field.
        Dim origNombre As String
        Dim origNemo As String
        Dim origCIF As String
        Dim origDir As String
        Dim origCiudad As String
        Dim origCP As String
        Dim origTramitadora As String

        origNombre = ExtraerValor(p_ValoresOriginales, SUMINISTRADOR_ALTA_FIELD_NOMBRE)
        origNemo = ExtraerValor(p_ValoresOriginales, SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO)
        origCIF = ExtraerValor(p_ValoresOriginales, SUMINISTRADOR_ALTA_FIELD_CIF)
        origDir = ExtraerValor(p_ValoresOriginales, SUMINISTRADOR_ALTA_FIELD_DIRECCION)
        origCiudad = ExtraerValor(p_ValoresOriginales, SUMINISTRADOR_ALTA_FIELD_CIUDAD)
        origCP = ExtraerValor(p_ValoresOriginales, SUMINISTRADOR_ALTA_FIELD_CP)
        origTramitadora = ExtraerValor(p_ValoresOriginales, SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS)

        If StrComp(currentNombre, origNombre, vbTextCompare) <> 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_NOMBRE) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_NOMBRE) = False
        End If

        If StrComp(currentNemo, origNemo, vbTextCompare) <> 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO) = False
        End If

        If StrComp(currentCIF, origCIF, vbTextCompare) <> 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_CIF) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_CIF) = False
        End If

        If StrComp(currentDir, origDir, vbTextCompare) <> 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_DIRECCION) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_DIRECCION) = False
        End If

        If StrComp(currentCiudad, origCiudad, vbTextCompare) <> 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_CIUDAD) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_CIUDAD) = False
        End If

        If StrComp(currentCP, origCP, vbTextCompare) <> 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_CP) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_CP) = False
        End If

        If StrComp(currentTramitadora, origTramitadora, vbTextCompare) <> 0 Then
            diffs(SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS) = True
            hayCambios = True
        Else
            diffs(SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS) = False
        End If

        logs(0) = "VerificarCambios: edicion mode, hayCambios=" & CStr(hayCambios)
    End If

    payload("hayCambios") = hayCambios
    Set payload("diffs") = diffs

    SuministradorAlta_VerificarCambios = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "SuministradorAlta_VerificarCambios: " & Err.Description
    End If
    SuministradorAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- SuministradorAlta_Registrar -----------------------------------------------
' Pure-data persist: given the final form values, persist via SuministradorOperaciones.
' The helper decides alta vs edicion based on the project-conventional global
' m_ObjSuministradorActivo:
'   - If m_ObjSuministradorActivo is Nothing  -> alta mode (insert new row)
'   - If m_ObjSuministradorActivo is set      -> edicion mode (update existing row)
'
' p_Valores: Dictionary {Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP,
'                       TramitadoraHPS} — final values to persist
'
' Returns JSON: {ok, payload:{ok, id, modo}, error, logs}.
Public Function SuministradorAlta_Registrar( _
    ByVal p_Valores As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Valores Is Nothing Then
        p_Error = "SuministradorAlta_Registrar: p_Valores is Nothing"
        SuministradorAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim currentNombre As String
    Dim currentNemo As String
    Dim currentCIF As String
    Dim currentDir As String
    Dim currentCiudad As String
    Dim currentCP As String
    Dim currentTramitadora As String

    currentNombre = ExtraerValor(p_Valores, SUMINISTRADOR_ALTA_FIELD_NOMBRE)
    currentNemo = ExtraerValor(p_Valores, SUMINISTRADOR_ALTA_FIELD_NEMOTECNICO)
    currentCIF = ExtraerValor(p_Valores, SUMINISTRADOR_ALTA_FIELD_CIF)
    currentDir = ExtraerValor(p_Valores, SUMINISTRADOR_ALTA_FIELD_DIRECCION)
    currentCiudad = ExtraerValor(p_Valores, SUMINISTRADOR_ALTA_FIELD_CIUDAD)
    currentCP = ExtraerValor(p_Valores, SUMINISTRADOR_ALTA_FIELD_CP)
    currentTramitadora = ExtraerValor(p_Valores, SUMINISTRADOR_ALTA_FIELD_TRAMITADORAHPS)

    ' Determine mode and capture id.
    Dim modo As String
    Dim idEntidad As String
    If m_ObjSuministradorActivo Is Nothing Then
        modo = SUMINISTRADOR_ALTA_MODO_ALTA
        idEntidad = ""
    Else
        modo = SUMINISTRADOR_ALTA_MODO_EDICION
        If m_ObjSuministradorActivo.IDSuministrador <> "" Then
            idEntidad = CStr(m_ObjSuministradorActivo.IDSuministrador)
        Else
            idEntidad = ""
        End If
    End If

    logs(0) = "Registrar: mode=" & modo & ", id=" & idEntidad

    ' Build the Suministrador class instance from p_Valores (and existing m_ObjSuministradorActivo).
    Dim sm As Suministrador
    If m_ObjSuministradorActivo Is Nothing Then
        Set sm = New Suministrador
    Else
        Set sm = m_ObjSuministradorActivo
    End If

    sm.Nombre = currentNombre
    sm.Nemotecnico = currentNemo
    sm.CIF = currentCIF
    sm.Direccion = currentDir
    sm.Ciudad = currentCiudad
    sm.CP = currentCP
    sm.TramitadoraHPS = currentTramitadora

    ' Update the project-conventional global so subsequent operations see the entity.
    Set m_ObjSuministradorActivo = sm

    ' Delegate to SuministradorOperaciones for the actual DB write.
    Dim m_SuministradorOP As New SuministradorOperaciones
    Dim m_AlInicio As Suministrador
    If modo = SUMINISTRADOR_ALTA_MODO_EDICION Then
        Dim loadErr As String
        Set m_AlInicio = constructor.getSuministrador(p_IDSuministrador:=idEntidad, p_Error:=loadErr)
    Else
        Set m_AlInicio = Nothing
    End If

    With m_SuministradorOP
        Set .Suministrador = sm
        Dim registrarErr As String
        .Registrar m_AlInicio, registrarErr
        If Len(registrarErr) > 0 Then
            p_Error = "SuministradorAlta_Registrar: " & registrarErr
            SuministradorAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
        End If
    End With

    ' Capture the resulting ID after Registrar has stamped it.
    Dim resultingId As String
    If sm.IDSuministrador <> "" Then
        resultingId = CStr(sm.IDSuministrador)
    Else
        resultingId = idEntidad
    End If

    payload("ok") = True
    payload("id") = resultingId
    payload("modo") = modo

    logs(1) = "Registrar: persisted, resultingId=" & resultingId

    SuministradorAlta_Registrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "SuministradorAlta_Registrar: " & Err.Description
    End If
    SuministradorAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- SuministradorAlta_Cerrar --------------------------------------------------
' Form cleanup. Releases the project-conventional m_ObjSuministradorActivo global
' so the next form open starts clean. Returns ok=true on success.
'
' Returns JSON: {ok, payload:{ok}, error, logs}.
Public Function SuministradorAlta_Cerrar( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    ' Release the project-conventional global.
    Set m_ObjSuministradorActivo = Nothing

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ok") = True

    logs(0) = "Cerrar: m_ObjSuministradorActivo cleared"

    SuministradorAlta_Cerrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "SuministradorAlta_Cerrar: " & Err.Description
    End If
    SuministradorAlta_Cerrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
