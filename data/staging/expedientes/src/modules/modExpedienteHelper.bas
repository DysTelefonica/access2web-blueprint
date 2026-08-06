Attribute VB_Name = "modExpedienteHelper"
Option Compare Database
Option Explicit

' modExpedienteHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data helpers for Form_FormExpediente (main expediente view/edit,
' tab manager with save-delegation).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix: Expediente_*),
' and rule #9 (no `ByRef p_Form`).
'
' Shape contract:
'   The form uses `m_ObjExpedienteDTOActivo` (ExpedienteDTO) globally. Helpers
'   accept p_Expediente or p_DTO as Object (Dictionary stub) so tests can pass
'   pure stubs.
'
' Anti-pattern removed (was in pre-PR-8b baseline):
'   - Inline Caption/AllowEdits decision in Form_Load (5 branches by OpenArgs/admin).
'   - Inline label-setting + button-visibility in EstablecerDatos.
'   - TabX_Click events had no decision logic — pure orchestration.
'
' New design (4 helpers, all pure-data, no Form refs):
'   1. Expediente_Form_Load_Init(p_OpenArgs, p_EsAdmin, p_BackendActivo,
'                                  p_HasExpediente, p_Error)
'      -> {caption, allowEdits, comandoRegistrarEnabled, insideHeight, insideWidth}
'   2. Expediente_EstablecerDatos(p_DTO, p_Error)
'      -> {lblTitulo, lblUltimaModificacion, lblUltimaModificacionVisible,
'           comandoActualizarCompletoVisible, esNuevo}
'   3. Expediente_DebeGuardarPestana(p_NombrePestana, p_AbiertoParaEditar,
'                                     p_AllowEdits, p_HayCambios, p_Error)
'      -> {debeGuardar, error}
'   4. Expediente_Tab_Seleccionar_Guardar(p_AbiertoParaEditar, p_AllowEdits,
'                                            p_Error)
'      -> {guardarAhora:Boolean, error}
'      (Wraps the GuardarPestanaActivaAntesDeNavegar logic.)
'
' UI orchestration that stays in the form (rule #1):
'   - 7 tabX_Click handlers — call the helper, do popup on error
'   - ComandoActualizarCompleto_Click — DAO call
'   - ComandoRegistrar_Click — popup + save
'   - ComandoAyuda_Click — AbrirAyuda
'   - cmdSalir_Click — save + close
'   - Form_Load — backend init + DAO lookup + form wiring
'   - Public method EstablecerDatos (signature preserved)
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.
'
' vba-access §10.1 declaration ordering: all Private Const / Private Function at top,
' Public Function atoms after. No mid-module consts.
'
' vba-access §1.6.1: split guards (no IIf/And short-circuit on the same object).

' === Module-level constants (all at top per vba-access §10.1) =====================

' InsideHeight/InsideWidth defaults set in Form_Load.
Private Const EXPEDIENTE_INSIDE_HEIGHT As Long = 9060
Private Const EXPEDIENTE_INSIDE_WIDTH As Long = 13980

' Sentinel for OpenArgs that disables editing.
Private Const EXPEDIENTE_OPENARGS_SOLOLECTURA As String = "Sololectura"

' Yes/No flag values used by the project (matches EnumSiNo.Sí/No).
Private Const EXPEDIENTE_FLAG_SI As Long = 1

' Test fixture ID base for entidades; not used by helpers, just for traceability.
Private Const EXPEDIENTE_TEST_ID_BASE As Long = 900840


' === Local helpers (all at top per vba-access §10.1) =============================

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

' --- Expediente_GetValueSafely ---------------------------------------------------
Private Function Expediente_GetValueSafely( _
    ByVal p_Exp As Object, _
    ByVal p_Key As String, _
    ByVal p_Default As Variant _
) As Variant
    On Error GoTo EH
    If p_Exp Is Nothing Then
        Expediente_GetValueSafely = p_Default
        Exit Function
    End If
    If Not p_Exp.Exists(p_Key) Then
        Expediente_GetValueSafely = p_Default
        Exit Function
    End If
    Expediente_GetValueSafely = p_Exp(p_Key)
    Exit Function
EH:
    Expediente_GetValueSafely = p_Default
End Function

' --- Expediente_EsPestanaGeneralOFechas -----------------------------------------
' Helper for Tab_Seleccionar_Guardar: returns True if p_NombrePestana is one of
' the General/Fechas tab names. Empty -> False.
Private Function Expediente_EsPestanaGeneralOFechas(ByVal p_NombrePestana As String) As Boolean
    Dim t As String
    t = "" & p_NombrePestana
    Expediente_EsPestanaGeneralOFechas = (t = "FormExpedienteGeneral" Or t = "FormExpedienteFechas")
End Function


' === Public API =================================================================

' --- Expediente_Form_Load_Init ---------------------------------------------------
' Pure-data init: given the OpenArgs, admin flag, backend activo, and whether
' there's a loaded expediente, returns the UI decisions (caption, allowEdits,
' comandoRegistrarEnabled, inside dimensions).
'
' Returns JSON: {ok, payload:{caption, allowEdits, comandoRegistrarEnabled,
'                              insideHeight, insideWidth, abiertoParaEditar}, error, logs}.
Public Function Expediente_Form_Load_Init( _
    ByVal p_OpenArgs As String, _
    ByVal p_EsAdmin As Boolean, _
    ByVal p_BackendActivo As String, _
    ByVal p_HasExpediente As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim openArgs As String
    openArgs = Trim$("" & p_OpenArgs)

    Dim abiertoParaEditar As Boolean
    If openArgs = EXPEDIENTE_OPENARGS_SOLOLECTURA Then
        abiertoParaEditar = False
    Else
        abiertoParaEditar = True
    End If

    Dim allowEdits As Boolean
    Dim comandoRegistrarEnabled As Boolean

    If abiertoParaEditar Then
        If p_EsAdmin Then
            allowEdits = True
            comandoRegistrarEnabled = True
        Else
            allowEdits = False
            comandoRegistrarEnabled = False
        End If
    Else
        allowEdits = False
        comandoRegistrarEnabled = False
    End If

    Dim backend As String
    backend = Trim$("" & p_BackendActivo)

    Dim m_TituloFormulario As String
    m_TituloFormulario = "" ' populated by caller via m_TituloFormulario global

    payload("abiertoParaEditar") = abiertoParaEditar
    payload("allowEdits") = allowEdits
    payload("comandoRegistrarEnabled") = comandoRegistrarEnabled
    payload("insideHeight") = EXPEDIENTE_INSIDE_HEIGHT
    payload("insideWidth") = EXPEDIENTE_INSIDE_WIDTH
    payload("backendActivo") = backend
    payload("hasExpediente") = p_HasExpediente
    payload("caption") = m_TituloFormulario & " [" & backend & "]"

    logs(0) = "Form_Load_Init: openArgs='" & openArgs & "' esAdmin=" & CStr(p_EsAdmin) & _
              " abiertoParaEditar=" & CStr(abiertoParaEditar)

    Expediente_Form_Load_Init = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Expediente_Form_Load_Init: " & Err.Description
    End If
    Expediente_Form_Load_Init = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Expediente_EstablecerDatos -------------------------------------------------
' Pure-data decision for the EstablecerDatos handler: returns the labels and
' button visibility flags based on the Expediente state.
'
' Returns JSON: {ok, payload:{lblTitulo, lblUltimaModificacion,
'                              lblUltimaModificacionVisible,
'                              comandoActualizarCompletoVisible}, error, logs}.
Public Function Expediente_EstablecerDatos( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    If p_DTO Is Nothing Then
        p_Error = "Expediente_EstablecerDatos: p_DTO is Nothing"
        Expediente_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = Expediente_GetValueSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "Expediente_EstablecerDatos: p_DTO.Expediente is Nothing"
        Expediente_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idExp As String
    Dim tituloFormulario As String
    Dim ultimaModificacionTexto As String

    idExp = "" & Expediente_GetValueSafely(expObj, "IDExpediente", "")
    tituloFormulario = "" & Expediente_GetValueSafely(expObj, "TituloFormulario", "")
    ultimaModificacionTexto = "" & Expediente_GetValueSafely(expObj, "UltimaModificacionTexto", "")

    Dim esNuevo As Boolean
    esNuevo = (Len(idExp) = 0)

    Dim lblUltimaModificacionVisible As Boolean
    If esNuevo Then
        lblUltimaModificacionVisible = False
    Else
        lblUltimaModificacionVisible = True
    End If

    Dim comandoActualizarCompletoVisible As Boolean
    comandoActualizarCompletoVisible = Not esNuevo

    payload("lblTitulo") = tituloFormulario
    payload("lblUltimaModificacion") = ultimaModificacionTexto
    payload("lblUltimaModificacionVisible") = lblUltimaModificacionVisible
    payload("comandoActualizarCompletoVisible") = comandoActualizarCompletoVisible
    payload("esNuevo") = esNuevo

    logs(0) = "EstablecerDatos: idExp='" & idExp & "' esNuevo=" & CStr(esNuevo)

    Expediente_EstablecerDatos = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Expediente_EstablecerDatos: " & Err.Description
    End If
    Expediente_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Expediente_DebeGuardarPestana -----------------------------------------------
' Pure-data decision: given the tab name, open-for-edit state, and dirty state,
' returns whether the tab change should trigger a save.
'
' Returns JSON: {ok, payload:{debeGuardar, error}, error, logs}.
Public Function Expediente_DebeGuardarPestana( _
    ByVal p_NombrePestana As String, _
    ByVal p_AbiertoParaEditar As Boolean, _
    ByVal p_AllowEdits As Boolean, _
    ByVal p_HayCambios As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim debeGuardar As Boolean
    debeGuardar = False

    ' Only General/Fechas tabs need save-before-navigate.
    If Not Expediente_EsPestanaGeneralOFechas(p_NombrePestana) Then
        payload("debeGuardar") = False
        Expediente_DebeGuardarPestana = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Must be open for edit AND AllowEdits AND have changes.
    If (Not p_AbiertoParaEditar) Or (Not p_AllowEdits) Then
        payload("debeGuardar") = False
        Expediente_DebeGuardarPestana = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    If Not p_HayCambios Then
        payload("debeGuardar") = False
        Expediente_DebeGuardarPestana = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    debeGuardar = True
    payload("debeGuardar") = True
    logs(0) = "DebeGuardarPestana: pestana='" & p_NombrePestana & "' debeGuardar=true"

    Expediente_DebeGuardarPestana = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Expediente_DebeGuardarPestana: " & Err.Description
    End If
    Expediente_DebeGuardarPestana = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Expediente_Tab_Seleccionar_Guardar ------------------------------------------
' Pure-data orchestrator: returns whether GuardarPestanaActivaAntesDeNavegar
' should attempt to save. Form does the actual call (which goes through
' getFormPestanaActiva + Helper_ExpedienteEdicion.GuardarPestanaActiva).
'
' Returns JSON: {ok, payload:{intentarGuardar, motivo}, error, logs}.
Public Function Expediente_Tab_Seleccionar_Guardar( _
    ByVal p_AbiertoParaEditar As Boolean, _
    ByVal p_AllowEdits As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim intentarGuardar As Boolean
    If p_AbiertoParaEditar And p_AllowEdits Then
        intentarGuardar = True
    Else
        intentarGuardar = False
    End If

    Dim motivo As String
    If intentarGuardar Then
        motivo = "editable + allowEdits -> attempt save"
    Else
        motivo = "not editable / allowEdits off -> skip save"
    End If

    payload("intentarGuardar") = intentarGuardar
    payload("motivo") = motivo
    logs(0) = "Tab_Seleccionar_Guardar: intentarGuardar=" & CStr(intentarGuardar)

    Expediente_Tab_Seleccionar_Guardar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Expediente_Tab_Seleccionar_Guardar: " & Err.Description
    End If
    Expediente_Tab_Seleccionar_Guardar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function