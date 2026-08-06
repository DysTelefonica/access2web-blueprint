Attribute VB_Name = "modExpedienteCambioTipoHelper"
Option Compare Database
Option Explicit

' modExpedienteCambioTipoHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data helpers for Form_FormExpedienteCambioTipo
' (cambio de tipo de expediente: 5-branch dispatch).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix:
' ExpedienteCambioTipo_*), and rule #9 (no `ByRef p_Form`).
'
' Shape contract:
'   The form uses `m_ObjExpedienteActivo` (Expediente, not DTO) globally.
'   Helpers accept p_Expediente as Object so tests can pass a Dictionary
'   stub with the relevant properties.
'
' Anti-pattern removed (was in pre-PR-8b baseline):
'   - 5-branch dispatch in ComandoRegistrar_Click with inline MsgBox.
'   - Inline pad validation logic in ComandoRegistrar.
'   - Tipo_AfterUpdate toggling Me.IDExpedientePadre.Locked/Enabled inline.
'
' New design (4 helpers, all pure-data, no Form refs):
'   1. ExpedienteCambioTipo_EstablecerDatos(p_Expediente, p_Error)
'      -> {tipoActual, opciones: ["Convertir a Lote", ...]}
'   2. ExpedienteCambioTipo_Tipo_AfterUpdate(p_TipoElegido, p_Error)
'      -> {idPadreLocked, idPadreEnabled}
'   3. ExpedienteCambioTipo_ComandoRegistrar(p_Expediente, p_TipoElegido,
'                                              p_IDExpedientePadre, p_Error)
'      -> {registered, validation, flagsApplied: {EsAM, EsLote, EsExpediente,
'                                                   EsBasado, IDExpedientePadre}}
'   4. ExpedienteCambioTipo_ValidarPadre(p_TipoElegido, p_IDExpedientePadre,
'                                          p_EsPadreAM As Boolean,
'                                          p_EsPadreLote As Boolean, p_Error)
'      -> {valido, error}
'
' UI orchestration that stays in the form (rule #1):
'   - cmdSalir_Click — DoCmd.Close
'   - ComandoAyuda_Click — AbrirAyuda
'   - ComandoBuscar_Click — DoCmd.OpenForm "FormExpedientesParaCambioTipo"
'   - Tipo_AfterUpdate — applies Locked/Enabled (helper returns the flags)
'   - m_Form_Seleccionado — UI pass-through (Me.IDExpedientePadre = m_expediente.IDExpediente)
'   - Form_Open — Ajustar Me + EstablecerDatos
'   - EstablecerDatos public wrapper — populates cmb.AddItem + TIPOACTUAL
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter. On failure: set p_Error,
' return fail JSON. On success: return canonical JSON envelope.
'
' vba-access §10.1 declaration ordering: all Private Const / Private Function at top,
' Public Function atoms after. No mid-module consts.
'
' vba-access §1.6.1: split guards (no IIf/And short-circuit on the same object).

' === Module-level constants (all at top per vba-access §10.1) =====================

' The 5 tipoElegido values the form dispatches on.
Private Const EXPEDIENTECAMBIOTIPO_TIPO_AM As String = "Convertir a Acuerdo Marco"
Private Const EXPEDIENTECAMBIOTIPO_TIPO_LOTE As String = "Convertir a Lote"
Private Const EXPEDIENTECAMBIOTIPO_TIPO_BASADO As String = "Convertir a Basado"
Private Const EXPEDIENTECAMBIOTIPO_TIPO_EXP_INDEP As String = "Convertir a Expediente Independiente"
Private Const EXPEDIENTECAMBIOTIPO_TIPO_LOTE_AM As String = "Convertir a Lote de Acuerdo Marco"

' Flag values used to write back into the Expediente entity.
Private Const EXPEDIENTECAMBIOTIPO_FLAG_SI As String = "Sí"
Private Const EXPEDIENTECAMBIOTIPO_FLAG_NO As String = "No"

' Test fixture ID base for entidades; not used by helpers, just for traceability.
Private Const EXPEDIENTECAMBIOTIPO_TEST_ID_BASE As Long = 900820


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

' --- ExpedienteCambioTipo_GetValueSafely -----------------------------------------
' Returns p_Exp(p_Key) if p_Exp is a Dictionary-like object with that key.
' Returns the provided default if p_Exp is Nothing, missing key, or access raised.
Private Function ExpedienteCambioTipo_GetValueSafely( _
    ByVal p_Exp As Object, _
    ByVal p_Key As String, _
    ByVal p_Default As Variant _
) As Variant
    On Error GoTo EH
    If p_Exp Is Nothing Then
        ExpedienteCambioTipo_GetValueSafely = p_Default
        Exit Function
    End If
    If Not p_Exp.Exists(p_Key) Then
        ExpedienteCambioTipo_GetValueSafely = p_Default
        Exit Function
    End If
    ExpedienteCambioTipo_GetValueSafely = p_Exp(p_Key)
    Exit Function
EH:
    ExpedienteCambioTipo_GetValueSafely = p_Default
End Function

' --- ExpedienteCambioTipo_DetectarTipoActual -------------------------------------
' Returns one of "AM", "LOTE", "BASADO", "EXP_INDEP", or "" if no flag is set.
' Real Expediente entities have .EsAM/.EsLote/.EsExpediente/.EsBasado as String
' properties with value "Sí" or "No". We honor the first one set to "Sí".
Private Function ExpedienteCambioTipo_DetectarTipoActual(ByVal p_Exp As Object) As String
    Dim esAM As String
    Dim esLote As String
    Dim esExp As String
    Dim esBas As String

    ' Dictionary stub path
    If Not p_Exp Is Nothing Then
        If TypeName(p_Exp) = "Dictionary" Then
            esAM = CStr(ExpedienteCambioTipo_GetValueSafely(p_Exp, "EsAM", ""))
            esLote = CStr(ExpedienteCambioTipo_GetValueSafely(p_Exp, "EsLote", ""))
            esExp = CStr(ExpedienteCambioTipo_GetValueSafely(p_Exp, "EsExpediente", ""))
            esBas = CStr(ExpedienteCambioTipo_GetValueSafely(p_Exp, "EsBasado", ""))
        Else
            ' Real entity path — best-effort, no short-circuit.
            On Error Resume Next
            esAM = p_Exp.EsAM
            esLote = p_Exp.EsLote
            esExp = p_Exp.EsExpediente
            esBas = p_Exp.EsBasado
            On Error GoTo 0
        End If
    End If

    If esAM = EXPEDIENTECAMBIOTIPO_FLAG_SI Then
        ExpedienteCambioTipo_DetectarTipoActual = "AM"
    ElseIf esLote = EXPEDIENTECAMBIOTIPO_FLAG_SI Then
        ExpedienteCambioTipo_DetectarTipoActual = "LOTE"
    ElseIf esExp = EXPEDIENTECAMBIOTIPO_FLAG_SI Then
        ExpedienteCambioTipo_DetectarTipoActual = "EXP_INDEP"
    ElseIf esBas = EXPEDIENTECAMBIOTIPO_FLAG_SI Then
        ExpedienteCambioTipo_DetectarTipoActual = "BASADO"
    Else
        ExpedienteCambioTipo_DetectarTipoActual = ""
    End If
End Function

' --- ExpedienteCambioTipo_BuildOpcionesList --------------------------------------
' Builds the array of tipo options based on the current tipo.
' Mirrors the legacy EstablecerDatos logic verbatim.
Private Function ExpedienteCambioTipo_BuildOpcionesList(ByVal p_TipoActual As String) As Object
    Dim opts As Object
    Set opts = CreateObject("Scripting.Dictionary")

    Select Case p_TipoActual
        Case "AM"
            opts("0") = EXPEDIENTECAMBIOTIPO_TIPO_LOTE
            opts("1") = EXPEDIENTECAMBIOTIPO_TIPO_BASADO
            opts("2") = EXPEDIENTECAMBIOTIPO_TIPO_EXP_INDEP
        Case "LOTE"
            opts("0") = EXPEDIENTECAMBIOTIPO_TIPO_LOTE_AM
            opts("1") = EXPEDIENTECAMBIOTIPO_TIPO_AM
            opts("2") = EXPEDIENTECAMBIOTIPO_TIPO_BASADO
            opts("3") = EXPEDIENTECAMBIOTIPO_TIPO_EXP_INDEP
        Case "BASADO"
            opts("0") = EXPEDIENTECAMBIOTIPO_TIPO_AM
            opts("1") = EXPEDIENTECAMBIOTIPO_TIPO_LOTE
            opts("2") = EXPEDIENTECAMBIOTIPO_TIPO_EXP_INDEP
            opts("3") = EXPEDIENTECAMBIOTIPO_TIPO_BASADO
        Case "EXP_INDEP"
            opts("0") = EXPEDIENTECAMBIOTIPO_TIPO_AM
            opts("1") = EXPEDIENTECAMBIOTIPO_TIPO_LOTE
            opts("2") = EXPEDIENTECAMBIOTIPO_TIPO_BASADO
        Case Else
            ' unknown — empty list
    End Select

    Set ExpedienteCambioTipo_BuildOpcionesList = opts
End Function


' === Public API =================================================================

' --- ExpedienteCambioTipo_EstablecerDatos ----------------------------------------
' Pure-data init: detects tipo actual, returns the list of allowed tipoElegido
' options. The form does the AddItem on the cmb.
'
' Returns JSON: {ok, payload:{tipoActual, opciones:{key:label}}, error, logs}.
Public Function ExpedienteCambioTipo_EstablecerDatos( _
    ByVal p_Expediente As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    If p_Expediente Is Nothing Then
        p_Error = "ExpedienteCambioTipo_EstablecerDatos: p_Expediente is Nothing"
        ExpedienteCambioTipo_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim tipoActual As String
    tipoActual = ExpedienteCambioTipo_DetectarTipoActual(p_Expediente)

    Dim opciones As Object
    Set opciones = ExpedienteCambioTipo_BuildOpcionesList(tipoActual)

    payload("tipoActual") = tipoActual
    Set payload("opciones") = opciones
    logs(0) = "EstablecerDatos: tipoActual=" & tipoActual & " count=" & CStr(opciones.Count)

    ExpedienteCambioTipo_EstablecerDatos = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteCambioTipo_EstablecerDatos: " & Err.Description
    End If
    ExpedienteCambioTipo_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteCambioTipo_Tipo_AfterUpdate ---------------------------------------
' Pure-data: given the tipoElegido selected, returns the Locked/Enabled state
' for the IDExpedientePadre combo. Form applies these.
'
' Returns JSON: {ok, payload:{idPadreLocked, idPadreEnabled}, error, logs}.
Public Function ExpedienteCambioTipo_Tipo_AfterUpdate( _
    ByVal p_TipoElegido As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idPadreLocked As Boolean
    Dim idPadreEnabled As Boolean

    Dim tipo As String
    tipo = "" & p_TipoElegido

    ' Default: locked & disabled (no input needed).
    idPadreLocked = True
    idPadreEnabled = False

    Select Case tipo
        Case EXPEDIENTECAMBIOTIPO_TIPO_AM, _
             EXPEDIENTECAMBIOTIPO_TIPO_EXP_INDEP
            ' no se puede meter idpadre
            idPadreLocked = True
            idPadreEnabled = False
        Case EXPEDIENTECAMBIOTIPO_TIPO_LOTE, _
             EXPEDIENTECAMBIOTIPO_TIPO_BASADO, _
             EXPEDIENTECAMBIOTIPO_TIPO_LOTE_AM
            ' editable
            idPadreLocked = False
            idPadreEnabled = True
        Case ""
            ' empty selection — keep default locked
        Case Else
            ' unknown tipo — keep default locked
            logs(1) = "Tipo_AfterUpdate: unknown tipo '" & tipo & "'"
    End Select

    payload("idPadreLocked") = idPadreLocked
    payload("idPadreEnabled") = idPadreEnabled
    logs(0) = "Tipo_AfterUpdate: tipo='" & tipo & "' locked=" & CStr(idPadreLocked) & _
              " enabled=" & CStr(idPadreEnabled)

    ExpedienteCambioTipo_Tipo_AfterUpdate = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteCambioTipo_Tipo_AfterUpdate: " & Err.Description
    End If
    ExpedienteCambioTipo_Tipo_AfterUpdate = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteCambioTipo_ValidarPadre -------------------------------------------
' Pure-data validation of the IDExpedientePadre for the chosen tipo. Caller
' passes the loaded padre's flags (p_EsPadreAM, p_EsPadreLote) so the helper
' stays a pure function (no DAO lookup). The form does the constructor.getExpediente
' call before invoking this.
'
' Returns JSON: {ok, payload:{valido, error}, error, logs}.
Public Function ExpedienteCambioTipo_ValidarPadre( _
    ByVal p_TipoElegido As String, _
    ByVal p_IDExpedientePadre As String, _
    ByVal p_EsPadreAM As Boolean, _
    ByVal p_EsPadreLote As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim valido As Boolean
    valido = False

    Dim idPadre As String
    idPadre = Trim$("" & p_IDExpedientePadre)

    Dim tipo As String
    tipo = "" & p_TipoElegido

    Select Case tipo
        Case EXPEDIENTECAMBIOTIPO_TIPO_AM
            ' AM does not require a padre; valid if padre empty.
            If Len(idPadre) = 0 Then
                valido = True
            Else
                p_Error = "Para Convertir en Acuerdo Marco no hemos de indicar un IDExpediente padre"
            End If
        Case EXPEDIENTECAMBIOTIPO_TIPO_LOTE
            ' Lote requiere padre AM.
            If Len(idPadre) = 0 Then
                p_Error = "Para Convertir en Lote, hemos de rellenar el ID del Acuerdo Marco"
            ElseIf Not p_EsPadreAM Then
                p_Error = "Un Lote puede tener un padre sólo de tipo Acuerdo Marco"
            Else
                valido = True
            End If
        Case EXPEDIENTECAMBIOTIPO_TIPO_BASADO
            ' Basado requiere padre AM o Lote.
            If Len(idPadre) = 0 Then
                p_Error = "Para Convertir en Basado, hemos de indicar un IDExpediente padre de tipo Acuerdo Marco o Lote"
            ElseIf (Not p_EsPadreAM) And (Not p_EsPadreLote) Then
                p_Error = "Para Convertir en Basado, hemos de indicar un IDExpediente padre de tipo Acuerdo Marco o Lote"
            Else
                valido = True
            End If
        Case EXPEDIENTECAMBIOTIPO_TIPO_EXP_INDEP
            ' EXPIndep must NOT have padre.
            If Len(idPadre) > 0 Then
                p_Error = "Para Convertir en Expediente Independiente, no hemos de indicar un IDExpediente padre"
            Else
                valido = True
            End If
        Case EXPEDIENTECAMBIOTIPO_TIPO_LOTE_AM
            ' Lote AM requiere padre (presumed AM).
            If Len(idPadre) = 0 Then
                p_Error = "Para Convertir a Lote de Acuerdo Marco. Hemos de rellenar el ID del Acuerdo Marco"
            Else
                valido = True
            End If
        Case Else
            p_Error = "Seleccione el cambio de tipo"
    End Select

    payload("valido") = valido
    payload("error") = p_Error
    logs(0) = "ValidarPadre: tipo='" & tipo & "' valido=" & CStr(valido) & _
              " idPadre='" & idPadre & "'"

    If valido Then
        ExpedienteCambioTipo_ValidarPadre = BuildJsonPayload(True, payload, "", logs)
    Else
        ExpedienteCambioTipo_ValidarPadre = BuildJsonPayload(False, payload, p_Error, logs)
    End If
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteCambioTipo_ValidarPadre: " & Err.Description
    End If
    ExpedienteCambioTipo_ValidarPadre = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteCambioTipo_ComandoRegistrar ---------------------------------------
' Pure-data 5-branch dispatch: validates the padre, computes the flag set,
' and returns the flags the form should apply to m_ObjExpedienteActivo.
' The form does the actual DAO call (ExpedienteOperaciones.RegistrarCambioTipo)
' and the RaiseEvent — those are NOT pure data.
'
' p_Expediente: Object (real Expediente or Dictionary stub with EsAM/EsLote/...)
'
' Returns JSON: {ok, payload:{registered, validation, flagsApplied:{...}}, error, logs}.
Public Function ExpedienteCambioTipo_ComandoRegistrar( _
    ByVal p_Expediente As Object, _
    ByVal p_TipoElegido As String, _
    ByVal p_IDExpedientePadre As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim tipo As String
    tipo = Trim$("" & p_TipoElegido)

    If Len(tipo) = 0 Then
        p_Error = "Seleccione el cambio de tipo"
        ExpedienteCambioTipo_ComandoRegistrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If p_Expediente Is Nothing Then
        p_Error = "ExpedienteCambioTipo_ComandoRegistrar: p_Expediente is Nothing"
        ExpedienteCambioTipo_ComandoRegistrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idPadre As String
    idPadre = Trim$("" & p_IDExpedientePadre)

    ' Compute the flag set per branch. These mirror the legacy .cls inline logic.
    Dim flags As Object
    Set flags = CreateObject("Scripting.Dictionary")

    Dim validation As String
    validation = ""

    Select Case tipo
        Case EXPEDIENTECAMBIOTIPO_TIPO_AM
            flags("IDExpedientePadre") = ""
            flags("EsAM") = EXPEDIENTECAMBIOTIPO_FLAG_SI
            flags("EsLote") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsExpediente") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsBasado") = EXPEDIENTECAMBIOTIPO_FLAG_NO
        Case EXPEDIENTECAMBIOTIPO_TIPO_LOTE
            ' Validation requires loading padre — done by the form before calling.
            ' This helper assumes the form has done the check and just computes flags
            ' when validation passed. The form must guard.
            flags("IDExpedientePadre") = idPadre
            flags("EsAM") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsLote") = EXPEDIENTECAMBIOTIPO_FLAG_SI
            flags("EsExpediente") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsBasado") = EXPEDIENTECAMBIOTIPO_FLAG_NO
        Case EXPEDIENTECAMBIOTIPO_TIPO_BASADO
            flags("IDExpedientePadre") = idPadre
            flags("EsAM") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsLote") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsExpediente") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsBasado") = EXPEDIENTECAMBIOTIPO_FLAG_SI
        Case EXPEDIENTECAMBIOTIPO_TIPO_EXP_INDEP
            flags("IDExpedientePadre") = ""
            flags("EsAM") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsLote") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsExpediente") = EXPEDIENTECAMBIOTIPO_FLAG_SI
            flags("EsBasado") = EXPEDIENTECAMBIOTIPO_FLAG_NO
        Case EXPEDIENTECAMBIOTIPO_TIPO_LOTE_AM
            flags("IDExpedientePadre") = idPadre
            flags("EsAM") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsLote") = EXPEDIENTECAMBIOTIPO_FLAG_SI
            flags("EsExpediente") = EXPEDIENTECAMBIOTIPO_FLAG_NO
            flags("EsBasado") = EXPEDIENTECAMBIOTIPO_FLAG_NO
        Case Else
            p_Error = "Tipo de cambio no reconocido: '" & tipo & "'"
            ExpedienteCambioTipo_ComandoRegistrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
    End Select

    payload("registered") = True
    payload("validation") = validation
    Set payload("flagsApplied") = flags
    logs(0) = "ComandoRegistrar: tipo='" & tipo & "' idPadre='" & idPadre & "' flagsApplied"
    logs(1) = "  EsAM=" & CStr(flags("EsAM")) & " EsLote=" & CStr(flags("EsLote")) & _
              " EsExpediente=" & CStr(flags("EsExpediente")) & _
              " EsBasado=" & CStr(flags("EsBasado")) & _
              " IDExpedientePadre='" & CStr(flags("IDExpedientePadre")) & "'"

    ExpedienteCambioTipo_ComandoRegistrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteCambioTipo_ComandoRegistrar: " & Err.Description
    End If
    ExpedienteCambioTipo_ComandoRegistrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function