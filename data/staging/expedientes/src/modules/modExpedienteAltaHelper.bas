Attribute VB_Name = "modExpedienteAltaHelper"
Option Compare Database
Option Explicit

' modExpedienteAltaHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data helpers for Form_FormExpedienteAlta
' (alta de nuevo expediente con OpenArgs "IDExpedientePadre|TipoEnum").
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix:
' ExpedienteAlta_*), and rule #9 (no `ByRef p_Form`).
'
' Anti-pattern removed (was in pre-PR-8b baseline):
'   - Inline visibility/ordinal decision tree in EstablecerDatos (4-branch logic).
'   - Inline TIpo->{EsAM,EsLote,EsExpediente} translation in EstablecerDatosConTipo.
'   - Inline Ambito->HPSAplica setting in Ambito_AfterUpdate.
'   - Inline ordinal conflict detection logic in Ordinal_AfterUpdate.
'
' New design (5 helpers, all pure-data, no Form refs):
'   1. ExpedienteAlta_EstablecerDatos(p_DTO, p_TipoEnum, p_IDExpedientePadre,
'                                       p_Error)
'      -> {ordinalVisible, lblNumeroVisible, comandoRellenarVisible,
'           ordinalValue, ordinalCaption}
'   2. ExpedienteAlta_EstablecerDatosConTipo(p_DTO, p_TipoTexto, p_Error)
'      -> {esAM, esLote, esExpediente} (string values "Sí" or "No" or "")
'   3. ExpedienteAlta_Ambito_AfterUpdate(p_AmbitoActual, p_Error)
'      -> {hpsAplica: "Sí"|""}
'   4. ExpedienteAlta_Ordinal_AfterUpdate(p_Ordinal, p_IDExpedientePadre,
'                                            p_Conflicto, p_Error)
'      p_Conflicto is a Dictionary stub for ExpedienteOrdinalUsado's result;
'      helper formats the conflict text.
'      -> {conflictoTexto, showWarning}
'   5. ExpedienteAlta_NotificarAltaTipo(p_HasOpenAltaTipo, p_Error)
'      -> {notified:Boolean}
'      (Wrapper around the form-side check for FormExpedienteAltaTipo being open.)
'
' UI orchestration that stays in the form (rule #1):
'   - Form_Load — DAO wiring (SetDatosDeExpedienteTipo, HeredarDatosDePadre)
'   - Form_Unload — snapshot-based "hay cambios sin guardar" prompt
'   - 4 ComandoElegir*_Click — DoCmd.OpenForm for selection forms (UI nav)
'   - 3 ComandoLimpiar*_Click — Me.IdX = Null (UI action)
'   - ComandoVerSharepoint_Click — Ejecutar (file open)
'   - 4 m_Form*_Seleccionar WithEvents callbacks — UI pass-through
'   - AccesoSharepoint_Exit — enable/disable button (UI wiring)
'   - ComandoRegistrar_Click — DAO call (ExpedienteOperaciones.RegistrarAlta)
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

' TipoEnum identifiers (mirrors the EnumTipoExpediente project enum values).
Private Const EXPEDIENTEALTA_TIPO_AM As Long = 1
Private Const EXPEDIENTEALTA_TIPO_LOTE As Long = 2
Private Const EXPEDIENTEALTA_TIPO_BASADO_DE_AM As Long = 3
Private Const EXPEDIENTEALTA_TIPO_BASADO_DE_LOTE As Long = 4
Private Const EXPEDIENTEALTA_TIPO_EXP_INDEP As Long = 5

' TIpo (texto) constants matching the legacy .cls EstablecerDatosConTipo.
Private Const EXPEDIENTEALTA_TIPO_TEXTO_AM As String = "Acuerdo Marco"
Private Const EXPEDIENTEALTA_TIPO_TEXTO_LOTE_SAM As String = "Lote sin Acuero Marco"
Private Const EXPEDIENTEALTA_TIPO_TEXTO_LOTE_AM As String = "Lote de Acuerdo Marco"
Private Const EXPEDIENTEALTA_TIPO_TEXTO_DERIVADO_LOTE As String = "Expediente Derivado de Lote"
Private Const EXPEDIENTEALTA_TIPO_TEXTO_DERIVADO_AM As String = "Expediente Derivado de Acuerdo Marco"
Private Const EXPEDIENTEALTA_TIPO_TEXTO_INDIVIDUAL As String = "Expediente individual"

' Ordinal caption constants.
Private Const EXPEDIENTEALTA_ORDINAL_CAPTION_LOTE As String = "Nº DE LOTE"
Private Const EXPEDIENTEALTA_ORDINAL_CAPTION_BASADO As String = "Nº BASADO"

Private Const EXPEDIENTEALTA_FLAG_SI As String = "Sí"
Private Const EXPEDIENTEALTA_FLAG_NO As String = "No"

' Test fixture ID base for entidades; not used by helpers, just for traceability.
Private Const EXPEDIENTEALTA_TEST_ID_BASE As Long = 900830


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

' --- ExpedienteAlta_GetValueSafely -----------------------------------------------
Private Function ExpedienteAlta_GetValueSafely( _
    ByVal p_Exp As Object, _
    ByVal p_Key As String, _
    ByVal p_Default As Variant _
) As Variant
    On Error GoTo EH
    If p_Exp Is Nothing Then
        ExpedienteAlta_GetValueSafely = p_Default
        Exit Function
    End If
    If Not p_Exp.Exists(p_Key) Then
        ExpedienteAlta_GetValueSafely = p_Default
        Exit Function
    End If
    ExpedienteAlta_GetValueSafely = p_Exp(p_Key)
    Exit Function
EH:
    ExpedienteAlta_GetValueSafely = p_Default
End Function

' --- ExpedienteAlta_DebeMostrarOrdinal -------------------------------------------
' Returns True if the ordinal field should be visible for this (TipoEnum, IDPadre)
' combination. Mirrors the legacy 4-branch logic in Form_FormExpedienteAlta.cls.
Private Function ExpedienteAlta_DebeMostrarOrdinal( _
    ByVal p_TipoEnum As Long, _
    ByVal p_IDExpedientePadre As String _
) As Boolean
    If p_TipoEnum = EXPEDIENTEALTA_TIPO_AM Or _
       p_TipoEnum = EXPEDIENTEALTA_TIPO_EXP_INDEP Then
        ExpedienteAlta_DebeMostrarOrdinal = False
        Exit Function
    End If

    If p_TipoEnum = EXPEDIENTEALTA_TIPO_LOTE And Len(p_IDExpedientePadre) = 0 Then
        ExpedienteAlta_DebeMostrarOrdinal = False
        Exit Function
    End If

    ExpedienteAlta_DebeMostrarOrdinal = True
End Function

' --- ExpedienteAlta_EsOrdinalNumerico --------------------------------------------
' Returns True if the Ordinal value is numeric.
Private Function ExpedienteAlta_EsOrdinalNumerico(ByVal p_Ordinal As String) As Boolean
    ExpedienteAlta_EsOrdinalNumerico = IsNumeric(p_Ordinal)
End Function

' --- ExpedienteAlta_OrdinalCaption ------------------------------------------------
' Returns the lblNumero.Caption value for the given TipoEnum.
Private Function ExpedienteAlta_OrdinalCaption(ByVal p_TipoEnum As Long) As String
    If p_TipoEnum = EXPEDIENTEALTA_TIPO_LOTE Then
        ExpedienteAlta_OrdinalCaption = EXPEDIENTEALTA_ORDINAL_CAPTION_LOTE
    ElseIf p_TipoEnum = EXPEDIENTEALTA_TIPO_BASADO_DE_AM Or _
           p_TipoEnum = EXPEDIENTEALTA_TIPO_BASADO_DE_LOTE Then
        ExpedienteAlta_OrdinalCaption = EXPEDIENTEALTA_ORDINAL_CAPTION_BASADO
    Else
        ExpedienteAlta_OrdinalCaption = ""
    End If
End Function


' === Public API =================================================================

' --- ExpedienteAlta_EstablecerDatos ----------------------------------------------
' Pure-data decision for the EstablecerDatos handler: returns the visibility
' flags and ordinal value to apply to the form's controls.
'
' p_DTO is the Dictionary wrapper with .Expediente (Dictionary stub with TIpo,
' Nemotecnico, IDExpedientePadre, IDExpediente, Ordinal, OrdinalCalculado).
'
' Returns JSON: {ok, payload:{ordinalVisible, lblNumeroVisible,
'                              comandoRellenarVisible, ordinalValue,
'                              ordinalCaption}, error, logs}.
Public Function ExpedienteAlta_EstablecerDatos( _
    ByVal p_DTO As Object, _
    ByVal p_TipoEnum As Long, _
    ByVal p_IDExpedientePadre As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    If p_DTO Is Nothing Then
        p_Error = "ExpedienteAlta_EstablecerDatos: p_DTO is Nothing"
        ExpedienteAlta_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim ordinalVisible As Boolean
    Dim lblNumeroVisible As Boolean
    Dim comandoRellenarVisible As Boolean
    Dim ordinalValue As String
    Dim ordinalCaption As String

    lblNumeroVisible = True
    ordinalVisible = True
    comandoRellenarVisible = True

    If Not ExpedienteAlta_DebeMostrarOrdinal(p_TipoEnum, p_IDExpedientePadre) Then
        lblNumeroVisible = False
        ordinalVisible = False
        comandoRellenarVisible = False
    End If

    ' Compute the ordinal value (mirrors legacy logic):
    '   if .IDExpediente = "":
    '     if not IsNumeric(.Ordinal): use .OrdinalCalculado
    '     else: use .Ordinal
    '   else:
    '     if IsNumeric(.Ordinal): use .Ordinal
    Dim expObj As Object
    Set expObj = ExpedienteAlta_GetValueSafely(p_DTO, "Expediente", Nothing)

    Dim idExp As String
    Dim ordinalStored As String
    Dim ordinalCalculado As String

    If Not expObj Is Nothing Then
        idExp = CStr(ExpedienteAlta_GetValueSafely(expObj, "IDExpediente", ""))
        ordinalStored = CStr(ExpedienteAlta_GetValueSafely(expObj, "Ordinal", ""))
        ordinalCalculado = CStr(ExpedienteAlta_GetValueSafely(expObj, "OrdinalCalculado", ""))
    Else
        idExp = ""
        ordinalStored = ""
        ordinalCalculado = ""
    End If

    If idExp = "" Then
        If Not ExpedienteAlta_EsOrdinalNumerico(ordinalStored) Then
            ordinalValue = ordinalCalculado
        Else
            ordinalValue = ordinalStored
        End If
    Else
        If ExpedienteAlta_EsOrdinalNumerico(ordinalStored) Then
            ordinalValue = ordinalStored
        Else
            ordinalValue = ""
        End If
    End If

    ordinalCaption = ExpedienteAlta_OrdinalCaption(p_TipoEnum)

    payload("ordinalVisible") = ordinalVisible
    payload("lblNumeroVisible") = lblNumeroVisible
    payload("comandoRellenarVisible") = comandoRellenarVisible
    payload("ordinalValue") = ordinalValue
    payload("ordinalCaption") = ordinalCaption

    logs(0) = "EstablecerDatos: tipoEnum=" & CStr(p_TipoEnum) & _
              " ordinalVisible=" & CStr(ordinalVisible) & _
              " ordinalValue='" & ordinalValue & "'"

    ExpedienteAlta_EstablecerDatos = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteAlta_EstablecerDatos: " & Err.Description
    End If
    ExpedienteAlta_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteAlta_EstablecerDatosConTipo ---------------------------------------
' Pure-data translation of TIpo (texto) to {EsAM, EsLote, EsExpediente} flags.
' Returns "" for each flag if TIpo is empty or unrecognized.
'
' Note: this helper does NOT mutate the DTO; it returns the computed flags so
' the form (or DAO caller) can apply them. Mirrors legacy logic verbatim.
'
' Returns JSON: {ok, payload:{esAM, esLote, esExpediente}, error, logs}.
Public Function ExpedienteAlta_EstablecerDatosConTipo( _
    ByVal p_DTO As Object, _
    ByVal p_TipoTexto As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim esAM As String
    Dim esLote As String
    Dim esExpediente As String

    Dim tipo As String
    tipo = Trim$("" & p_TipoTexto)

    Select Case tipo
        Case EXPEDIENTEALTA_TIPO_TEXTO_AM
            esAM = EXPEDIENTEALTA_FLAG_SI
            esLote = EXPEDIENTEALTA_FLAG_NO
            esExpediente = EXPEDIENTEALTA_FLAG_NO
        Case EXPEDIENTEALTA_TIPO_TEXTO_LOTE_SAM, _
             EXPEDIENTEALTA_TIPO_TEXTO_LOTE_AM
            esAM = EXPEDIENTEALTA_FLAG_NO
            esLote = EXPEDIENTEALTA_FLAG_SI
            esExpediente = EXPEDIENTEALTA_FLAG_NO
        Case EXPEDIENTEALTA_TIPO_TEXTO_DERIVADO_LOTE, _
             EXPEDIENTEALTA_TIPO_TEXTO_DERIVADO_AM, _
             EXPEDIENTEALTA_TIPO_TEXTO_INDIVIDUAL
            esAM = EXPEDIENTEALTA_FLAG_NO
            esLote = EXPEDIENTEALTA_FLAG_NO
            esExpediente = EXPEDIENTEALTA_FLAG_SI
        Case ""
            esAM = ""
            esLote = ""
            esExpediente = ""
        Case Else
            p_Error = "ExpedienteAlta_EstablecerDatosConTipo: unknown tipo '" & tipo & "'"
            ExpedienteAlta_EstablecerDatosConTipo = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
    End Select

    payload("esAM") = esAM
    payload("esLote") = esLote
    payload("esExpediente") = esExpediente
    logs(0) = "EstablecerDatosConTipo: tipo='" & tipo & "' esAM=" & esAM & _
              " esLote=" & esLote & " esExpediente=" & esExpediente

    ExpedienteAlta_EstablecerDatosConTipo = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteAlta_EstablecerDatosConTipo: " & Err.Description
    End If
    ExpedienteAlta_EstablecerDatosConTipo = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteAlta_Ambito_AfterUpdate -------------------------------------------
' Pure-data decision for Ambito_AfterUpdate: if p_AmbitoActual = "HPS", returns
' hpsAplica="Sí". Otherwise returns hpsAplica="" (no change).
'
' Returns JSON: {ok, payload:{hpsAplica}, error, logs}.
Public Function ExpedienteAlta_Ambito_AfterUpdate( _
    ByVal p_AmbitoActual As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim hpsAplica As String
    If UCase$("" & p_AmbitoActual) = "HPS" Then
        hpsAplica = EXPEDIENTEALTA_FLAG_SI
    Else
        hpsAplica = ""
    End If

    payload("hpsAplica") = hpsAplica
    logs(0) = "Ambito_AfterUpdate: ambito='" & p_AmbitoActual & "' hpsAplica='" & hpsAplica & "'"

    ExpedienteAlta_Ambito_AfterUpdate = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteAlta_Ambito_AfterUpdate: " & Err.Description
    End If
    ExpedienteAlta_Ambito_AfterUpdate = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteAlta_Ordinal_AfterUpdate ------------------------------------------
' Pure-data decision for Ordinal_AfterUpdate: validates the ordinal, builds the
' conflict text if there is a conflict, and returns whether to show a warning.
'
' p_Ordinal: the new ordinal entered by the user.
' p_IDExpedientePadre: the parent expediente ID (empty = no conflict possible).
' p_Conflicto: Dictionary stub for ExpedienteOrdinalUsado's result (None or a
'   Dictionary with CodExp/Nemotecnico/Titulo). Passing Nothing = no conflict.
'
' Returns JSON: {ok, payload:{conflictoTexto, showWarning}, error, logs}.
Public Function ExpedienteAlta_Ordinal_AfterUpdate( _
    ByVal p_Ordinal As String, _
    ByVal p_IDExpedientePadre As String, _
    ByVal p_Conflicto As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim showWarning As Boolean
    showWarning = False

    Dim conflictoTexto As String
    conflictoTexto = ""

    Dim ordinalVal As String
    ordinalVal = Trim$("" & p_Ordinal)

    ' Empty/non-numeric ordinal -> no validation needed.
    If Not ExpedienteAlta_EsOrdinalNumerico(ordinalVal) Then
        payload("conflictoTexto") = ""
        payload("showWarning") = False
        ExpedienteAlta_Ordinal_AfterUpdate = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' No padre -> no conflict possible.
    If Len(Trim$("" & p_IDExpedientePadre)) = 0 Then
        payload("conflictoTexto") = ""
        payload("showWarning") = False
        ExpedienteAlta_Ordinal_AfterUpdate = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' No conflict reported by caller -> no warning.
    If p_Conflicto Is Nothing Then
        payload("conflictoTexto") = ""
        payload("showWarning") = False
        ExpedienteAlta_Ordinal_AfterUpdate = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Build the conflict text from p_Conflicto fields (CodExp > Nemotecnico > Titulo).
    Dim codExp As String
    Dim nemotecnico As String
    Dim titulo As String

    If TypeName(p_Conflicto) = "Dictionary" Then
        codExp = CStr(ExpedienteAlta_GetValueSafely(p_Conflicto, "CodExp", ""))
        nemotecnico = CStr(ExpedienteAlta_GetValueSafely(p_Conflicto, "Nemotecnico", ""))
        titulo = CStr(ExpedienteAlta_GetValueSafely(p_Conflicto, "Titulo", ""))
    Else
        On Error Resume Next
        codExp = p_Conflicto.CodExp
        nemotecnico = p_Conflicto.Nemotecnico
        titulo = p_Conflicto.Titulo
        On Error GoTo errores
    End If

    Dim textoExp As String
    If Len(codExp) > 0 Then
        textoExp = codExp
    ElseIf Len(nemotecnico) > 0 Then
        textoExp = nemotecnico
    Else
        textoExp = titulo
    End If

    showWarning = True
    conflictoTexto = "El número ya ha sido usado para el Expediente : " & textoExp

    payload("conflictoTexto") = conflictoTexto
    payload("showWarning") = showWarning
    logs(0) = "Ordinal_AfterUpdate: conflict detected textoExp='" & textoExp & "'"

    ExpedienteAlta_Ordinal_AfterUpdate = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteAlta_Ordinal_AfterUpdate: " & Err.Description
    End If
    ExpedienteAlta_Ordinal_AfterUpdate = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteAlta_NotificarAltaTipo --------------------------------------------
' Pure-data wrapper around the form-side notification to FormExpedienteAltaTipo.
' The form's Form_Registrar handler passes p_HasOpenAltaTipo (a Boolean the form
' derives from FormularioAbierto). The helper returns the structured decision.
'
' Returns JSON: {ok, payload:{notified, reason}, error, logs}.
Public Function ExpedienteAlta_NotificarAltaTipo( _
    ByVal p_HasOpenAltaTipo As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim notified As Boolean
    notified = p_HasOpenAltaTipo

    Dim reason As String
    If notified Then
        reason = "FormExpedienteAltaTipo is open — notified"
    Else
        reason = "FormExpedienteAltaTipo not open — skipped"
    End If

    payload("notified") = notified
    payload("reason") = reason
    logs(0) = "NotificarAltaTipo: notified=" & CStr(notified)

    ExpedienteAlta_NotificarAltaTipo = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteAlta_NotificarAltaTipo: " & Err.Description
    End If
    ExpedienteAlta_NotificarAltaTipo = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function