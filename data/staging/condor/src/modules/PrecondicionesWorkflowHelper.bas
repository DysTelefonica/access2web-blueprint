Attribute VB_Name = "PrecondicionesWorkflowHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MODULE: PrecondicionesWorkflowHelper
' PURPOSE: Shared pure helper that absorbs the BR-002..BR-006 precondiciones
'          branches of WorkflowServicio.CumplePasoA*. Mirrors the pure-helper
'          pattern of DecisionFinalHelper (Slice 3.1) and
'          DictamenRACDefaultsHelper (Slice 3.2).
'
' Contract documented in:
'   changes/e2e-methodology-exhaustive-rollout/specs/workflow-precondicion-transicion/spec.md
'
' Constraints (per spec):
'   * Pure helper: NO DAO.Database, NO control access, NO MsgBox.
'   * Value-type inputs including Variant to absorb Null DAO returns.
'   * Trim/Nz of string inputs mirrors DecisionFinalHelper style.
'   * 100% deterministic (no Date, no Random, no Env).
'
' BR mapping:
'   * BR-002 (parte tecnica completa) -> _ParteTecnicaCompleta
'   * BR-003 (RAC completo + aprobacion suministrador) -> _RACCompleto
'   * BR-004 (adjunto de borrador en Validacion) + BR-006 (adjunto de cierre)
'     -> _AdjuntoEnEtapa
'   * BR-005 (resultado validacion APROBADO) -> _ResultadoValidacionAprobado
'
' D3 (slice 3.3, 2026-07-13): _AdjuntoEnEtapa acepta tanto el nombre canónico
' sin tilde ("Validacion") como la variante localizada con tilde
' ("Validación") que devuelve GetNombreEstadoSafe(estadoValidacion). El
' helper normaliza internamente los caracteres acentuados antes de comparar
' con "VALIDACION" / "CIERRE". Firma pública y atom set existentes no
' cambian; solo se ampla la tolerancia del input.
' ==========================================================================

' p_Entidad shape (Variant):
'   * Dictionary / Object with key "ParteTecnicaCompleta" -> Boolean (preferred)
'   * Plain Boolean (pre-computed by caller)
'   * Empty / Nothing / other -> treated as incomplete (False)
'
' Writes p_RazonFallo with CAP-007 §2 wording when the result is False.
Public Function PrecondicionesWorkflowHelper_ParteTecnicaCompleta( _
    ByVal p_TipoSolicitud As Variant, _
    ByVal p_EstadoOrigen As Variant, _
    ByVal p_Entidad As Variant, _
    ByRef p_RazonFallo As String) As Boolean

    Dim tipo As String
    Dim origen As String
    Dim completa As Boolean

    PrecondicionesWorkflowHelper_ParteTecnicaCompleta = False
    p_RazonFallo = ""

    tipo = UCase$(Trim$(Nz(p_TipoSolicitud, vbNullString)))
    origen = UCase$(Trim$(Nz(p_EstadoOrigen, vbNullString)))

    ' Pre-read shape resolution. Caller is responsible for the DAO read;
    ' the helper only evaluates the pre-read flag.
    completa = PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve(p_Entidad)

    If completa Then
        PrecondicionesWorkflowHelper_ParteTecnicaCompleta = True
        Exit Function
    End If

    ' Wording mirrors CAP-007 §2 BR-002: "Faltan campos obligatorios en las
    ' pestañas 'Propuesta' o 'Impacto'."
    p_RazonFallo = "Faltan campos obligatorios en las pestanas 'Propuesta' o 'Impacto' " & _
                   "(tipo=" & tipo & ", origen=" & origen & ")."
End Function

' BR-003 + BR-008 exemption under RECHAZADO.
'   * codigo required unless decision = "RECHAZADO" (case-insensitive)
'   * nombre required always
'   * decision required (treated as missing if Null/Empty)
'   * AprobacionSuministradorCompleta must be True (strict Boolean contract)
'
' Variant on the first three absorbs DAO Null returns without raising 94.
Public Function PrecondicionesWorkflowHelper_RACCompleto( _
    ByVal p_racCodigo As Variant, _
    ByVal p_racNombre As Variant, _
    ByVal p_racDecision As Variant, _
    ByVal p_AprobacionSuministradorCompleta As Boolean) As Boolean

    Dim codigo As String
    Dim nombre As String
    Dim decision As String

    PrecondicionesWorkflowHelper_RACCompleto = False

    ' Trim + Nz to absorb DAO Null without error 94.
    nombre = Trim$(Nz(p_racNombre, vbNullString))
    codigo = Trim$(Nz(p_racCodigo, vbNullString))
    decision = UCase$(Trim$(Nz(p_racDecision, vbNullString)))

    ' Nombre is required regardless of decision (BR-008).
    If Len(nombre) = 0 Then Exit Function

    ' BR-008: under RECHAZADO the codigo may be empty; everything else requires it.
    If decision = "RECHAZADO" Then
        PrecondicionesWorkflowHelper_RACCompleto = p_AprobacionSuministradorCompleta
        Exit Function
    End If

    ' Non-RECHAZADO: codigo + decision + aprobacion all required.
    If Len(codigo) = 0 Then Exit Function
    If Len(decision) = 0 Then Exit Function

    PrecondicionesWorkflowHelper_RACCompleto = p_AprobacionSuministradorCompleta
End Function

' BR-004 (adjunto de borrador en Validacion) + BR-006 (adjunto de cierre).
' Pre-read boolean from AdjuntosServicio.ExisteAdjuntoEtapa. The helper
' stays pure: it only validates the canonical etapa name + the pre-read
' flag, never touches DAO.
'
' Valid etapa names (case-insensitive): "Validacion", "Cierre".
' Unknown etapa -> False (safe-by-default per spec scenario AdjuntoEnEtapa_UnknownEtapa_False).
' D3: toleramos tambien la variante con tilde que devuelve GetNombreEstadoSafe
' (p.ej. "Validación") normalizando los caracteres acentuados antes de comparar.
Public Function PrecondicionesWorkflowHelper_AdjuntoEnEtapa( _
    ByVal p_Etapa As Variant, _
    ByVal p_AdjuntoExiste As Variant) As Boolean

    Dim etapa As String

    PrecondicionesWorkflowHelper_AdjuntoEnEtapa = False

    etapa = PrecondicionesWorkflowHelper_NormalizarTildes(UCase$(Trim$(Nz(p_Etapa, vbNullString))))
    If Len(etapa) = 0 Then Exit Function

    ' Canonical etapas that demand an attached document.
    If etapa <> "VALIDACION" And etapa <> "CIERRE" Then Exit Function

    ' The boolean contract accepts Variant so callers can pass the DAO
    ' value (Long 0/1 or actual Boolean) without an extra conversion.
    If IsNumeric(p_AdjuntoExiste) Then
        PrecondicionesWorkflowHelper_AdjuntoEnEtapa = (CLng(p_AdjuntoExiste) <> 0)
        Exit Function
    End If

    PrecondicionesWorkflowHelper_AdjuntoEnEtapa = CBool(p_AdjuntoExiste)
End Function

' BR-005: True iff pre-read p_UltimoResultado = "APROBADO" (case-insensitive,
' trimmed, Null-tolerant). The comentarios argument is reserved for the
' spec's "non-empty comentarios" requirement but the canonical atom set
' (per Slice 3.3 brief) only depends on resultado; the helper leaves
' the comentarios check to the caller so it stays a single-argument
' pure function as documented in the brief.
Public Function PrecondicionesWorkflowHelper_ResultadoValidacionAprobado( _
    ByVal p_UltimoResultado As Variant) As Boolean

    Dim resultado As String

    PrecondicionesWorkflowHelper_ResultadoValidacionAprobado = False

    resultado = UCase$(Trim$(Nz(p_UltimoResultado, vbNullString)))
    PrecondicionesWorkflowHelper_ResultadoValidacionAprobado = (resultado = "APROBADO")
End Function

' ----------------------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------------------

' Resolves the pre-read "parte tecnica completa" flag from a Variant that
' may carry a Dictionary / Object / Boolean / Long / Empty payload. Returns
' False for any unrecognized shape (safe-by-default per spec).
Private Function PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve( _
    ByVal p_Entidad As Variant) As Boolean

    Dim key As Variant
    Dim value As Variant

    PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve = False

    If IsEmpty(p_Entidad) Then Exit Function
    If IsNull(p_Entidad) Then Exit Function

    ' Plain Boolean (pre-computed by caller).
    If VarType(p_Entidad) = vbBoolean Then
        PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve = CBool(p_Entidad)
        Exit Function
    End If

    ' Plain Long (0/1) when caller forwards a DAO-style flag.
    If VarType(p_Entidad) = vbLong Or VarType(p_Entidad) = vbInteger Then
        PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve = (CLng(p_Entidad) <> 0)
        Exit Function
    End If

    ' Object case (Dictionary, custom class). The helper reads a single
    ' canonical key so callers from all four gemelos stay symmetric.
    If IsObject(p_Entidad) Then
        If p_Entidad Is Nothing Then Exit Function
        ' Defensive: Dictionary lookup raises 32811 if key missing.
        On Error Resume Next
        value = p_Entidad("ParteTecnicaCompleta")
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

        If IsEmpty(value) Or IsNull(value) Then Exit Function
        If VarType(value) = vbBoolean Then
            PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve = CBool(value)
            Exit Function
        End If
        If IsNumeric(value) Then
            PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve = (CLng(value) <> 0)
            Exit Function
        End If
        Exit Function
    End If

    ' String fallback: "S"/"N" or "1"/"0" honoring common patterns from the
    ' caller (caller may pass a stringified flag from the form layer).
    If VarType(p_Entidad) = vbString Then
        Dim s As String
        s = UCase$(Trim$(CStr(p_Entidad)))
        PrecondicionesWorkflowHelper_ParteTecnicaCompleta_Resolve = _
            (s = "S" Or s = "SI" Or s = "TRUE" Or s = "1")
        Exit Function
    End If
End Function

' Sustituye los caracteres acentuados más comunes (ISO-8859-1 / Latin-1) por
' su equivalente ASCII. Usado por _AdjuntoEnEtapa para tolerar tanto
' "Validacion" como "Validación" que devuelve GetNombreEstadoSafe.
' Encoding-independent: se basa en Chr$() en vez de literales con tildes para
' sobrevivir a la grabación en UTF-8 sin BOM del source.
Private Function PrecondicionesWorkflowHelper_NormalizarTildes(ByVal p_Valor As String) As String
    Dim result As String
    result = p_Valor
    result = Replace(result, Chr$(193), "A")  ' Á
    result = Replace(result, Chr$(201), "E")  ' É
    result = Replace(result, Chr$(205), "I")  ' Í
    result = Replace(result, Chr$(211), "O")  ' Ó
    result = Replace(result, Chr$(218), "U")  ' Ú
    result = Replace(result, Chr$(225), "a")  ' á
    result = Replace(result, Chr$(233), "e")  ' é
    result = Replace(result, Chr$(237), "i")  ' í
    result = Replace(result, Chr$(243), "o")  ' ó
    result = Replace(result, Chr$(250), "u")  ' ú
    PrecondicionesWorkflowHelper_NormalizarTildes = result
End Function
