Attribute VB_Name = "DictamenRACDefaultsHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MODULE: DictamenRACDefaultsHelper
' PURPOSE: Shared pure helper for RAC defaults-loading plan and BR-008
'          completeness rule across PC, PCSUB, CDCA, CDCASUB gemelos.
'
' Contract documented in:
'   changes/e2e-methodology-exhaustive-rollout/specs/dictamen-rac-helper-contract/spec.md
'
' Constraints (per spec):
'   * Pure helper: NO DAO.Database, NO control access, NO MsgBox.
'   * Value-type inputs including Variant to absorb Null DAO returns.
'   * Trim/Nz of string inputs mirrors DecisionFinalHelper style.
'
' Asymmetry resolution:
'   * PCSUB resolves its destination control to racDelegadoNombre
'     instead of racNombre; the helper encodes the target field name
'     inside p_MessagePlan with a pipe separator
'     ("<targetField>|<message>") so the four subforms can dispatch
'     without gemelo-specific branches in their handlers.
'   * PC / CDCA / CDCASUB resolve to racNombre. CDCASUB also writes to
'     racNombreDelegador; that second target stays inside the Form
'     wiring (Slice 3.4) and is not encoded here because it does not
'     depend on whether the expediente has a RAC.
' ==========================================================================

' ConstruirPlan returns True when the defaults should be applied
' (i.e. there is an editable block AND the expediente has an assigned
' RAC). Returns False otherwise, with a p_PromptResult seam the caller
' renders instead of a MsgBox.
'
' When apply=True, p_MessagePlan carries the target control name and
' the message text in the format "<targetField>|<message>" so the four
' subforms can dispatch without gemelo branches.
Public Function DictamenRACDefaultsHelper_ConstruirPlan( _
    ByVal p_TipoSolicitud As String, _
    ByVal p_RolUsuario As String, _
    ByVal p_RacAsignadoByExpediente As Variant, _
    ByVal p_BloqueEditable As Boolean, _
    ByVal p_racDecisionActual As Variant, _
    ByRef p_PromptResult As String, _
    ByRef p_MessagePlan As String) As Boolean

    DictamenRACDefaultsHelper_ConstruirPlan = False
    p_PromptResult = ""
    p_MessagePlan = ""

    ' --- Guard: workflow / role says this block is read-only
    If Not p_BloqueEditable Then
        p_PromptResult = "BloqueNoEditable"
        p_MessagePlan = "El bloque no es editable por permisos/workflow"
        Exit Function
    End If

    ' --- Guard: expediente has no RAC assigned ("" or Null both denied)
    If Len(Trim$(Nz(p_RacAsignadoByExpediente, vbNullString))) = 0 Then
        p_PromptResult = "SinRACAsignado"
        p_MessagePlan = "No hay RAC asignado al expediente; complete manualmente o asigne uno"
        Exit Function
    End If

    ' --- Happy path: defaults apply; signal target field by gemelo
    Dim targetField As String
    targetField = DictamenRACDefaultsHelper_TargetFieldFor(p_TipoSolicitud)
    p_PromptResult = "OK"
    p_MessagePlan = targetField & "|RAC cargado"
    DictamenRACDefaultsHelper_ConstruirPlan = True
End Function

' ReglaEsCompleta centralizes BR-008 (CAP-001) and its homologs in
' CAP-003 / CAP-004 / CAP-005 for EsDictamenRACCompleto delegation.
'
' Rules:
'   * racDecision="RECHAZADO" (case-insensitive) exempts racCodigo
'     (but requires racNombre to be non-blank).
'   * Any other non-empty racDecision requires both codigo AND nombre.
'   * Empty racDecision requires codigo AND nombre AND decision all
'     non-blank (legacy completeness default).
Public Function DictamenRACDefaultsHelper_ReglaEsCompleta( _
    ByVal p_racCodigo As Variant, _
    ByVal p_racNombre As Variant, _
    ByVal p_racDecision As Variant) As Boolean

    Dim nombre As String
    Dim codigo As String
    Dim decision As String

    nombre = Trim$(Nz(p_racNombre, vbNullString))
    codigo = Trim$(Nz(p_racCodigo, vbNullString))
    decision = UCase$(Trim$(Nz(p_racDecision, vbNullString)))

    DictamenRACDefaultsHelper_ReglaEsCompleta = False

    If decision = "RECHAZADO" Then
        ' BR-008: codigo exempt; nombre still required.
        If Len(nombre) > 0 Then DictamenRACDefaultsHelper_ReglaEsCompleta = True
        Exit Function
    End If

    If Len(decision) > 0 Then
        ' Non-empty decision other than RECHAZADO: codigo AND nombre.
        If Len(codigo) > 0 And Len(nombre) > 0 Then _
            DictamenRACDefaultsHelper_ReglaEsCompleta = True
        Exit Function
    End If

    ' Empty decision -> all three required.
    If Len(codigo) > 0 And Len(nombre) > 0 And Len(decision) > 0 Then _
        DictamenRACDefaultsHelper_ReglaEsCompleta = True
End Function

' Private target field resolution. PCSUB uses racDelegadoNombre; the
' other three gemelos use racNombre. This is where the PCSUB asymmetry
' is centralized per Spec-005_Alignment (gemelos alignment).
Private Function DictamenRACDefaultsHelper_TargetFieldFor(ByVal p_TipoSolicitud As String) As String
    Select Case UCase$(Trim$(Nz(p_TipoSolicitud, vbNullString)))
        Case "PCSUB"
            DictamenRACDefaultsHelper_TargetFieldFor = "racDelegadoNombre"
        Case Else
            DictamenRACDefaultsHelper_TargetFieldFor = "racNombre"
    End Select
End Function