Attribute VB_Name = "NotificacionHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MODULE: NotificacionHelper
' PURPOSE: Thin wrapper around NotificacionServicio.EnviarNotificacion with
'          an On Error Resume Next block. Encodes the non-blocking policy
'          (BR-001: a failed email must not revert a committed transition).
'
' Contract documented in:
'   changes/e2e-methodology-exhaustive-rollout/specs/workflow-precondicion-transicion/spec.md
'
' Real service signature this helper forwards to
' (src/classes/NotificacionServicio.cls):
'   Public Sub EnviarNotificacion( _
'       ByRef vm As SolicitudViewModel, _
'       ByVal asunto As String, _
'       ByVal destinatarioEmail As String, _
'       ByVal copiaEmail As String, _
'       Optional ByVal nombreEstadoNotif As String = "")
'
' Constraints (per spec):
'   * Pure helper: NO DAO.Database, NO control access, NO MsgBox.
'   * Returns "OK" on success or a "NOTIF_ERROR_<numero>_<desc>" marker on
'     swallowed failure.
'   * Never propagates an error to the caller (BR-001 policy).
'   * Variant on every input to absorb DAO Null gracefully (same pattern as
'     DictamenRACDefaultsHelper).
'   * Stub-friendly signature (Object for the service) so tests can inject a
'     MockNotifServ without spinning up the real backend.
'
' Nothing / invalid service:
'   When p_NotifServ Is Nothing, invoking .EnviarNotificacion raises a
'   runtime error (VBA reports 91 "Object variable or With block variable
'   not set"; some hosts surface 438 "Object doesn't support this property
'   or method"). Either way the On Error Resume Next block SWALLOWS it and
'   returns "NOTIF_ERROR_<numero>_...". The helper never crashes the caller.
' ==========================================================================

' Pre-computed prefix used in the swallowed-error return value. Kept as
' a Private Const so callers can pattern-match on the marker.
Private Const NOTIF_HELPER_ERROR_PREFIX As String = "NOTIF_ERROR_"

Public Function NotificacionHelper_NotificarSiPosible( _
    ByVal p_NotifServ As Object, _
    ByVal p_vm As Variant, _
    ByVal p_asunto As Variant, _
    ByVal p_destinatarioEmail As Variant, _
    ByVal p_copiaEmail As Variant, _
    Optional ByVal p_nombreEstadoNotif As Variant = "") As String

    On Error Resume Next

    p_NotifServ.EnviarNotificacion p_vm, p_asunto, p_destinatarioEmail, p_copiaEmail, p_nombreEstadoNotif

    If Err.Number <> 0 Then
        ' BR-001: swallow + capture motivo. Caller may log this string.
        ' Err.Clear is CRITICAL: it clears the error so it does not
        ' propagate to the caller once we return.
        NotificacionHelper_NotificarSiPosible = _
            NOTIF_HELPER_ERROR_PREFIX & Err.Number & "_" & Left$(Err.Description, 60)
        Err.Clear
    Else
        NotificacionHelper_NotificarSiPosible = "OK"
    End If

    On Error GoTo 0
End Function
