VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublicacionCalidadExecutionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' PublicacionCalidadExecutionService (issue #39, slice 1)
' Pure decision/orchestration helpers extracted from
' Form_FormPublicacionCalidadPublicarEjecutar.
'
' Slice 1: ResolveRefreshRiskID — the precondition that decides
' whether the publication workflow should refresh quality tasks
' against a specific riesgo or fall through to a global reset.
'
' Form keeps UI (MsgBox, DoCmd, Hourglass, Ejecutar).
' Service keeps: pure comparisons and result shapes.
' ============================================================

Public Function ResolveRefreshRiskID( _
    ByVal p_RiesgoActivo As riesgo, _
    ByVal p_EdicionActiva As Edicion, _
    Optional ByRef p_Error As String _
) As String
    On Error GoTo EH

    p_Error = vbNullString

    If p_RiesgoActivo Is Nothing Then
        ResolveRefreshRiskID = vbNullString
        Exit Function
    End If

    If Trim$(Nz(p_RiesgoActivo.IDRiesgo, vbNullString)) = vbNullString Then
        ResolveRefreshRiskID = vbNullString
        Exit Function
    End If

    If Not p_EdicionActiva Is Nothing Then
        If Trim$(Nz(p_RiesgoActivo.IDEdicion, vbNullString)) <> vbNullString Then
            If CStr(p_RiesgoActivo.IDEdicion) <> CStr(p_EdicionActiva.IDEdicion) Then
                ResolveRefreshRiskID = vbNullString
                Exit Function
            End If
        End If
    End If

    ResolveRefreshRiskID = CStr(p_RiesgoActivo.IDRiesgo)
    Exit Function

EH:
    p_Error = "PublicacionCalidadExecutionService.ResolveRefreshRiskID: " & Err.Description
End Function
