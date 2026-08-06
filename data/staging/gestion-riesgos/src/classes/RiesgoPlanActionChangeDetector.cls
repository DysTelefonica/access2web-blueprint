VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoPlanActionChangeDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' RiesgoPlanActionChangeDetector (issue #42)
' Pure field-by-field comparison for risk plans and actions.
' Extracted from plan/action form HaHabidoCambios methods so the
' comparison contract can be tested without UI, DAO or Access forms.
' ============================================================

Public Function HasPMChanges( _
    ByVal p_Original As PM, _
    ByVal p_Current As PM, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    p_Error = vbNullString
    If p_Original Is Nothing Then
        p_Error = "HasPMChanges: p_Original is Nothing"
        Exit Function
    End If
    If p_Current Is Nothing Then
        p_Error = "HasPMChanges: p_Current is Nothing"
        Exit Function
    End If

    HasPMChanges = Not TextEqual(p_Original.DisparadorDelPlan, p_Current.DisparadorDelPlan)
    Exit Function

EH:
    p_Error = "RiesgoPlanActionChangeDetector.HasPMChanges: " & Err.Description
End Function

Public Function HasPCChanges( _
    ByVal p_Original As PC, _
    ByVal p_Current As PC, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    p_Error = vbNullString
    If p_Original Is Nothing Then
        p_Error = "HasPCChanges: p_Original is Nothing"
        Exit Function
    End If
    If p_Current Is Nothing Then
        p_Error = "HasPCChanges: p_Current is Nothing"
        Exit Function
    End If

    HasPCChanges = Not TextEqual(p_Original.DisparadorDelPlan, p_Current.DisparadorDelPlan)
    Exit Function

EH:
    p_Error = "RiesgoPlanActionChangeDetector.HasPCChanges: " & Err.Description
End Function

Public Function HasPMActionChanges( _
    ByVal p_Original As PMAccion, _
    ByVal p_Current As PMAccion, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    p_Error = vbNullString
    If p_Original Is Nothing Then
        p_Error = "HasPMActionChanges: p_Original is Nothing"
        Exit Function
    End If
    If p_Current Is Nothing Then
        p_Error = "HasPMActionChanges: p_Current is Nothing"
        Exit Function
    End If

    With p_Original
        If Not TextEqual(.Accion, p_Current.Accion) Then
            HasPMActionChanges = True
            Exit Function
        End If
        If Not TextEqual(.ResponsableAccion, p_Current.ResponsableAccion) Then
            HasPMActionChanges = True
            Exit Function
        End If
        If Not DateFieldEqual(.FechaInicio, p_Current.FechaInicio) Then
            HasPMActionChanges = True
            Exit Function
        End If
        If Not DateFieldEqual(.FechaFinPrevista, p_Current.FechaFinPrevista) Then
            HasPMActionChanges = True
            Exit Function
        End If
        If Not DateFieldEqual(.FechaFinReal, p_Current.FechaFinReal) Then
            HasPMActionChanges = True
            Exit Function
        End If
    End With

    HasPMActionChanges = False
    Exit Function

EH:
    p_Error = "RiesgoPlanActionChangeDetector.HasPMActionChanges: " & Err.Description
End Function

Public Function HasPCActionChanges( _
    ByVal p_Original As PCAccion, _
    ByVal p_Current As PCAccion, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    p_Error = vbNullString
    If p_Original Is Nothing Then
        p_Error = "HasPCActionChanges: p_Original is Nothing"
        Exit Function
    End If
    If p_Current Is Nothing Then
        p_Error = "HasPCActionChanges: p_Current is Nothing"
        Exit Function
    End If

    With p_Original
        If Not TextEqual(.Accion, p_Current.Accion) Then
            HasPCActionChanges = True
            Exit Function
        End If
        If Not TextEqual(.ResponsableAccion, p_Current.ResponsableAccion) Then
            HasPCActionChanges = True
            Exit Function
        End If
        If Not DateFieldEqual(.FechaInicio, p_Current.FechaInicio) Then
            HasPCActionChanges = True
            Exit Function
        End If
        If Not DateFieldEqual(.FechaFinPrevista, p_Current.FechaFinPrevista) Then
            HasPCActionChanges = True
            Exit Function
        End If
        If Not DateFieldEqual(.FechaFinReal, p_Current.FechaFinReal) Then
            HasPCActionChanges = True
            Exit Function
        End If
    End With

    HasPCActionChanges = False
    Exit Function

EH:
    p_Error = "RiesgoPlanActionChangeDetector.HasPCActionChanges: " & Err.Description
End Function

Private Function TextEqual(ByVal p_Left As String, ByVal p_Right As String) As Boolean
    TextEqual = (Nz(p_Left, vbNullString) = Nz(p_Right, vbNullString))
End Function

Private Function DateFieldEqual(ByVal p_Left As String, ByVal p_Right As String) As Boolean
    If IsDate(p_Left) Then
        If Not IsDate(Nz(p_Right, vbNullString)) Then
            DateFieldEqual = False
            Exit Function
        End If
        DateFieldEqual = (Format(p_Left, "dd/mm/yyyy") = Format(p_Right, "dd/mm/yyyy"))
    Else
        DateFieldEqual = Not IsDate(Nz(p_Right, vbNullString))
    End If
End Function
