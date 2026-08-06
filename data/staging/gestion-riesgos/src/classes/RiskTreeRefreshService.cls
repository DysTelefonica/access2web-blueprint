VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiskTreeRefreshService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function BuildPlan( _
    ByVal p_Scope As String, _
    ByVal p_IDEdicion As String, _
    ByVal p_IDRiesgo As String, _
    ByVal p_RequestedKey As String, _
    ByVal p_CandidateKeys As String, _
    ByVal p_ExistingKeys As String, _
    Optional ByRef p_Error As String _
) As Scripting.Dictionary
    Dim plan As Scripting.Dictionary
    Dim resolver As RiskTreeSelectionResolver
    Dim normalizedScope As String
    Dim invalidationOrder As String
    Dim fallbackKey As String

    On Error GoTo EH

    p_Error = vbNullString
    normalizedScope = NormalizeScope(p_Scope, p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000

    invalidationOrder = GetInvalidationOrder(normalizedScope, p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000

    Set resolver = New RiskTreeSelectionResolver
    fallbackKey = resolver.ResolveFallbackKey( _
        p_RequestedKey, _
        p_CandidateKeys, _
        p_ExistingKeys, _
        p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000

    Set plan = New Scripting.Dictionary
    plan.CompareMode = TextCompare
    plan.Add "normalizedScope", normalizedScope
    plan.Add "invalidationOrder", invalidationOrder
    plan.Add "fullReload", IsFullReload(normalizedScope, p_Error)
    plan.Add "fallbackKey", fallbackKey
    plan.Add "requestedKey", Trim$(Nz(p_RequestedKey, vbNullString))
    plan.Add "candidateKeys", Nz(p_CandidateKeys, vbNullString)
    plan.Add "existingKeys", Nz(p_ExistingKeys, vbNullString)
    plan.Add "IDEdicion", Trim$(Nz(p_IDEdicion, vbNullString))
    plan.Add "IDRiesgo", Trim$(Nz(p_IDRiesgo, vbNullString))

    Set BuildPlan = plan
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeRefreshService.BuildPlan: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeRefreshService.BuildPlan", p_Error
End Function

Public Function NormalizeScope( _
    ByVal p_Scope As String, _
    Optional ByRef p_Error As String _
) As String
    Dim scope As String

    On Error GoTo EH

    p_Error = vbNullString
    scope = LCase$(Trim$(Nz(p_Scope, vbNullString)))

    Select Case scope
        Case "risk", "plan", "action", "action-delete", "full"
            NormalizeScope = scope
        Case Else
            NormalizeScope = "risk"
    End Select
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeRefreshService.NormalizeScope: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeRefreshService.NormalizeScope", p_Error
End Function

Public Function GetInvalidationOrder( _
    ByVal p_Scope As String, _
    Optional ByRef p_Error As String _
) As String
    Dim scope As String

    On Error GoTo EH

    p_Error = vbNullString
    scope = NormalizeScope(p_Scope, p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000

    Select Case scope
        Case "action", "action-delete", "plan"
            GetInvalidationOrder = "plan|risk"
        Case "full"
            GetInvalidationOrder = "plan|risk|edition"
        Case Else
            GetInvalidationOrder = "risk"
    End Select
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeRefreshService.GetInvalidationOrder: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeRefreshService.GetInvalidationOrder", p_Error
End Function

Public Function IsFullReload( _
    ByVal p_Scope As String, _
    Optional ByRef p_Error As String _
) As Boolean
    Dim scope As String

    On Error GoTo EH

    p_Error = vbNullString
    scope = NormalizeScope(p_Scope, p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000

    IsFullReload = (StrComp(scope, "full", vbTextCompare) = 0)
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeRefreshService.IsFullReload: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeRefreshService.IsFullReload", p_Error
End Function

Public Function ResolveManualRefreshScope( _
    Optional ByRef p_Error As String _
) As String
    On Error GoTo EH

    p_Error = vbNullString
    ResolveManualRefreshScope = "full"
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeRefreshService.ResolveManualRefreshScope: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeRefreshService.ResolveManualRefreshScope", p_Error
End Function

Public Function ResolveDetailRefreshRoute( _
    ByVal p_NodeResolved As Boolean, _
    Optional ByRef p_Error As String _
) As String
    On Error GoTo EH

    p_Error = vbNullString
    If p_NodeResolved Then
        ResolveDetailRefreshRoute = "risk"
    Else
        ResolveDetailRefreshRoute = "full-fallback"
    End If
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeRefreshService.ResolveDetailRefreshRoute: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeRefreshService.ResolveDetailRefreshRoute", p_Error
End Function

Public Function ResolveQualityRiskRefreshScope( _
    ByVal p_Mutation As String, _
    Optional ByRef p_Error As String _
) As String
    Dim mutation As String

    On Error GoTo EH

    p_Error = vbNullString
    mutation = LCase$(Trim$(Nz(p_Mutation, vbNullString)))

    Select Case mutation
        Case "aceptacion", "rechazo", "retiro", "materializacion", "calidad"
            ResolveQualityRiskRefreshScope = "full"
        Case Else
            ResolveQualityRiskRefreshScope = "risk"
    End Select
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeRefreshService.ResolveQualityRiskRefreshScope: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeRefreshService.ResolveQualityRiskRefreshScope", p_Error
End Function
