VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoPriorizacionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Issue #36 service seam for validating a complete proposed prioritization.
Public Function ValidatePriorities( _
    ByVal p_Priorities As Scripting.Dictionary, _
    ByVal p_RiskCount As Long, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    Dim seenPriorities As Scripting.Dictionary
    Dim key As Variant
    Dim rawPriority As String
    Dim numericPriority As Double
    Dim priorityValue As Long
    Dim rangeMessage As String

    p_Error = vbNullString
    rangeMessage = "Se ha de incluir un número entre el 1 y el " & CStr(p_RiskCount)

    If p_RiskCount < 1 Then
        p_Error = rangeMessage
        ValidatePriorities = False
        Exit Function
    End If

    If p_Priorities Is Nothing Then
        p_Error = rangeMessage
        ValidatePriorities = False
        Exit Function
    End If

    If p_Priorities.Count <> p_RiskCount Then
        p_Error = rangeMessage
        ValidatePriorities = False
        Exit Function
    End If

    Set seenPriorities = New Scripting.Dictionary
    seenPriorities.CompareMode = TextCompare

    For Each key In p_Priorities.Keys
        If IsNull(p_Priorities(key)) Then
            p_Error = rangeMessage
            ValidatePriorities = False
            Exit Function
        End If

        rawPriority = Trim$(CStr(p_Priorities(key)))
        If rawPriority = vbNullString Then
            p_Error = rangeMessage
            ValidatePriorities = False
            Exit Function
        End If

        If Not IsNumeric(rawPriority) Then
            p_Error = rangeMessage
            ValidatePriorities = False
            Exit Function
        End If

        numericPriority = CDbl(rawPriority)
        If numericPriority <> Fix(numericPriority) Then
            p_Error = rangeMessage
            ValidatePriorities = False
            Exit Function
        End If

        priorityValue = CLng(numericPriority)
        If priorityValue < 1 Or priorityValue > p_RiskCount Then
            p_Error = rangeMessage
            ValidatePriorities = False
            Exit Function
        End If

        If seenPriorities.Exists(CStr(priorityValue)) Then
            p_Error = "La priorización " & CStr(priorityValue) & " está repetida"
            ValidatePriorities = False
            Exit Function
        End If

        seenPriorities.Add CStr(priorityValue), True
    Next key

    ValidatePriorities = True
    Exit Function

EH:
    If p_Error = "" Then
        p_Error = "RiesgoPriorizacionService.ValidatePriorities: " & Err.Description
    End If
    ValidatePriorities = False
End Function

' Issue #44 pure seam for calculating which risks changed priority between
' the initial snapshot and the final proposed prioritization.
Public Function GetChangedRiskIds( _
    ByVal p_InitialPriorities As Scripting.Dictionary, _
    ByVal p_FinalPriorities As Scripting.Dictionary, _
    Optional ByRef p_Error As String _
) As Scripting.Dictionary
    On Error GoTo EH

    Dim changedRiskIds As Scripting.Dictionary
    Dim key As Variant
    Dim riskId As String
    Dim initialPriority As String
    Dim finalPriority As String

    p_Error = vbNullString

    If p_InitialPriorities Is Nothing Then
        p_Error = "No se estableció el listado inicial"
        Exit Function
    End If
    If p_FinalPriorities Is Nothing Then
        p_Error = "No se estableció el listado final"
        Exit Function
    End If

    Set changedRiskIds = New Scripting.Dictionary
    changedRiskIds.CompareMode = TextCompare

    For Each key In p_FinalPriorities.Keys
        riskId = CStr(key)
        If Not p_InitialPriorities.Exists(riskId) Then
            p_Error = "No se encontró la priorización inicial del riesgo " & riskId
            Set changedRiskIds = Nothing
            Exit Function
        End If

        initialPriority = CStr(p_InitialPriorities(riskId))
        finalPriority = CStr(p_FinalPriorities(riskId))
        If finalPriority <> initialPriority Then
            changedRiskIds.Add riskId, finalPriority
        End If
    Next key

    Set GetChangedRiskIds = changedRiskIds
    Exit Function

EH:
    If p_Error = "" Then
        p_Error = "RiesgoPriorizacionService.GetChangedRiskIds: " & Err.Description
    End If
    Set GetChangedRiskIds = Nothing
End Function
