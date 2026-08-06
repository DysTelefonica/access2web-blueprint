VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiskTreeDataLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function PreloadForEdition( _
    ByVal p_IDEdicion As String, _
    ByVal p_Riesgos As Scripting.Dictionary, _
    Optional ByRef p_Error As String _
) As Scripting.Dictionary
    Dim result As Scripting.Dictionary
    Dim warnings As Collection
    Dim riskIds As Collection
    Dim pmIds As Collection
    Dim pcIds As Collection
    Dim dicPMsBatch As Scripting.Dictionary
    Dim dicPCsBatch As Scripting.Dictionary
    Dim dicPMAccionesBatch As Scripting.Dictionary
    Dim dicPCAccionesBatch As Scripting.Dictionary
    Dim batchError As String

    On Error GoTo EH

    p_Error = vbNullString
    Set warnings = New Collection
    Set result = New Scripting.Dictionary
    result.CompareMode = TextCompare

    Set riskIds = CollectRiskIds(p_Riesgos)
    If riskIds.Count = 0 Then
        AddResult result, True, warnings, 0, 0, 0, 0, 0
        Set PreloadForEdition = result
        Exit Function
    End If

    batchError = vbNullString
    On Error Resume Next
    Set dicPMsBatch = Constructor.getPMsPorEdicionBatch(p_IDEdicion, riskIds, batchError)
    If Err.Number <> 0 Then
        AddWarning warnings, "PM batch failed: " & Err.Description
        Err.Clear
        Set dicPMsBatch = Nothing
    End If
    On Error GoTo EH
    If batchError <> vbNullString Then
        AddWarning warnings, "PM batch failed: " & batchError
        Set dicPMsBatch = Nothing
    End If

    batchError = vbNullString
    On Error Resume Next
    Set dicPCsBatch = Constructor.getPCsPorEdicionBatch(p_IDEdicion, riskIds, batchError)
    If Err.Number <> 0 Then
        AddWarning warnings, "PC batch failed: " & Err.Description
        Err.Clear
        Set dicPCsBatch = Nothing
    End If
    On Error GoTo EH
    If batchError <> vbNullString Then
        AddWarning warnings, "PC batch failed: " & batchError
        Set dicPCsBatch = Nothing
    End If

    Set pmIds = New Collection
    Set pcIds = New Collection
    InjectPlans p_Riesgos, dicPMsBatch, dicPCsBatch, pmIds, pcIds

    batchError = vbNullString
    On Error Resume Next
    Set dicPMAccionesBatch = Constructor.getPMAccionesBatch(pmIds, batchError)
    If Err.Number <> 0 Then
        AddWarning warnings, "PM actions batch failed: " & Err.Description
        Err.Clear
        Set dicPMAccionesBatch = Nothing
    End If
    On Error GoTo EH
    If batchError <> vbNullString Then
        AddWarning warnings, "PM actions batch failed: " & batchError
        Set dicPMAccionesBatch = Nothing
    End If
    InjectPMActions dicPMsBatch, dicPMAccionesBatch

    batchError = vbNullString
    On Error Resume Next
    Set dicPCAccionesBatch = Constructor.getPCAccionesBatch(pcIds, batchError)
    If Err.Number <> 0 Then
        AddWarning warnings, "PC actions batch failed: " & Err.Description
        Err.Clear
        Set dicPCAccionesBatch = Nothing
    End If
    On Error GoTo EH
    If batchError <> vbNullString Then
        AddWarning warnings, "PC actions batch failed: " & batchError
        Set dicPCAccionesBatch = Nothing
    End If
    InjectPCActions dicPCsBatch, dicPCAccionesBatch

    AddResult result, True, warnings, riskIds.Count, CountNested(dicPMsBatch), _
        CountNested(dicPCsBatch), CountNested(dicPMAccionesBatch), CountNested(dicPCAccionesBatch)
    Set PreloadForEdition = result
    Exit Function

EH:
    p_Error = "RiskTreeDataLoader.PreloadForEdition: " & Err.Description
    Err.Raise Err.Number, "RiskTreeDataLoader.PreloadForEdition", p_Error
End Function

Private Function CollectRiskIds(ByVal p_Riesgos As Scripting.Dictionary) As Collection
    Dim ids As Collection
    Dim key As Variant
    Dim risk As riesgo

    Set ids = New Collection
    If p_Riesgos Is Nothing Then
        Set CollectRiskIds = ids
        Exit Function
    End If

    For Each key In p_Riesgos
        Set risk = p_Riesgos(key)
        If Not risk Is Nothing Then
            If Trim$(Nz(risk.IDRiesgo, vbNullString)) <> vbNullString Then
                ids.Add risk.IDRiesgo
            End If
        End If
        Set risk = Nothing
    Next key

    Set CollectRiskIds = ids
End Function

Private Sub InjectPlans( _
    ByVal p_Riesgos As Scripting.Dictionary, _
    ByVal p_DicPMsBatch As Scripting.Dictionary, _
    ByVal p_DicPCsBatch As Scripting.Dictionary, _
    ByVal p_PMIds As Collection, _
    ByVal p_PCIds As Collection _
)
    Dim riskId As Variant
    Dim planId As Variant
    Dim risk As riesgo
    Dim plans As Scripting.Dictionary

    If Not p_DicPMsBatch Is Nothing Then
        For Each riskId In p_DicPMsBatch
            If Not p_Riesgos Is Nothing Then
                If p_Riesgos.Exists(riskId) Then
                    Set risk = p_Riesgos(riskId)
                    Set plans = p_DicPMsBatch(riskId)
                    risk.InjectarColPMs plans
                    For Each planId In plans
                        p_PMIds.Add planId
                    Next planId
                End If
            End If
            Set plans = Nothing
            Set risk = Nothing
        Next riskId
    End If

    If Not p_DicPCsBatch Is Nothing Then
        For Each riskId In p_DicPCsBatch
            If Not p_Riesgos Is Nothing Then
                If p_Riesgos.Exists(riskId) Then
                    Set risk = p_Riesgos(riskId)
                    Set plans = p_DicPCsBatch(riskId)
                    risk.InjectarColPCs plans
                    For Each planId In plans
                        p_PCIds.Add planId
                    Next planId
                End If
            End If
            Set plans = Nothing
            Set risk = Nothing
        Next riskId
    End If
End Sub

Private Sub InjectPMActions( _
    ByVal p_DicPMsBatch As Scripting.Dictionary, _
    ByVal p_DicPMAccionesBatch As Scripting.Dictionary _
)
    Dim riskId As Variant
    Dim pmId As Variant
    Dim plans As Scripting.Dictionary
    Dim actions As Scripting.Dictionary
    Dim plan As pm

    If p_DicPMsBatch Is Nothing Or p_DicPMAccionesBatch Is Nothing Then Exit Sub

    For Each pmId In p_DicPMAccionesBatch
        Set actions = p_DicPMAccionesBatch(pmId)
        For Each riskId In p_DicPMsBatch
            Set plans = p_DicPMsBatch(riskId)
            If plans.Exists(pmId) Then
                Set plan = plans(pmId)
                plan.InjectarColAcciones actions
                Exit For
            End If
        Next riskId
        Set plan = Nothing
        Set actions = Nothing
        Set plans = Nothing
    Next pmId
End Sub

Private Sub InjectPCActions( _
    ByVal p_DicPCsBatch As Scripting.Dictionary, _
    ByVal p_DicPCAccionesBatch As Scripting.Dictionary _
)
    Dim riskId As Variant
    Dim pcId As Variant
    Dim plans As Scripting.Dictionary
    Dim actions As Scripting.Dictionary
    Dim plan As pc

    If p_DicPCsBatch Is Nothing Or p_DicPCAccionesBatch Is Nothing Then Exit Sub

    For Each pcId In p_DicPCAccionesBatch
        Set actions = p_DicPCAccionesBatch(pcId)
        For Each riskId In p_DicPCsBatch
            Set plans = p_DicPCsBatch(riskId)
            If plans.Exists(pcId) Then
                Set plan = plans(pcId)
                plan.InjectarColAcciones actions
                Exit For
            End If
        Next riskId
        Set plan = Nothing
        Set actions = Nothing
        Set plans = Nothing
    Next pcId
End Sub

Private Function CountNested(ByVal p_Dic As Scripting.Dictionary) As Long
    Dim key As Variant
    Dim child As Scripting.Dictionary

    If p_Dic Is Nothing Then Exit Function

    For Each key In p_Dic
        Set child = p_Dic(key)
        If Not child Is Nothing Then CountNested = CountNested + child.Count
        Set child = Nothing
    Next key
End Function

Private Sub AddWarning(ByVal p_Warnings As Collection, ByVal p_Message As String)
    If Trim$(Nz(p_Message, vbNullString)) <> vbNullString Then
        p_Warnings.Add p_Message
    End If
End Sub

Private Sub AddResult( _
    ByVal p_Result As Scripting.Dictionary, _
    ByVal p_Ok As Boolean, _
    ByVal p_Warnings As Collection, _
    ByVal p_RiskCount As Long, _
    ByVal p_PMCount As Long, _
    ByVal p_PCCount As Long, _
    ByVal p_PMActionCount As Long, _
    ByVal p_PCActionCount As Long _
)
    p_Result.RemoveAll
    p_Result.Add "ok", p_Ok
    p_Result.Add "warningCount", p_Warnings.Count
    p_Result.Add "warnings", JoinWarnings(p_Warnings)
    p_Result.Add "riskCount", p_RiskCount
    p_Result.Add "pmCount", p_PMCount
    p_Result.Add "pcCount", p_PCCount
    p_Result.Add "pmActionCount", p_PMActionCount
    p_Result.Add "pcActionCount", p_PCActionCount
End Sub

Private Function JoinWarnings(ByVal p_Warnings As Collection) As String
    Dim index As Long
    Dim joined As String

    For index = 1 To p_Warnings.Count
        If joined <> vbNullString Then joined = joined & " | "
        joined = joined & CStr(p_Warnings(index))
    Next index
    JoinWarnings = joined
End Function
