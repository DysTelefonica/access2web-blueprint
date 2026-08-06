Attribute VB_Name = "Test_RiesgoPlanActionChangeDetector"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: RiesgoPlanActionChangeDetector (issue #42)
' Pure logic: compare plan/action objects field-by-field.
' No DB, no UI, no fixtures.
' ============================================================

Private Const RPACD_BASE As String = "BASE"
Private Const RPACD_CHANGED As String = "CHANGED"
Private Const RPACD_DATE As String = "01/02/2026"
Private Const RPACD_DATE_WITH_TIME As String = "01/02/2026 18:45:00"
Private Const RPACD_DATE_CHANGED As String = "02/02/2026"
Private Const RPACD_NOT_DATE As String = "not-a-date"

Private Sub RpacdBuildPMPair(ByRef p_Original As PM, ByRef p_Current As PM)
    Set p_Original = New PM
    Set p_Current = New PM

    p_Original.DisparadorDelPlan = RPACD_BASE
    p_Current.DisparadorDelPlan = RPACD_BASE
End Sub

Private Sub RpacdBuildPCPair(ByRef p_Original As PC, ByRef p_Current As PC)
    Set p_Original = New PC
    Set p_Current = New PC

    p_Original.DisparadorDelPlan = RPACD_BASE
    p_Current.DisparadorDelPlan = RPACD_BASE
End Sub

Private Sub RpacdBuildPMActionPair(ByRef p_Original As PMAccion, ByRef p_Current As PMAccion)
    Set p_Original = New PMAccion
    Set p_Current = New PMAccion

    With p_Original
        .Accion = RPACD_BASE
        .ResponsableAccion = RPACD_BASE
        .FechaInicio = RPACD_DATE
        .FechaFinPrevista = RPACD_DATE
        .FechaFinReal = RPACD_DATE
    End With

    With p_Current
        .Accion = RPACD_BASE
        .ResponsableAccion = RPACD_BASE
        .FechaInicio = RPACD_DATE
        .FechaFinPrevista = RPACD_DATE
        .FechaFinReal = RPACD_DATE
    End With
End Sub

Private Sub RpacdBuildPCActionPair(ByRef p_Original As PCAccion, ByRef p_Current As PCAccion)
    Set p_Original = New PCAccion
    Set p_Current = New PCAccion

    With p_Original
        .Accion = RPACD_BASE
        .ResponsableAccion = RPACD_BASE
        .FechaInicio = RPACD_DATE
        .FechaFinPrevista = RPACD_DATE
        .FechaFinReal = RPACD_DATE
    End With

    With p_Current
        .Accion = RPACD_BASE
        .ResponsableAccion = RPACD_BASE
        .FechaInicio = RPACD_DATE
        .FechaFinPrevista = RPACD_DATE
        .FechaFinReal = RPACD_DATE
    End With
End Sub

Private Function RpacdRunPMActionChangedField(ByVal p_FieldName As String) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: equal pair of PMAccion objects"
    logs(1) = "2. Act: mutate single field: " & p_FieldName
    logs(2) = "3. Act: detector.HasPMActionChanges(original, current)"
    logs(3) = "4. Assert: returns True"

    Dim original As PMAccion
    Dim current As PMAccion
    RpacdBuildPMActionPair original, current

    Select Case p_FieldName
        Case "Accion": current.Accion = RPACD_CHANGED
        Case "ResponsableAccion": current.ResponsableAccion = RPACD_CHANGED
        Case "FechaInicio": current.FechaInicio = RPACD_DATE_CHANGED
        Case "FechaFinPrevista": current.FechaFinPrevista = RPACD_DATE_CHANGED
        Case "FechaFinReal": current.FechaFinReal = RPACD_DATE_CHANGED
        Case Else
            RpacdRunPMActionChangedField = BuildJsonFail("unknown field: " & p_FieldName, logs)
            Exit Function
    End Select

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If Not detector.HasPMActionChanges(original, current, errMsg) Then
        RpacdRunPMActionChangedField = BuildJsonFail("changed field not detected: " & p_FieldName, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        RpacdRunPMActionChangedField = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    RpacdRunPMActionChangedField = BuildJsonOk("pm_action_field_changed", logs)
    Exit Function

EH:
    RpacdRunPMActionChangedField = BuildJsonFail("RpacdRunPMActionChangedField: " & Err.Description, logs)
End Function

Private Function RpacdRunPCActionChangedField(ByVal p_FieldName As String) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: equal pair of PCAccion objects"
    logs(1) = "2. Act: mutate single field: " & p_FieldName
    logs(2) = "3. Act: detector.HasPCActionChanges(original, current)"
    logs(3) = "4. Assert: returns True"

    Dim original As PCAccion
    Dim current As PCAccion
    RpacdBuildPCActionPair original, current

    Select Case p_FieldName
        Case "Accion": current.Accion = RPACD_CHANGED
        Case "ResponsableAccion": current.ResponsableAccion = RPACD_CHANGED
        Case "FechaInicio": current.FechaInicio = RPACD_DATE_CHANGED
        Case "FechaFinPrevista": current.FechaFinPrevista = RPACD_DATE_CHANGED
        Case "FechaFinReal": current.FechaFinReal = RPACD_DATE_CHANGED
        Case Else
            RpacdRunPCActionChangedField = BuildJsonFail("unknown field: " & p_FieldName, logs)
            Exit Function
    End Select

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If Not detector.HasPCActionChanges(original, current, errMsg) Then
        RpacdRunPCActionChangedField = BuildJsonFail("changed field not detected: " & p_FieldName, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        RpacdRunPCActionChangedField = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    RpacdRunPCActionChangedField = BuildJsonOk("pc_action_field_changed", logs)
    Exit Function

EH:
    RpacdRunPCActionChangedField = BuildJsonFail("RpacdRunPCActionChangedField: " & Err.Description, logs)
End Function

Private Function RpacdRunPMActionDateSemantics( _
    ByVal p_OriginalFechaInicio As String, _
    ByVal p_CurrentFechaInicio As String, _
    ByVal p_ExpectedChanged As Boolean, _
    ByVal p_CaseName As String _
) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: PMAccion date values for legacy semantics"
    logs(1) = "2. Act: detector.HasPMActionChanges(original, current)"
    logs(2) = "3. Assert: change result matches legacy Format(dd/mm/yyyy) behavior"
    logs(3) = "4. Assert: p_Error remains empty"

    Dim original As PMAccion
    Dim current As PMAccion
    RpacdBuildPMActionPair original, current
    original.FechaInicio = p_OriginalFechaInicio
    current.FechaInicio = p_CurrentFechaInicio

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    Dim changed As Boolean
    changed = detector.HasPMActionChanges(original, current, errMsg)

    If changed <> p_ExpectedChanged Then
        RpacdRunPMActionDateSemantics = BuildJsonFail("unexpected PM date change result: " & p_CaseName, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        RpacdRunPMActionDateSemantics = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    RpacdRunPMActionDateSemantics = BuildJsonOk("pm_action_date_semantics", logs)
    Exit Function

EH:
    RpacdRunPMActionDateSemantics = BuildJsonFail("RpacdRunPMActionDateSemantics: " & Err.Description, logs)
End Function

Private Function RpacdRunPCActionDateSemantics( _
    ByVal p_OriginalFechaInicio As String, _
    ByVal p_CurrentFechaInicio As String, _
    ByVal p_ExpectedChanged As Boolean, _
    ByVal p_CaseName As String _
) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: PCAccion date values for legacy semantics"
    logs(1) = "2. Act: detector.HasPCActionChanges(original, current)"
    logs(2) = "3. Assert: change result matches legacy Format(dd/mm/yyyy) behavior"
    logs(3) = "4. Assert: p_Error remains empty"

    Dim original As PCAccion
    Dim current As PCAccion
    RpacdBuildPCActionPair original, current
    original.FechaInicio = p_OriginalFechaInicio
    current.FechaInicio = p_CurrentFechaInicio

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    Dim changed As Boolean
    changed = detector.HasPCActionChanges(original, current, errMsg)

    If changed <> p_ExpectedChanged Then
        RpacdRunPCActionDateSemantics = BuildJsonFail("unexpected PC date change result: " & p_CaseName, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        RpacdRunPCActionDateSemantics = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    RpacdRunPCActionDateSemantics = BuildJsonOk("pc_action_date_semantics", logs)
    Exit Function

EH:
    RpacdRunPCActionDateSemantics = BuildJsonFail("RpacdRunPCActionDateSemantics: " & Err.Description, logs)
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMNoChanges_ReturnsFalse() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of PM objects"
    logs(1) = "2. Act: detector.HasPMChanges(original, current)"
    logs(2) = "3. Assert: returns False"

    Dim original As PM
    Dim current As PM
    RpacdBuildPMPair original, current

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If detector.HasPMChanges(original, current, errMsg) Then
        Test_RiesgoPlanActionChangeDetector_PMNoChanges_ReturnsFalse = BuildJsonFail("PM equal pair must return False", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_RiesgoPlanActionChangeDetector_PMNoChanges_ReturnsFalse = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    Test_RiesgoPlanActionChangeDetector_PMNoChanges_ReturnsFalse = BuildJsonOk("pm_no_changes", logs)
    Exit Function

EH:
    Test_RiesgoPlanActionChangeDetector_PMNoChanges_ReturnsFalse = BuildJsonFail("Test_RiesgoPlanActionChangeDetector_PMNoChanges_ReturnsFalse: " & Err.Description, logs)
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCNoChanges_ReturnsFalse() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of PC objects"
    logs(1) = "2. Act: detector.HasPCChanges(original, current)"
    logs(2) = "3. Assert: returns False"

    Dim original As PC
    Dim current As PC
    RpacdBuildPCPair original, current

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If detector.HasPCChanges(original, current, errMsg) Then
        Test_RiesgoPlanActionChangeDetector_PCNoChanges_ReturnsFalse = BuildJsonFail("PC equal pair must return False", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_RiesgoPlanActionChangeDetector_PCNoChanges_ReturnsFalse = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    Test_RiesgoPlanActionChangeDetector_PCNoChanges_ReturnsFalse = BuildJsonOk("pc_no_changes", logs)
    Exit Function

EH:
    Test_RiesgoPlanActionChangeDetector_PCNoChanges_ReturnsFalse = BuildJsonFail("Test_RiesgoPlanActionChangeDetector_PCNoChanges_ReturnsFalse: " & Err.Description, logs)
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMDisparadorChanged_ReturnsTrue() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of PM objects"
    logs(1) = "2. Act: mutate DisparadorDelPlan"
    logs(2) = "3. Assert: detector returns True"

    Dim original As PM
    Dim current As PM
    RpacdBuildPMPair original, current
    current.DisparadorDelPlan = RPACD_CHANGED

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If Not detector.HasPMChanges(original, current, errMsg) Then
        Test_RiesgoPlanActionChangeDetector_PMDisparadorChanged_ReturnsTrue = BuildJsonFail("PM DisparadorDelPlan change not detected", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_RiesgoPlanActionChangeDetector_PMDisparadorChanged_ReturnsTrue = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    Test_RiesgoPlanActionChangeDetector_PMDisparadorChanged_ReturnsTrue = BuildJsonOk("pm_disparador_changed", logs)
    Exit Function

EH:
    Test_RiesgoPlanActionChangeDetector_PMDisparadorChanged_ReturnsTrue = BuildJsonFail("Test_RiesgoPlanActionChangeDetector_PMDisparadorChanged_ReturnsTrue: " & Err.Description, logs)
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCDisparadorChanged_ReturnsTrue() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of PC objects"
    logs(1) = "2. Act: mutate DisparadorDelPlan"
    logs(2) = "3. Assert: detector returns True"

    Dim original As PC
    Dim current As PC
    RpacdBuildPCPair original, current
    current.DisparadorDelPlan = RPACD_CHANGED

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If Not detector.HasPCChanges(original, current, errMsg) Then
        Test_RiesgoPlanActionChangeDetector_PCDisparadorChanged_ReturnsTrue = BuildJsonFail("PC DisparadorDelPlan change not detected", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_RiesgoPlanActionChangeDetector_PCDisparadorChanged_ReturnsTrue = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    Test_RiesgoPlanActionChangeDetector_PCDisparadorChanged_ReturnsTrue = BuildJsonOk("pc_disparador_changed", logs)
    Exit Function

EH:
    Test_RiesgoPlanActionChangeDetector_PCDisparadorChanged_ReturnsTrue = BuildJsonFail("Test_RiesgoPlanActionChangeDetector_PCDisparadorChanged_ReturnsTrue: " & Err.Description, logs)
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionNoChanges_ReturnsFalse() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of PMAccion objects"
    logs(1) = "2. Act: detector.HasPMActionChanges(original, current)"
    logs(2) = "3. Assert: returns False"

    Dim original As PMAccion
    Dim current As PMAccion
    RpacdBuildPMActionPair original, current

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If detector.HasPMActionChanges(original, current, errMsg) Then
        Test_RiesgoPlanActionChangeDetector_PMActionNoChanges_ReturnsFalse = BuildJsonFail("PMAccion equal pair must return False", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_RiesgoPlanActionChangeDetector_PMActionNoChanges_ReturnsFalse = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    Test_RiesgoPlanActionChangeDetector_PMActionNoChanges_ReturnsFalse = BuildJsonOk("pm_action_no_changes", logs)
    Exit Function

EH:
    Test_RiesgoPlanActionChangeDetector_PMActionNoChanges_ReturnsFalse = BuildJsonFail("Test_RiesgoPlanActionChangeDetector_PMActionNoChanges_ReturnsFalse: " & Err.Description, logs)
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionNoChanges_ReturnsFalse() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of PCAccion objects"
    logs(1) = "2. Act: detector.HasPCActionChanges(original, current)"
    logs(2) = "3. Assert: returns False"

    Dim original As PCAccion
    Dim current As PCAccion
    RpacdBuildPCActionPair original, current

    Dim detector As RiesgoPlanActionChangeDetector
    Set detector = New RiesgoPlanActionChangeDetector

    Dim errMsg As String
    If detector.HasPCActionChanges(original, current, errMsg) Then
        Test_RiesgoPlanActionChangeDetector_PCActionNoChanges_ReturnsFalse = BuildJsonFail("PCAccion equal pair must return False", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_RiesgoPlanActionChangeDetector_PCActionNoChanges_ReturnsFalse = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    Test_RiesgoPlanActionChangeDetector_PCActionNoChanges_ReturnsFalse = BuildJsonOk("pc_action_no_changes", logs)
    Exit Function

EH:
    Test_RiesgoPlanActionChangeDetector_PCActionNoChanges_ReturnsFalse = BuildJsonFail("Test_RiesgoPlanActionChangeDetector_PCActionNoChanges_ReturnsFalse: " & Err.Description, logs)
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionAccionChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PMActionAccionChanged_ReturnsTrue = RpacdRunPMActionChangedField("Accion")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionResponsableChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PMActionResponsableChanged_ReturnsTrue = RpacdRunPMActionChangedField("ResponsableAccion")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionFechaInicioChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PMActionFechaInicioChanged_ReturnsTrue = RpacdRunPMActionChangedField("FechaInicio")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionFechaFinPrevistaChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PMActionFechaFinPrevistaChanged_ReturnsTrue = RpacdRunPMActionChangedField("FechaFinPrevista")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionFechaFinRealChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PMActionFechaFinRealChanged_ReturnsTrue = RpacdRunPMActionChangedField("FechaFinReal")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionAccionChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PCActionAccionChanged_ReturnsTrue = RpacdRunPCActionChangedField("Accion")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionResponsableChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PCActionResponsableChanged_ReturnsTrue = RpacdRunPCActionChangedField("ResponsableAccion")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionFechaInicioChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PCActionFechaInicioChanged_ReturnsTrue = RpacdRunPCActionChangedField("FechaInicio")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionFechaFinPrevistaChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PCActionFechaFinPrevistaChanged_ReturnsTrue = RpacdRunPCActionChangedField("FechaFinPrevista")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionFechaFinRealChanged_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PCActionFechaFinRealChanged_ReturnsTrue = RpacdRunPCActionChangedField("FechaFinReal")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionSameDateDifferentTime_ReturnsFalse() As String
    Test_RiesgoPlanActionChangeDetector_PMActionSameDateDifferentTime_ReturnsFalse = _
        RpacdRunPMActionDateSemantics(RPACD_DATE, RPACD_DATE_WITH_TIME, False, "same date different time")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionSameDateDifferentTime_ReturnsFalse() As String
    Test_RiesgoPlanActionChangeDetector_PCActionSameDateDifferentTime_ReturnsFalse = _
        RpacdRunPCActionDateSemantics(RPACD_DATE, RPACD_DATE_WITH_TIME, False, "same date different time")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionValidDateVsBlank_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PMActionValidDateVsBlank_ReturnsTrue = _
        RpacdRunPMActionDateSemantics(RPACD_DATE, vbNullString, True, "valid date vs blank")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionValidDateVsBlank_ReturnsTrue() As String
    Test_RiesgoPlanActionChangeDetector_PCActionValidDateVsBlank_ReturnsTrue = _
        RpacdRunPCActionDateSemantics(RPACD_DATE, vbNullString, True, "valid date vs blank")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PMActionBlankVsNonDate_ReturnsFalse() As String
    Test_RiesgoPlanActionChangeDetector_PMActionBlankVsNonDate_ReturnsFalse = _
        RpacdRunPMActionDateSemantics(vbNullString, RPACD_NOT_DATE, False, "blank vs non-date")
End Function

Public Function Test_RiesgoPlanActionChangeDetector_PCActionBlankVsNonDate_ReturnsFalse() As String
    Test_RiesgoPlanActionChangeDetector_PCActionBlankVsNonDate_ReturnsFalse = _
        RpacdRunPCActionDateSemantics(vbNullString, RPACD_NOT_DATE, False, "blank vs non-date")
End Function
