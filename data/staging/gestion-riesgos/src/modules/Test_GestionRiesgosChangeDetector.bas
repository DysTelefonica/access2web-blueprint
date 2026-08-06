Attribute VB_Name = "Test_GestionRiesgosChangeDetector"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: GestionRiesgosChangeDetector (issue #41)
' Pure logic: compare two Proyecto objects field-by-field.
' No DB, no UI, no fixtures. Schema lives in Proyecto.cls.
' Returns JSON {ok, value, payload, error, logs}
' ============================================================

Private Const GRCD_VAL As String = "X"

Private Function GrcdBuildEqualPair(ByRef p_Original As Proyecto, ByRef p_Current As Proyecto) As Boolean
    On Error GoTo EH

    Set p_Original = New Proyecto
    Set p_Current = New Proyecto

    With p_Original
        .ParaInformeAvisos = GRCD_VAL
        .EnUTE = GRCD_VAL
        .NombreUsuarioCalidad = GRCD_VAL
        .CorreoRAC = GRCD_VAL
        .Elaborado = GRCD_VAL
        .Revisado = GRCD_VAL
        .Aprobado = GRCD_VAL
    End With

    With p_Current
        .ParaInformeAvisos = GRCD_VAL
        .EnUTE = GRCD_VAL
        .NombreUsuarioCalidad = GRCD_VAL
        .CorreoRAC = GRCD_VAL
        .Elaborado = GRCD_VAL
        .Revisado = GRCD_VAL
        .Aprobado = GRCD_VAL
    End With

    GrcdBuildEqualPair = True
    Exit Function

EH:
    GrcdBuildEqualPair = False
End Function

Private Function GrcdRunChangedField(ByVal p_FieldName As String) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: equal pair of Proyecto objects"
    logs(1) = "2. Act: mutate single field: " & p_FieldName
    logs(2) = "3. Act: detector.HaHabidoCambios(original, current)"
    logs(3) = "4. Assert: returns True (single-field change detected)"

    Dim original As Proyecto
    Dim current As Proyecto
    If Not GrcdBuildEqualPair(original, current) Then
        GrcdRunChangedField = BuildJsonFail("setup failed: equal pair could not be created", logs)
        Exit Function
    End If

    Select Case p_FieldName
        Case "ParaInformeAvisos": current.ParaInformeAvisos = "Y"
        Case "EnUTE": current.EnUTE = "Y"
        Case "NombreUsuarioCalidad": current.NombreUsuarioCalidad = "Y"
        Case "CorreoRAC": current.CorreoRAC = "Y"
        Case "Elaborado": current.Elaborado = "Y"
        Case "Revisado": current.Revisado = "Y"
        Case "Aprobado": current.Aprobado = "Y"
        Case Else
            GrcdRunChangedField = BuildJsonFail("unknown field: " & p_FieldName, logs)
            Exit Function
    End Select

    Dim detector As GestionRiesgosChangeDetector
    Set detector = New GestionRiesgosChangeDetector

    Dim errMsg As String
    Dim changed As Boolean
    changed = detector.HaHabidoCambios(original, current, errMsg)
    If errMsg <> "" Then
        GrcdRunChangedField = BuildJsonFail("HaHabidoCambios raised error: " & errMsg, logs)
        Exit Function
    End If
    If Not changed Then
        GrcdRunChangedField = BuildJsonFail("HaHabidoCambios returned False for changed field " & p_FieldName, logs)
        Exit Function
    End If

    GrcdRunChangedField = BuildJsonOk("field_changed_detected", logs)
    Exit Function

EH:
    GrcdRunChangedField = BuildJsonFail("GrcdRunChangedField: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosChangeDetector_InitialNothing_ReturnsTrue() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: initial project is Nothing"
    logs(1) = "2. Act: detector.HaHabidoCambios(Nothing, Nothing)"
    logs(2) = "3. Assert: returns True before reading active project"

    Dim detector As GestionRiesgosChangeDetector
    Set detector = New GestionRiesgosChangeDetector

    Dim errMsg As String
    If Not detector.HaHabidoCambios(Nothing, Nothing, errMsg) Then
        Test_GestionRiesgosChangeDetector_InitialNothing_ReturnsTrue = BuildJsonFail("initial Nothing must return True", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_GestionRiesgosChangeDetector_InitialNothing_ReturnsTrue = BuildJsonFail("initial Nothing must not surface error: " & errMsg, logs)
        Exit Function
    End If

    Test_GestionRiesgosChangeDetector_InitialNothing_ReturnsTrue = BuildJsonOk("initial_nothing_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosChangeDetector_InitialNothing_ReturnsTrue = BuildJsonFail("Test_GestionRiesgosChangeDetector_InitialNothing_ReturnsTrue: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosChangeDetector_NoChanges_ReturnsFalse() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of Proyecto objects"
    logs(1) = "2. Act: detector.HaHabidoCambios(original, current)"
    logs(2) = "3. Assert: returns False (no field changed)"

    Dim original As Proyecto
    Dim current As Proyecto
    If Not GrcdBuildEqualPair(original, current) Then
        Test_GestionRiesgosChangeDetector_NoChanges_ReturnsFalse = BuildJsonFail("setup failed", logs)
        Exit Function
    End If

    Dim detector As GestionRiesgosChangeDetector
    Set detector = New GestionRiesgosChangeDetector

    Dim errMsg As String
    If detector.HaHabidoCambios(original, current, errMsg) Then
        Test_GestionRiesgosChangeDetector_NoChanges_ReturnsFalse = BuildJsonFail("identical pair must return False", logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_GestionRiesgosChangeDetector_NoChanges_ReturnsFalse = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If

    Test_GestionRiesgosChangeDetector_NoChanges_ReturnsFalse = BuildJsonOk("no_changes_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosChangeDetector_NoChanges_ReturnsFalse = BuildJsonFail("Test_GestionRiesgosChangeDetector_NoChanges_ReturnsFalse: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosChangeDetector_ActiveNothing_ReturnsError() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: initial project exists, active project is Nothing"
    logs(1) = "2. Act: detector.HaHabidoCambios(original, Nothing)"
    logs(2) = "3. Assert: returns False and populates original error style"

    Dim original As Proyecto
    Dim current As Proyecto
    If Not GrcdBuildEqualPair(original, current) Then
        Test_GestionRiesgosChangeDetector_ActiveNothing_ReturnsError = BuildJsonFail("setup failed", logs)
        Exit Function
    End If
    Set current = Nothing

    Dim detector As GestionRiesgosChangeDetector
    Set detector = New GestionRiesgosChangeDetector

    Dim errMsg As String
    Dim changed As Boolean
    changed = detector.HaHabidoCambios(original, current, errMsg)
    If changed Then
        Test_GestionRiesgosChangeDetector_ActiveNothing_ReturnsError = BuildJsonFail("active Nothing with existing initial must not return True", logs)
        Exit Function
    End If
    If InStr(1, errMsg, "Error en el método HaHabidoCambios", vbTextCompare) = 0 Then
        Test_GestionRiesgosChangeDetector_ActiveNothing_ReturnsError = BuildJsonFail("expected original error prefix, got: " & errMsg, logs)
        Exit Function
    End If

    Test_GestionRiesgosChangeDetector_ActiveNothing_ReturnsError = BuildJsonOk("active_nothing_error_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosChangeDetector_ActiveNothing_ReturnsError = BuildJsonFail("Test_GestionRiesgosChangeDetector_ActiveNothing_ReturnsError: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosChangeDetector_ParaInformeAvisosChanged_ReturnsTrue() As String
    Test_GestionRiesgosChangeDetector_ParaInformeAvisosChanged_ReturnsTrue = GrcdRunChangedField("ParaInformeAvisos")
End Function

Public Function Test_GestionRiesgosChangeDetector_EnUTEChanged_ReturnsTrue() As String
    Test_GestionRiesgosChangeDetector_EnUTEChanged_ReturnsTrue = GrcdRunChangedField("EnUTE")
End Function

Public Function Test_GestionRiesgosChangeDetector_NombreUsuarioCalidadChanged_ReturnsTrue() As String
    Test_GestionRiesgosChangeDetector_NombreUsuarioCalidadChanged_ReturnsTrue = GrcdRunChangedField("NombreUsuarioCalidad")
End Function

Public Function Test_GestionRiesgosChangeDetector_CorreoRACChanged_ReturnsTrue() As String
    Test_GestionRiesgosChangeDetector_CorreoRACChanged_ReturnsTrue = GrcdRunChangedField("CorreoRAC")
End Function

Public Function Test_GestionRiesgosChangeDetector_ElaboradoChanged_ReturnsTrue() As String
    Test_GestionRiesgosChangeDetector_ElaboradoChanged_ReturnsTrue = GrcdRunChangedField("Elaborado")
End Function

Public Function Test_GestionRiesgosChangeDetector_RevisadoChanged_ReturnsTrue() As String
    Test_GestionRiesgosChangeDetector_RevisadoChanged_ReturnsTrue = GrcdRunChangedField("Revisado")
End Function

Public Function Test_GestionRiesgosChangeDetector_AprobadoChanged_ReturnsTrue() As String
    Test_GestionRiesgosChangeDetector_AprobadoChanged_ReturnsTrue = GrcdRunChangedField("Aprobado")
End Function
