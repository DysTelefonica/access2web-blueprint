Attribute VB_Name = "Test_RiesgoChangeDetector"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: RiesgoChangeDetector (issue #38)
' Pure logic: compare two riesgo objects field-by-field.
' No DB, no UI, no fixtures. Schema lives in Riesgo.cls.
' Returns JSON {ok, value, payload, error, logs}
' Run all: dysflow.test_vba with Test_RiesgoChangeDetector_* procedures
' ============================================================

Private Const RC_VAL As String = "X"

Private Function RcBuildEqualPair(ByRef p_Original As riesgo, ByRef p_Current As riesgo) As Boolean
    On Error GoTo EH
    Set p_Original = New riesgo
    Set p_Current = New riesgo
    With p_Original
        .Priorizacion = RC_VAL
        .Descripcion = RC_VAL
        .CausaRaiz = RC_VAL
        .EntidadDetecta = RC_VAL
        .DetectadoPor = RC_VAL
        .FechaDetectado = RC_VAL
        .FechaMaterializado = RC_VAL
        .FechaRetirado = RC_VAL
        .Plazo = RC_VAL
        .Coste = RC_VAL
        .Calidad = RC_VAL
        .Vulnerabilidad = RC_VAL
        .Mitigacion = RC_VAL
        .JustificacionAceptacionRiesgo = RC_VAL
        .JustificacionRetiroRiesgo = RC_VAL
        .CodRiesgoBiblioteca = RC_VAL
    End With
    With p_Current
        .Priorizacion = RC_VAL
        .Descripcion = RC_VAL
        .CausaRaiz = RC_VAL
        .EntidadDetecta = RC_VAL
        .DetectadoPor = RC_VAL
        .FechaDetectado = RC_VAL
        .FechaMaterializado = RC_VAL
        .FechaRetirado = RC_VAL
        .Plazo = RC_VAL
        .Coste = RC_VAL
        .Calidad = RC_VAL
        .Vulnerabilidad = RC_VAL
        .Mitigacion = RC_VAL
        .JustificacionAceptacionRiesgo = RC_VAL
        .JustificacionRetiroRiesgo = RC_VAL
        .CodRiesgoBiblioteca = RC_VAL
    End With
    RcBuildEqualPair = True
    Exit Function
EH:
    RcBuildEqualPair = False
End Function

Private Function RcRun(ByVal p_ChangedFieldSetter As String) As String
    On Error GoTo EH
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: equal pair of riesgo objects"
    logs(1) = "2. Act: mutate single field via setter: " & p_ChangedFieldSetter
    logs(2) = "3. Act: detector.HasChanges(original, current)"
    logs(3) = "4. Assert: returns True (single-field change detected)"

    Dim original As riesgo
    Dim current As riesgo
    If Not RcBuildEqualPair(original, current) Then
        RcRun = BuildJsonFail("setup failed: equal pair could not be created", logs)
        Exit Function
    End If

    Select Case p_ChangedFieldSetter
        Case "Priorizacion": current.Priorizacion = "Y"
        Case "Descripcion": current.Descripcion = "Y"
        Case "CausaRaiz": current.CausaRaiz = "Y"
        Case "EntidadDetecta": current.EntidadDetecta = "Y"
        Case "DetectadoPor": current.DetectadoPor = "Y"
        Case "FechaDetectado": current.FechaDetectado = "Y"
        Case "FechaMaterializado": current.FechaMaterializado = "Y"
        Case "FechaRetirado": current.FechaRetirado = "Y"
        Case "Plazo": current.Plazo = "Y"
        Case "Coste": current.Coste = "Y"
        Case "Calidad": current.Calidad = "Y"
        Case "Vulnerabilidad": current.Vulnerabilidad = "Y"
        Case "Mitigacion": current.Mitigacion = "Y"
        Case "JustificacionAceptacionRiesgo": current.JustificacionAceptacionRiesgo = "Y"
        Case "JustificacionRetiroRiesgo": current.JustificacionRetiroRiesgo = "Y"
        Case "CodRiesgoBiblioteca": current.CodRiesgoBiblioteca = "Y"
        Case Else
            RcRun = BuildJsonFail("unknown field: " & p_ChangedFieldSetter, logs)
            Exit Function
    End Select

    Dim detector As RiesgoChangeDetector
    Set detector = New RiesgoChangeDetector
    Dim errMsg As String
    Dim changed As Boolean
    changed = detector.HasChanges(original, current, errMsg)
    If errMsg <> "" Then
        RcRun = BuildJsonFail("HasChanges raised error: " & errMsg, logs)
        Exit Function
    End If
    If Not changed Then
        RcRun = BuildJsonFail("HasChanges returned False but field '" & p_ChangedFieldSetter & "' was mutated", logs)
        Exit Function
    End If

    RcRun = BuildJsonOk("field_changed_detected", logs)
    Exit Function
EH:
    RcRun = BuildJsonFail("RcRun: " & Err.Description, logs)
End Function

' --- Tests ---

Public Function Test_RiesgoChangeDetector_NoChanges_ReturnsFalse() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of riesgo objects"
    logs(1) = "2. Act: detector.HasChanges(original, current)"
    logs(2) = "3. Assert: returns False (no field changed)"

    Dim original As riesgo
    Dim current As riesgo
    If Not RcBuildEqualPair(original, current) Then
        Test_RiesgoChangeDetector_NoChanges_ReturnsFalse = BuildJsonFail("setup failed", logs)
        Exit Function
    End If

    Dim detector As RiesgoChangeDetector
    Set detector = New RiesgoChangeDetector
    Dim errMsg As String
    If detector.HasChanges(original, current, errMsg) Then
        Test_RiesgoChangeDetector_NoChanges_ReturnsFalse = BuildJsonFail("HasChanges must return False for identical pair", logs)
        Exit Function
    End If

    Test_RiesgoChangeDetector_NoChanges_ReturnsFalse = BuildJsonOk("no_changes_ok", logs)
    Exit Function
EH:
    Test_RiesgoChangeDetector_NoChanges_ReturnsFalse = BuildJsonFail("Test_RiesgoChangeDetector_NoChanges_ReturnsFalse: " & Err.Description, logs)
End Function

Public Function Test_RiesgoChangeDetector_NothingOriginal_RaisesError() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: original is Nothing, current is a real riesgo"
    logs(1) = "2. Act: detector.HasChanges(Nothing, current)"
    logs(2) = "3. Assert: error string is populated"

    Dim current As riesgo
    Set current = New riesgo

    Dim detector As RiesgoChangeDetector
    Set detector = New RiesgoChangeDetector
    Dim errMsg As String
    Dim changed As Boolean
    changed = detector.HasChanges(Nothing, current, errMsg)
    If errMsg = "" Then
        Test_RiesgoChangeDetector_NothingOriginal_RaisesError = BuildJsonFail("HasChanges must surface error when original is Nothing", logs)
        Exit Function
    End If
    If changed Then
        Test_RiesgoChangeDetector_NothingOriginal_RaisesError = BuildJsonFail("HasChanges must return False when error is reported", logs)
        Exit Function
    End If

    Test_RiesgoChangeDetector_NothingOriginal_RaisesError = BuildJsonOk("nothing_original_error_ok", logs)
    Exit Function
EH:
    Test_RiesgoChangeDetector_NothingOriginal_RaisesError = BuildJsonFail("Test_RiesgoChangeDetector_NothingOriginal_RaisesError: " & Err.Description, logs)
End Function

Public Function Test_RiesgoChangeDetector_NothingCurrent_RaisesError() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: original is a real riesgo, current is Nothing"
    logs(1) = "2. Act: detector.HasChanges(original, Nothing)"
    logs(2) = "3. Assert: error string is populated"

    Dim original As riesgo
    Set original = New riesgo

    Dim detector As RiesgoChangeDetector
    Set detector = New RiesgoChangeDetector
    Dim errMsg As String
    Dim changed As Boolean
    changed = detector.HasChanges(original, Nothing, errMsg)
    If errMsg = "" Then
        Test_RiesgoChangeDetector_NothingCurrent_RaisesError = BuildJsonFail("HasChanges must surface error when current is Nothing", logs)
        Exit Function
    End If
    If changed Then
        Test_RiesgoChangeDetector_NothingCurrent_RaisesError = BuildJsonFail("HasChanges must return False when error is reported", logs)
        Exit Function
    End If

    Test_RiesgoChangeDetector_NothingCurrent_RaisesError = BuildJsonOk("nothing_current_error_ok", logs)
    Exit Function
EH:
    Test_RiesgoChangeDetector_NothingCurrent_RaisesError = BuildJsonFail("Test_RiesgoChangeDetector_NothingCurrent_RaisesError: " & Err.Description, logs)
End Function

Public Function Test_RiesgoChangeDetector_PriorizacionChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_PriorizacionChanged_ReturnsTrue = RcRun("Priorizacion")
End Function

Public Function Test_RiesgoChangeDetector_DescripcionChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_DescripcionChanged_ReturnsTrue = RcRun("Descripcion")
End Function

Public Function Test_RiesgoChangeDetector_CausaRaizChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_CausaRaizChanged_ReturnsTrue = RcRun("CausaRaiz")
End Function

Public Function Test_RiesgoChangeDetector_EntidadDetectaChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_EntidadDetectaChanged_ReturnsTrue = RcRun("EntidadDetecta")
End Function

Public Function Test_RiesgoChangeDetector_DetectadoPorChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_DetectadoPorChanged_ReturnsTrue = RcRun("DetectadoPor")
End Function

Public Function Test_RiesgoChangeDetector_FechaDetectadoChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_FechaDetectadoChanged_ReturnsTrue = RcRun("FechaDetectado")
End Function

Public Function Test_RiesgoChangeDetector_FechaMaterializadoChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_FechaMaterializadoChanged_ReturnsTrue = RcRun("FechaMaterializado")
End Function

Public Function Test_RiesgoChangeDetector_FechaRetiradoChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_FechaRetiradoChanged_ReturnsTrue = RcRun("FechaRetirado")
End Function

Public Function Test_RiesgoChangeDetector_PlazoChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_PlazoChanged_ReturnsTrue = RcRun("Plazo")
End Function

Public Function Test_RiesgoChangeDetector_CosteChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_CosteChanged_ReturnsTrue = RcRun("Coste")
End Function

Public Function Test_RiesgoChangeDetector_CalidadChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_CalidadChanged_ReturnsTrue = RcRun("Calidad")
End Function

Public Function Test_RiesgoChangeDetector_VulnerabilidadChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_VulnerabilidadChanged_ReturnsTrue = RcRun("Vulnerabilidad")
End Function

Public Function Test_RiesgoChangeDetector_MitigacionChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_MitigacionChanged_ReturnsTrue = RcRun("Mitigacion")
End Function

Public Function Test_RiesgoChangeDetector_JustificacionAceptacionRiesgoChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_JustificacionAceptacionRiesgoChanged_ReturnsTrue = RcRun("JustificacionAceptacionRiesgo")
End Function

Public Function Test_RiesgoChangeDetector_JustificacionRetiroRiesgoChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_JustificacionRetiroRiesgoChanged_ReturnsTrue = RcRun("JustificacionRetiroRiesgo")
End Function

Public Function Test_RiesgoChangeDetector_CodRiesgoBibliotecaChanged_ReturnsTrue() As String
    Test_RiesgoChangeDetector_CodRiesgoBibliotecaChanged_ReturnsTrue = RcRun("CodRiesgoBiblioteca")
End Function

Public Function Test_RiesgoChangeDetector_MultipleFieldsChanged_ReturnsTrue() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: equal pair of riesgo objects"
    logs(1) = "2. Act: mutate 3 fields (Priorizacion, Descripcion, Plazo)"
    logs(2) = "3. Assert: HasChanges returns True (any field changed)"

    Dim original As riesgo
    Dim current As riesgo
    If Not RcBuildEqualPair(original, current) Then
        Test_RiesgoChangeDetector_MultipleFieldsChanged_ReturnsTrue = BuildJsonFail("setup failed", logs)
        Exit Function
    End If

    current.Priorizacion = "A"
    current.Descripcion = "B"
    current.Plazo = "C"

    Dim detector As RiesgoChangeDetector
    Set detector = New RiesgoChangeDetector
    Dim errMsg As String
    If Not detector.HasChanges(original, current, errMsg) Then
        Test_RiesgoChangeDetector_MultipleFieldsChanged_ReturnsTrue = BuildJsonFail("HasChanges must return True when multiple fields differ", logs)
        Exit Function
    End If

    Test_RiesgoChangeDetector_MultipleFieldsChanged_ReturnsTrue = BuildJsonOk("multiple_changes_ok", logs)
    Exit Function
EH:
    Test_RiesgoChangeDetector_MultipleFieldsChanged_ReturnsTrue = BuildJsonFail("Test_RiesgoChangeDetector_MultipleFieldsChanged_ReturnsTrue: " & Err.Description, logs)
End Function
