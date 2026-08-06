Attribute VB_Name = "Test_GestionRiesgosSaveDecisionService"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: GestionRiesgosSaveDecisionService (issue #41)
' Pure logic only. No DB, no UI, no fixtures.
' Returns JSON {ok, value, payload, error, logs}
' ============================================================

Private Function GrsdsBuildProyecto(ByVal p_ParaInformeAvisos As String) As Proyecto
    Dim proyecto As Proyecto
    Set proyecto = New Proyecto
    proyecto.ParaInformeAvisos = p_ParaInformeAvisos
    Set GrsdsBuildProyecto = proyecto
End Function

Private Function GrsdsRunDebeAvisarNoInforme( _
    ByVal p_TestName As String, _
    ByVal p_SinEco As EnumSiNo, _
    ByVal p_InitialValue As String, _
    ByVal p_UseInitial As Boolean, _
    ByVal p_ActiveValue As String, _
    ByVal p_Expected As Boolean _
) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: build initial/active Proyecto objects"
    logs(1) = "2. Arrange: p_SinEco and ParaInformeAvisos scenario"
    logs(2) = "3. Act: service.DebeAvisarNoInforme"
    logs(3) = "4. Assert: decision matches legacy warning rule"

    Dim initial As Proyecto
    Dim active As Proyecto
    If p_UseInitial Then
        Set initial = GrsdsBuildProyecto(p_InitialValue)
    End If
    Set active = GrsdsBuildProyecto(p_ActiveValue)

    Dim service As GestionRiesgosSaveDecisionService
    Set service = New GestionRiesgosSaveDecisionService

    Dim errMsg As String
    Dim result As Boolean
    result = service.DebeAvisarNoInforme(p_SinEco, initial, active, errMsg)

    If errMsg <> "" Then
        GrsdsRunDebeAvisarNoInforme = BuildJsonFail(p_TestName & ": unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If result <> p_Expected Then
        GrsdsRunDebeAvisarNoInforme = BuildJsonFail(p_TestName & ": expected " & CStr(p_Expected) & ", got " & CStr(result), logs)
        Exit Function
    End If

    GrsdsRunDebeAvisarNoInforme = BuildJsonOk(p_TestName & "_ok", logs)
    Exit Function

EH:
    GrsdsRunDebeAvisarNoInforme = BuildJsonFail(p_TestName & ": " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosSaveDecisionService_SinEcoSi_SuppressesWarning() As String
    Test_GestionRiesgosSaveDecisionService_SinEcoSi_SuppressesWarning = _
        GrsdsRunDebeAvisarNoInforme("sin_eco_si_suppresses_warning", EnumSiNo.Sí, "Sí", True, "No", False)
End Function

Public Function Test_GestionRiesgosSaveDecisionService_ActiveNoInitialNothing_ReturnsTrue() As String
    Test_GestionRiesgosSaveDecisionService_ActiveNoInitialNothing_ReturnsTrue = _
        GrsdsRunDebeAvisarNoInforme("active_no_initial_nothing", EnumSiNo.No, vbNullString, False, "No", True)
End Function

Public Function Test_GestionRiesgosSaveDecisionService_ActiveNoChangedFromSi_ReturnsTrue() As String
    Test_GestionRiesgosSaveDecisionService_ActiveNoChangedFromSi_ReturnsTrue = _
        GrsdsRunDebeAvisarNoInforme("active_no_changed_from_si", EnumSiNo.No, "Sí", True, "No", True)
End Function

Public Function Test_GestionRiesgosSaveDecisionService_ActiveNoUnchangedNo_ReturnsFalse() As String
    Test_GestionRiesgosSaveDecisionService_ActiveNoUnchangedNo_ReturnsFalse = _
        GrsdsRunDebeAvisarNoInforme("active_no_unchanged_no", EnumSiNo.No, "No", True, "No", False)
End Function

Public Function Test_GestionRiesgosSaveDecisionService_ActiveSi_ReturnsFalse() As String
    Test_GestionRiesgosSaveDecisionService_ActiveSi_ReturnsFalse = _
        GrsdsRunDebeAvisarNoInforme("active_si_returns_false", EnumSiNo.No, "No", True, "Sí", False)
End Function
