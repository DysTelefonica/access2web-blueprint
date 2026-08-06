Attribute VB_Name = "Test_GestionRiesgosEdicionDeleteDecisionService"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: GestionRiesgosEdicionDeleteDecisionService (issue #45)
' Pure logic only. No DB, no UI, no filesystem, no fixtures.
' Returns JSON {ok, value, payload, error, logs}
' ============================================================

Private Function GreddsRunPuedeIniciarBorrado( _
    ByVal p_TestName As String, _
    ByVal p_HayEdicionActiva As Boolean, _
    ByVal p_EsPrimeraEdicion As EnumSiNo, _
    ByVal p_EsTecnico As EnumSiNo, _
    ByVal p_Expected As Boolean, _
    ByVal p_ExpectedMessage As String _
) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: define active edition, first-edition and role flags"
    logs(1) = "2. Arrange: expected legacy delete precondition result"
    logs(2) = "3. Act: service.PuedeIniciarBorrado"
    logs(3) = "4. Assert: decision and message match legacy rule"

    Dim service As GestionRiesgosEdicionDeleteDecisionService
    Set service = New GestionRiesgosEdicionDeleteDecisionService

    Dim message As String
    Dim result As Boolean
    result = service.PuedeIniciarBorrado(p_HayEdicionActiva, p_EsPrimeraEdicion, p_EsTecnico, message)

    If result <> p_Expected Then
        GreddsRunPuedeIniciarBorrado = BuildJsonFail(p_TestName & ": expected " & CStr(p_Expected) & ", got " & CStr(result), logs)
        Exit Function
    End If
    If message <> p_ExpectedMessage Then
        GreddsRunPuedeIniciarBorrado = BuildJsonFail(p_TestName & ": expected message '" & p_ExpectedMessage & "', got '" & message & "'", logs)
        Exit Function
    End If

    GreddsRunPuedeIniciarBorrado = BuildJsonOk(p_TestName & "_ok", logs)
    Exit Function

EH:
    GreddsRunPuedeIniciarBorrado = BuildJsonFail(p_TestName & ": " & Err.Description, logs)
End Function

Private Function GreddsRunInformeAnteriorBloqueaBorrado( _
    ByVal p_TestName As String, _
    ByVal p_TieneDocumentoAnterior As Boolean, _
    ByVal p_ExisteArchivoInformeAnterior As Boolean, _
    ByVal p_InformeAnteriorAbierto As Boolean, _
    ByVal p_Expected As Boolean, _
    ByVal p_ExpectedMessage As String _
) As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: define previous report document/file/open flags"
    logs(1) = "2. Arrange: expected legacy delete blocking result"
    logs(2) = "3. Act: service.InformeAnteriorBloqueaBorrado"
    logs(3) = "4. Assert: decision and message match legacy rule"

    Dim service As GestionRiesgosEdicionDeleteDecisionService
    Set service = New GestionRiesgosEdicionDeleteDecisionService

    Dim message As String
    Dim result As Boolean
    result = service.InformeAnteriorBloqueaBorrado( _
            p_TieneDocumentoAnterior, _
            p_ExisteArchivoInformeAnterior, _
            p_InformeAnteriorAbierto, _
            message)

    If result <> p_Expected Then
        GreddsRunInformeAnteriorBloqueaBorrado = BuildJsonFail(p_TestName & ": expected " & CStr(p_Expected) & ", got " & CStr(result), logs)
        Exit Function
    End If
    If message <> p_ExpectedMessage Then
        GreddsRunInformeAnteriorBloqueaBorrado = BuildJsonFail(p_TestName & ": expected message '" & p_ExpectedMessage & "', got '" & message & "'", logs)
        Exit Function
    End If

    GreddsRunInformeAnteriorBloqueaBorrado = BuildJsonOk(p_TestName & "_ok", logs)
    Exit Function

EH:
    GreddsRunInformeAnteriorBloqueaBorrado = BuildJsonFail(p_TestName & ": " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_NoActiveEdition_BlocksDelete() As String
    Test_GestionRiesgosEdicionDeleteDecisionService_NoActiveEdition_BlocksDelete = _
        GreddsRunPuedeIniciarBorrado("no_active_edition_blocks_delete", False, EnumSiNo.No, EnumSiNo.No, False, "No se conoce la edición activa")
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_FirstEdition_BlocksDelete() As String
    Test_GestionRiesgosEdicionDeleteDecisionService_FirstEdition_BlocksDelete = _
        GreddsRunPuedeIniciarBorrado("first_edition_blocks_delete", True, EnumSiNo.Sí, EnumSiNo.No, False, "Es la primera edición y una gestión de riesgos al menos ha de tener una. Si lo desea, debe borrar la gestión de riesgos completa")
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_Tecnico_BlocksDelete() As String
    Test_GestionRiesgosEdicionDeleteDecisionService_Tecnico_BlocksDelete = _
        GreddsRunPuedeIniciarBorrado("tecnico_blocks_delete", True, EnumSiNo.No, EnumSiNo.Sí, False, "Desde el 17/12/2019 sólo un miembro de Calidad puede borrar una edición")
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_CalidadNonFirst_AllowsDelete() As String
    Test_GestionRiesgosEdicionDeleteDecisionService_CalidadNonFirst_AllowsDelete = _
        GreddsRunPuedeIniciarBorrado("calidad_non_first_allows_delete", True, EnumSiNo.No, EnumSiNo.No, True, vbNullString)
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_WithChanges_RequiresConfirmation() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: changes with previous edition exist"
    logs(1) = "2. Act: service.DebeConfirmarCambiosConEdicionAnterior"
    logs(2) = "3. Assert: confirmation is required"

    Dim service As GestionRiesgosEdicionDeleteDecisionService
    Set service = New GestionRiesgosEdicionDeleteDecisionService

    If Not service.DebeConfirmarCambiosConEdicionAnterior(True) Then
        Test_GestionRiesgosEdicionDeleteDecisionService_WithChanges_RequiresConfirmation = BuildJsonFail("with_changes_should_require_confirmation", logs)
        Exit Function
    End If

    Test_GestionRiesgosEdicionDeleteDecisionService_WithChanges_RequiresConfirmation = BuildJsonOk("with_changes_requires_confirmation_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosEdicionDeleteDecisionService_WithChanges_RequiresConfirmation = BuildJsonFail("with_changes_requires_confirmation: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_WithoutChanges_SkipsConfirmation() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: no changes with previous edition exist"
    logs(1) = "2. Act: service.DebeConfirmarCambiosConEdicionAnterior"
    logs(2) = "3. Assert: confirmation is skipped"

    Dim service As GestionRiesgosEdicionDeleteDecisionService
    Set service = New GestionRiesgosEdicionDeleteDecisionService

    If service.DebeConfirmarCambiosConEdicionAnterior(False) Then
        Test_GestionRiesgosEdicionDeleteDecisionService_WithoutChanges_SkipsConfirmation = BuildJsonFail("without_changes_should_skip_confirmation", logs)
        Exit Function
    End If

    Test_GestionRiesgosEdicionDeleteDecisionService_WithoutChanges_SkipsConfirmation = BuildJsonOk("without_changes_skips_confirmation_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosEdicionDeleteDecisionService_WithoutChanges_SkipsConfirmation = BuildJsonFail("without_changes_skips_confirmation: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_OpenPreviousReport_BlocksDelete() As String
    Test_GestionRiesgosEdicionDeleteDecisionService_OpenPreviousReport_BlocksDelete = _
        GreddsRunInformeAnteriorBloqueaBorrado("open_previous_report_blocks_delete", True, True, True, True, "Se tiene abierto el informe de la edición anterior")
End Function

Public Function Test_GestionRiesgosEdicionDeleteDecisionService_MissingPreviousReportFile_AllowsDelete() As String
    Test_GestionRiesgosEdicionDeleteDecisionService_MissingPreviousReportFile_AllowsDelete = _
        GreddsRunInformeAnteriorBloqueaBorrado("missing_previous_report_file_allows_delete", True, False, True, False, vbNullString)
End Function
