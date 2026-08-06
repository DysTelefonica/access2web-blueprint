Attribute VB_Name = "Test_LanzaderaSuite"
Option Compare Database
Option Explicit

Public Function Test_Lanzadera_RunAll() As String
    On Error GoTo EH

    Dim logs(0 To 10) As String
    Dim failures As New Collection
    Dim resultText As String
    Dim errMsg As String

    logs(0) = "1. SuiteSetup: ForceLocalBackend"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_Lanzadera_RunAll = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Manifest main atomic"
    resultText = Test_Manifest_MainAtomic_NoRunAll()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_Manifest_MainAtomic_NoRunAll", resultText)

    logs(2) = "3. Manifest smoke only runall"
    resultText = Test_Manifest_SmokeManifest_OnlyRunAll()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_Manifest_SmokeManifest_OnlyRunAll", resultText)

    logs(3) = "4. Manifest no duplicate procedures"
    resultText = Test_Manifest_NoDuplicateProcedures()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_Manifest_NoDuplicateProcedures", resultText)

    logs(4) = "5. Manifest slices unique names"
    resultText = Test_Manifest_Slices_NoDuplicateNames()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_Manifest_Slices_NoDuplicateNames", resultText)

    logs(5) = "6. Manifest slices source main"
    resultText = Test_Manifest_Slices_SourceManifest_IsMain()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_Manifest_Slices_SourceManifest_IsMain", resultText)

    logs(6) = "7. Backend bootstrap respects config"
    resultText = Test_BackendBootstrap_LeeConfiguracionLocal_RespetaBackendConfigurado()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendBootstrap_LeeConfiguracionLocal_RespetaBackendConfigurado", resultText)

    logs(7) = "8. Backend bootstrap force local with parameter"
    resultText = Test_BackendBootstrap_LeeConfiguracionLocal_ForceLocalSoloConParametro()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendBootstrap_LeeConfiguracionLocal_ForceLocalSoloConParametro", resultText)

    resultText = Test_BackendBootstrap_LeeConfiguracionLocal_NoFuerzaLocalPorEnPruebas()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendBootstrap_LeeConfiguracionLocal_NoFuerzaLocalPorEnPruebas", resultText)

    resultText = Test_BackendBootstrap_NombreAplicacionSegunConfiguracion_UsaEnPruebas()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendBootstrap_NombreAplicacionSegunConfiguracion_UsaEnPruebas", resultText)

    resultText = Test_BackendConfigHelper_CreaFilaDesdeTablaVacia()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendConfigHelper_CreaFilaDesdeTablaVacia", resultText)

    resultText = Test_BackendConfigHelper_ActualizaFilaExistente()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendConfigHelper_ActualizaFilaExistente", resultText)

    resultText = Test_BackendConfigHelper_BloqueaDuplicados()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendConfigHelper_BloqueaDuplicados", resultText)

    resultText = Test_BackendConfigHelper_RechazaBackendActivoInvalido()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_BackendConfigHelper_RechazaBackendActivoInvalido", resultText)

    logs(8) = "9. Testing harness lifecycle"
    resultText = Test_TestingHarness_ForceLocalBackend_ActivaModoTest()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_TestingHarness_ForceLocalBackend_ActivaModoTest", resultText)

    resultText = Test_TestingHarness_ResetTestSession_DesactivaModoTest()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_TestingHarness_ResetTestSession_DesactivaModoTest", resultText)

    resultText = Test_TestingHarness_Getdb_UsaSandboxEnModoTest()
    If Not Test_Helper.JsonResultOk(resultText) Then failures.Add Test_Helper.BuildChildFailureSummary("Test_TestingHarness_Getdb_UsaSandboxEnModoTest", resultText)

    logs(9) = "10. SuiteTeardown: ResetTestSession"
    If Not Test_Helper.ResetTestSession(errMsg) Then failures.Add "SuiteTeardown: " & errMsg

    logs(10) = "11. Assert: suite sin fallos"
    If failures.Count > 0 Then
        Test_Lanzadera_RunAll = Test_Helper.BuildJsonFail("Fallaron: " & Test_Helper.JoinCollection(failures), logs)
        Exit Function
    End If

    Test_Lanzadera_RunAll = Test_Helper.BuildJsonOk("lanzadera_runall", logs)
    Exit Function

EH:
    errMsg = Err.Description
    Test_Helper.ResetTestSession errMsg
    Test_Lanzadera_RunAll = Test_Helper.BuildJsonFail(Err.Description, logs)
End Function
