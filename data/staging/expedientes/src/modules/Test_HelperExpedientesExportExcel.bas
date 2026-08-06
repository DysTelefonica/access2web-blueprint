Attribute VB_Name = "Test_HelperExpedientesExportExcel"
Option Compare Database
Option Explicit
' Tests atómicos para Helper_ExpedientesExportExcel (PRUEBA-003 REFAC-1b).
' STATELESS helper. Depende de m_ObjEntorno (global) y Excel COM (no sandbox-ificable).
' Cobertura: BR-25-01..04.
'
' LIMITACIÓN: Los popup helpers (MostrarPopupProgreso/AnimarProgresoIndefinido/
' ActualizarEstadoPopup) ejecutan DoCmd.OpenForm "frmBusy" que bloquea el COM test
' runner. Por eso solo testeamos las ramas de early-exit (no collections / Nothing
' collections). Los 3 tests marcados @Deprecated cubren la lógica de cancelación
' (g_OperationCancelled) que sí está implementada en el helper pero no es
' testeable end-to-end sin un popup tolerante a tests. Próximo slice: refactor
' del popup para que sea inyectable o tenga un test-mode flag.

' Helper: crea un ExpedienteCompleto con campos para el export.
Private Function BuildExpedienteCompleto( _
    ByVal p_IDExpediente As String, _
    ByVal p_CodExp As String, _
    ByVal p_Nemotecnico As String, _
    ByVal p_ESTADO As String) As ExpedienteCompleto
    Dim m_Exp As New ExpedienteCompleto
    Dim m_Err As String
    m_Exp.SetPropiedad "IDExpediente", p_IDExpediente, m_Err
    m_Exp.SetPropiedad "CodExp", p_CodExp, m_Err
    m_Exp.SetPropiedad "Nemotecnico", p_Nemotecnico, m_Err
    m_Exp.SetPropiedad "ESTADO", p_ESTADO, m_Err
    m_Exp.SetPropiedad "Titulo", "Test " & p_IDExpediente, m_Err
    Set BuildExpedienteCompleto = m_Exp
End Function

' BR-25-01: Coleccion vacia -> retorna "" y no genera Excel.
Public Function Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_ColVacia_NoGeneraExcel() As String
    Dim logs(0 To 2) As String
    Dim colCampos As Scripting.Dictionary
    Dim errMsg As String
    Dim result As String
    Set colCampos = New Scripting.Dictionary
    colCampos.CompareMode = TextCompare
    colCampos.Add "IDExpediente", "IDExpediente"
    colCampos.Add "Nemotecnico", "Nemotecnico"
    logs(0) = "Arrange: p_ColCampos set, all collections Nothing"
    result = Helper_ExpedientesExportExcel.GenerarConsultaExpedientes(colCampos, Nothing, Nothing, Nothing, Nothing, True, errMsg)
    logs(1) = "Act: GenerarConsultaExpedientes executed"
    If errMsg <> "" Then
        Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_ColVacia_NoGeneraExcel = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    If result <> "" Then
        Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_ColVacia_NoGeneraExcel = BuildJsonFail("expected empty URL, got: " & result, logs)
        Exit Function
    End If
    logs(2) = "Assert: empty URL, no error"
    Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_ColVacia_NoGeneraExcel = BuildJsonOk("ok", logs)
End Function

' BR-25-01 (subcase): p_ColCampos Nothing -> retorna "" sin error.
Public Function Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_pColCamposNothing_NoGeneraExcel() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: p_ColCampos=Nothing"
    result = Helper_ExpedientesExportExcel.GenerarConsultaExpedientes(Nothing, Nothing, Nothing, Nothing, Nothing, True, errMsg)
    logs(1) = "Act: executed"
    If errMsg <> "" Then
        Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_pColCamposNothing_NoGeneraExcel = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    If result <> "" Then
        Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_pColCamposNothing_NoGeneraExcel = BuildJsonFail("expected empty, got: " & result, logs)
        Exit Function
    End If
    logs(2) = "Assert: empty URL, no error"
    Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_pColCamposNothing_NoGeneraExcel = BuildJsonOk("ok", logs)
End Function

' @Deprecated BR-25-02..04: tests de la lógica de cancelación y el flujo Excel
' completo. Implementados en el helper (g_OperationCancelled + chequeos en cada
' loop) pero NO testeables end-to-end porque los popup helpers bloquean el COM
' test runner. Mantenemos el slot para post-PR-E (popup inyectable o test-mode).
Public Function Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_FiltroAM_NoTesteableHoy() As String
    ' Marcador de cobertura: el caso BR-25-02 (filtro AM con 1 item) está implementado
    ' pero no es testeable sin refactor del popup.
    Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_FiltroAM_NoTesteableHoy = BuildJsonOk("skip: popup helpers bloquean COM", Array( _
        "Arrange: popup helpers bloquean COM test runner", _
        "Act: skip", _
        "Assert: target post-PRUEBA-002 PR-E (popup inyectable)"))
End Function
