Attribute VB_Name = "modFrmNCAuditoriaGeneralHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Form helper for Form_FormNCAuditoriaGeneral (slice 1 of form-thin-helper-refactor).
' Mirror of modFrmNCProyectoGeneralHelper for the NCAuditoria side.
' =============================================================================

' ----- Message contract (byte-for-byte match) -----
Private Const MSG_MOTIVO_LIMPIADO As String = "El campo 'Motivo de control de eficacia no requerido' se ha limpiado porque ahora sí requiere control de eficacia."

' =============================================================================
' PUBLIC API
' =============================================================================

Public Function modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos( _
                                            ByRef p_NcActual As NCAuditoria, _
                                            ByRef p_NcInicial As NCAuditoria, _
                                            Optional ByRef db As DAO.Database = Nothing, _
                                            Optional ByRef p_PromptResult As Long = -1, _
                                            Optional ByRef p_MessageText As String = "", _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim motivoBefore As String
    Dim motivoWasCleared As Boolean
    Dim op As NCaUDITORIAOperaciones
    Dim errMsgLocal As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()

    ' Pre-flight: detect if motivo WILL be cleared by the operations class
    ' (No -> Sí transition with motivo set in the initial state). Decoupled from
    ' operations success/failure (see mirror comment in modFrmNCProyectoGeneralHelper).
    If IsNull(p_NcInicial.MotivoNoRequiereControlEficacia) Then
        motivoBefore = ""
    Else
        motivoBefore = p_NcInicial.MotivoNoRequiereControlEficacia
    End If
    motivoWasCleared = (motivoBefore <> "" And _
                        p_NcInicial.RequiereControlEficacia = "No" And _
                        p_NcActual.RequiereControlEficacia = "Sí")
    TestHelper.AddLog logs, "Arrange: motivoInitial='" & motivoBefore & _
                            "', requiereInitial='" & p_NcInicial.RequiereControlEficacia & _
                            "', requiereCurrent='" & p_NcActual.RequiereControlEficacia & _
                            "', willBeCleared=" & motivoWasCleared

    Set op = New NCaUDITORIAOperaciones
    Set op.nc = p_NcActual

    errMsgLocal = ""
    resultStr = op.RegistrarDatosUnicos(p_NcInicial, errMsgLocal, EnumSino.No)

    If motivoWasCleared Then
        p_MessageText = MSG_MOTIVO_LIMPIADO
        TestHelper.AddLog logs, "Assert: p_MessageText populated with MSG_MOTIVO_LIMPIADO"
    Else
        p_MessageText = ""
        TestHelper.AddLog logs, "Assert: p_MessageText empty (no No->Sí transition with motivo)"
    End If

    If errMsgLocal <> "" Then
        p_Error = errMsgLocal
        modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos = TestHelper.BuildJsonFail(errMsgLocal, logs)
        TestHelper.AddLog logs, "Result: FAIL (" & errMsgLocal & ")"
        Exit Function
    End If

    p_Error = ""
    modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos = TestHelper.BuildJsonOk(logs, "registrar_datos_unicos_ok")
    TestHelper.AddLog logs, "Result: OK"
    Exit Function

EH:
    modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos = TestHelper.BuildJsonFail( _
        "modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos: " & Err.Description, logs)
End Function
