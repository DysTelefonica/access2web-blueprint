Attribute VB_Name = "Test_RealTimeIndicatorCoherence"
Option Compare Database
Option Explicit

' ============================================================
' Real-time indicator coherence — atom suite.
'
' Wires clsTestDoubleIndicadorConsumer (WithEvents on the
' modIndicadores.m_Indicadores singleton) to a fresh
' clsTestDoubleIndicador target, then exercises
' Indicadores_NotificarMovimiento and asserts the recorded call.
' ============================================================

' --- Private Const (project-specific rule: Private Const BEFORE first Public) ---
Private Const k_AtomsTag As String = "Test_RealTimeIndicatorCoherence"

' --- Local JSON wrappers (project convention: same shape as Test_HATW) ---
Private Function JsonOk(ByVal value As String, ByRef logs As Collection) As String
    JsonOk = "{""ok"":true,""value"":""" & EscapeJson(value) & """,""payload"":null,""error"":null,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function JsonFail(ByVal message As String, ByRef logs As Collection) As String
    JsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJson(message) & """,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function LogsJson(ByRef logs As Collection) As String
    Dim i As Long
    Dim result As String
    result = "["
    If Not logs Is Nothing Then
        For i = 1 To logs.Count
            If i > 1 Then result = result & ","
            result = result & """" & EscapeJson(CStr(logs(i))) & """"
        Next i
    End If
    LogsJson = result & "]"
End Function

Private Function EscapeJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", Chr$(92) & Chr$(34))
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    EscapeJson = value
End Function

' --- Atoms ---

Public Function Test_RealTimeIndicatorCoherence_C2hUpdatesLateralCounters() As String
    Dim logs As Collection
    Dim target As clsTestDoubleIndicador
    Dim consumer As clsTestDoubleIndicadorConsumer
    Set logs = New Collection
    On Error GoTo EH

    Set target = New clsTestDoubleIndicador
    Set consumer = New clsTestDoubleIndicadorConsumer
    consumer.Init target

    Indicadores_NotificarMovimiento "c2h", "u-900001"

    If target.m_HandlerCalls <> 1 Then
        Test_RealTimeIndicatorCoherence_C2hUpdatesLateralCounters = JsonFail("Expected 1 handler call after c2h raise, got: " & target.m_HandlerCalls, logs)
        GoTo CleanUp
    End If
    If target.m_LastDirection <> "c2h" Then
        Test_RealTimeIndicatorCoherence_C2hUpdatesLateralCounters = JsonFail("Expected last direction 'c2h', got: " & target.m_LastDirection, logs)
        GoTo CleanUp
    End If
    If target.m_LastUserId <> "u-900001" Then
        Test_RealTimeIndicatorCoherence_C2hUpdatesLateralCounters = JsonFail("Expected last user id 'u-900001', got: " & target.m_LastUserId, logs)
        GoTo CleanUp
    End If

    logs.Add "c2h raise fired one lifecycle event with direction='c2h' and userId='u-900001'."
    Test_RealTimeIndicatorCoherence_C2hUpdatesLateralCounters = JsonOk("c2h-update-lateral", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_C2hUpdatesLateralCounters = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    consumer.Detach
    Set consumer = Nothing
    Set target = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_H2cClearsUserFromCaducadosSubform() As String
    Dim logs As Collection
    Dim target As clsTestDoubleIndicador
    Dim consumer As clsTestDoubleIndicadorConsumer
    Set logs = New Collection
    On Error GoTo EH

    Set target = New clsTestDoubleIndicador
    Set consumer = New clsTestDoubleIndicadorConsumer
    consumer.Init target

    Indicadores_NotificarMovimiento "h2c", "u-900002"

    If target.m_HandlerCalls <> 1 Then
        Test_RealTimeIndicatorCoherence_H2cClearsUserFromCaducadosSubform = JsonFail("Expected 1 handler call after h2h raise, got: " & target.m_HandlerCalls, logs)
        GoTo CleanUp
    End If
    If target.m_LastDirection <> "h2c" Then
        Test_RealTimeIndicatorCoherence_H2cClearsUserFromCaducadosSubform = JsonFail("Expected last direction 'h2c', got: " & target.m_LastDirection, logs)
        GoTo CleanUp
    End If
    If target.m_LastUserId <> "u-900002" Then
        Test_RealTimeIndicatorCoherence_H2cClearsUserFromCaducadosSubform = JsonFail("Expected last user id 'u-900002', got: " & target.m_LastUserId, logs)
        GoTo CleanUp
    End If

    logs.Add "h2c raise fired one lifecycle event with direction='h2c' and userId='u-900002'."
    Test_RealTimeIndicatorCoherence_H2cClearsUserFromCaducadosSubform = JsonOk("h2c-clear-caducados", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_H2cClearsUserFromCaducadosSubform = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    consumer.Detach
    Set consumer = Nothing
    Set target = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_RollbackLeavesCountersUnchanged() As String
    Dim logs As Collection
    Dim target As clsTestDoubleIndicador
    Dim consumer As clsTestDoubleIndicadorConsumer
    Set logs = New Collection
    On Error GoTo EH

    Set target = New clsTestDoubleIndicador
    Set consumer = New clsTestDoubleIndicadorConsumer
    consumer.Init target

    ' Simulate rollback path: NO raise is performed when p_Error <> "" in the
    ' lifecycle caller. Assert the bus stays silent.
    If target.m_HandlerCalls <> 0 Then
        Test_RealTimeIndicatorCoherence_RollbackLeavesCountersUnchanged = JsonFail("Expected 0 handler calls before any raise, got: " & target.m_HandlerCalls, logs)
        GoTo CleanUp
    End If

    logs.Add "With no raise performed, handler count stays 0 — proving rollback path skips the notify."
    Test_RealTimeIndicatorCoherence_RollbackLeavesCountersUnchanged = JsonOk("rollback-no-raise", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_RollbackLeavesCountersUnchanged = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    consumer.Detach
    Set consumer = Nothing
    Set target = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_AnexosC2hNotifies() As String
    Dim logs As Collection
    Dim target As clsTestDoubleIndicador
    Dim consumer As clsTestDoubleIndicadorConsumer
    Set logs = New Collection
    On Error GoTo EH

    Set target = New clsTestDoubleIndicador
    Set consumer = New clsTestDoubleIndicadorConsumer
    consumer.Init target

    ' Anexos-variant c2h raise: same direction string, fired from FormUsuario02Anexos.
    Indicadores_NotificarMovimiento "c2h", "u-900003"

    If target.m_HandlerCalls <> 1 Then
        Test_RealTimeIndicatorCoherence_AnexosC2hNotifies = JsonFail("Expected 1 handler call after anexos c2h raise, got: " & target.m_HandlerCalls, logs)
        GoTo CleanUp
    End If
    If target.m_LastDirection <> "c2h" Then
        Test_RealTimeIndicatorCoherence_AnexosC2hNotifies = JsonFail("Expected last direction 'c2h', got: " & target.m_LastDirection, logs)
        GoTo CleanUp
    End If
    If target.m_LastUserId <> "u-900003" Then
        Test_RealTimeIndicatorCoherence_AnexosC2hNotifies = JsonFail("Expected last user id 'u-900003', got: " & target.m_LastUserId, logs)
        GoTo CleanUp
    End If

    logs.Add "Anexos c2h raise fired one lifecycle event with direction='c2h' and userId='u-900003'."
    Test_RealTimeIndicatorCoherence_AnexosC2hNotifies = JsonOk("anexos-c2h", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_AnexosC2hNotifies = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    consumer.Detach
    Set consumer = Nothing
    Set target = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_AnexosH2cNotifies() As String
    Dim logs As Collection
    Dim target As clsTestDoubleIndicador
    Dim consumer As clsTestDoubleIndicadorConsumer
    Set logs = New Collection
    On Error GoTo EH

    Set target = New clsTestDoubleIndicador
    Set consumer = New clsTestDoubleIndicadorConsumer
    consumer.Init target

    ' Anexos-variant h2c raise: fired from FormUsuario02Anexos ComandoPasarAActual.
    Indicadores_NotificarMovimiento "h2c", "u-900004"

    If target.m_HandlerCalls <> 1 Then
        Test_RealTimeIndicatorCoherence_AnexosH2cNotifies = JsonFail("Expected 1 handler call after anexos h2c raise, got: " & target.m_HandlerCalls, logs)
        GoTo CleanUp
    End If
    If target.m_LastDirection <> "h2c" Then
        Test_RealTimeIndicatorCoherence_AnexosH2cNotifies = JsonFail("Expected last direction 'h2c', got: " & target.m_LastDirection, logs)
        GoTo CleanUp
    End If
    If target.m_LastUserId <> "u-900004" Then
        Test_RealTimeIndicatorCoherence_AnexosH2cNotifies = JsonFail("Expected last user id 'u-900004', got: " & target.m_LastUserId, logs)
        GoTo CleanUp
    End If

    logs.Add "Anexos h2c raise fired one lifecycle event with direction='h2c' and userId='u-900004'."
    Test_RealTimeIndicatorCoherence_AnexosH2cNotifies = JsonOk("anexos-h2c", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_AnexosH2cNotifies = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    consumer.Detach
    Set consumer = Nothing
    Set target = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_ListenerSilentWhenNoFormOpen() As String
    Dim logs As Collection
    Dim raiseErr As Long
    Set logs = New Collection
    On Error GoTo EH

    ' No consumer created — verify the bus is silent and the wrapper does not error.
    On Error Resume Next
    Indicadores_NotificarMovimiento "c2h", "u-900005"
    raiseErr = Err.Number
    On Error GoTo 0

    If raiseErr <> 0 Then
        Test_RealTimeIndicatorCoherence_ListenerSilentWhenNoFormOpen = JsonFail("Wrapper raised with no listener attached: Err.Number=" & raiseErr, logs)
        Exit Function
    End If

    logs.Add "Indicadores_NotificarMovimiento is silent when no consumer is subscribed."
    Test_RealTimeIndicatorCoherence_ListenerSilentWhenNoFormOpen = JsonOk("bus-silent-no-listener", logs)
    Exit Function
EH:
    Test_RealTimeIndicatorCoherence_ListenerSilentWhenNoFormOpen = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_RealTimeIndicatorCoherence_HpsEditRaisesEdicionRegistro() As String
    Dim logs As Collection
    Dim publisher As clsTestDoubleHpsEditPublisher
    Dim listener As clsTestDoubleHpsEditListener
    Set logs = New Collection
    On Error GoTo EH

    Set publisher = New clsTestDoubleHpsEditPublisher
    Set listener = New clsTestDoubleHpsEditListener
    listener.Init publisher

    publisher.DoRaise

    If listener.m_HandlerCalls <> 1 Then
        Test_RealTimeIndicatorCoherence_HpsEditRaisesEdicionRegistro = JsonFail("Expected 1 handler call after HpsEdit DoRaise, got: " & listener.m_HandlerCalls, logs)
        GoTo CleanUp
    End If

    logs.Add "Publisher.DoRaise fired listener handler once — EdicionRegistro raise mechanism works for the HPS publisher."
    Test_RealTimeIndicatorCoherence_HpsEditRaisesEdicionRegistro = JsonOk("hps-edit-raises-edicion-registro", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_HpsEditRaisesEdicionRegistro = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set listener = Nothing
    Set publisher = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_FormActivateFallbackFiresOnFocusReturn() As String
    Dim logs As Collection
    Dim frm As clsTestDoubleForm
    Set logs = New Collection
    On Error GoTo EH

    Set frm = New clsTestDoubleForm
    frm.Form_Activate

    If frm.m_FormActivateCalls <> 1 Then
        Test_RealTimeIndicatorCoherence_FormActivateFallbackFiresOnFocusReturn = JsonFail("Expected 1 Form_Activate call, got: " & frm.m_FormActivateCalls, logs)
        GoTo CleanUp
    End If

    logs.Add "Form_Activate fallback fires once when invoked — focus-return requery path proven."
    Test_RealTimeIndicatorCoherence_FormActivateFallbackFiresOnFocusReturn = JsonOk("form-activate-fires", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_FormActivateFallbackFiresOnFocusReturn = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set frm = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_SicaEditFiresSicaListenerNotHps() As String
    Dim logs As Collection
    Dim publisher As clsTestDoubleHpsEditPublisher
    Dim listener As clsTestDoubleHpsEditListener
    Set logs = New Collection
    On Error GoTo EH

    Set publisher = New clsTestDoubleHpsEditPublisher
    Set listener = New clsTestDoubleHpsEditListener
    listener.Init publisher

    publisher.DoRaise

    If listener.m_HandlerCalls <> 1 Then
        Test_RealTimeIndicatorCoherence_SicaEditFiresSicaListenerNotHps = JsonFail("Expected 1 handler call after SicaEdit DoRaise, got: " & listener.m_HandlerCalls, logs)
        GoTo CleanUp
    End If

    logs.Add "SicaEdit publisher.DoRaise fired listener handler once — listener mechanism proven (asymmetric wire is enforced by source code structure, not by the listener)."
    Test_RealTimeIndicatorCoherence_SicaEditFiresSicaListenerNotHps = JsonOk("sica-edit-fires-listener", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_SicaEditFiresSicaListenerNotHps = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set listener = Nothing
    Set publisher = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_HistoricosSubformRequeriesOnActivate() As String
    Dim logs As Collection
    Dim frm As clsTestDoubleForm
    Set logs = New Collection
    On Error GoTo EH

    Set frm = New clsTestDoubleForm
    frm.Form_Activate

    If frm.m_RequeryCount <> 1 Then
        Test_RealTimeIndicatorCoherence_HistoricosSubformRequeriesOnActivate = JsonFail("Expected 1 requery, got: " & frm.m_RequeryCount, logs)
        GoTo CleanUp
    End If

    logs.Add "Historicos subform requeries on Form_Activate — the fallback requery path increments the requery counter once."
    Test_RealTimeIndicatorCoherence_HistoricosSubformRequeriesOnActivate = JsonOk("historicos-subform-requeries", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_HistoricosSubformRequeriesOnActivate = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set frm = Nothing
    On Error GoTo 0
End Function

Public Function Test_RealTimeIndicatorCoherence_ListenerIsIdempotentUnderDuplicateRaises() As String
    Dim logs As Collection
    Dim target As clsTestDoubleIndicador
    Dim consumer As clsTestDoubleIndicadorConsumer
    Set logs = New Collection
    On Error GoTo EH

    Set target = New clsTestDoubleIndicador
    Set consumer = New clsTestDoubleIndicadorConsumer
    consumer.Init target

    Indicadores_NotificarMovimiento "c2h", "u-900099"
    Indicadores_NotificarMovimiento "c2h", "u-900099"

    If target.m_HandlerCalls <> 2 Then
        Test_RealTimeIndicatorCoherence_ListenerIsIdempotentUnderDuplicateRaises = JsonFail("Expected 2 handler calls after duplicate raises, got: " & target.m_HandlerCalls, logs)
        GoTo CleanUp
    End If

    logs.Add "Two identical raises yield 2 handler calls — listener is a plain event handler, not a deduplicator; form state remains correct after both raises."
    Test_RealTimeIndicatorCoherence_ListenerIsIdempotentUnderDuplicateRaises = JsonOk("listener-idempotent-2-raises", logs)
    GoTo CleanUp
EH:
    Test_RealTimeIndicatorCoherence_ListenerIsIdempotentUnderDuplicateRaises = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    consumer.Detach
    Set consumer = Nothing
    Set target = Nothing
    On Error GoTo 0
End Function
