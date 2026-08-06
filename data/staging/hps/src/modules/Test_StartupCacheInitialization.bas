Attribute VB_Name = "Test_StartupCacheInitialization"
Option Compare Database
Option Explicit

Public Function Test_SCI_DefaultPlan_UsesIncrementalDriftAwareStartup() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim plan As String
    Set logs = New Collection
    On Error GoTo EH

    plan = DescribirPlanSincronizarCachesLocalesFrontend(p_Error:=errMsg)
    If errMsg <> vbNullString Then Err.Raise 1000, , errMsg

    If InStr(1, plan, "HPS_FULL", vbTextCompare) > 0 Then
        Test_SCI_DefaultPlan_UsesIncrementalDriftAwareStartup = JsonFail("Default startup must not use full HPS rebuild: " & plan, logs)
        Exit Function
    End If

    If plan <> "HPS_INCREMENTAL|HPS_DERIVED_DRIFT_TbHPS_HPS_SOLICITUDES_TbUsuarios|SICA_CONTROLLED_FULL_CAVEAT|INDICATORS_REFRESH" Then
        Test_SCI_DefaultPlan_UsesIncrementalDriftAwareStartup = JsonFail("Unexpected default startup cache plan: " & plan, logs)
        Exit Function
    End If

    logs.Add "Default frontend cache startup is incremental plus HPS-derived drift check, controlled SICA caveat, then indicator refresh."
    Test_SCI_DefaultPlan_UsesIncrementalDriftAwareStartup = JsonOk("startup-cache-default-plan", logs)
    Exit Function
EH:
    Test_SCI_DefaultPlan_UsesIncrementalDriftAwareStartup = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_SCI_DefaultPlan_IncludesHpsDerivedDriftInvalidator() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim plan As String
    Set logs = New Collection
    On Error GoTo EH

    plan = DescribirPlanSincronizarCachesLocalesFrontend(p_Error:=errMsg)
    If errMsg <> vbNullString Then Err.Raise 1000, , errMsg

    If InStr(1, plan, "HPS_DERIVED_DRIFT", vbTextCompare) = 0 Then
        Test_SCI_DefaultPlan_IncludesHpsDerivedDriftInvalidator = JsonFail("Default plan misses HPS-derived drift invalidator: " & plan, logs)
        Exit Function
    End If

    If InStr(1, plan, "TbHPS", vbTextCompare) = 0 Then
        Test_SCI_DefaultPlan_IncludesHpsDerivedDriftInvalidator = JsonFail("Default plan does not name TbHPS as drift source: " & plan, logs)
        Exit Function
    End If

    If InStr(1, plan, "HPS_SOLICITUDES_TbUsuarios", vbTextCompare) = 0 Then
        Test_SCI_DefaultPlan_IncludesHpsDerivedDriftInvalidator = JsonFail("Default plan does not document HPS_solicitudes consequences on TbUsuarios: " & plan, logs)
        Exit Function
    End If

    logs.Add "Default plan includes TbHPS and HPS_solicitudes/TbUsuarios as HPS-derived cache invalidators."
    Test_SCI_DefaultPlan_IncludesHpsDerivedDriftInvalidator = JsonOk("startup-cache-hps-derived-drift", logs)
    Exit Function
EH:
    Test_SCI_DefaultPlan_IncludesHpsDerivedDriftInvalidator = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_SCI_FormInicialStartupPlan_ExcludesHistoricalFullRebuild() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim plan As String
    Set logs = New Collection
    On Error GoTo EH

    plan = DescribirPlanCargaInicialFormInicial(p_Error:=errMsg)
    If errMsg <> vbNullString Then Err.Raise 1000, , errMsg

    If InStr(1, plan, "HISTORICAL_FULL_REBUILD", vbTextCompare) > 0 Then
        Test_SCI_FormInicialStartupPlan_ExcludesHistoricalFullRebuild = JsonFail("FormInicial startup must not run historical full rebuild by default: " & plan, logs)
        Exit Function
    End If

    If InStr(1, plan, "HPS_FULL", vbTextCompare) > 0 Then
        Test_SCI_FormInicialStartupPlan_ExcludesHistoricalFullRebuild = JsonFail("FormInicial startup must not run full HPS rebuild by default: " & plan, logs)
        Exit Function
    End If

    If InStr(1, plan, "STARTUP_FAST_SYNC", vbTextCompare) = 0 Then
        Test_SCI_FormInicialStartupPlan_ExcludesHistoricalFullRebuild = JsonFail("FormInicial startup plan must identify the real fast-sync startup seam: " & plan, logs)
        Exit Function
    End If

    If InStr(1, plan, "HPS_INCREMENTAL", vbTextCompare) = 0 Then
        Test_SCI_FormInicialStartupPlan_ExcludesHistoricalFullRebuild = JsonFail("FormInicial startup plan must keep incremental HPS cache sync: " & plan, logs)
        Exit Function
    End If

    logs.Add "FormInicial startup delegates to the fast default cache sync seam and excludes the manual historical full rebuild."
    Test_SCI_FormInicialStartupPlan_ExcludesHistoricalFullRebuild = JsonOk("startup-forminicial-no-historical-full-rebuild", logs)
    Exit Function
EH:
    Test_SCI_FormInicialStartupPlan_ExcludesHistoricalFullRebuild = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_SCI_ManualFullPlan_UsesFullHpsRebuildBeforeSicaAndIndicators() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim plan As String
    Set logs = New Collection
    On Error GoTo EH

    plan = DescribirPlanRegenerarCachesLocalesFrontend(p_Error:=errMsg)
    If errMsg <> vbNullString Then Err.Raise 1000, , errMsg

    If plan <> "HPS_FULL|SICA_FULL|INDICATORS_REFRESH" Then
        Test_SCI_ManualFullPlan_UsesFullHpsRebuildBeforeSicaAndIndicators = JsonFail("Unexpected manual full startup cache plan: " & plan, logs)
        Exit Function
    End If

    logs.Add "Manual compatibility wrapper keeps the full HPS, full SICA, then indicator refresh plan."
    Test_SCI_ManualFullPlan_UsesFullHpsRebuildBeforeSicaAndIndicators = JsonOk("startup-cache-manual-full-plan", logs)
    Exit Function
EH:
    Test_SCI_ManualFullPlan_UsesFullHpsRebuildBeforeSicaAndIndicators = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_SCI_ManualHistoricalRefreshPlan_IsExplicitOnly() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim plan As String
    Set logs = New Collection
    On Error GoTo EH

    plan = DescribirPlanRegenerarUsuariosHistoricosLocalesFrontend(p_Error:=errMsg)
    If errMsg <> vbNullString Then Err.Raise 1000, , errMsg

    If plan <> "HISTORICAL_FULL_REBUILD" Then
        Test_SCI_ManualHistoricalRefreshPlan_IsExplicitOnly = JsonFail("Unexpected manual historical refresh plan: " & plan, logs)
        Exit Function
    End If

    logs.Add "Historical users cache full rebuild remains available only through an explicit manual helper."
    Test_SCI_ManualHistoricalRefreshPlan_IsExplicitOnly = JsonOk("startup-cache-manual-historical-refresh", logs)
    Exit Function
EH:
    Test_SCI_ManualHistoricalRefreshPlan_IsExplicitOnly = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_SCI_PrincipalRefreshButton_UsesFastSyncAndCounterRefresh() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim plan As String
    Set logs = New Collection
    On Error GoTo EH

    plan = DescribirPlanBotonPrincipalActualizarCachesFrontend(p_Error:=errMsg)
    If errMsg <> vbNullString Then Err.Raise 1000, , errMsg

    If plan <> "FAST_SYNC|COUNTERS_REFRESH" Then
        Test_SCI_PrincipalRefreshButton_UsesFastSyncAndCounterRefresh = JsonFail("Unexpected principal refresh button plan: " & plan, logs)
        Exit Function
    End If

    If InStr(1, plan, "FULL", vbTextCompare) > 0 Then
        Test_SCI_PrincipalRefreshButton_UsesFastSyncAndCounterRefresh = JsonFail("Principal refresh button must not request full rebuild by default: " & plan, logs)
        Exit Function
    End If

    logs.Add "Principal refresh button uses fast cache sync, then refreshes main counters."
    Test_SCI_PrincipalRefreshButton_UsesFastSyncAndCounterRefresh = JsonOk("startup-cache-principal-refresh-button", logs)
    Exit Function
EH:
    Test_SCI_PrincipalRefreshButton_UsesFastSyncAndCounterRefresh = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_SCI_IndicatorsRefreshButton_UsesFastSyncIndicatorsAndSafeRequery() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim plan As String
    Set logs = New Collection
    On Error GoTo EH

    plan = DescribirPlanBotonIndicadoresActualizarCachesFrontend(p_Error:=errMsg)
    If errMsg <> vbNullString Then Err.Raise 1000, , errMsg

    If plan <> "FAST_SYNC|INDICATORS_REFRESH|SUBFORM_SAFE_REQUERY" Then
        Test_SCI_IndicatorsRefreshButton_UsesFastSyncIndicatorsAndSafeRequery = JsonFail("Unexpected indicators refresh button plan: " & plan, logs)
        Exit Function
    End If

    If InStr(1, plan, "FULL", vbTextCompare) > 0 Then
        Test_SCI_IndicatorsRefreshButton_UsesFastSyncIndicatorsAndSafeRequery = JsonFail("Indicators refresh button must not request full rebuild by default: " & plan, logs)
        Exit Function
    End If

    logs.Add "Indicators refresh button uses fast cache sync, refreshes numeric indicators, then safely requeries the subform."
    Test_SCI_IndicatorsRefreshButton_UsesFastSyncIndicatorsAndSafeRequery = JsonOk("startup-cache-indicators-refresh-button", logs)
    Exit Function
EH:
    Test_SCI_IndicatorsRefreshButton_UsesFastSyncIndicatorsAndSafeRequery = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

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
