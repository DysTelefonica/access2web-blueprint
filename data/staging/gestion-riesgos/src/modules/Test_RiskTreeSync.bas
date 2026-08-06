Attribute VB_Name = "Test_RiskTreeSync"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_RiskTreeSync_DescriptionMode_Si() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: valor de combo='Sí'"
    logs(1) = "2. Act: ResolverModoVerDescripcion"
    logs(2) = "3. Assert: devuelve EnumSiNo.Sí"

    Dim m_Resultado As EnumSiNo
    m_Resultado = ResolverModoVerDescripcion("Sí")
    If m_Resultado <> EnumSiNo.Sí Then
        Test_RiskTreeSync_DescriptionMode_Si = BuildFail("ResolverModoVerDescripcion debe devolver Sí para valor 'Sí'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_DescriptionMode_Si = BuildOk("description_mode_si", logs)
    Exit Function
EH:
    Test_RiskTreeSync_DescriptionMode_Si = BuildFail("Test_RiskTreeSync_DescriptionMode_Si: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_DescriptionMode_No() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: valor de combo='No'"
    logs(1) = "2. Act: ResolverModoVerDescripcion"
    logs(2) = "3. Assert: devuelve EnumSiNo.No"

    Dim m_Resultado As EnumSiNo
    m_Resultado = ResolverModoVerDescripcion("No")
    If m_Resultado <> EnumSiNo.No Then
        Test_RiskTreeSync_DescriptionMode_No = BuildFail("ResolverModoVerDescripcion debe devolver No para valor 'No'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_DescriptionMode_No = BuildOk("description_mode_no", logs)
    Exit Function
EH:
    Test_RiskTreeSync_DescriptionMode_No = BuildFail("Test_RiskTreeSync_DescriptionMode_No: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ScopeNormalizer_Plan() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='plan'"
    logs(1) = "2. Act: NormalizarScopeRefrescoArbolRiesgos"
    logs(2) = "3. Assert: devuelve 'plan'"

    Dim m_Scope As String
    m_Scope = NormalizarScopeRefrescoArbolRiesgos("plan")
    If StrComp(m_Scope, "plan", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ScopeNormalizer_Plan = BuildFail("NormalizarScopeRefrescoArbolRiesgos debe devolver 'plan'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ScopeNormalizer_Plan = BuildOk("scope_plan", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ScopeNormalizer_Plan = BuildFail("Test_RiskTreeSync_ScopeNormalizer_Plan: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ScopeNormalizer_UnknownFallsRisk() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='desconocido'"
    logs(1) = "2. Act: NormalizarScopeRefrescoArbolRiesgos"
    logs(2) = "3. Assert: fallback='risk'"

    Dim m_Scope As String
    m_Scope = NormalizarScopeRefrescoArbolRiesgos("desconocido")
    If StrComp(m_Scope, "risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ScopeNormalizer_UnknownFallsRisk = BuildFail("NormalizarScopeRefrescoArbolRiesgos debe fallback a 'risk'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ScopeNormalizer_UnknownFallsRisk = BuildOk("scope_fallback_risk", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ScopeNormalizer_UnknownFallsRisk = BuildFail("Test_RiskTreeSync_ScopeNormalizer_UnknownFallsRisk: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ScopeNormalizer_Action() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='action'"
    logs(1) = "2. Act: NormalizarScopeRefrescoArbolRiesgos"
    logs(2) = "3. Assert: devuelve 'action'"

    Dim m_Scope As String
    m_Scope = NormalizarScopeRefrescoArbolRiesgos("action")
    If StrComp(m_Scope, "action", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ScopeNormalizer_Action = BuildFail("NormalizarScopeRefrescoArbolRiesgos debe devolver 'action'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ScopeNormalizer_Action = BuildOk("scope_action", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ScopeNormalizer_Action = BuildFail("Test_RiskTreeSync_ScopeNormalizer_Action: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ScopeNormalizer_ActionDelete() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='action-delete'"
    logs(1) = "2. Act: NormalizarScopeRefrescoArbolRiesgos"
    logs(2) = "3. Assert: devuelve 'action-delete'"

    Dim m_Scope As String
    m_Scope = NormalizarScopeRefrescoArbolRiesgos("action-delete")
    If StrComp(m_Scope, "action-delete", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ScopeNormalizer_ActionDelete = BuildFail("NormalizarScopeRefrescoArbolRiesgos debe devolver 'action-delete'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ScopeNormalizer_ActionDelete = BuildOk("scope_action_delete", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ScopeNormalizer_ActionDelete = BuildFail("Test_RiskTreeSync_ScopeNormalizer_ActionDelete: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_InvalidationOrder_Action() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='action'"
    logs(1) = "2. Act: ObtenerOrdenInvalidacionRiesgosPorScope"
    logs(2) = "3. Assert: orden='plan|risk'"

    Dim m_Orden As String
    m_Orden = ObtenerOrdenInvalidacionRiesgosPorScope("action")
    If StrComp(m_Orden, "plan|risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_InvalidationOrder_Action = BuildFail("ObtenerOrdenInvalidacionRiesgosPorScope debe devolver 'plan|risk' para scope action", logs)
        Exit Function
    End If

    Test_RiskTreeSync_InvalidationOrder_Action = BuildOk("invalidation_order_action", logs)
    Exit Function
EH:
    Test_RiskTreeSync_InvalidationOrder_Action = BuildFail("Test_RiskTreeSync_InvalidationOrder_Action: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_InvalidationOrder_ActionDelete() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='action-delete'"
    logs(1) = "2. Act: ObtenerOrdenInvalidacionRiesgosPorScope"
    logs(2) = "3. Assert: orden='plan|risk'"

    Dim m_Orden As String
    m_Orden = ObtenerOrdenInvalidacionRiesgosPorScope("action-delete")
    If StrComp(m_Orden, "plan|risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_InvalidationOrder_ActionDelete = BuildFail("ObtenerOrdenInvalidacionRiesgosPorScope debe devolver 'plan|risk' para scope action-delete", logs)
        Exit Function
    End If

    Test_RiskTreeSync_InvalidationOrder_ActionDelete = BuildOk("invalidation_order_action_delete", logs)
    Exit Function
EH:
    Test_RiskTreeSync_InvalidationOrder_ActionDelete = BuildFail("Test_RiskTreeSync_InvalidationOrder_ActionDelete: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_InvalidationOrder_Full() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='full'"
    logs(1) = "2. Act: ObtenerOrdenInvalidacionRiesgosPorScope"
    logs(2) = "3. Assert: orden='plan|risk|edition'"

    Dim m_Orden As String
    m_Orden = ObtenerOrdenInvalidacionRiesgosPorScope("full")
    If StrComp(m_Orden, "plan|risk|edition", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_InvalidationOrder_Full = BuildFail("ObtenerOrdenInvalidacionRiesgosPorScope debe devolver 'plan|risk|edition' para scope full", logs)
        Exit Function
    End If

    Test_RiskTreeSync_InvalidationOrder_Full = BuildOk("invalidation_order_full", logs)
    Exit Function
EH:
    Test_RiskTreeSync_InvalidationOrder_Full = BuildFail("Test_RiskTreeSync_InvalidationOrder_Full: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_FullReloadFlag_ActionDeleteIsMinimal() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='action-delete'"
    logs(1) = "2. Act: EsRefrescoCompletoArbolRiesgosScope"
    logs(2) = "3. Assert: no requiere CargarArbol completo"

    If EsRefrescoCompletoArbolRiesgosScope("action-delete") Then
        Test_RiskTreeSync_FullReloadFlag_ActionDeleteIsMinimal = BuildFail("action-delete no debe usar recarga completa del arbol", logs)
        Exit Function
    End If

    Test_RiskTreeSync_FullReloadFlag_ActionDeleteIsMinimal = BuildOk("action_delete_minimal_refresh", logs)
    Exit Function
EH:
    Test_RiskTreeSync_FullReloadFlag_ActionDeleteIsMinimal = BuildFail("Test_RiskTreeSync_FullReloadFlag_ActionDeleteIsMinimal: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_FullReloadFlag_FullIsFullReload() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope='full'"
    logs(1) = "2. Act: EsRefrescoCompletoArbolRiesgosScope"
    logs(2) = "3. Assert: requiere CargarArbol completo"

    If Not EsRefrescoCompletoArbolRiesgosScope("full") Then
        Test_RiskTreeSync_FullReloadFlag_FullIsFullReload = BuildFail("full debe usar recarga completa del arbol", logs)
        Exit Function
    End If

    Test_RiskTreeSync_FullReloadFlag_FullIsFullReload = BuildOk("full_refresh", logs)
    Exit Function
EH:
    Test_RiskTreeSync_FullReloadFlag_FullIsFullReload = BuildFail("Test_RiskTreeSync_FullReloadFlag_FullIsFullReload: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_FallbackKey_ActionDeleteFallsToPlan() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: requested action deleted, parent plan still exists"
    logs(1) = "2. Act: ResolverClaveFallbackArbolRiesgosConExistentes"
    logs(2) = "3. Assert: devuelve parent plan"

    Dim m_Candidates As String
    Dim m_Existing As String
    Dim m_Result As String

    m_Candidates = "PMACCION|20|30" & vbLf & "PM|10|20" & vbLf & "RIESGO|10" & vbLf & "EDICION|1"
    m_Existing = "PM|10|20" & vbLf & "RIESGO|10" & vbLf & "EDICION|1"
    m_Result = ResolverClaveFallbackArbolRiesgosConExistentes("PMACCION|20|30", m_Candidates, m_Existing)
    If StrComp(m_Result, "PM|10|20", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_FallbackKey_ActionDeleteFallsToPlan = BuildFail("Debe caer al plan padre cuando la accion borrada ya no existe", logs)
        Exit Function
    End If

    Test_RiskTreeSync_FallbackKey_ActionDeleteFallsToPlan = BuildOk("fallback_action_delete_plan", logs)
    Exit Function
EH:
    Test_RiskTreeSync_FallbackKey_ActionDeleteFallsToPlan = BuildFail("Test_RiskTreeSync_FallbackKey_ActionDeleteFallsToPlan: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_QualityScopeResolver_Materializacion() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: mutacion='materializacion'"
    logs(1) = "2. Act: ResolverScopeRefrescoCalidadRiesgo"
    logs(2) = "3. Assert: scope='full'"

    Dim m_Scope As String
    m_Scope = ResolverScopeRefrescoCalidadRiesgo("materializacion")
    If StrComp(m_Scope, "full", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_QualityScopeResolver_Materializacion = BuildFail("ResolverScopeRefrescoCalidadRiesgo debe devolver 'full' para materializacion", logs)
        Exit Function
    End If

    Test_RiskTreeSync_QualityScopeResolver_Materializacion = BuildOk("quality_scope_materializacion", logs)
    Exit Function
EH:
    Test_RiskTreeSync_QualityScopeResolver_Materializacion = BuildFail("Test_RiskTreeSync_QualityScopeResolver_Materializacion: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_QualityScopeResolver_UnknownFallsRisk() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: mutacion='desconocida'"
    logs(1) = "2. Act: ResolverScopeRefrescoCalidadRiesgo"
    logs(2) = "3. Assert: fallback='risk'"

    Dim m_Scope As String
    m_Scope = ResolverScopeRefrescoCalidadRiesgo("desconocida")
    If StrComp(m_Scope, "risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_QualityScopeResolver_UnknownFallsRisk = BuildFail("ResolverScopeRefrescoCalidadRiesgo debe fallback a 'risk'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_QualityScopeResolver_UnknownFallsRisk = BuildOk("quality_scope_fallback_risk", logs)
    Exit Function
EH:
    Test_RiskTreeSync_QualityScopeResolver_UnknownFallsRisk = BuildFail("Test_RiskTreeSync_QualityScopeResolver_UnknownFallsRisk: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_FallbackKey_PrioritizesRequested() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: requested='PM|10|20', parent='RIESGO|10', root='EDICION|1'"
    logs(1) = "2. Act: ResolverClaveFallbackArbolRiesgos con requested existente"
    logs(2) = "3. Assert: devuelve requested"

    Dim m_Result As String
    m_Result = ResolverClaveFallbackArbolRiesgos("PM|10|20", "PM|10|20" & vbLf & "RIESGO|10" & vbLf & "EDICION|1")
    If StrComp(m_Result, "PM|10|20", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_FallbackKey_PrioritizesRequested = BuildFail("Debe priorizar requested cuando existe en candidatos", logs)
        Exit Function
    End If

    Test_RiskTreeSync_FallbackKey_PrioritizesRequested = BuildOk("fallback_requested", logs)
    Exit Function
EH:
    Test_RiskTreeSync_FallbackKey_PrioritizesRequested = BuildFail("Test_RiskTreeSync_FallbackKey_PrioritizesRequested: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_FallbackKey_FallsToParent() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: requested vacio, parent='RIESGO|10', root='EDICION|1'"
    logs(1) = "2. Act: ResolverClaveFallbackArbolRiesgos"
    logs(2) = "3. Assert: devuelve parent"

    Dim m_Result As String
    m_Result = ResolverClaveFallbackArbolRiesgos("", "RIESGO|10" & vbLf & "EDICION|1")
    If StrComp(m_Result, "RIESGO|10", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_FallbackKey_FallsToParent = BuildFail("Debe caer a parent cuando requested no está", logs)
        Exit Function
    End If

    Test_RiskTreeSync_FallbackKey_FallsToParent = BuildOk("fallback_parent", logs)
    Exit Function
EH:
    Test_RiskTreeSync_FallbackKey_FallsToParent = BuildFail("Test_RiskTreeSync_FallbackKey_FallsToParent: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_FallbackKey_FallsToRoot() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: requested y parent vacios, root='EDICION|1'"
    logs(1) = "2. Act: ResolverClaveFallbackArbolRiesgos"
    logs(2) = "3. Assert: devuelve root"

    Dim m_Result As String
    m_Result = ResolverClaveFallbackArbolRiesgos("", vbLf & "EDICION|1")
    If StrComp(m_Result, "EDICION|1", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_FallbackKey_FallsToRoot = BuildFail("Debe caer a root cuando requested/parent no están", logs)
        Exit Function
    End If

    Test_RiskTreeSync_FallbackKey_FallsToRoot = BuildOk("fallback_root", logs)
    Exit Function
EH:
    Test_RiskTreeSync_FallbackKey_FallsToRoot = BuildFail("Test_RiskTreeSync_FallbackKey_FallsToRoot: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_SelectionResolver_FallsToExistingAncestor() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    Dim resolver As RiskTreeSelectionResolver
    Dim candidates As String
    Dim existingKeys As String
    Dim result As String
    Dim errMsg As String

    logs(0) = "1. Arrange: requested action is unavailable and parent plan exists"
    logs(1) = "2. Act: RiskTreeSelectionResolver.ResolveFallbackKey"
    logs(2) = "3. Assert: selected key falls back to the existing ancestor plan"

    Set resolver = New RiskTreeSelectionResolver
    candidates = "PMACCION|20|30" & vbLf & "PM|10|20" & vbLf & "RIESGO|10" & vbLf & "EDICION|1"
    existingKeys = "PM|10|20" & vbLf & "RIESGO|10" & vbLf & "EDICION|1"
    result = resolver.ResolveFallbackKey("PMACCION|20|30", candidates, existingKeys, errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_SelectionResolver_FallsToExistingAncestor = BuildFail("ResolveFallbackKey returned error: " & errMsg, logs)
        Exit Function
    End If
    If StrComp(result, "PM|10|20", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_SelectionResolver_FallsToExistingAncestor = BuildFail("Expected fallback key PM|10|20, got " & result, logs)
        Exit Function
    End If

    Test_RiskTreeSync_SelectionResolver_FallsToExistingAncestor = BuildOk("selection_resolver_existing_ancestor", logs)
    Exit Function
EH:
    Test_RiskTreeSync_SelectionResolver_FallsToExistingAncestor = BuildFail("Test_RiskTreeSync_SelectionResolver_FallsToExistingAncestor: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_SelectionResolver_ReturnsBlankWithoutExistingCandidate() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    Dim resolver As RiskTreeSelectionResolver
    Dim result As String
    Dim errMsg As String

    logs(0) = "1. Arrange: requested and candidates are absent from existing keys"
    logs(1) = "2. Act: RiskTreeSelectionResolver.ResolveFallbackKey"
    logs(2) = "3. Assert: no fallback key is selected"

    Set resolver = New RiskTreeSelectionResolver
    result = resolver.ResolveFallbackKey("PMACCION|20|30", "PMACCION|20|30" & vbLf & "PM|10|20", "EDICION|1", errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_SelectionResolver_ReturnsBlankWithoutExistingCandidate = BuildFail("ResolveFallbackKey returned error: " & errMsg, logs)
        Exit Function
    End If
    If result <> "" Then
        Test_RiskTreeSync_SelectionResolver_ReturnsBlankWithoutExistingCandidate = BuildFail("Expected blank fallback when no candidate exists, got " & result, logs)
        Exit Function
    End If

    Test_RiskTreeSync_SelectionResolver_ReturnsBlankWithoutExistingCandidate = BuildOk("selection_resolver_blank_without_existing_candidate", logs)
    Exit Function
EH:
    Test_RiskTreeSync_SelectionResolver_ReturnsBlankWithoutExistingCandidate = BuildFail("Test_RiskTreeSync_SelectionResolver_ReturnsBlankWithoutExistingCandidate: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim service As RiskTreeRefreshService
    Dim plan As Scripting.Dictionary
    Dim candidates As String
    Dim existingKeys As String
    Dim errMsg As String

    logs(0) = "1. Arrange: full refresh with requested action unavailable"
    logs(1) = "2. Act: RiskTreeRefreshService.BuildPlan"
    logs(2) = "3. Assert: normalized scope and invalidation order are full-refresh contracts"
    logs(3) = "4. Assert: fullReload flag is true"
    logs(4) = "5. Assert: fallback key resolves without UI objects"

    Set service = New RiskTreeRefreshService
    candidates = "PMACCION|20|30" & vbLf & "PM|10|20" & vbLf & "RIESGO|10" & vbLf & "EDICION|1"
    existingKeys = "PM|10|20" & vbLf & "RIESGO|10" & vbLf & "EDICION|1"
    Set plan = service.BuildPlan("full", "1", "10", "PMACCION|20|30", candidates, existingKeys, errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback = BuildFail("BuildPlan returned error: " & errMsg, logs)
        Exit Function
    End If
    If StrComp(CStr(plan("normalizedScope")), "full", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback = BuildFail("Expected normalizedScope=full", logs)
        Exit Function
    End If
    If StrComp(CStr(plan("invalidationOrder")), "plan|risk|edition", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback = BuildFail("Expected invalidationOrder=plan|risk|edition", logs)
        Exit Function
    End If
    If CBool(plan("fullReload")) <> True Then
        Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback = BuildFail("Expected fullReload=True", logs)
        Exit Function
    End If
    If StrComp(CStr(plan("fallbackKey")), "PM|10|20", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback = BuildFail("Expected fallbackKey=PM|10|20", logs)
        Exit Function
    End If

    Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback = BuildOk("refresh_service_full_plan", logs)
    Exit Function
EH:
    Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback = BuildFail("Test_RiskTreeSync_RefreshService_BuildPlanFullReloadFallback: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim service As RiskTreeRefreshService
    Dim plan As Scripting.Dictionary
    Dim errMsg As String

    logs(0) = "1. Arrange: unknown scope and existing requested risk key"
    logs(1) = "2. Act: RiskTreeRefreshService.BuildPlan"
    logs(2) = "3. Assert: normalized scope defaults to risk with risk-only invalidation"
    logs(3) = "4. Assert: fullReload flag is false"

    Set service = New RiskTreeRefreshService
    Set plan = service.BuildPlan("unknown", "1", "10", "RIESGO|10", "RIESGO|10" & vbLf & "EDICION|1", "RIESGO|10" & vbLf & "EDICION|1", errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk = BuildFail("BuildPlan returned error: " & errMsg, logs)
        Exit Function
    End If
    If StrComp(CStr(plan("normalizedScope")), "risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk = BuildFail("Expected normalizedScope=risk", logs)
        Exit Function
    End If
    If StrComp(CStr(plan("invalidationOrder")), "risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk = BuildFail("Expected invalidationOrder=risk", logs)
        Exit Function
    End If
    If CBool(plan("fullReload")) <> False Then
        Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk = BuildFail("Expected fullReload=False", logs)
        Exit Function
    End If

    Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk = BuildOk("refresh_service_unknown_scope_defaults_risk", logs)
    Exit Function
EH:
    Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk = BuildFail("Test_RiskTreeSync_RefreshService_BuildPlanUnknownScopeDefaultsRisk: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim errMsg As String
    Dim normalizedScope As String
    Dim order As String
    Dim fallbackKey As String

    logs(0) = "1. Arrange: wrapper inputs for refresh scope and fallback selection"
    logs(1) = "2. Act: NormalizarScopeRefrescoArbolRiesgos"
    logs(2) = "3. Act: ObtenerOrdenInvalidacionRiesgosPorScope"
    logs(3) = "4. Act: ResolverClaveFallbackArbolRiesgosConExistentes"
    logs(4) = "5. Assert: wrapper-compatible outputs are preserved"

    normalizedScope = NormalizarScopeRefrescoArbolRiesgos("ACTION", errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildFail("NormalizarScopeRefrescoArbolRiesgos returned error: " & errMsg, logs)
        Exit Function
    End If
    order = ObtenerOrdenInvalidacionRiesgosPorScope(normalizedScope, errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildFail("ObtenerOrdenInvalidacionRiesgosPorScope returned error: " & errMsg, logs)
        Exit Function
    End If
    fallbackKey = ResolverClaveFallbackArbolRiesgosConExistentes("PMACCION|20|30", "PMACCION|20|30" & vbLf & "PM|10|20", "PM|10|20", errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildFail("ResolverClaveFallbackArbolRiesgosConExistentes returned error: " & errMsg, logs)
        Exit Function
    End If
    If StrComp(normalizedScope, "action", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildFail("Expected normalized scope action", logs)
        Exit Function
    End If
    If StrComp(order, "plan|risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildFail("Expected invalidation order plan|risk", logs)
        Exit Function
    End If
    If StrComp(fallbackKey, "PM|10|20", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildFail("Expected fallback key PM|10|20", logs)
        Exit Function
    End If

    Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildOk("wrappers_route_to_services", logs)
    Exit Function
EH:
    Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices = BuildFail("Test_RiskTreeSync_Wrappers_RouteToRefreshAndSelectionServices: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ManualRefreshScope_IsFull() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: refresh manual desde ComandoActualizarContador"
    logs(1) = "2. Act: ResolverScopeRefrescoManualRiesgos"
    logs(2) = "3. Assert: scope canonical='full'"

    Dim m_Error As String
    Dim m_Scope As String

    m_Error = ""
    m_Scope = ResolverScopeRefrescoManualRiesgos(m_Error)
    If m_Error <> "" Then
        Test_RiskTreeSync_ManualRefreshScope_IsFull = BuildFail("ResolverScopeRefrescoManualRiesgos devolvio error: " & m_Error, logs)
        Exit Function
    End If

    If StrComp(m_Scope, "full", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ManualRefreshScope_IsFull = BuildFail("El scope manual debe ser 'full' para refresco canonical", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ManualRefreshScope_IsFull = BuildOk("manual_refresh_scope_full", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ManualRefreshScope_IsFull = BuildFail("Test_RiskTreeSync_ManualRefreshScope_IsFull: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ManualRefreshScope_UsesFullInvalidationOrder() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: scope manual canonical"
    logs(1) = "2. Act: ObtenerOrdenInvalidacionRiesgosPorScope(scope manual)"
    logs(2) = "3. Assert: orden='plan|risk|edition'"

    Dim m_Error As String
    Dim m_Scope As String
    Dim m_Orden As String

    m_Error = ""
    m_Scope = ResolverScopeRefrescoManualRiesgos(m_Error)
    If m_Error <> "" Then
        Test_RiskTreeSync_ManualRefreshScope_UsesFullInvalidationOrder = BuildFail("ResolverScopeRefrescoManualRiesgos devolvio error: " & m_Error, logs)
        Exit Function
    End If

    m_Orden = ObtenerOrdenInvalidacionRiesgosPorScope(m_Scope, m_Error)
    If m_Error <> "" Then
        Test_RiskTreeSync_ManualRefreshScope_UsesFullInvalidationOrder = BuildFail("ObtenerOrdenInvalidacionRiesgosPorScope devolvio error: " & m_Error, logs)
        Exit Function
    End If

    If StrComp(m_Orden, "plan|risk|edition", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ManualRefreshScope_UsesFullInvalidationOrder = BuildFail("El refresco manual debe invalidar plan|risk|edition", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ManualRefreshScope_UsesFullInvalidationOrder = BuildOk("manual_refresh_full_invalidation", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ManualRefreshScope_UsesFullInvalidationOrder = BuildFail("Test_RiskTreeSync_ManualRefreshScope_UsesFullInvalidationOrder: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ManualDetailRefreshRoute_IsRiskPunctual() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: refresh manual desde Form_FormRiesgo.ComandoActualizar"
    logs(1) = "2. Act: contrato de refresco puntual de detalle"
    logs(2) = "3. Assert: route='risk'"

    Dim m_Error As String
    Dim m_Route As String

    m_Error = ""
    m_Route = ResolverRutaRefrescoDetalleRiesgo(True, m_Error)
    If m_Error <> "" Then
        Test_RiskTreeSync_ManualDetailRefreshRoute_IsRiskPunctual = BuildFail("ResolverRutaRefrescoDetalleRiesgo devolvio error: " & m_Error, logs)
        Exit Function
    End If
    If StrComp(m_Route, "risk", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ManualDetailRefreshRoute_IsRiskPunctual = BuildFail("El refresco manual del detalle debe enrutar por 'risk'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ManualDetailRefreshRoute_IsRiskPunctual = BuildOk("manual_detail_refresh_scope_risk", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ManualDetailRefreshRoute_IsRiskPunctual = BuildFail("Test_RiskTreeSync_ManualDetailRefreshRoute_IsRiskPunctual: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_ManualDetailRefreshRoute_FallbackSupportsFull() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: nodo actual de riesgo no resoluble"
    logs(1) = "2. Act: resolver route de fallback"
    logs(2) = "3. Assert: fallback='full-fallback'"

    Dim m_Error As String
    Dim m_Route As String

    m_Error = ""
    m_Route = ResolverRutaRefrescoDetalleRiesgo(False, m_Error)
    If m_Error <> "" Then
        Test_RiskTreeSync_ManualDetailRefreshRoute_FallbackSupportsFull = BuildFail("ResolverRutaRefrescoDetalleRiesgo devolvio error: " & m_Error, logs)
        Exit Function
    End If
    If StrComp(m_Route, "full-fallback", vbTextCompare) <> 0 Then
        Test_RiskTreeSync_ManualDetailRefreshRoute_FallbackSupportsFull = BuildFail("Cuando no hay nodo válido, el detalle debe caer a 'full-fallback'", logs)
        Exit Function
    End If

    Test_RiskTreeSync_ManualDetailRefreshRoute_FallbackSupportsFull = BuildOk("manual_detail_refresh_full_fallback", logs)
    Exit Function
EH:
    Test_RiskTreeSync_ManualDetailRefreshRoute_FallbackSupportsFull = BuildFail("Test_RiskTreeSync_ManualDetailRefreshRoute_FallbackSupportsFull: " & Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_PublicableIcon_MissingOptionalMaterializationPlan() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim riesgoFixture As riesgo
    Dim lineas As Scripting.Dictionary
    Dim errMsg As String

    logs(0) = "1. Arrange: materialized risk with active PM and active PC"
    logs(1) = "2. Arrange: current materialization has no associated contingency plan"
    logs(2) = "3. Act: getColLineasPublicabilidadRiesgo"
    logs(3) = "4. Assert: legacy tree/publicability lines do not include red missing-plan line"
    logs(4) = "5. Assert: optional missing association is not non-publicable"

    Set riesgoFixture = BuildIssue26RiskTreeRisk(False)
    Set lineas = getColLineasPublicabilidadRiesgo(riesgoFixture, errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_PublicableIcon_MissingOptionalMaterializationPlan = _
            BuildFail("getColLineasPublicabilidadRiesgo devolvio error: " & errMsg, logs)
        Exit Function
    End If
    If PublicabilityLinesContain(lineas, _
        "Materialización con plan de contingencia asociado: No||Rojo") Then
        Test_RiskTreeSync_PublicableIcon_MissingOptionalMaterializationPlan = _
            BuildFail("La publicabilidad legacy no debe marcar rojo si falta el plan asociado opcional", logs)
        Exit Function
    End If

    Test_RiskTreeSync_PublicableIcon_MissingOptionalMaterializationPlan = _
        BuildOk("risk_tree_missing_optional_materialization_plan_not_red", logs)
    Exit Function

EH:
    Test_RiskTreeSync_PublicableIcon_MissingOptionalMaterializationPlan = _
        BuildFail("Test_RiskTreeSync_PublicableIcon_MissingOptionalMaterializationPlan: " & _
            Err.Description, logs)
End Function

Public Function Test_RiskTreeSync_NonPublicableIcon_WithMaterializationPlan() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim riesgoFixture As riesgo
    Dim lineas As Scripting.Dictionary
    Dim errMsg As String

    logs(0) = "1. Arrange: materialized risk with active PM and active PC"
    logs(1) = "2. Arrange: current materialization has associated contingency plan"
    logs(2) = "3. Act: getColLineasPublicabilidadRiesgo"
    logs(3) = "4. Assert: legacy tree/publicability lines include positive associated-plan line"
    logs(4) = "5. Assert: missing-plan red line is absent"

    Set riesgoFixture = BuildIssue26RiskTreeRisk(True)
    Set lineas = getColLineasPublicabilidadRiesgo(riesgoFixture, errMsg)
    If errMsg <> "" Then
        Test_RiskTreeSync_NonPublicableIcon_WithMaterializationPlan = _
            BuildFail("getColLineasPublicabilidadRiesgo devolvio error: " & errMsg, logs)
        Exit Function
    End If
    If PublicabilityLinesContain(lineas, _
        "Materialización con plan de contingencia asociado: No||Rojo") Then
        Test_RiskTreeSync_NonPublicableIcon_WithMaterializationPlan = _
            BuildFail("No debe marcar rojo cuando la materializacion vigente tiene plan", logs)
        Exit Function
    End If
    If Not PublicabilityLinesContain(lineas, _
        "Materialización con plan de contingencia asociado: Sí") Then
        Test_RiskTreeSync_NonPublicableIcon_WithMaterializationPlan = _
            BuildFail("Debe informar que la materializacion tiene plan asociado", logs)
        Exit Function
    End If

    Test_RiskTreeSync_NonPublicableIcon_WithMaterializationPlan = _
        BuildOk("risk_tree_materialization_plan_ok", logs)
    Exit Function

EH:
    Test_RiskTreeSync_NonPublicableIcon_WithMaterializationPlan = _
        BuildFail("Test_RiskTreeSync_NonPublicableIcon_WithMaterializationPlan: " & _
            Err.Description, logs)
End Function

Private Function BuildIssue26RiskTreeRisk( _
    ByVal hasAssociatedPlan As Boolean _
) As riesgo
    Dim riesgoFixture As riesgo
    Dim colPMs As Scripting.Dictionary
    Dim colPCs As Scripting.Dictionary
    Dim colMaterializaciones As Scripting.Dictionary

    Set riesgoFixture = New riesgo
    With riesgoFixture
        .IDRiesgo = "926001"
        .CodigoRiesgo = "I26"
        .Descripcion = "Fixture issue 26 icon propagation"
        .CausaRaiz = "Fixture"
        .Priorizacion = "3"
        .Valoracion = "Medio"
        .FechaMaterializado = CStr(Date)
        .RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No
    End With

    Set colPMs = New Scripting.Dictionary
    colPMs.Add "pm", BuildIssue26ActivePM()
    Set riesgoFixture.ColPMs = colPMs

    Set colPCs = New Scripting.Dictionary
    colPCs.Add "pc", BuildIssue26ActivePC()
    Set riesgoFixture.ColPCs = colPCs

    Set colMaterializaciones = New Scripting.Dictionary
    colMaterializaciones.Add "old", BuildIssue26Materialization( _
        "926010", DateAdd("d", -1, Date), "910212")
    colMaterializaciones.Add "current", BuildIssue26Materialization( _
        "926011", Date, IIf(hasAssociatedPlan, "910212", ""))
    Set riesgoFixture.ColMaterializaciones = colMaterializaciones

    Set BuildIssue26RiskTreeRisk = riesgoFixture
End Function

Private Function BuildIssue26ActivePM() As PM
    Dim pmFixture As PM
    Dim accion As PMAccion
    Dim acciones As Scripting.Dictionary

    Set pmFixture = New PM
    Set accion = New PMAccion
    accion.FechaInicio = CStr(Date)
    Set acciones = New Scripting.Dictionary
    acciones.Add "accion", accion
    Set pmFixture.colAcciones = acciones
    Set BuildIssue26ActivePM = pmFixture
End Function

Private Function BuildIssue26ActivePC() As PC
    Dim pcFixture As PC
    Dim accion As PCAccion
    Dim acciones As Scripting.Dictionary

    Set pcFixture = New PC
    Set accion = New PCAccion
    accion.FechaInicio = CStr(Date)
    Set acciones = New Scripting.Dictionary
    acciones.Add "accion", accion
    Set pcFixture.colAcciones = acciones
    Set BuildIssue26ActivePC = pcFixture
End Function

Private Function BuildIssue26Materialization( _
    ByVal materializacionId As String, _
    ByVal materializacionDate As Date, _
    ByVal planId As String _
) As RiesgoMaterializacion
    Dim materializacion As RiesgoMaterializacion

    Set materializacion = New RiesgoMaterializacion
    materializacion.ID = materializacionId
    materializacion.Fecha = CStr(materializacionDate)
    materializacion.EsMaterializacion = "Sí"
    materializacion.IDPlanContingencia = planId
    Set BuildIssue26Materialization = materializacion
End Function

Private Function PublicabilityLinesContain( _
    ByVal lineas As Scripting.Dictionary, _
    ByVal expectedText As String _
) As Boolean
    Dim lineKey As Variant

    If lineas Is Nothing Then Exit Function
    For Each lineKey In lineas
        If InStr(1, CStr(lineas(lineKey)), expectedText, vbTextCompare) > 0 Then
            PublicabilityLinesContain = True
            Exit Function
        End If
    Next lineKey
End Function
