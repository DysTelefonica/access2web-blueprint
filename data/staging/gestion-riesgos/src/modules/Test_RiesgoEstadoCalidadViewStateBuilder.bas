Attribute VB_Name = "Test_RiesgoEstadoCalidadViewStateBuilder"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: RiesgoEstadoCalidadViewStateBuilder (issue #43)
' Pure logic only. No DB, no UI, no fixtures.
' Returns JSON {ok, value, payload, error, logs}
' ============================================================

Public Function Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: accepted flow states with active non-technician risk"
    logs(1) = "2. Act: Build acceptance quality button view states"
    logs(2) = "3. Assert: captions and enabled flags match legacy form logic"

    Dim errMsg As String
    Dim builder As RiesgoEstadoCalidadViewStateBuilder
    Set builder = New RiesgoEstadoCalidadViewStateBuilder

    If Not AssertAceptacionState(builder, EnumRiesgoEstado.AceptadoSinJustificar, "Aceptar", "Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertAceptacionState(builder, EnumRiesgoEstado.AceptadoSinVisar, "Aceptar", "Rechazar", True, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertAceptacionState(builder, EnumRiesgoEstado.Aceptado, "Quitar Aceptar", "Rechazar", True, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertAceptacionState(builder, EnumRiesgoEstado.AceptadoRechazado, "Aceptar", "Quitar Rechazar", False, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad = BuildJsonOk("aceptacion_estados_calidad_ok", logs)
    Exit Function

EH:
    Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad = BuildJsonFail("Test_RiesgoEstadoCalidadViewStateBuilder_AceptacionEstadosCalidad: " & Err.Description, logs)
End Function

Public Function Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: retired flow states with active non-technician risk"
    logs(1) = "2. Act: Build retirement quality button view states"
    logs(2) = "3. Assert: captions and enabled flags match legacy form logic"

    Dim errMsg As String
    Dim builder As RiesgoEstadoCalidadViewStateBuilder
    Set builder = New RiesgoEstadoCalidadViewStateBuilder

    If Not AssertRetiradaState(builder, EnumRiesgoEstado.RetiradoSinJustificar, "Aceptar", "Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaState(builder, EnumRiesgoEstado.RetiradoSinVisar, "Aceptar", "Rechazar", True, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaState(builder, EnumRiesgoEstado.Retirado, "Quitar Aceptar", "Rechazar", True, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaState(builder, EnumRiesgoEstado.RetiradoRechazado, "Aceptar", "Quitar Rechazar", False, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad = BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad = BuildJsonOk("retirada_estados_calidad_ok", logs)
    Exit Function

EH:
    Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad = BuildJsonFail("Test_RiesgoEstadoCalidadViewStateBuilder_RetiradaEstadosCalidad: " & Err.Description, logs)
End Function

Public Function Test_RiesgoEstadoCalidadViewStateBuilder_TecnicoOcultaBotones() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: issue-43 technician quality scenarios"
    logs(1) = "2. Act: Build acceptance and retirement states with EsTecnico = Sí"
    logs(2) = "3. Assert: legacy quality buttons remain hidden without UI/DB/forms"

    Dim errMsg As String
    Dim builder As RiesgoEstadoCalidadViewStateBuilder
    Set builder = New RiesgoEstadoCalidadViewStateBuilder

    If Not AssertAceptacionHiddenState(builder, EnumRiesgoEstado.AceptadoSinVisar, EnumSiNo.Sí, True, True, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_TecnicoOcultaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaHiddenState(builder, EnumRiesgoEstado.RetiradoSinVisar, EnumSiNo.Sí, True, True, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_TecnicoOcultaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Test_RiesgoEstadoCalidadViewStateBuilder_TecnicoOcultaBotones = BuildJsonOk("issue_43_tecnico_oculta_botones_ok", logs)
    Exit Function

EH:
    Test_RiesgoEstadoCalidadViewStateBuilder_TecnicoOcultaBotones = BuildJsonFail("Test_RiesgoEstadoCalidadViewStateBuilder_TecnicoOcultaBotones: " & Err.Description, logs)
End Function

Public Function Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: issue-43 inactive risk quality states"
    logs(1) = "2. Act: Build acceptance and retirement states with p_RiesgoActivo = False"
    logs(2) = "3. Assert: legacy buttons stay visible but disabled"

    Dim errMsg As String
    Dim builder As RiesgoEstadoCalidadViewStateBuilder
    Set builder = New RiesgoEstadoCalidadViewStateBuilder

    If Not AssertAceptacionStateWithActive(builder, EnumRiesgoEstado.AceptadoSinVisar, False, "Aceptar", "Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertAceptacionStateWithActive(builder, EnumRiesgoEstado.Aceptado, False, "Quitar Aceptar", "Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertAceptacionStateWithActive(builder, EnumRiesgoEstado.AceptadoRechazado, False, "Aceptar", "Quitar Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaStateWithActive(builder, EnumRiesgoEstado.RetiradoSinVisar, False, "Aceptar", "Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaStateWithActive(builder, EnumRiesgoEstado.Retirado, False, "Quitar Aceptar", "Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaStateWithActive(builder, EnumRiesgoEstado.RetiradoRechazado, False, "Aceptar", "Quitar Rechazar", False, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonOk("issue_43_riesgo_inactivo_deshabilita_botones_ok", logs)
    Exit Function

EH:
    Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones = BuildJsonFail("Test_RiesgoEstadoCalidadViewStateBuilder_RiesgoInactivoDeshabilitaBotones: " & Err.Description, logs)
End Function

Public Function Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: issue-43 missing initial risk/date scenarios"
    logs(1) = "2. Act: Build states without initial risk or registered quality date"
    logs(2) = "3. Assert: legacy quality buttons remain hidden"

    Dim errMsg As String
    Dim builder As RiesgoEstadoCalidadViewStateBuilder
    Set builder = New RiesgoEstadoCalidadViewStateBuilder

    If Not AssertAceptacionHiddenState(builder, EnumRiesgoEstado.AceptadoSinVisar, EnumSiNo.No, True, False, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertAceptacionHiddenState(builder, EnumRiesgoEstado.AceptadoSinVisar, EnumSiNo.No, True, True, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaHiddenState(builder, EnumRiesgoEstado.RetiradoSinVisar, EnumSiNo.No, True, False, True, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaHiddenState(builder, EnumRiesgoEstado.RetiradoSinVisar, EnumSiNo.No, True, True, False, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones = BuildJsonOk("issue_43_sin_inicial_o_sin_fecha_oculta_botones_ok", logs)
    Exit Function

EH:
    Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones = BuildJsonFail("Test_RiesgoEstadoCalidadViewStateBuilder_SinInicialOSinFechaOcultaBotones: " & Err.Description, logs)
End Function

Public Function Test_RiesgoEstadoCalidadViewStateBuilder_EstadosCruzadosSinDecisionBotones() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: issue-43 crossed acceptance/retirement states"
    logs(1) = "2. Act: Build acceptance with retirement state and retirement with acceptance state"
    logs(2) = "3. Assert: buttons are visible but no caption/enabled decision is emitted"

    Dim errMsg As String
    Dim builder As RiesgoEstadoCalidadViewStateBuilder
    Set builder = New RiesgoEstadoCalidadViewStateBuilder

    If Not AssertAceptacionVisibleWithoutDecision(builder, EnumRiesgoEstado.RetiradoSinVisar, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_EstadosCruzadosSinDecisionBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not AssertRetiradaVisibleWithoutDecision(builder, EnumRiesgoEstado.AceptadoSinVisar, errMsg) Then
        Test_RiesgoEstadoCalidadViewStateBuilder_EstadosCruzadosSinDecisionBotones = BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Test_RiesgoEstadoCalidadViewStateBuilder_EstadosCruzadosSinDecisionBotones = BuildJsonOk("issue_43_estados_cruzados_sin_decision_botones_ok", logs)
    Exit Function

EH:
    Test_RiesgoEstadoCalidadViewStateBuilder_EstadosCruzadosSinDecisionBotones = BuildJsonFail("Test_RiesgoEstadoCalidadViewStateBuilder_EstadosCruzadosSinDecisionBotones: " & Err.Description, logs)
End Function

Private Function AssertAceptacionState( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByVal p_AceptarCaption As String, _
    ByVal p_RechazarCaption As String, _
    ByVal p_AceptarEnabled As Boolean, _
    ByVal p_RechazarEnabled As Boolean, _
    ByRef p_Error As String _
) As Boolean
    AssertAceptacionState = AssertAceptacionStateWithActive( _
        p_Builder, _
        p_Estado, _
        True, _
        p_AceptarCaption, _
        p_RechazarCaption, _
        p_AceptarEnabled, _
        p_RechazarEnabled, _
        p_Error)
End Function

Private Function AssertRetiradaState( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByVal p_AceptarCaption As String, _
    ByVal p_RechazarCaption As String, _
    ByVal p_AceptarEnabled As Boolean, _
    ByVal p_RechazarEnabled As Boolean, _
    ByRef p_Error As String _
) As Boolean
    AssertRetiradaState = AssertRetiradaStateWithActive( _
        p_Builder, _
        p_Estado, _
        True, _
        p_AceptarCaption, _
        p_RechazarCaption, _
        p_AceptarEnabled, _
        p_RechazarEnabled, _
        p_Error)
End Function

Private Function AssertAceptacionStateWithActive( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByVal p_RiesgoActivo As Boolean, _
    ByVal p_AceptarCaption As String, _
    ByVal p_RechazarCaption As String, _
    ByVal p_AceptarEnabled As Boolean, _
    ByVal p_RechazarEnabled As Boolean, _
    ByRef p_Error As String _
) As Boolean
    Dim state As RiesgoEstadoCalidadViewState
    p_Error = ""
    Set state = p_Builder.BuildAceptacionCalidad(p_Estado, EnumSiNo.No, p_RiesgoActivo, True, True, p_Error)
    AssertAceptacionStateWithActive = AssertButtonState(state, p_AceptarCaption, p_RechazarCaption, p_AceptarEnabled, p_RechazarEnabled, p_Error)
End Function

Private Function AssertRetiradaStateWithActive( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByVal p_RiesgoActivo As Boolean, _
    ByVal p_AceptarCaption As String, _
    ByVal p_RechazarCaption As String, _
    ByVal p_AceptarEnabled As Boolean, _
    ByVal p_RechazarEnabled As Boolean, _
    ByRef p_Error As String _
) As Boolean
    Dim state As RiesgoEstadoCalidadViewState
    p_Error = ""
    Set state = p_Builder.BuildRetiradaCalidad(p_Estado, EnumSiNo.No, p_RiesgoActivo, True, True, p_Error)
    AssertRetiradaStateWithActive = AssertButtonState(state, p_AceptarCaption, p_RechazarCaption, p_AceptarEnabled, p_RechazarEnabled, p_Error)
End Function

Private Function AssertAceptacionHiddenState( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByVal p_EsTecnico As EnumSiNo, _
    ByVal p_RiesgoActivo As Boolean, _
    ByVal p_HayRiesgoInicial As Boolean, _
    ByVal p_FechaRegistrada As Boolean, _
    ByRef p_Error As String _
) As Boolean
    Dim state As RiesgoEstadoCalidadViewState
    p_Error = ""
    Set state = p_Builder.BuildAceptacionCalidad(p_Estado, p_EsTecnico, p_RiesgoActivo, p_HayRiesgoInicial, p_FechaRegistrada, p_Error)
    AssertAceptacionHiddenState = AssertHiddenWithoutDecision(state, p_Error)
End Function

Private Function AssertRetiradaHiddenState( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByVal p_EsTecnico As EnumSiNo, _
    ByVal p_RiesgoActivo As Boolean, _
    ByVal p_HayRiesgoInicial As Boolean, _
    ByVal p_FechaRegistrada As Boolean, _
    ByRef p_Error As String _
) As Boolean
    Dim state As RiesgoEstadoCalidadViewState
    p_Error = ""
    Set state = p_Builder.BuildRetiradaCalidad(p_Estado, p_EsTecnico, p_RiesgoActivo, p_HayRiesgoInicial, p_FechaRegistrada, p_Error)
    AssertRetiradaHiddenState = AssertHiddenWithoutDecision(state, p_Error)
End Function

Private Function AssertAceptacionVisibleWithoutDecision( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByRef p_Error As String _
) As Boolean
    Dim state As RiesgoEstadoCalidadViewState
    p_Error = ""
    Set state = p_Builder.BuildAceptacionCalidad(p_Estado, EnumSiNo.No, True, True, True, p_Error)
    AssertAceptacionVisibleWithoutDecision = AssertVisibleWithoutDecision(state, p_Error)
End Function

Private Function AssertRetiradaVisibleWithoutDecision( _
    ByVal p_Builder As RiesgoEstadoCalidadViewStateBuilder, _
    ByVal p_Estado As EnumRiesgoEstado, _
    ByRef p_Error As String _
) As Boolean
    Dim state As RiesgoEstadoCalidadViewState
    p_Error = ""
    Set state = p_Builder.BuildRetiradaCalidad(p_Estado, EnumSiNo.No, True, True, True, p_Error)
    AssertRetiradaVisibleWithoutDecision = AssertVisibleWithoutDecision(state, p_Error)
End Function

Private Function AssertButtonState( _
    ByVal p_State As RiesgoEstadoCalidadViewState, _
    ByVal p_AceptarCaption As String, _
    ByVal p_RechazarCaption As String, _
    ByVal p_AceptarEnabled As Boolean, _
    ByVal p_RechazarEnabled As Boolean, _
    ByRef p_Error As String _
) As Boolean
    If p_Error <> "" Then
        Exit Function
    End If
    If Not p_State.ComandoAceptarVisible Or Not p_State.ComandoRechazarVisible Then
        p_Error = "quality buttons must be visible"
        Exit Function
    End If
    If Not p_State.HasButtonDecision Then
        p_Error = "state must include a button decision"
        Exit Function
    End If
    If p_State.ComandoAceptarCaption <> p_AceptarCaption Then
        p_Error = "unexpected accept caption: " & p_State.ComandoAceptarCaption
        Exit Function
    End If
    If p_State.ComandoRechazarCaption <> p_RechazarCaption Then
        p_Error = "unexpected reject caption: " & p_State.ComandoRechazarCaption
        Exit Function
    End If
    If p_State.ComandoAceptarEnabled <> p_AceptarEnabled Then
        p_Error = "unexpected accept enabled flag"
        Exit Function
    End If
    If p_State.ComandoRechazarEnabled <> p_RechazarEnabled Then
        p_Error = "unexpected reject enabled flag"
        Exit Function
    End If

    AssertButtonState = True
End Function

Private Function AssertHiddenWithoutDecision( _
    ByVal p_State As RiesgoEstadoCalidadViewState, _
    ByRef p_Error As String _
) As Boolean
    If p_Error <> "" Then
        Exit Function
    End If
    If p_State.ComandoAceptarVisible Or p_State.ComandoRechazarVisible Then
        p_Error = "quality buttons must be hidden"
        Exit Function
    End If
    If p_State.HasButtonDecision Then
        p_Error = "hidden state must not include a button decision"
        Exit Function
    End If

    AssertHiddenWithoutDecision = True
End Function

Private Function AssertVisibleWithoutDecision( _
    ByVal p_State As RiesgoEstadoCalidadViewState, _
    ByRef p_Error As String _
) As Boolean
    If p_Error <> "" Then
        Exit Function
    End If
    If Not p_State.ComandoAceptarVisible Or Not p_State.ComandoRechazarVisible Then
        p_Error = "crossed quality buttons must stay visible"
        Exit Function
    End If
    If p_State.HasButtonDecision Then
        p_Error = "crossed state must not include a button decision"
        Exit Function
    End If

    AssertVisibleWithoutDecision = True
End Function
