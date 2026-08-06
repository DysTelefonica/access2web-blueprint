VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoEstadoCalidadViewStateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' RiesgoEstadoCalidadViewStateBuilder (issue #43)
' Pure button-state decisions for acceptance/retirement quality flows.
' No Access UI, no DAO, no cache, no prompts and no persistence.
' ============================================================

Public Function BuildAceptacionCalidad( _
    ByVal p_EstadoRiesgo As EnumRiesgoEstado, _
    ByVal p_EsTecnico As EnumSiNo, _
    ByVal p_RiesgoActivo As Boolean, _
    ByVal p_HayRiesgoInicial As Boolean, _
    ByVal p_FechaMitigacionAceptarRegistrada As Boolean, _
    Optional ByRef p_Error As String _
) As RiesgoEstadoCalidadViewState
    On Error GoTo errores

    Dim state As RiesgoEstadoCalidadViewState
    Set state = New RiesgoEstadoCalidadViewState

    If p_EsTecnico = EnumSiNo.Sí Or _
        Not p_HayRiesgoInicial Or _
        Not p_FechaMitigacionAceptarRegistrada Then
        Set BuildAceptacionCalidad = state
        Exit Function
    End If

    state.ComandoAceptarVisible = True
    state.ComandoRechazarVisible = True

    Select Case p_EstadoRiesgo
        Case EnumRiesgoEstado.AceptadoSinJustificar
            SetAcceptanceDecision state, "Aceptar", "Rechazar", False, False
        Case EnumRiesgoEstado.AceptadoSinVisar
            SetAcceptanceDecision state, "Aceptar", "Rechazar", p_RiesgoActivo, p_RiesgoActivo
        Case EnumRiesgoEstado.Aceptado
            SetAcceptanceDecision state, "Quitar Aceptar", "Rechazar", p_RiesgoActivo, False
        Case EnumRiesgoEstado.AceptadoRechazado
            SetAcceptanceDecision state, "Aceptar", "Quitar Rechazar", False, p_RiesgoActivo
        Case EnumRiesgoEstado.Retirado, _
             EnumRiesgoEstado.RetiradoSinJustificar, _
             EnumRiesgoEstado.RetiradoSinVisar, _
             EnumRiesgoEstado.RetiradoRechazado
            ' Preserve legacy behavior: visible quality buttons, no caption/enabled mutation.
        Case Else
            p_Error = "El riesgo no está ni en estado de aceptación ni de retirada"
            state.ErrorMessage = p_Error
    End Select

    Set BuildAceptacionCalidad = state
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en RiesgoEstadoCalidadViewStateBuilder.BuildAceptacionCalidad" & vbNewLine & Err.Description
    End If
End Function

Public Function BuildRetiradaCalidad( _
    ByVal p_EstadoRiesgo As EnumRiesgoEstado, _
    ByVal p_EsTecnico As EnumSiNo, _
    ByVal p_RiesgoActivo As Boolean, _
    ByVal p_HayRiesgoInicial As Boolean, _
    ByVal p_FechaRetiradoRegistrada As Boolean, _
    Optional ByRef p_Error As String _
) As RiesgoEstadoCalidadViewState
    On Error GoTo errores

    Dim state As RiesgoEstadoCalidadViewState
    Set state = New RiesgoEstadoCalidadViewState

    If p_EsTecnico = EnumSiNo.Sí Or _
        Not p_HayRiesgoInicial Or _
        Not p_FechaRetiradoRegistrada Then
        Set BuildRetiradaCalidad = state
        Exit Function
    End If

    state.ComandoAceptarVisible = True
    state.ComandoRechazarVisible = True

    Select Case p_EstadoRiesgo
        Case EnumRiesgoEstado.RetiradoSinJustificar
            SetAcceptanceDecision state, "Aceptar", "Rechazar", False, False
        Case EnumRiesgoEstado.RetiradoSinVisar
            SetAcceptanceDecision state, "Aceptar", "Rechazar", p_RiesgoActivo, p_RiesgoActivo
        Case EnumRiesgoEstado.Retirado
            SetAcceptanceDecision state, "Quitar Aceptar", "Rechazar", p_RiesgoActivo, False
        Case EnumRiesgoEstado.RetiradoRechazado
            SetAcceptanceDecision state, "Aceptar", "Quitar Rechazar", False, p_RiesgoActivo
        Case EnumRiesgoEstado.Aceptado, _
             EnumRiesgoEstado.AceptadoSinJustificar, _
             EnumRiesgoEstado.AceptadoSinVisar, _
             EnumRiesgoEstado.AceptadoRechazado
            ' Preserve legacy behavior: visible quality buttons, no caption/enabled mutation.
        Case Else
            p_Error = "El riesgo no está ni en estado de aceptación ni de retirada"
            state.ErrorMessage = p_Error
    End Select

    Set BuildRetiradaCalidad = state
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en RiesgoEstadoCalidadViewStateBuilder.BuildRetiradaCalidad" & vbNewLine & Err.Description
    End If
End Function

Private Sub SetAcceptanceDecision( _
    ByVal p_State As RiesgoEstadoCalidadViewState, _
    ByVal p_AceptarCaption As String, _
    ByVal p_RechazarCaption As String, _
    ByVal p_AceptarEnabled As Boolean, _
    ByVal p_RechazarEnabled As Boolean _
)
    p_State.ComandoAceptarCaption = p_AceptarCaption
    p_State.ComandoRechazarCaption = p_RechazarCaption
    p_State.ComandoAceptarEnabled = p_AceptarEnabled
    p_State.ComandoRechazarEnabled = p_RechazarEnabled
    p_State.HasButtonDecision = True
End Sub
