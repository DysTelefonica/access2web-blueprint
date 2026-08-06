VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublicacionCalidadStatusPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' PublicacionCalidadStatusPresenter (issue #46)
' Pure visual/status decisions extracted from
' Form_FormPublicacionCalidadPublicar.EstablecerDatos.
' No UI, no DAO, no cache, no prompts and no persistence.
' ============================================================

Private Const COLOR_ROLE_COMPLEMENTADO As String = "complementado"
Private Const COLOR_ROLE_NO_COMPLEMENTADO As String = "no-complementado"

Public Function Build( _
    ByVal p_ProyectoEnUTE As Boolean, _
    ByVal p_TieneAnexoEvidenciaUTE As Boolean, _
    ByVal p_HaySuministradores As Boolean, _
    ByVal p_EvidenciasSuministradoresCompletadas As Boolean, _
    ByVal p_EsActivo As Boolean, _
    ByVal p_FechaPreparadaParaPublicar As String, _
    ByVal p_PropuestaRechazadaPorCalidadFecha As String, _
    ByVal p_FechaPublicacion As String _
) As PublicacionCalidadStatusViewState
    Dim state As PublicacionCalidadStatusViewState
    Set state = New PublicacionCalidadStatusViewState

    BuildUteState state, p_ProyectoEnUTE, p_TieneAnexoEvidenciaUTE
    BuildSuministradoresState state, p_HaySuministradores, p_EvidenciasSuministradoresCompletadas
    BuildPublicacionState state, _
            p_EsActivo, _
            p_FechaPreparadaParaPublicar, _
            p_PropuestaRechazadaPorCalidadFecha, _
            p_FechaPublicacion

    Set Build = state
End Function

Private Sub BuildUteState( _
    ByVal p_State As PublicacionCalidadStatusViewState, _
    ByVal p_ProyectoEnUTE As Boolean, _
    ByVal p_TieneAnexoEvidenciaUTE As Boolean _
)
    If p_ProyectoEnUTE Then
        p_State.UteRequeridoCaption = "Requerido"
        If p_TieneAnexoEvidenciaUTE Then
            p_State.UteComplementadoCaption = "Complementado"
            p_State.UteComplementadoColorRole = COLOR_ROLE_COMPLEMENTADO
        Else
            p_State.UteComplementadoCaption = "No complementado"
            p_State.UteComplementadoColorRole = COLOR_ROLE_NO_COMPLEMENTADO
        End If
    Else
        p_State.UteRequeridoCaption = "No Requerido"
        p_State.UteComplementadoCaption = "N/A"
        p_State.UteComplementadoColorRole = COLOR_ROLE_COMPLEMENTADO
    End If
End Sub

Private Sub BuildSuministradoresState( _
    ByVal p_State As PublicacionCalidadStatusViewState, _
    ByVal p_HaySuministradores As Boolean, _
    ByVal p_EvidenciasSuministradoresCompletadas As Boolean _
)
    If p_HaySuministradores Then
        p_State.SuministradoresRequeridoCaption = "Requerido"
        If p_EvidenciasSuministradoresCompletadas Then
            p_State.SuministradoresComplementadoCaption = "Complementado"
            p_State.SuministradoresComplementadoColorRole = COLOR_ROLE_COMPLEMENTADO
        Else
            p_State.SuministradoresComplementadoCaption = "No complementado"
            p_State.SuministradoresComplementadoColorRole = COLOR_ROLE_NO_COMPLEMENTADO
        End If
    Else
        p_State.SuministradoresRequeridoCaption = "No requerido"
        p_State.SuministradoresComplementadoCaption = "N/A"
        p_State.SuministradoresComplementadoColorRole = COLOR_ROLE_COMPLEMENTADO
    End If
End Sub

Private Sub BuildPublicacionState( _
    ByVal p_State As PublicacionCalidadStatusViewState, _
    ByVal p_EsActivo As Boolean, _
    ByVal p_FechaPreparadaParaPublicar As String, _
    ByVal p_PropuestaRechazadaPorCalidadFecha As String, _
    ByVal p_FechaPublicacion As String _
)
    If p_EsActivo Then
        BuildActivePublicacionState p_State, p_FechaPreparadaParaPublicar, p_PropuestaRechazadaPorCalidadFecha
    Else
        BuildPublishedPublicacionState p_State, p_FechaPreparadaParaPublicar, p_FechaPublicacion
    End If
End Sub

Private Sub BuildActivePublicacionState( _
    ByVal p_State As PublicacionCalidadStatusViewState, _
    ByVal p_FechaPreparadaParaPublicar As String, _
    ByVal p_PropuestaRechazadaPorCalidadFecha As String _
)
    p_State.ComandoRechazarEnabled = True
    p_State.ComandoPublicarEnabled = True
    p_State.ComandoPublicarCaption = "Publicar"
    p_State.ComandoGenerarInformePublicabilidadEnabled = True
    p_State.ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled = True
    p_State.PublicacionVisible = True

    If Not IsDate(p_FechaPreparadaParaPublicar) Then
        p_State.PublicacionCaption = "SIN PROPUESTA TÉCNICA"
        p_State.PublicacionColorRole = COLOR_ROLE_NO_COMPLEMENTADO
    ElseIf p_PropuestaRechazadaPorCalidadFecha = vbNullString Then
        p_State.PublicacionCaption = "PROPUESTA (" & p_FechaPreparadaParaPublicar & ") ESPERANDO A CALIDAD"
        p_State.PublicacionColorRole = COLOR_ROLE_COMPLEMENTADO
        p_State.ComandoVerRechazoVisible = False
        p_State.HasComandoVerRechazoVisibleDecision = True
    Else
        p_State.PublicacionCaption = "PROPUESTA (" & p_FechaPreparadaParaPublicar & ") RECHAZADA (" & p_PropuestaRechazadaPorCalidadFecha & ")"
        p_State.PublicacionColorRole = COLOR_ROLE_NO_COMPLEMENTADO
        p_State.ComandoVerRechazoVisible = True
        p_State.HasComandoVerRechazoVisibleDecision = True
    End If
End Sub

Private Sub BuildPublishedPublicacionState( _
    ByVal p_State As PublicacionCalidadStatusViewState, _
    ByVal p_FechaPreparadaParaPublicar As String, _
    ByVal p_FechaPublicacion As String _
)
    p_State.ComandoRechazarEnabled = False
    p_State.ComandoPublicarEnabled = False
    p_State.ComandoGenerarInformePublicabilidadEnabled = False
    p_State.ComandoVerRechazoVisible = False
    p_State.HasComandoVerRechazoVisibleDecision = True
    p_State.ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled = False
    p_State.ComandoPublicarCaption = "Publicado"

    If p_FechaPublicacion <> vbNullString Then
        p_State.PublicacionVisible = True
        If IsDate(p_FechaPreparadaParaPublicar) Then
            p_State.PublicacionCaption = "PROPUESTA (" & p_FechaPreparadaParaPublicar & ") PUBLICADO EL " & p_FechaPublicacion
        Else
            p_State.PublicacionCaption = "PROPUESTA PUBLICADA EL " & p_FechaPublicacion
        End If
        p_State.PublicacionColorRole = COLOR_ROLE_COMPLEMENTADO
    Else
        p_State.PublicacionVisible = False
    End If
End Sub
