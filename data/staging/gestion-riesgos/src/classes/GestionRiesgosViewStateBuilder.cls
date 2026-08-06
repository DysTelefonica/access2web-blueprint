VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestionRiesgosViewStateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' GestionRiesgosViewStateBuilder (issue #41)
' Pure view-state decisions extracted from Form_FormGestionRiesgos.
' No Access UI, no DAO, no object construction/loading.
' ============================================================

Public Function Build( _
    ByVal p_EsAlta As EnumSiNo, _
    ByVal p_EsTecnico As EnumSiNo, _
    ByVal p_NombreProyecto As String, _
    Optional ByRef p_Error As String _
) As GestionRiesgosViewState
    On Error GoTo errores

    Dim state As GestionRiesgosViewState
    Set state = New GestionRiesgosViewState

    state.NavDatosGeneralesEnabled = True
    state.PermitidoEditar = (p_EsTecnico <> EnumSiNo.Sí)
    state.ComandoGrabarEnabled = state.PermitidoEditar

    If p_EsAlta = Empty Then
        p_Error = "No se sabe si es para alta o no"
        state.ErrorMessage = p_Error
        Set Build = state
        Exit Function
    End If

    If p_EsAlta = EnumSiNo.Sí Then
        If p_EsTecnico = EnumSiNo.Sí Then
            p_Error = "Es un usuario no autorizado"
            state.ErrorMessage = p_Error
            Set Build = state
            Exit Function
        End If

        state.NavResponsablesEnabled = False
        state.NavRiesgosOfertaEnabled = False
        state.NavSuministradoresEnabled = False
        state.Titulo = "ALTA DE GESTIÓN DE RIESGOS"
    Else
        state.NavResponsablesEnabled = True
        state.NavRiesgosOfertaEnabled = True
        state.NavSuministradoresEnabled = True
        state.Titulo = "EDICIÓN DE LA GESTIÓN DE RIESGOS: " & p_NombreProyecto
    End If

    Set Build = state
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en GestionRiesgosViewStateBuilder.Build" & vbNewLine & Err.Description
    End If
End Function
