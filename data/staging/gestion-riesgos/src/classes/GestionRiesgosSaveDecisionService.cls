VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestionRiesgosSaveDecisionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' GestionRiesgosSaveDecisionService (issue #41)
' Pure save-decision rules extracted from Form_FormGestionRiesgos.grabar.
' No UI, no DAO, no persistence/events.
' ============================================================

Public Function DebeAvisarNoInforme( _
    ByVal p_SinEco As EnumSiNo, _
    ByVal p_ProyectoAlInicio As Proyecto, _
    ByVal p_ProyectoActivo As Proyecto, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo errores

    If p_SinEco <> EnumSiNo.No Then
        Exit Function
    End If

    If p_ProyectoActivo.ParaInformeAvisos <> "No" Then
        Exit Function
    End If

    If p_ProyectoAlInicio Is Nothing Then
        DebeAvisarNoInforme = True
        Exit Function
    End If

    DebeAvisarNoInforme = (p_ProyectoAlInicio.ParaInformeAvisos <> p_ProyectoActivo.ParaInformeAvisos)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en el método DebeAvisarNoInforme" & vbNewLine & Err.Description
    End If
End Function
