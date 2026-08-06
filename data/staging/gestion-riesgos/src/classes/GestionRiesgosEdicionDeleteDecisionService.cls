VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestionRiesgosEdicionDeleteDecisionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' GestionRiesgosEdicionDeleteDecisionService (issue #45)
' Pure delete-decision rules extracted from Form_FormRiesgosGestionEdicion.
' No UI, no DAO, no filesystem, no cache, no persistence.
' ============================================================

Public Function PuedeIniciarBorrado( _
    ByVal p_HayEdicionActiva As Boolean, _
    ByVal p_EsPrimeraEdicion As EnumSiNo, _
    ByVal p_EsTecnico As EnumSiNo, _
    Optional ByRef p_Mensaje As String _
) As Boolean
    On Error GoTo errores

    p_Mensaje = vbNullString

    If Not p_HayEdicionActiva Then
        p_Mensaje = "No se conoce la edición activa"
        Exit Function
    End If

    If p_EsPrimeraEdicion = EnumSiNo.Sí Then
        p_Mensaje = "Es la primera edición y una gestión de riesgos al menos ha de tener una. Si lo desea, debe borrar la gestión de riesgos completa"
        Exit Function
    End If

    If p_EsTecnico = EnumSiNo.Sí Then
        p_Mensaje = "Desde el 17/12/2019 sólo un miembro de Calidad puede borrar una edición"
        Exit Function
    End If

    PuedeIniciarBorrado = True
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Mensaje = "Error en el método PuedeIniciarBorrado" & vbNewLine & Err.Description
    End If
End Function

Public Function DebeConfirmarCambiosConEdicionAnterior( _
    ByVal p_HayCambiosConEdicionAnterior As Boolean _
) As Boolean
    DebeConfirmarCambiosConEdicionAnterior = p_HayCambiosConEdicionAnterior
End Function

Public Function InformeAnteriorBloqueaBorrado( _
    ByVal p_TieneDocumentoAnterior As Boolean, _
    ByVal p_ExisteArchivoInformeAnterior As Boolean, _
    ByVal p_InformeAnteriorAbierto As Boolean, _
    Optional ByRef p_Mensaje As String _
) As Boolean
    p_Mensaje = vbNullString

    If Not p_TieneDocumentoAnterior Then
        Exit Function
    End If

    If Not p_ExisteArchivoInformeAnterior Then
        Exit Function
    End If

    If p_InformeAnteriorAbierto Then
        p_Mensaje = "Se tiene abierto el informe de la edición anterior"
        InformeAnteriorBloqueaBorrado = True
    End If
End Function
