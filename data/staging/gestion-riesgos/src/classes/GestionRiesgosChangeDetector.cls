VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestionRiesgosChangeDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' GestionRiesgosChangeDetector (issue #41)
' Pure field-by-field comparison of two Proyecto objects.
' Extracted from Form_FormGestionRiesgos.HaHabidoCambios to make
' save-decision logic testable without UI or DAO.
' ============================================================

Public Function HaHabidoCambios( _
    ByVal p_ProyectoAlInicio As Proyecto, _
    ByVal p_ProyectoActivo As Proyecto, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo errores

    If p_ProyectoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If

    With p_ProyectoAlInicio
        If .ParaInformeAvisos <> p_ProyectoActivo.ParaInformeAvisos Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .EnUTE <> p_ProyectoActivo.EnUTE Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreUsuarioCalidad <> p_ProyectoActivo.NombreUsuarioCalidad Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .CorreoRAC <> p_ProyectoActivo.CorreoRAC Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Elaborado <> p_ProyectoActivo.Elaborado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Revisado <> p_ProyectoActivo.Revisado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Aprobado <> p_ProyectoActivo.Aprobado Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en el método HaHabidoCambios" & vbNewLine & Err.Description
    End If
End Function
