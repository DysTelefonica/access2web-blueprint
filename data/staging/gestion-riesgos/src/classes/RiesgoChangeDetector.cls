VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoChangeDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' RiesgoChangeDetector (issue #38)
' Pure field-by-field comparison of two riesgo objects.
' Extracted from Form_FormRiesgo.HaHabidoCambios to make
' save-decision logic testable without UI.
'
' The 16 fields compared mirror the original HaHabidoCambios
' contract (Priorizacion, Descripcion, CausaRaiz, EntidadDetecta,
' DetectadoPor, FechaDetectado, FechaMaterializado, FechaRetirado,
' Plazo, Coste, Calidad, Vulnerabilidad, Mitigacion,
' JustificacionAceptacionRiesgo, JustificacionRetiroRiesgo,
' CodRiesgoBiblioteca).
'
' Simplification note: the original guard "if proyecto requires
' CausaRaiz then compare it" is removed. In practice, RellenarObjeto
' writes CausaRaiz on both objects consistently, so unconditional
' comparison yields the same observable result.
' ============================================================

Public Function HasChanges( _
    ByVal p_Original As riesgo, _
    ByVal p_Current As riesgo, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    p_Error = vbNullString

    If p_Original Is Nothing Then
        p_Error = "HasChanges: p_Original is Nothing"
        Exit Function
    End If
    If p_Current Is Nothing Then
        p_Error = "HasChanges: p_Current is Nothing"
        Exit Function
    End If

    With p_Original
        If Not FieldEqual(.Priorizacion, p_Current.Priorizacion) Then HasChanges = True: Exit Function
        If Not FieldEqual(.Descripcion, p_Current.Descripcion) Then HasChanges = True: Exit Function
        If Not FieldEqual(.CausaRaiz, p_Current.CausaRaiz) Then HasChanges = True: Exit Function
        If Not FieldEqual(.EntidadDetecta, p_Current.EntidadDetecta) Then HasChanges = True: Exit Function
        If Not FieldEqual(.DetectadoPor, p_Current.DetectadoPor) Then HasChanges = True: Exit Function
        If Not FieldEqual(.FechaDetectado, p_Current.FechaDetectado) Then HasChanges = True: Exit Function
        If Not FieldEqual(.FechaMaterializado, p_Current.FechaMaterializado) Then HasChanges = True: Exit Function
        If Not FieldEqual(.FechaRetirado, p_Current.FechaRetirado) Then HasChanges = True: Exit Function
        If Not FieldEqual(.Plazo, p_Current.Plazo) Then HasChanges = True: Exit Function
        If Not FieldEqual(.Coste, p_Current.Coste) Then HasChanges = True: Exit Function
        If Not FieldEqual(.Calidad, p_Current.Calidad) Then HasChanges = True: Exit Function
        If Not FieldEqual(.Vulnerabilidad, p_Current.Vulnerabilidad) Then HasChanges = True: Exit Function
        If Not FieldEqual(.Mitigacion, p_Current.Mitigacion) Then HasChanges = True: Exit Function
        If Not FieldEqual(.JustificacionAceptacionRiesgo, p_Current.JustificacionAceptacionRiesgo) Then HasChanges = True: Exit Function
        If Not FieldEqual(.JustificacionRetiroRiesgo, p_Current.JustificacionRetiroRiesgo) Then HasChanges = True: Exit Function
        If Not FieldEqual(.CodRiesgoBiblioteca, p_Current.CodRiesgoBiblioteca) Then HasChanges = True: Exit Function
    End With

    HasChanges = False
    Exit Function

EH:
    p_Error = "RiesgoChangeDetector.HasChanges: " & Err.description
End Function

Private Function FieldEqual(ByVal p_Left As String, ByVal p_Right As String) As Boolean
    FieldEqual = (Trim$(Nz(p_Left, vbNullString)) = Trim$(Nz(p_Right, vbNullString)))
End Function

