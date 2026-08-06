Attribute VB_Name = "DecisionFinalHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DecisionFinalHelper
' RESPONSABILIDAD: contrato puro compartido de completitud de Decisión Final.
' ==========================================================================

Public Function DecisionFinalHelper_EsCompleta( _
    ByVal p_DecisionFinal As Variant, _
    ByVal p_NombreFirmanteFinal As Variant) As Boolean

    DecisionFinalHelper_EsCompleta = False

    If Len(Trim$(Nz(p_DecisionFinal, vbNullString))) = 0 Then Exit Function
    If Len(Trim$(Nz(p_NombreFirmanteFinal, vbNullString))) = 0 Then Exit Function

    DecisionFinalHelper_EsCompleta = True
End Function
