Attribute VB_Name = "Helper_ExpedienteFechasLogic"
Option Compare Database
Option Explicit
' Helper REFAC-4a slice 1: logica pura de fechas de expediente.
' Stateless, no popup, no UI. Extrae las reglas de negocio del form
' Form_FormExpedienteFechas.cls, donde la logica DateAdd estaba duplicada
' en 4 event handlers (ComandoFECHACERTIFICACION_Click,
' FECHACERTIFICACION_BeforeUpdate, FechaFinContrato_BeforeUpdate,
' GARANTIAMESES_BeforeUpdate).
'
' Cobertura: BR-19-01..06 (PR-D del coverage matrix).
' Tests en src/modules/Test_Helper_ExpedienteFechasLogic.bas.

' BR-19-01..06: Calcula FechaFinGarantia a partir de GarantiaMeses,
' FechaCertificacion y FechaFinContrato. Logica equivalente a la del form
' original (DateAdd "m", GarantiaMeses, fecha elegida).
' Reglas:
'   - Si GarantiaMeses no es numerico -> FechaFinGarantia = Null
'   - Si FechaFinContrato no es fecha -> FechaFinGarantia = Null
'   - Si FechaCertificacion es fecha -> DateAdd sobre FechaCertificacion
'   - Si FechaCertificacion NO es fecha -> DateAdd sobre FechaFinContrato (fallback)
Public Function CalcularFechaFinGarantia( _
    ByVal p_GarantiaMeses As Variant, _
    ByVal p_FechaCertificacion As Variant, _
    ByVal p_FechaFinContrato As Variant, _
    ByRef p_FechaFinGarantia As Variant, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores

    ' BR-19-01: GarantiaMeses no numerico (Null, "", error de parseo) -> fechaFin = Null
    ' IsNumeric cubre Null y string vacio devolviendo False.
    If Not IsNumeric(p_GarantiaMeses) Then
        p_FechaFinGarantia = Null
        CalcularFechaFinGarantia = ""
        Exit Function
    End If

    ' BR-19-02: FechaFinContrato no es fecha (Null, "") -> fechaFin = Null
    ' Sin FechaFinContrato no hay base para el calculo (incluso si FechaCertificacion
    ' es valida, la regla del form original es que FechaFinContrato es obligatoria).
    If Not IsDate(p_FechaFinContrato) Then
        p_FechaFinGarantia = Null
        CalcularFechaFinGarantia = ""
        Exit Function
    End If

    ' BR-19-03/04: DateAdd sobre Certificacion o fallback a FinContrato.
    ' Si FechaCertificacion es fecha valida -> se usa como base.
    ' Si FechaCertificacion NO es fecha -> fallback a FechaFinContrato.
    ' DateAdd acepta GarantiaMeses negativo (BR-19-06) y cero (BR-19-05).
    If IsDate(p_FechaCertificacion) Then
        p_FechaFinGarantia = DateAdd("m", CDbl(p_GarantiaMeses), CDate(p_FechaCertificacion))
    Else
        p_FechaFinGarantia = DateAdd("m", CDbl(p_GarantiaMeses), CDate(p_FechaFinContrato))
    End If

    CalcularFechaFinGarantia = ""
    Exit Function

errores:
    p_Error = "CalcularFechaFinGarantia: " & Err.Description
    CalcularFechaFinGarantia = "ERR"
End Function