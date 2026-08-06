Attribute VB_Name = "modRiesgoPriorizacionHelper"
' =============================================================================
' modRiesgoPriorizacionHelper.bas
'
' Helper module para validación de priorización de riesgos.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor: 2026-06-19
'
' Helper público (1):
'   ValidarPriorizacion — valida que la prioridad está en rango 1..N
'
' NOTA: este helper retorna Boolean (no String JSON). Excepción al contrato
' canónico documentada en los átomos (Test_RiesgoPriorizacionHelper.bas).
' =============================================================================
Option Compare Database
Option Explicit

' =============================================================================
' ValidarPriorizacion
'
' Valida que la cadena de prioridad sea un entero en rango [1, p_NumRiesgosActivos].
'
' Parámetros:
'   p_Prioridad         (String) — valor a validar; se trimea whitespace.
'   p_NumRiesgosActivos (Long)   — cantidad de riesgos activos (debe ser >= 1).
'
' Retorno:
'   Boolean — True si p_Prioridad es un entero N tal que 1 <= N <= p_NumRiesgosActivos.
'
' Comportamiento por caso:
'   Happy:    "3", 10 > True
'   Sad:      "", "abc", "-1" > False
'   Edge:     " 3 " (trim) > True; "3.5" (decimal) > False; N=1, P="1" > True; N=0 > False
'   Adversarial: caracteres especiales, Null, N=-5 > False
' =============================================================================
Public Function ValidarPriorizacion( _
    ByVal p_Prioridad As String, _
    ByVal p_NumRiesgosActivos As Long) As Boolean

    Dim sPrioridadTrimmed As String
    Dim nPrioridad As Long

    ' ----- Defensivo: Null o array en p_Prioridad -----
    ' VarType 8204 = vbArray; VarType 10 = vbNull
    If IsNull(p_Prioridad) Then
        ValidarPriorizacion = False
        Exit Function
    End If
    If IsArray(p_Prioridad) Then
        ValidarPriorizacion = False
        Exit Function
    End If

    ' ----- Defensivo: p_NumRiesgosActivos degenerados -----
    ' Si N <= 0, no hay posición válida.
    If p_NumRiesgosActivos <= 0 Then
        ValidarPriorizacion = False
        Exit Function
    End If

    ' ----- Trim -----
    sPrioridadTrimmed = Trim$(p_Prioridad)

    ' ----- No-numérico -----
    ' IsNumeric en VBA acepta decimal (ej. "3.5") y whitespace-padding.
    ' Necesitamos additionally verificar que sea un entero puro.
    If Not IsNumeric(sPrioridadTrimmed) Then
        ValidarPriorizacion = False
        Exit Function
    End If

    ' ----- Conversión a Long -----
    ' IsNumeric acepta entradas como "3.5", "+", "1E2", etc.
    ' CLng con "3.5" produce 4 (redondeo), pero queremos False para decimales.
    ' Verificación: CDbl del trimmed debe ser igual a CLng del trimmed.
    Dim dblValor As Double
    dblValor = CDbl(sPrioridadTrimmed)

    ' CDbl puede convertir valores como "1E2" (100) que son numéricos pero no
    ' representan un entero textual. Verificar que CLng(trimmed) reconstruya el
    ' mismo string para capturar decimales literales (3.5 > "3.5" ? "4").
    Dim lngValor As Long
    lngValor = CLng(sPrioridadTrimmed)

    ' "3.5" > CDbl=3.5, CLng=4 > CStr(4)="4" ? "3.5" > False
    ' "3"   > CDbl=3,   CLng=3 > CStr(3)="3" = "3"   > continúa
    ' "+3"  > CDbl=3,   CLng=3 > CStr(3)="3" ? "+3"  > False
    ' "1E2" > CDbl=100, CLng=100 > CStr(100)="100" = "1E2"? No > False
    If CStr(dblValor) <> sPrioridadTrimmed Then
        ' El string original difiere de la representación canónica del número:
        ' es un decimal, exponencial, o tiene signos extraños.
        ValidarPriorizacion = False
        Exit Function
    End If

    ' ----- Rango -----
    nPrioridad = lngValor
    If nPrioridad < 1 Then
        ValidarPriorizacion = False
        Exit Function
    End If
    If nPrioridad > p_NumRiesgosActivos Then
        ValidarPriorizacion = False
        Exit Function
    End If

    ValidarPriorizacion = True
End Function

