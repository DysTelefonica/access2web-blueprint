Attribute VB_Name = "modIndicadorOpenArgsParser"
Option Compare Database
Option Explicit

' ============================================================
' modIndicadorOpenArgsParser - Pure parser for the FormIndicador
'                               OpenArgs contract.
'
' Project: gestion_riesgos
' Consumer: Form_formIndicadorProyectos.cls (Form_Open -> ParseOpenArgs)
' Producer: Form_FormIndicador.cls (cmdCargarProyectos_Click)
'
' Canonical contract (NEW - primary):
'   "ANIO=<year>;SEM=<value>"
'   <year>  is a positive Long in [1900..2100] (e.g. 2025)
'   <value> is "1", "2", or "" (empty = annual)
'   Segments are separated by ";"; order is free; unknown
'   segments are ignored for forward compatibility.
'
'   Examples:
'     "ANIO=2025;SEM=1"   -> p_Anio=2025, p_Semestre="1"
'     "ANIO=2025;SEM=2"   -> p_Anio=2025, p_Semestre="2"
'     "ANIO=2025;SEM="    -> p_Anio=2025, p_Semestre=""   (annual)
'     "SEM=1;ANIO=2025"   -> same (order-independent)
'
' Legacy pipe contract (defensive fallback). The pipe branch only
' runs when hasPipe is True, so the parser accepts exactly:
'   "2025|S1"             -> p_Anio=2025, p_Semestre="1"
'   "2025|S2"             -> p_Anio=2025, p_Semestre="2"
'   "2025|1"              -> p_Anio=2025, p_Semestre="1"   (short form)
'   "2025|2"              -> p_Anio=2025, p_Semestre="2"   (short form)
'   "2025|"               -> p_Anio=2025, p_Semestre=""    (annual)
' A bare year without a pipe (e.g. "2025") is NOT part of the
' legacy contract and falls through to the unknown-format error.
'
' Returns:
'   True  on success - p_Anio and p_Semestre populated, p_Error empty.
'   False on malformed input - p_Anio=0, p_Semestre="", p_Error populated
'         in Spanish. Callers MUST NOT consume stale values on failure.
'
' Pure logic, no DAO, no Access controls, no globals - directly
' unit-testable from a Test_* module.
'
' Issue:   DysTelefonica/dysflow#1006 (tooling gap).
'          Application fix proceeds independently of the tooling fix.
' ============================================================

Public Function ParseIndicadorOpenArgs( _
    ByVal p_OpenArgs As String, _
    ByRef p_Anio As Long, _
    ByRef p_Semestre As String, _
    Optional ByRef p_Error As String _
) As Boolean
    Dim s As String
    Dim hasNamed As Boolean
    Dim hasPipe As Boolean

    On Error GoTo errores

    p_Anio = 0
    p_Semestre = ""
    p_Error = ""
    ParseIndicadorOpenArgs = False

    s = Trim$(Nz(p_OpenArgs, ""))
    If Len(s) = 0 Then
        p_Error = "Falta el argumento OpenArgs (esperado 'ANIO=<year>;SEM=<value>')"
        Exit Function
    End If

    hasNamed = (InStr(1, s, "ANIO=", vbTextCompare) > 0)
    hasPipe = (InStr(1, s, "|") > 0)

    If hasNamed Then
        If ParseNamed(s, p_Anio, p_Semestre, p_Error) Then ParseIndicadorOpenArgs = True
        If Not ParseIndicadorOpenArgs Then
            p_Anio = 0
            p_Semestre = ""
        End If
        Exit Function
    End If

    If hasPipe Then
        If ParseLegacyPipe(s, p_Anio, p_Semestre, p_Error) Then ParseIndicadorOpenArgs = True
        If Not ParseIndicadorOpenArgs Then
            p_Anio = 0
            p_Semestre = ""
        End If
        Exit Function
    End If

    p_Error = "Formato de OpenArgs no reconocido (esperado 'ANIO=<year>;SEM=<value>')"
    Exit Function

errores:
    p_Anio = 0
    p_Semestre = ""
    p_Error = "ParseIndicadorOpenArgs: " & Err.Number & " - " & Err.Description
    ParseIndicadorOpenArgs = False
End Function

Private Function ParseNamed( _
    ByVal p_S As String, _
    ByRef p_Anio As Long, _
    ByRef p_Semestre As String, _
    ByRef p_Error As String _
) As Boolean
    Dim segments() As String
    Dim i As Long
    Dim seg As String
    Dim eqPos As Long
    Dim key As String
    Dim val As String
    Dim sawAnio As Boolean

    p_Anio = 0
    p_Semestre = ""
    p_Error = ""
    ParseNamed = False

    segments = Split(p_S, ";")
    For i = LBound(segments) To UBound(segments)
        seg = Trim$(Nz(segments(i), ""))
        If Len(seg) > 0 Then
            eqPos = InStr(1, seg, "=")
            If eqPos <= 0 Then
                p_Error = "Segmento OpenArgs sin '=': '" & seg & "'"
                Exit Function
            End If
            key = UCase$(Trim$(Left$(seg, eqPos - 1)))
            val = Trim$(Mid$(seg, eqPos + 1))

            Select Case key
                Case "ANIO"
                    If Len(val) = 0 Or Not IsNumeric(val) Then
                        p_Error = "ANIO no numerico o vacio: '" & val & "'"
                        Exit Function
                    End If
                    p_Anio = CLng(val)
                    sawAnio = True
                Case "SEM"
                    p_Semestre = NormalizeSemestre(val, p_Error)
                    If p_Error <> "" Then Exit Function
                Case Else
                    ' Unknown key - ignored for forward compatibility.
            End Select
        End If
    Next i

    If Not sawAnio Then
        p_Error = "Falta ANIO= en OpenArgs"
        Exit Function
    End If

    If p_Anio < 1900 Or p_Anio > 2100 Then
        p_Error = "ANIO fuera de rango (1900..2100): " & CStr(p_Anio)
        p_Anio = 0
        Exit Function
    End If

    ParseNamed = True
End Function

Private Function ParseLegacyPipe( _
    ByVal p_S As String, _
    ByRef p_Anio As Long, _
    ByRef p_Semestre As String, _
    ByRef p_Error As String _
) As Boolean
    Dim parts() As String
    Dim yRaw As String
    Dim sRaw As String

    p_Anio = 0
    p_Semestre = ""
    p_Error = ""
    ParseLegacyPipe = False

    parts = Split(p_S, "|")
    If UBound(parts) < 1 Then
        p_Error = "Formato pipe OpenArgs incompleto (esperado 'year|sem')"
        Exit Function
    End If

    yRaw = Trim$(Nz(parts(0), ""))
    sRaw = Trim$(Nz(parts(1), ""))

    If Len(yRaw) = 0 Then
        p_Error = "Falta el anyo en OpenArgs (formato pipe)"
        Exit Function
    End If
    If Not IsNumeric(yRaw) Then
        p_Error = "Anyo no numerico en OpenArgs (formato pipe): '" & yRaw & "'"
        Exit Function
    End If
    p_Anio = CLng(yRaw)

    If p_Anio < 1900 Or p_Anio > 2100 Then
        p_Error = "Anyo fuera de rango (1900..2100): " & CStr(p_Anio)
        p_Anio = 0
        Exit Function
    End If

    p_Semestre = NormalizeSemestre(sRaw, p_Error)
    If p_Error <> "" Then
        p_Anio = 0
        Exit Function
    End If

    ParseLegacyPipe = True
End Function

Private Function NormalizeSemestre( _
    ByVal p_Raw As String, _
    ByRef p_Error As String _
) As String
    Dim s As String
    s = UCase$(Trim$(Nz(p_Raw, "")))

    Select Case s
        Case "1", "S1"
            NormalizeSemestre = "1"
        Case "2", "S2"
            NormalizeSemestre = "2"
        Case "", "ANUAL", "A"
            NormalizeSemestre = ""
        Case Else
            p_Error = "Semestre invalido (esperado 1/S1, 2/S2 o vacio/anual): '" & p_Raw & "'"
            NormalizeSemestre = ""
    End Select
End Function
