Attribute VB_Name = "modIndicadorDashboard"
Option Compare Database
Option Explicit

'========================
' CONFIG
'========================
Public Const TVAR_IDSCSV As String = "Indicador_IdsCsv"
Public Const TVAR_COUNT As String = "Indicador_ProyectosCount"
Public Const TVAR_OPENED As String = "Indicador_ProyectosDialogOpened"


Public Enum IndicadorTile
    itIdentificados = 1
    itRetirados = 2
    itEnOferta = 3
    itMaterializados = 4
    itOfertaTrasladar = 5
    itVigentesEnPeriodo = 6
End Enum

'========================
' RANGO FECHAS
'========================
Public Function CalcularRangoSemestre( _
    ByVal p_Semestre As String, _
    ByVal p_Anio As Long, _
    ByRef p_dIni As Date, _
    ByRef p_dFin As Date, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo errores
    p_Error = ""
    CalcularRangoSemestre = False

    If p_Anio < 1900 Or p_Anio > 2100 Then
        p_Error = "Año inválido."
        Exit Function
    End If

    Select Case UCase$(Trim$(p_Semestre))
        Case "S1", "1"
            p_dIni = DateSerial(p_Anio, 1, 1)
            p_dFin = DateSerial(p_Anio, 6, 30)

        Case "S2", "2"
            p_dIni = DateSerial(p_Anio, 7, 1)
            p_dFin = DateSerial(p_Anio, 12, 31)

        Case "ANUAL", ""
            p_dIni = DateSerial(p_Anio, 1, 1)
            p_dFin = DateSerial(p_Anio, 12, 31)

        Case Else
            p_Error = "Semestre inválido (S1/S2/Anual)."
            Exit Function
    End Select

    CalcularRangoSemestre = True
    Exit Function

errores:
    p_Error = "CalcularRangoSemestre: " & Err.Number & vbCrLf & Err.Description
End Function

Public Function SqlDateUS(ByVal d As Date) As String
    SqlDateUS = Format$(d, "mm\/dd\/yyyy")
End Function

'========================
' TEMPVARS PROYECTOS
'========================
Public Sub LimpiarTempVarsIndicador()
    On Error Resume Next
    TempVars.Remove TVAR_IDSCSV
    TempVars.Remove TVAR_COUNT
    TempVars.Remove TVAR_OPENED
End Sub


Public Function GetIdsCsvSeleccionados(ByRef p_Error As String) As String
    p_Error = ""
    If IsNull(TempVars(TVAR_IDSCSV)) Then
        p_Error = "No hay proyectos seleccionados."
        GetIdsCsvSeleccionados = ""
        Exit Function
    End If
    GetIdsCsvSeleccionados = Nz(TempVars(TVAR_IDSCSV), "")
    If Len(GetIdsCsvSeleccionados) = 0 Then p_Error = "No hay proyectos seleccionados."
End Function

Public Function GetCountSeleccionados() As Long
    On Error Resume Next
    If Not IsNull(TempVars(TVAR_COUNT)) Then
        GetCountSeleccionados = CLng(Nz(TempVars(TVAR_COUNT), 0))
    Else
        GetCountSeleccionados = 0
    End If
End Function

'========================
' EXCEL EXPORT
'========================
Public Function ExportarSQLaExcel( _
                                    ByVal p_SQL As String, _
                                    ByVal p_Titulo As String, _
                                    ByRef p_Error As String _
                                ) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xl As Object, wb As Object, ws As Object
    Dim i As Long

    On Error GoTo errores
    p_Error = ""
    ExportarSQLaExcel = False

    Set db = getdb()
    Set rs = db.OpenRecordset(p_SQL, dbOpenSnapshot)

    If rs.EOF And rs.BOF Then
        p_Error = "No hay registros para exportar."
        GoTo salir
    End If

    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Name = "Datos"

    ' Cabeceras
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).value = rs.Fields(i).Name
        ws.Cells(1, i + 1).Font.Bold = True
    Next

    ' Datos
    ws.Range("A2").CopyFromRecordset rs

    ws.Columns.AutoFit
    xl.Visible = True

    ExportarSQLaExcel = True

salir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    Exit Function

errores:
    p_Error = "ExportarSQLaExcel: " & Err.Number & vbCrLf & Err.Description
    Resume salir
End Function
Public Function HaPasadoPorDialogo() As Boolean
    On Error Resume Next
    If Not (IsNull(TempVars(TVAR_IDSCSV)) Or IsNull(TempVars(TVAR_COUNT))) Then
        HaPasadoPorDialogo = True

    End If


End Function

' ============================================================
' IndicadorDashboard_BuildOpenArgs
' Pure producer for the canonical FormIndicador OpenArgs contract.
'
' Input  : p_Anio (String, "2025"-shape expected), p_Sem (String, "S1"/"S2"/"1"/"2"/"Anual"/"A"/"").
' Output : canonical "ANIO=<year>;SEM=<value>" contract when inputs are valid.
'         Empty string + p_Error populated when input cannot be safely consumed
'         by the child form (FormIndicadorProyectos / ParseIndicadorOpenArgs).
'
' Pure logic, no DAO, no Access controls, no globals - directly unit-testable.
' ============================================================
Public Function IndicadorDashboard_BuildOpenArgs( _
    ByVal p_Anio As String, _
    ByVal p_Sem As String, _
    ByRef p_Error As String _
) As String
    Dim anio As Long
    Dim semNorm As String
    Dim semErr As String

    On Error GoTo errores
    p_Error = ""
    IndicadorDashboard_BuildOpenArgs = ""

    If Not IsNumeric(Nz(p_Anio, "")) Then
        p_Error = "Anyo invalido (IndicadorDashboard_BuildOpenArgs): '" & Nz(p_Anio, "") & "'"
        Exit Function
    End If
    anio = CLng(Nz(p_Anio, "0"))
    If anio < 1900 Or anio > 2100 Then
        p_Error = "Anyo fuera de rango (1900..2100): " & CStr(anio)
        Exit Function
    End If

    semNorm = IndicadorDashboard_NormalizeSemestre(p_Sem, semErr)
    If semErr <> "" Then
        p_Error = semErr
        Exit Function
    End If

    IndicadorDashboard_BuildOpenArgs = "ANIO=" & CStr(anio) & ";SEM=" & semNorm
    Exit Function

errores:
    p_Error = "IndicadorDashboard_BuildOpenArgs: " & Err.Number & vbCrLf & Err.Description
    IndicadorDashboard_BuildOpenArgs = ""
End Function

' Pure helper for IndicadorDashboard_BuildOpenArgs. Mirrors the parser-side
' NormalizeSemestre semantics so producer and consumer stay in sync.
Private Function IndicadorDashboard_NormalizeSemestre( _
    ByVal p_Sem As String, _
    ByRef p_Error As String _
) As String
    Dim s As String
    p_Error = ""
    s = UCase$(Trim$(Nz(p_Sem, "")))
    Select Case s
        Case "1", "S1"
            IndicadorDashboard_NormalizeSemestre = "1"
        Case "2", "S2"
            IndicadorDashboard_NormalizeSemestre = "2"
        Case "", "ANUAL", "A"
            IndicadorDashboard_NormalizeSemestre = ""
        Case Else
            p_Error = "Semestre invalido (esperado S1/S2/Anual): '" & Nz(p_Sem, "") & "'"
            IndicadorDashboard_NormalizeSemestre = ""
    End Select
End Function

' ============================================================
' IndicadorDashboard_BuildExportSql
' Returns the detail SQL for a given tile only when the calculation context
' is fresh (p_CalculoExitoso = True). When stale, returns "" + p_Error so
' the caller is forced to recalculate before exporting - never exports
' silent garbage from a previous run with different proyectos/year.
'
' Pure wrapper around IndicadorRiesgosV2_SqlDetalle; the gating policy lives
' here so the form does not have to track "calculo exitoso" itself.
' ============================================================
Public Function IndicadorDashboard_BuildExportSql( _
    ByVal p_Tile As IndicadorTile, _
    ByVal p_dIni As Date, _
    ByVal p_dFin As Date, _
    ByVal p_IdsCsv As String, _
    ByVal p_CalculoExitoso As Boolean, _
    ByRef p_Error As String _
) As String
    On Error GoTo errores
    p_Error = ""
    IndicadorDashboard_BuildExportSql = ""

    If Not p_CalculoExitoso Then
        p_Error = "No hay un calculo exitoso reciente; recalcula antes de exportar."
        Exit Function
    End If

    IndicadorDashboard_BuildExportSql = IndicadorRiesgosV2_SqlDetalle( _
        p_Tile, p_dIni, p_dFin, p_IdsCsv)
    Exit Function

errores:
    p_Error = "IndicadorDashboard_BuildExportSql: " & Err.Number & vbCrLf & Err.Description
    IndicadorDashboard_BuildExportSql = ""
End Function

' ============================================================
' IndicadorDashboard_CanExport
' Tile-indexed export gate: only tiles whose displayed value is > 0 unlock
' the per-tile XLS export. The six indexes (itIdentificados..itVigentesEnPeriodo)
' map 1:1 with the lblTile{N}Valor captions in FormIndicador.
' ============================================================
Public Function IndicadorDashboard_CanExport( _
    ByRef p_Values() As Long, _
    ByVal p_Tile As IndicadorTile _
) As Boolean
    On Error GoTo errores
    IndicadorDashboard_CanExport = False

    If LBound(p_Values) > 1 Or UBound(p_Values) < 1 Then Exit Function
    Select Case p_Tile
        Case itIdentificados, itRetirados, itEnOferta, _
             itMaterializados, itOfertaTrasladar, itVigentesEnPeriodo
            IndicadorDashboard_CanExport = (p_Values(CLng(p_Tile)) > 0)
        Case Else
            ' Unknown tile - intentionally false; defensive against future enum drift.
            IndicadorDashboard_CanExport = False
    End Select
    Exit Function

errores:
    IndicadorDashboard_CanExport = False
End Function





