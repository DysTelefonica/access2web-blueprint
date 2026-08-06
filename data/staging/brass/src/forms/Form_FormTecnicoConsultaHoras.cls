Option Compare Database
Option Explicit

Private m_Error As String
Private fso As New FileSystemObject


Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub




Private Sub ComandoInforme_Click()

    Dim m_FechaInicial As String
    Dim m_FechaFinal As String
    Dim m_Tecnico As String
    Dim m_ColTenicosHoras As Collection
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_FechaInicial = Nz(Me.FechaInicial, "")
    m_FechaFinal = Nz(Me.FechaFinal, "")
    m_Tecnico = Nz(Me.ComboTecnico, "")
    
    m_Linea = "Informe Horas de los Técnicos....."
    Avance m_Linea
    Set m_ColTenicosHoras = getTecnicosHoras(m_FechaInicial, m_FechaFinal, m_Tecnico, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Linea = "Informe Horas Excel....."
    Avance m_Linea
    
    PasarAExcel m_ColTenicosHoras, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblEstado.Visible = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoInforme_Click se ha producido el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub Form_Load()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
  
    AjustarTamaño Me
    If Not m_ObjTecnicoActivo Is Nothing Then
        Me.ComboTecnico = m_ObjTecnicoActivo.Alias
    End If
    Me.Caption = Me.lblTitulo1.Caption
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.No, EnumSino.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblEstado.Visible = False
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Function getTecnicosHoras( _
                                    p_FechaInicial As String, _
                                    p_FechaFinal As String, _
                                    Optional p_Alias As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Collection
                                   
    Dim m_ObjTecnicoHoras As TecnicoHoras
    Dim m_ColTecnicos As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjTecnico As Tecnico
    Dim i As Integer
    Dim m_Fecha As String
    Dim m_DiasTotales As Integer
    
    On Error GoTo errores
    If Not IsDate(p_FechaInicial) Then
        p_Error = "No se ha indicado la fecha inicial"
        Err.Raise 1000
    End If
    If Not IsDate(p_FechaFinal) Then
        p_Error = "No se ha indicado la fecha Final"
        Err.Raise 1000
    End If
    If CDate(p_FechaFinal) < CDate(p_FechaInicial) Then
        p_Error = "La fecha final no puede ser anterior a la inicial"
        Err.Raise 1000
    End If
    If p_Alias = "" Then
        Set m_ColTecnicos = m_ObjEntorno.ColTecnicosActivos
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set m_ObjTecnico = Constructor.getTecnico(p_Alias, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTecnico Is Nothing Then
            p_Error = "No se encuentra el Técnico " & p_Alias
            Err.Raise 1000
        End If
        Set m_ColTecnicos = New Scripting.Dictionary
        m_ColTecnicos.CompareMode = TextCompare
        m_ColTecnicos.Add m_ObjTecnico.Alias, m_ObjTecnico
    End If
    If m_ColTecnicos Is Nothing Then
        Exit Function
    End If
    m_DiasTotales = DateDiff("d", CDate(p_FechaInicial), CDate(p_FechaFinal))
    For i = 0 To m_DiasTotales - 1
        m_Fecha = DateAdd("d", CDbl(i), CDate(p_FechaInicial))
        For Each m_ID In m_ColTecnicos
            Set m_ObjTecnico = m_ColTecnicos(m_ID)
            m_Linea = "Informe Horas de los Técnicos....." & " " & m_ObjTecnico.Nombre
            Avance m_Linea
            Set m_ObjTecnicoHoras = New TecnicoHoras
            With m_ObjTecnicoHoras
                .Alias = m_ObjTecnico.Alias
                .Fecha = m_Fecha
            End With
            If getTecnicosHoras Is Nothing Then
                Set getTecnicosHoras = New Collection
            End If
            getTecnicosHoras.Add m_ObjTecnicoHoras
            Set m_ObjTecnicoHoras = Nothing
            Set m_ObjTecnico = Nothing
        Next
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getTecnicosHoras ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function




Private Function PasarAExcel( _
                                p_ObjColTecnicosHoras As Collection, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim m_NombreInforme As String
    Dim m_ObjTecnicoHora As TecnicoHoras
    Dim m_ID As Variant
    Dim m_CadenaActividadesLab As String
    Dim m_CadenaActividadesExt As String
    Dim m_CadenaActividades As String
    Dim WbLibro As excel.Workbook
    Dim wbHoja As excel.Worksheet
    Dim intFila As Integer
    Dim m_NombreArchivo As String
    Dim m_URLArchivo As String
    Dim MiRango As excel.Range
    Dim m_Alias As String
    Dim m_Fecha As String
    Dim m_TipoDia As String
    Dim m_Libranza As EnumSino
    Dim m_Fiesta As EnumSino
    Dim m_HorasLab As String
    Dim m_HorasExt As String
    Dim m_Horas As Double
    Dim m_Observaciones As String
    Dim m_DiaDeLaSemana As String
    Dim intFilaFinal As Integer
    Dim intFilaInicial As Integer
    
    On Error GoTo errores
    If p_ObjColTecnicosHoras Is Nothing Then
        Exit Function
    End If
    m_NombreInforme = "Informe de Horas"
    m_NombreArchivo = fso.GetBaseName(m_NombreInforme) & "_" & Hour(Now()) & "_" & Minute(Now()) & "_" & Second(Time) & ".xlsx"
    m_URLArchivo = m_ObjEntorno.URLDirectorioLocal & m_NombreArchivo
    If fso.FileExists(m_URLArchivo) Then
        If FicheroAbierto(m_URLArchivo) Then
            p_Error = "Tiene una consulta abierta"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLArchivo, True
    End If
    Set appexcel = New excel.Application
    appexcel.Visible = False
    Set WbLibro = appexcel.Workbooks.Add
    WbLibro.SaveAs m_URLArchivo
    Set wbHoja = WbLibro.Worksheets(1)
    intFila = 1
    intFilaInicial = intFila
    With wbHoja
        .Cells(intFila, 1).Value = "Técnico"
        .Cells(intFila, 2).Value = "Fecha"
        .Cells(intFila, 3).Value = "HorasLab"
        .Cells(intFila, 4).Value = "HorasExt"
        .Cells(intFila, 5).Value = "CadenaIDs"
        .Cells(intFila, 6).Value = "Día"
        .Cells(intFila, 7).Value = "Tipo"
        .Cells(intFila, 8).Value = "Libranza"
        .Cells(intFila, 9).Value = "Festivo"
        .Cells(intFila, 10).Value = "Observaciones"
        Set MiRango = .Range(.Cells(intFila, 1), .Cells(intFila, 10))
        MiRango.Font.Bold = True
    End With
    intFila = intFila + 1
    For Each m_ObjTecnicoHora In p_ObjColTecnicosHoras
        m_Linea = "Informe Horas Excel..... " & m_ObjTecnicoHora.Alias
        Avance m_Linea
        With m_ObjTecnicoHora
            m_Alias = .Alias
            m_Fecha = .Fecha
            m_TipoDia = .TipoDia
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_Libranza = .LIbranza
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_Fiesta = .Festivo
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_DiaDeLaSemana = .DiaDeLaSemana
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_DiaDeLaSemana <> "" Then
                If m_DiaDeLaSemana = "sábado" Or m_DiaDeLaSemana = "domingo" Then
                    GoTo siguiente
                End If
            End If
            m_CadenaActividadesLab = .ActividadesTextoLab
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_CadenaActividadesExt = .ActividadesTextoExt
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_CadenaActividadesLab <> "" Then
                m_CadenaActividades = m_CadenaActividadesLab
            End If
            If m_CadenaActividadesExt <> "" Then
                If m_CadenaActividades = "" Then
                    m_CadenaActividades = m_CadenaActividadesExt
                Else
                    m_CadenaActividades = m_CadenaActividades & ";" & m_CadenaActividadesExt
                End If
            End If
            m_HorasLab = .Hlab
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not IsNumeric(m_HorasLab) Then
                m_HorasLab = "0"
            End If
            m_HorasExt = .HExt
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not IsNumeric(m_HorasExt) Then
                m_HorasExt = "0"
            End If
            m_Horas = CDbl(m_HorasLab) + CDbl(m_HorasExt)
        End With
        
        With wbHoja
            
            .Cells(intFila, 1).Value = m_Alias
            .Cells(intFila, 2).Value = Format(m_Fecha, "mm/dd/yyyy")
            .Cells(intFila, 3).Value = m_HorasLab
            .Cells(intFila, 4).Value = m_HorasExt
            If m_CadenaActividades <> "" Then
                .Cells(intFila, 5).Value = m_CadenaActividades
            End If
            
            .Cells(intFila, 6).Value = m_DiaDeLaSemana
            .Cells(intFila, 7).Value = m_TipoDia
            .Cells(intFila, 8).Value = IIf(m_Libranza = EnumSino.Sí, "Sí", "No")
            .Cells(intFila, 9).Value = IIf(m_Fiesta = EnumSino.Sí, "Sí", "No")
            If m_TipoDia = "Laborable" Then
                If m_Libranza = EnumSino.No Then
                    If m_Horas <> 7.5 Then
                        m_Observaciones = "Día Laborable y no cumple las 7.5 horas registradas"
                    End If
                Else
                    If m_Horas <> 0 Then
                        m_Observaciones = "Día de Libranza con horas registradas"
                    End If
                End If
            Else
                If m_Horas <> 0 Then
                    m_Observaciones = "Día no laborable con horas registradas"
                End If
            End If
            If m_Observaciones <> "" Then
                .Cells(intFila, 10).Value = m_Observaciones
                Set MiRango = .Range(.Cells(intFila, 1), .Cells(intFila, 10))
                MiRango.Font.Bold = True
                MiRango.Font.Color = RGB(255, 0, 0)
            End If
        End With
        intFila = intFila + 1
siguiente:
        m_Observaciones = ""
        m_CadenaActividadesLab = ""
        m_CadenaActividadesExt = ""
        m_HorasLab = ""
        m_HorasExt = ""
        m_Libranza = 0
        m_Horas = 0
        m_Fiesta = 0
        m_TipoDia = ""
        
    Next
    
   
    
    intFilaFinal = intFila - 1
    
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        
    End With
    
    Set MiRango = wbHoja.Range(wbHoja.Cells(intFilaInicial, 1), wbHoja.Cells(intFilaFinal, 10))
    With MiRango
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    Set MiRango = wbHoja.Range(wbHoja.Cells(intFilaInicial, 1), wbHoja.Cells(intFilaInicial, 10))
    With MiRango
        .HorizontalAlignment = xlCenter
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
    WbLibro.Close True
    Set WbLibro = Nothing
    appexcel.Quit
    Set appexcel = Nothing
    Ejecutar Me.hWnd, "open", m_URLArchivo, "", "", 1
    PasarAExcel = m_URLArchivo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PasarAExcel ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not appexcel Is Nothing Then
        appexcel.Quit
        Set appexcel = Nothing
    End If
    
End Function

