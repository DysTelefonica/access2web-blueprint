VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Informe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New FileSystemObject
Private m_AppExcel As excel.Application
Private m_ObjLibro As excel.Workbook
Private m_ObjHoja As excel.Worksheet
Private m_Fila As Long
Private m_ObjPlantilla As excel.Workbook
Private m_ObjHojaPlantilla As excel.Worksheet
Private m_IDEventoActual As String


Public Function EventosRAC( _
                            ByRef p_ColEventos As Scripting.Dictionary, _
                            ByRef p_URLInforme As String, _
                            Optional p_ConHojaResumen As Boolean = False, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_ObjEvento As Evento
    Dim m_IDEvento As Variant
    Dim m_IntFilaFinalActividades As Integer
    Dim m_intFilaFinalMateriales As Integer
    Dim m_ColActividades As Scripting.Dictionary
    Dim m_ColMateriales As Scripting.Dictionary
    Dim m_ColMaterialesTotales As Scripting.Dictionary
    Dim m_IDMaterial As Variant
    Dim m_ObjMaterial As Material
    Dim m_IDActividad As Variant
    Dim m_ObjActividad As Actividad
    Dim intNumero As Integer
    Dim intNumeroTotal As Integer
    On Error GoTo errores
    
    p_Error = ""
    If p_ColEventos Is Nothing Then
        p_Error = "No se ha seleccionado ningún Parte de Trabajo para el Informe"
        Err.Raise 1000
    End If
    If p_ColEventos.count = 0 Then
        p_Error = "No se ha seleccionado ningún Parte de Trabajo para el Informe"
        Err.Raise 1000
    End If
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    m_Linea = "Abriendo excel ....."
    Avance m_Linea
    
    AbrirNuevoExcel p_URLInforme, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    intNumero = 1
    intNumeroTotal = p_ColEventos.count
    'copiamos para cada evento
    For Each m_IDEvento In p_ColEventos.Keys
        'If m_IDEvento = "SAN2008009" Then Stop
        Set m_ColMaterialesTotales = Nothing
        Set m_ObjEvento = p_ColEventos(m_IDEvento)
        m_IDEventoActual = m_ObjEvento.IDEVENTO
        m_AppExcel.Workbooks(m_ObjPlantilla.Name).Activate
        m_Linea = "Copiando plantila de excel para " & CStr(m_IDEvento) & " (" & intNumero & "\" & intNumeroTotal & ")"
        Avance m_Linea
        
       ' m_ObjPlantilla.Worksheets(1).Select
        m_AppExcel.Workbooks(m_ObjPlantilla.Name).Worksheets(1).Copy before:=m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1)
        
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Name = m_IDEventoActual
        RellenarDatosEvento m_ObjEvento, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ColActividades = m_ObjEvento.ColActividades
        p_Error = m_ObjEvento.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Linea = "Rellenando Actividades para " & CStr(m_IDEvento)
        Avance m_Linea
        RellenarActividades m_ColActividades, m_IntFilaFinalActividades, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColActividades Is Nothing Then
            If m_ColActividades.count > 0 Then
                For Each m_IDActividad In m_ColActividades.Keys
                    Set m_ObjActividad = m_ColActividades(m_IDActividad)
                    Set m_ColMateriales = m_ObjActividad.ColMateriales
                    p_Error = m_ObjActividad.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_ColMateriales Is Nothing Then
                        If m_ColMateriales.count > 0 Then
                            For Each m_IDMaterial In m_ColMateriales.Keys
                                If Nz(m_IDMaterial, "") <> "" Then
                                    Set m_ObjMaterial = m_ColMateriales(m_IDMaterial)
                                    If m_ColMaterialesTotales Is Nothing Then
                                        Set m_ColMaterialesTotales = New Scripting.Dictionary
                                        m_ColMaterialesTotales.Add m_IDMaterial & "|" & m_ObjActividad.IDActividad, m_ObjMaterial
                                    Else
                                        If Not m_ColMaterialesTotales.Exists(m_IDMaterial) Then
                                            m_ColMaterialesTotales.Add m_IDMaterial & "|" & m_ObjActividad.IDActividad, m_ObjMaterial
                                        End If
                                    End If
                                    Set m_ObjMaterial = Nothing
                                End If
                            Next
                        End If
                    End If
                    Set m_ObjActividad = Nothing
                Next
            End If
            
        End If
        m_Linea = "Rellenando Materiales para " & CStr(m_IDEvento)
        Avance m_Linea
        
        RellenarMateriales m_ColMaterialesTotales, m_IntFilaFinalActividades, m_intFilaFinalMateriales, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        intNumero = intNumero + 1
    Next
    m_Linea = "Formateando Excel"
    Avance m_Linea
    PonerBordeFinal m_intFilaFinalMateriales, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_ConHojaResumen = True Then
        GeneraHojaResumenInformeRAC m_AppExcel.Workbooks(m_ObjLibro.Name), p_ColEventos, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        PonerEnlaceAResumen m_AppExcel.Workbooks(m_ObjLibro.Name), p_ColEventos, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_AppExcel.Workbooks(m_ObjPlantilla.Name).Close False
    Set m_ObjPlantilla = Nothing
    m_AppExcel.Workbooks(m_ObjLibro.Name).SaveAs p_URLInforme
    m_AppExcel.Workbooks(m_ObjLibro.Name).Close False
    Set m_ObjLibro = Nothing
    m_AppExcel.Quit
    Set m_AppExcel = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EventosRAC ha devuelto: " & vbNewLine & Err.Description
    End If
    If Not m_AppExcel Is Nothing Then
        If Not m_ObjHoja Is Nothing Then
            Set m_ObjHoja = Nothing
        End If
        If Not m_ObjLibro Is Nothing Then
            m_ObjLibro.Close False
            Set m_ObjLibro = Nothing
        End If
        If Not m_ObjHojaPlantilla Is Nothing Then
            Set m_ObjHojaPlantilla = Nothing
        End If
        If Not m_ObjPlantilla Is Nothing Then
            m_ObjPlantilla.Close False
            Set m_ObjPlantilla = Nothing
        End If
        m_AppExcel.Quit
        Set m_AppExcel = Nothing
    End If
    
End Function
Private Function GeneraHojaResumenInformeRAC( _
                                            p_ObjLibro As excel.Workbook, _
                                            ByRef p_ObjColEventos As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                            
    Dim m_IDEvento As Variant
    Dim m_ObjEvento As Evento
    
    Dim intFila As Integer
    Dim lngNumeroRegistro As Long
    Dim lngNumerRegistrosTotales As Long
    On Error GoTo errores
    
    If p_ObjColEventos Is Nothing Then
        Exit Function
    End If
    m_Linea = "Generando Hoja Resúmen"
    Avance m_Linea
            
    m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets.Add m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(1)
    m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(1).Name = "Resumen"
    '-----------
    ' parte inicial de la cabecera
    '--------------
    intFila = 1
    With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
        .Cells(intFila, 1).Value = "PARTE TRABAJO"
        .Cells(intFila, 2).Value = "NODO"
        .Cells(intFila, 3).Value = "BUI"
        .Cells(intFila, 4).Value = "DESCRIPCIÓN"
        .Cells(intFila, 5).Value = "SUBSISTEMA"
        .Cells(intFila, 6).Value = "EQUIPO"
        .Cells(intFila, 7).Value = "PM/PR"
        .Cells(intFila, 8).Value = "FECHA INICIO"
        .Cells(intFila, 9).Value = "FECHA FIN"
        .Cells(intFila, 10).Value = "NºHORAS_LAB"
        .Cells(intFila, 11).Value = "NºHORAS_EXT"
    End With
    With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
        .Range(.Cells(1, 1), .Cells(1, 9)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 9)).Interior.ColorIndex = 15
        .Range(.Cells(1, 1), .Cells(1, 9)).HorizontalAlignment = xlCenter
    End With
    lngNumerRegistrosTotales = p_ObjColEventos.count
    lngNumeroRegistro = 1
    intFila = intFila + 1
    For Each m_IDEvento In p_ObjColEventos.Keys
        Set m_ObjEvento = p_ObjColEventos(m_IDEvento)
        m_Linea = "Generando Hoja Resúmen... " & "Evento " & m_ObjEvento.IDEVENTO & " ( " & lngNumeroRegistro & "\" & lngNumerRegistrosTotales & ")"
        Avance m_Linea
        With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
            If m_ObjEvento.IDEVENTO <> "" Then
                .Hyperlinks.Add Anchor:=.Range("A" & CStr(intFila)), _
                                     Address:="", SubAddress:=CStr(m_IDEvento) & "!A1", _
                                     TextToDisplay:=CStr(m_IDEvento)
             End If
             If m_ObjEvento.NODO <> "" Then
                 .Cells(intFila, 2).Value = m_ObjEvento.NODO
             End If
             If m_ObjEvento.BUI <> "" Then
                 .Cells(intFila, 3).Value = m_ObjEvento.BUI
             End If
             If m_ObjEvento.BUI <> "" Then
                 .Cells(intFila, 4).Value = m_ObjEvento.Descripcion
             End If
             If m_ObjEvento.SubSistema <> "" Then
                 .Cells(intFila, 5).Value = m_ObjEvento.SubSistema
             End If
             If m_ObjEvento.NombreEquipo <> "" Then
                 .Cells(intFila, 6).Value = m_ObjEvento.NombreEquipo
             End If
            
             If m_ObjEvento.PMPR <> "" Then
                 .Cells(intFila, 7).Value = "'" & m_ObjEvento.PMPR
             End If
             If m_ObjEvento.FECHAALTAEVENTO <> "" Then
                 .Cells(intFila, 8).Value = Format(m_ObjEvento.FECHAALTAEVENTO, "mm/dd/yyyy")
             End If
             If m_ObjEvento.FechaFinal <> "" Then
                 .Cells(intFila, 9).Value = Format(m_ObjEvento.FechaFinal, "mm/dd/yyyy")
             End If
             If m_ObjEvento.HorasLabTotales <> "" Then
                 .Cells(intFila, 10).Value = "'" & m_ObjEvento.HorasLabTotales
             End If
             If m_ObjEvento.HorasExtTotales <> "" Then
                 .Cells(intFila, 11).Value = "'" & m_ObjEvento.HorasExtTotales
             End If
         End With
        Set m_ObjEvento = Nothing
        intFila = intFila + 1
        lngNumeroRegistro = lngNumeroRegistro + 1
    Next
    With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
        RecuadrarRango .Range(.Cells(1, 1), .Cells(intFila - 1, 11)), "Multiple"
    
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
        .Columns("K:K").EntireColumn.AutoFit
    End With
    
    
    GeneraHojaResumenInformeRAC = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Certificacion.GeneraHojaResumenInformeRAC ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Function AbrirNuevoExcel( _
                                ByRef p_URLInforme As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim m_Alcanzable As Boolean
    Dim m_NombreInforme As String
    
    On Error GoTo errores
    If Not fso.FolderExists(m_ObjEntorno.URLDirectorioLocal) Then
        p_Error = "No se ha definido la plantilla para el informe" & vbNewLine & m_ObjEntorno.URLPlantillaInforme
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_ObjEntorno.URLPlantillaInforme) Then
        p_Error = "No se ha definido la plantilla para el informe" & vbNewLine & m_ObjEntorno.URLPlantillaInforme
        Err.Raise 1000
    End If
    
    m_NombreInforme = "Informe de Partes de Trabajo.xlsx"
    
    p_URLInforme = m_ObjEntorno.URLDirectorioLocal & m_NombreInforme
    If fso.FileExists(p_URLInforme) Then
        If FicheroAbierto(p_URLInforme) Then
            p_Error = "Tiene abierto otro informe previo"
            Err.Raise 1000
        End If
        fso.DeleteFile p_URLInforme, True
    End If
    
    
    If m_AppExcel Is Nothing Then
        Set m_AppExcel = New excel.Application
    End If
    m_AppExcel.Visible = False
    Set m_ObjPlantilla = m_AppExcel.Workbooks.Open(m_ObjEntorno.URLPlantillaInforme)
    Set m_ObjHojaPlantilla = m_ObjPlantilla.Worksheets(1)
    Set m_ObjLibro = m_AppExcel.Workbooks.Add
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EventosRAC ha devuelto: " & vbNewLine & Err.Description
    End If
    On Error Resume Next
    If Not m_AppExcel Is Nothing Then
        If Not m_ObjHoja Is Nothing Then
            Set m_ObjHoja = Nothing
        End If
        If Not m_ObjLibro Is Nothing Then
            m_ObjLibro.Close False
            Set m_ObjLibro = Nothing
        End If
        m_AppExcel.Quit
        Set m_AppExcel = Nothing
    End If
    
End Function


Private Function RellenarDatosEvento( _
                                    p_Evento As Evento, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim m_int As Integer
    Dim m_NombreCampoParaInforme As String
    Dim m_NombreCampoParaDato As String
    Dim m_intFilaTitular As Integer
    Dim m_IntFila As Integer
    Dim m_intFilaInicial As Integer
    Dim m_intFilaFinal As Integer
    
    Dim m_ObjRango As excel.Range
    Dim m_ValorEnRango As String
    Dim m_ValorDato As Variant
    Dim m_IDNombreCampo As Long
    On Error GoTo errores
    
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    
    m_intFilaTitular = 2
    Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("G2")
    m_ValorDato = p_Evento.getPropiedad("IDEvento", p_Error)
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjRango.Value = m_ValorDato
    
    Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("C9")
    m_ValorDato = p_Evento.HorasLabTotales
    p_Error = p_Evento.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjRango.Value = m_ValorDato
    
    Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("H9")
    m_ValorDato = p_Evento.HorasExtTotales
    p_Error = p_Evento.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjRango.Value = m_ValorDato
    
    
    m_intFilaInicial = m_intFilaTitular + 2
    m_intFilaFinal = 8
    
    For m_IntFila = m_intFilaInicial To m_intFilaFinal
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFila))
        m_ValorEnRango = m_ObjRango.Value
        m_ValorEnRango = Replace(m_ValorEnRango, "{{", "")
        m_ValorEnRango = Replace(m_ValorEnRango, "}}", "")
        If IsNumeric(m_ValorEnRango) Then
            
            m_IDNombreCampo = CLng(m_ValorEnRango)
            m_NombreCampoParaInforme = m_ColParaEnumNombreCamposEnInforme(CStr(m_IDNombreCampo))
            dato = Split(m_NombreCampoParaInforme, ".")
            m_ObjRango.Value = dato(1)
            m_NombreCampoParaDato = m_ColParaNombreCamposInformeTabla(m_NombreCampoParaInforme)
            'If m_NombreCampoParaDato = "NombreEquipo" Then Stop
            dato = Split(m_NombreCampoParaDato, ".")
            m_NombreCampoParaDato = dato(1)
            'If m_NombreCampoParaDato = "NombreEquipo" Then Stop
            m_ValorDato = p_Evento.getPropiedad(m_NombreCampoParaDato, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("C" & CStr(m_IntFila))
            If m_NombreCampoParaDato = "FECHAALTAEVENTO" Or m_NombreCampoParaDato = "FechaFinalCalculada" Or m_NombreCampoParaDato = "FechaFinal" Then
                If IsDate(m_ValorDato) Then
                    m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                    m_ObjRango.Value = m_ValorDato
                    m_ObjRango.NumberFormat = "dd/mm/yyyy"
                End If
            Else
                If m_NombreCampoParaDato = "PMPR" Then
                    m_ValorDato = "'" & m_ValorDato
                
                End If
                m_ObjRango.Value = m_ValorDato
            End If
            
            
        End If
        
    Next
    For m_IntFila = m_intFilaInicial To m_intFilaFinal
        
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("F" & CStr(m_IntFila))
        m_ValorEnRango = m_ObjRango.Value
        m_ValorEnRango = Replace(m_ValorEnRango, "{{", "")
        m_ValorEnRango = Replace(m_ValorEnRango, "}}", "")
        If IsNumeric(m_ValorEnRango) Then
            
            m_IDNombreCampo = CLng(m_ValorEnRango)
            m_NombreCampoParaInforme = m_ColParaEnumNombreCamposEnInforme(CStr(m_IDNombreCampo))
            dato = Split(m_NombreCampoParaInforme, ".")
            m_ObjRango.Value = dato(1)
            m_NombreCampoParaDato = m_ColParaNombreCamposInformeTabla(m_NombreCampoParaInforme)
            'If m_NombreCampoParaDato = "TbEventos.FranqueadoParaInforme" Then Stop
            dato = Split(m_NombreCampoParaDato, ".")
            m_NombreCampoParaDato = dato(1)
            'If p_Evento = "SAN2103019" And m_NombreCampoParaDato = "FechaFinal" Then Stop
            m_ValorDato = p_Evento.getPropiedad(m_NombreCampoParaDato, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("H" & CStr(m_IntFila))
            If m_NombreCampoParaDato = "FECHAALTAEVENTO" Or m_NombreCampoParaDato = "FechaFinalCalculada" Or m_NombreCampoParaDato = "FechaFinal" Then
                If IsDate(m_ValorDato) Then
                    m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                    m_ObjRango.Value = m_ValorDato
                    m_ObjRango.NumberFormat = "dd/mm/yyyy"
                End If
            Else
                m_ObjRango.Value = m_ValorDato
            End If
            
            
        End If
        
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarDatosEvento ha devuelto: " & vbNewLine & Err.Description
    End If
    
    
End Function
Private Function PonerEnlaceAResumen( _
                                    p_ObjLibro As excel.Workbook, _
                                    p_ObjColEventos As Scripting.Dictionary, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    
    Dim m_IDEvento As Variant
    Dim m_NombreHoja As String
       On Error GoTo errores
    
    If p_ObjColEventos Is Nothing Then
        Exit Function
    End If
    For Each m_IDEvento In p_ObjColEventos.Keys
        m_NombreHoja = CStr(m_IDEvento)
        m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(m_NombreHoja).Hyperlinks.Add m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(m_NombreHoja).Range("B1:B1"), "", m_AppExcel.Workbooks(p_ObjLibro.Name).Sheets(1).Name & "!A1", , "Ir a Resumen"
    Next
        
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PonerEnlaceAResumen ha devuelto: " & vbNewLine & Err.Description
    End If
    
    
End Function

Private Function RellenarActividades( _
                                        p_ColActividades As Scripting.Dictionary, _
                                        p_IntFilaUltimaActividad As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_ColColumnasNombreCampos As Scripting.Dictionary
    Dim m_ColNombreCamposLetraColumna As Scripting.Dictionary
    Dim m_IntFila As Integer
    Dim m_IDActividad As Variant
    Dim m_ObjActividad As Actividad
    Dim m_NombreCampoParaInformeCompleto As String
    Dim m_NombreCampoParaInforme As String
    Dim m_NombreCampoParaDatoCompleto As String
    Dim m_NombreCampoParaDato As String
    Dim m_ObjRango As excel.Range
    Dim m_ValorEnRango As String
    Dim m_ValorDato As Variant
    Dim m_NumeroActividad As Integer
    Dim m_Letra As Variant
    Dim m_Resto As String
    Dim m_intFilaInicial As Integer
    Dim m_intFilaTitulosActividades As Integer
    Dim m_intNumeroActividades As Integer
    Dim m_inti As Integer
    Dim m_IntFilaPrimeraActividad As Integer
    Dim m_IntFilaFinalActividad As Integer
    On Error GoTo errores
    
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    Set m_ColNombreCamposLetraColumna = New Scripting.Dictionary
    With m_ColNombreCamposLetraColumna
        .Add "B", m_ColParaEnumNombreCamposEnInforme("29") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("29"))
        .Add "D", m_ColParaEnumNombreCamposEnInforme("13") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("13"))
        .Add "E", m_ColParaEnumNombreCamposEnInforme("16") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("16"))
        .Add "F", m_ColParaEnumNombreCamposEnInforme("11") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("11"))
        .Add "G", m_ColParaEnumNombreCamposEnInforme("14") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("14"))
        .Add "H", m_ColParaEnumNombreCamposEnInforme("15") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("15"))
        .Add "I", m_ColParaEnumNombreCamposEnInforme("12") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("12"))
        .Add "J", m_ColParaEnumNombreCamposEnInforme("17") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("17"))
    End With
    m_intFilaTitulosActividades = 12
    'titulares
    For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_intFilaTitulosActividades))
        m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
        dato = Split(m_Resto, "|")
        m_NombreCampoParaInformeCompleto = dato(0)
        If InStr(1, m_NombreCampoParaInformeCompleto, ".") <> 0 Then
            dato = Split(m_NombreCampoParaInformeCompleto, ".")
            m_NombreCampoParaInforme = dato(1)
        Else
            m_NombreCampoParaInforme = m_NombreCampoParaInformeCompleto
        End If
        m_ObjRango.Value = m_NombreCampoParaInforme
    Next
    If p_ColActividades Is Nothing Then
        p_IntFilaUltimaActividad = m_intFilaTitulosActividades + 1
        Exit Function
    End If
    If p_ColActividades.count = 0 Then
        p_IntFilaUltimaActividad = m_intFilaTitulosActividades + 1
        Exit Function
    End If
    If p_ColActividades.count > 1 Then
        GenerarFilasDeActividades p_ColActividades, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_IntFila = m_intFilaTitulosActividades + 1
    m_NumeroActividad = 1
    m_IntFilaPrimeraActividad = m_IntFila
    For Each m_IDActividad In p_ColActividades.Keys
        m_Linea = "Rellenando Actividad: " & CStr(m_IDActividad)
        Avance m_Linea
        Set m_ObjActividad = p_ColActividades(m_IDActividad)
        For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
            Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_IntFila))
            m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
            dato = Split(m_Resto, "|")
            m_NombreCampoParaDatoCompleto = dato(1)
            If InStr(1, m_NombreCampoParaDatoCompleto, ".") <> 0 Then
                dato = Split(m_NombreCampoParaDatoCompleto, ".")
                m_NombreCampoParaDato = dato(1)
            Else
                m_NombreCampoParaDato = m_NombreCampoParaDatoCompleto
            End If
            
            m_ValorDato = m_ObjActividad.getPropiedad(m_NombreCampoParaDato, p_Error)
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_NombreCampoParaDato = "FechaAlta" Then
                If IsDate(m_ValorDato) Then
                    m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                    m_ObjRango.Value = m_ValorDato
                    m_ObjRango.NumberFormat = "dd/mm/yyyy"
                End If
            Else
                m_ObjRango.Value = m_ValorDato
            End If
        Next
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).EntireRow.AutoFit
        
        If m_NumeroActividad <> p_ColActividades.count Then
            m_IntFila = m_IntFila + 1
        End If
        Set m_ObjActividad = Nothing
        m_NumeroActividad = m_NumeroActividad + 1
    Next
    m_IntFilaFinalActividad = m_IntFila
    If p_ColActividades.count > 1 Then
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaFinalActividad) & ":M" & CStr(m_IntFilaFinalActividad)).Copy
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaPrimeraActividad) & _
                            ":M" & CStr(m_IntFilaFinalActividad - 1)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                            SkipBlanks:=False, Transpose:=False
        m_AppExcel.CutCopyMode = False
    End If
    
    p_IntFilaUltimaActividad = m_IntFilaFinalActividad
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarActividades ha devuelto: " & vbNewLine & Err.Description
    End If
    
    
End Function
Private Function GenerarFilasDeActividades( _
                                            ByRef p_ColActividades As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_IntFila As Integer
    Dim m_intFilaTitulosActividades As Integer
    Dim m_intNumeroActividades As Integer
    Dim m_inti As Integer
    On Error GoTo errores
    
    m_intFilaTitulosActividades = 11
    m_IntFila = m_intFilaTitulosActividades + 1
    'AHORA SE HA DE COPIAR Y PEGAR TANTAS VECES LA FILA 12 PARA CADA ACTIVIDAD QUE HAYA MENOS UNA
    m_intNumeroActividades = p_ColActividades.count
    If m_intNumeroActividades > 1 Then
        For m_inti = 1 To m_intNumeroActividades - 1
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).Copy
            
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila + 1) & ":" & CStr(m_IntFila + 1)).Insert Shift:=xlDown
            'Selection.Insert Shift:=xlDown
            m_IntFila = m_IntFila + 1
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarFilasDeActividades ha devuelto: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function GenerarFilasDeMateriales( _
                                            ByRef p_ColMateriales As Scripting.Dictionary, _
                                            P_IntFilaMaterialesTitulares As Integer, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_IntFila As Integer
    Dim m_intFilaTitulosMateriales As Integer
    Dim m_intNumeroMateriales As Integer
    Dim m_inti As Integer
    
    On Error GoTo errores
    
    m_IntFila = P_IntFilaMaterialesTitulares + 1
    'AHORA SE HA DE COPIAR Y PEGAR TANTAS VECES LA FILA 12 PARA CADA Material QUE HAYA MENOS UNA
    m_intNumeroMateriales = p_ColMateriales.count
    If m_intNumeroMateriales > 1 Then
        For m_inti = 1 To m_intNumeroMateriales - 1
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).Copy
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila + 1) & ":" & CStr(m_IntFila + 1)).Insert Shift:=xlDown
            
            m_IntFila = m_IntFila + 1
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarFilasDeMateriales ha devuelto: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function RellenarMateriales( _
                                        p_ColMateriales As Scripting.Dictionary, _
                                        p_IntFilaUltimaMaterial As Integer, _
                                        ByRef p_intFilaUltimaMateriales As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    Dim m_ColColumnasNombreCampos As Scripting.Dictionary
    Dim m_ColNombreCamposLetraColumna As Scripting.Dictionary
    Dim m_IntFila As Integer
    Dim m_IDMaterial As String
    Dim m_VarITem As Variant
    
    Dim m_ObjMaterial As Material
    Dim m_NombreCampoParaInformeCompleto As String
    Dim m_NombreCampoParaInforme As String
    Dim m_NombreCampoParaDatoCompleto As String
    Dim m_NombreCampoParaDato As String
    Dim m_ObjRango As excel.Range
    Dim m_ValorEnRango As String
    Dim m_ValorDato As Variant
    Dim m_NumeroMaterial As Integer
    Dim m_Letra As Variant
    Dim m_Resto As String
    Dim m_intFilaInicial As Integer
    Dim dato1
    Dim m_IntFilaTitularesMateriales As Integer
    Dim m_IntFilaPrimeraMaterial As Integer
    Dim m_IntFilaFinalMaterial As Integer
    On Error GoTo errores
    
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    Set m_ColNombreCamposLetraColumna = New Scripting.Dictionary
    With m_ColNombreCamposLetraColumna
        .Add "B", m_ColParaEnumNombreCamposEnInforme("29") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("29"))
        .Add "C", m_ColParaEnumNombreCamposEnInforme("30") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("30"))
        .Add "D", m_ColParaEnumNombreCamposEnInforme("18") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("18"))
        .Add "E", m_ColParaEnumNombreCamposEnInforme("21") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("21"))
        .Add "F", m_ColParaEnumNombreCamposEnInforme("22") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("22"))
        .Add "G", m_ColParaEnumNombreCamposEnInforme("23") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("23"))
        .Add "H", m_ColParaEnumNombreCamposEnInforme("19") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("19"))
        .Add "I", m_ColParaEnumNombreCamposEnInforme("26") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("26"))
        .Add "J", m_ColParaEnumNombreCamposEnInforme("20") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("20"))
        .Add "L", m_ColParaEnumNombreCamposEnInforme("24") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("24"))
        .Add "M", m_ColParaEnumNombreCamposEnInforme("25") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("25"))
    End With
    
    
    m_IntFilaTitularesMateriales = p_IntFilaUltimaMaterial + 3
    'titulares
    For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_IntFilaTitularesMateriales))
        m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
        dato = Split(m_Resto, "|")
        m_NombreCampoParaInformeCompleto = dato(0)
        If InStr(1, m_NombreCampoParaInformeCompleto, ".") <> 0 Then
            dato = Split(m_NombreCampoParaInformeCompleto, ".")
            m_NombreCampoParaInforme = dato(1)
        Else
            m_NombreCampoParaInforme = m_NombreCampoParaInformeCompleto
        End If
        m_ObjRango.Value = m_NombreCampoParaInforme
    Next
    If p_ColMateriales Is Nothing Then
        p_intFilaUltimaMateriales = m_IntFilaTitularesMateriales + 1
        Exit Function
    End If
    If p_ColMateriales.count = 0 Then
        p_intFilaUltimaMateriales = m_IntFilaTitularesMateriales + 1
        Exit Function
    End If
    If p_ColMateriales.count > 1 Then
        GenerarFilasDeMateriales p_ColMateriales, m_IntFilaTitularesMateriales, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_NumeroMaterial = 1
    m_IntFila = m_IntFilaTitularesMateriales + 1
    m_IntFilaPrimeraMaterial = m_IntFila
    For Each m_VarITem In p_ColMateriales.Keys
        dato = Split(m_VarITem, "|")
        m_IDMaterial = dato(0)
        
        If p_ColMateriales.Exists(m_VarITem) Then
            Set m_ObjMaterial = p_ColMateriales(m_VarITem)
            For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
                Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_IntFila))
                If m_Letra = "B" Then
                    m_ObjRango.Value = m_IDMaterial
                Else
                    m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
                    dato1 = Split(m_Resto, "|")
                    m_NombreCampoParaDatoCompleto = dato1(1)
                    If InStr(1, m_NombreCampoParaDatoCompleto, ".") <> 0 Then
                        
                        dato1 = Split(m_NombreCampoParaDatoCompleto, ".")
                        m_NombreCampoParaDato = dato1(1)
                    Else
                        m_NombreCampoParaDato = m_NombreCampoParaDatoCompleto
                    End If
                    m_ValorDato = m_ObjMaterial.getPropiedad(m_NombreCampoParaDato, p_Error)
                    
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If m_NombreCampoParaDato = "FechaIncio" Or m_NombreCampoParaDato = "FechaEntrega" Then
                        If IsDate(m_ValorDato) Then
                            m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                            m_ObjRango.Value = m_ValorDato
                            m_ObjRango.NumberFormat = "dd/mm/yyyy"
                        End If
                    Else
                        m_ObjRango.Value = m_ValorDato
                    End If
                End If
            Next
        End If
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).EntireRow.AutoFit
        If m_NumeroMaterial <> p_ColMateriales.count Then
            m_IntFila = m_IntFila + 1
        End If
        Set m_ObjMaterial = Nothing
        m_NumeroMaterial = m_NumeroMaterial + 1
    Next
    m_IntFilaFinalMaterial = m_IntFila
    If p_ColMateriales.count > 1 Then
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaFinalMaterial) & ":M" & CStr(m_IntFilaFinalMaterial)).Copy
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaPrimeraMaterial) & _
                            ":M" & CStr(m_IntFilaFinalMaterial - 1)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                            SkipBlanks:=False, Transpose:=False
        m_AppExcel.CutCopyMode = False
    End If
    p_intFilaUltimaMateriales = m_IntFilaFinalMaterial
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarMateriales ha devuelto: " & m_ObjRango.Address & vbNewLine & Err.Description
    End If
    
    
    
End Function

Public Function PonerBordeFinal( _
                                    ByRef m_intFilaFinalMateriales, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_intFilaInicial As Integer
    Dim m_intFilaFinal As Integer
    Dim m_Hoja As excel.Worksheet
    Dim m_Rango As excel.Range
    
    On Error GoTo errores
    
    Set m_Hoja = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1)
    m_intFilaInicial = 2
    
    m_intFilaFinal = m_intFilaFinalMateriales
    If m_intFilaFinalMateriales < m_intFilaInicial Then
        p_Error = "La fila final de los materiales es menor que la inicial"
        Err.Raise 1000
    End If
    
    Set m_Rango = m_Hoja.Range("B2:M" & CStr(m_intFilaFinal))
    With m_Rango
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    Set m_Rango = m_Hoja.Range("B2:M3")
    With m_Rango
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PonerBordeFinal ha devuelto: " & vbNewLine & Err.Description
    End If

End Function



Private Sub Class_Terminate()
    If Not m_AppExcel Is Nothing Then
        m_AppExcel.Quit
        Set m_AppExcel = Nothing
    End If
End Sub

