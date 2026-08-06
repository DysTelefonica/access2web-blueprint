Attribute VB_Name = "Helper_ExpedientesExportExcel"
Option Compare Database
Option Explicit
' SDD: staging-alignment-prueba-003 / PR-REFAC-1b
' Helper de exportación a Excel (.xlsx vía COM). Mueve la lógica de
' FUNCIONES UTILES (GenerarConsultaExpedientes + RellenarLinea + ConvertirATabla +
' AjustarCeldas) a un módulo dedicado, testeable end-to-end con Dysflow.
' STATELESS: depende de globales (m_ObjEntorno, fso) y de Excel COM,
' pero NO de controles de formulario.

' ==============================================================================
' GenerarConsultaExpedientes - entry principal. Abre Excel COM, escribe la
' cabecera desde p_ColCampos, itera las colecciones (AM/Lotes/Basados/Tecnica)
' llamando a RellenarLinea, ajusta celdas y convierte a tabla.
' Retorna: la URL local del .xlsx generado, o "" + p_Error si falla.
' ==============================================================================
Public Function GenerarConsultaExpedientes( _
                                            p_ColCampos As Scripting.Dictionary, _
                                            Optional p_ColAM As Scripting.Dictionary, _
                                            Optional p_ColLotes As Scripting.Dictionary, _
                                            Optional p_ColBasados As Scripting.Dictionary, _
                                            Optional p_ColTecnica As Scripting.Dictionary, _
                                            Optional p_IncluirDerivados As Boolean = True, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    Dim m_ID As Variant
    Dim m_ColExpUsados As Scripting.Dictionary
    Dim m_ExpC As ExpedienteCompleto
    Dim m_Campo As Variant
    Dim m_NombreCampo As String
    Dim intFila As Integer
    Dim columna As Integer
    Dim m_NombreArchivo As String
    Dim m_URLExcel As String
    Dim appExcel As Excel.Application
    Dim wbLibro As Excel.Workbook
    Dim wbHoja As Excel.Worksheet
    On Error GoTo errores

    ' REFAC-1b: reset cancel flag al inicio (CerrarPopupProgreso lo vuelve a setear
    ' a False al cerrar la popup; pero si se llama sin popup abierta, igual reseteamos).
    g_OperationCancelled = False

    If p_ColCampos Is Nothing Then
        GenerarConsultaExpedientes = ""
        Exit Function
    End If
    If p_ColAM Is Nothing And p_ColLotes Is Nothing And p_ColBasados Is Nothing And p_ColTecnica Is Nothing Then
        GenerarConsultaExpedientes = ""
        Exit Function
    End If
    If Not p_ColTecnica Is Nothing Then
        Set p_ColBasados = p_ColTecnica
    End If

    ' REFAC-1b: chequeo de cancel ANTES de abrir la popup. Tests setean
    ' g_OperationCancelled = True antes de invocar para abortar sin colgar el COM.
    If g_OperationCancelled Then
        GenerarConsultaExpedientes = ""
        Exit Function
    End If

    MostrarPopupProgreso "Exportando a Excel", "Preparando consulta..."
    AnimarProgresoIndefinido

    ' REFAC-1b: chequeo de cancel despues de abrir la popup. Si el usuario hace
    ' click en un futuro boton Cancel, aborta antes de abrir Excel.
    If g_OperationCancelled Then
        CerrarPopupProgreso
        GenerarConsultaExpedientes = ""
        Exit Function
    End If

    Avance "Abriendo excel para rellenar ..."
    m_NombreArchivo = fso.GetTempName & ".xlsx"
    m_URLExcel = m_ObjEntorno.URLDirectorioLocal & m_NombreArchivo
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Tiene una consulta abierta"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    Set appExcel = New Excel.Application
    appExcel.Visible = False
    Set wbLibro = appExcel.Workbooks.Add
    wbLibro.SaveAs m_URLExcel
    Set wbHoja = wbLibro.Worksheets(1)
    intFila = 1
    columna = 0
    'CABECERA
    With wbHoja
        For Each m_Campo In p_ColCampos
            m_NombreCampo = p_ColCampos(m_Campo)
            columna = columna + 1
            .Cells(intFila, columna).value = m_NombreCampo
        Next

    End With
    If Not p_ColAM Is Nothing Then
        For Each m_ID In p_ColAM
            If g_OperationCancelled Then
                wbLibro.Close False
                appExcel.Quit
                CerrarPopupProgreso
                GenerarConsultaExpedientes = ""
                Exit Function
            End If
            ActualizarEstadoPopup  "Procesando lote AM: " & m_ID
            DoEvents
            Set m_ExpC = p_ColAM(m_ID)
            RellenarLinea p_ColCampos, m_ColExpUsados, wbHoja, intFila, m_ExpC, p_IncluirDerivados, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ExpC = Nothing
        Next
    End If
    If Not p_ColLotes Is Nothing Then
        For Each m_ID In p_ColLotes
            If g_OperationCancelled Then
                wbLibro.Close False
                appExcel.Quit
                CerrarPopupProgreso
                GenerarConsultaExpedientes = ""
                Exit Function
            End If
            ActualizarEstadoPopup  "Procesando lote Lotes: " & m_ID
            DoEvents
            Set m_ExpC = p_ColLotes(m_ID)
            RellenarLinea p_ColCampos, m_ColExpUsados, wbHoja, intFila, m_ExpC, p_IncluirDerivados, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ExpC = Nothing
        Next
    End If
    If Not p_ColBasados Is Nothing Then
        For Each m_ID In p_ColBasados
            If g_OperationCancelled Then
                wbLibro.Close False
                appExcel.Quit
                CerrarPopupProgreso
                GenerarConsultaExpedientes = ""
                Exit Function
            End If
            ActualizarEstadoPopup  "Procesando lote Basados: " & m_ID
            DoEvents
            Set m_ExpC = p_ColBasados(m_ID)
            RellenarLinea p_ColCampos, m_ColExpUsados, wbHoja, intFila, m_ExpC, p_IncluirDerivados, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ExpC = Nothing
        Next
    End If
    AjustarCeldas p_Hoja:=wbHoja, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ConvertirATabla p_Hoja:=wbHoja, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For columna = 1 To p_ColCampos.Count
        With wbHoja
            .Columns(columna).EntireColumn.AutoFit
        End With
    Next
    wbLibro.Close True
    Set wbLibro = Nothing
    appExcel.Quit
    Set appExcel = Nothing

    GenerarConsultaExpedientes = m_URLExcel
    CerrarPopupProgreso
    Exit Function

errores:
    p_Error = "El método Helper_ExpedientesExportExcel.GenerarConsultaExpedientes ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    On Error Resume Next
    CerrarPopupProgreso
    If Not wbLibro Is Nothing Then wbLibro.Close False
    If Not appExcel Is Nothing Then appExcel.Quit
    GenerarConsultaExpedientes = ""
End Function

' ==============================================================================
' RellenarLinea - escribe una fila por cada ExpedienteCompleto + recursivo para
' Derivados (basados en p_ExpC.Derivados si p_IncluirDerivados=True).
' m_ColExpUsados evita duplicados de IDExpediente.
' ==============================================================================
Public Function RellenarLinea( _
                                p_ColCampos As Scripting.Dictionary, _
                                ByRef m_ColExpUsados As Scripting.Dictionary, _
                                ByRef wbHoja As Excel.Worksheet, _
                                ByRef intFila As Integer, _
                                p_ExpC As ExpedienteCompleto, _
                                Optional p_IncluirDerivados As Boolean = True, _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim m_ID As Variant
    Dim m_ExpCDerivado As ExpedienteCompleto
    Dim m_Campo As Variant
    Dim m_NombreCampo As String
    Dim columna As Integer
    Dim m_Valor As String
    On Error GoTo errores

    If p_ColCampos Is Nothing Then
        Exit Function
    End If
    If Not m_ColExpUsados Is Nothing Then
        If m_ColExpUsados.Exists(CStr(p_ExpC.IDExpediente)) Then
            Exit Function
        End If
    End If
    intFila = intFila + 1

    Avance p_ExpC.IDExpediente & "..........." & p_ExpC.Nemotecnico

    VBA.DoEvents
    With wbHoja
        columna = 0
        For Each m_Campo In p_ColCampos
            m_Valor = p_ExpC.getPropiedad(m_Campo, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_Campo = "Estado" Then
                m_Valor = p_ExpC.ESTADOCalculadoTexto
            End If
            If CStr(m_Campo) = "ImporteContratacion" And IsNumeric(m_Valor) Then
                m_Valor = Replace(m_Valor, ",", ".")
            ElseIf InStr(1, m_Campo, "Fecha") <> 0 And IsDate(m_Valor) Then
                m_Valor = Format(m_Valor, "mm/dd/yyyy")
            ElseIf CStr(m_Campo) = "ResponsableSeguridad" Then
                If m_Valor = "0" Or m_Valor = "" Then
                    m_Valor = "N/A"
                End If
            End If
            columna = columna + 1
            If m_Valor <> "" Then
                .Cells(intFila, columna).value = m_Valor
            End If
        Next

    End With
    If m_ColExpUsados Is Nothing Then
        Set m_ColExpUsados = New Scripting.Dictionary
        m_ColExpUsados.CompareMode = TextCompare
    End If
    If Not m_ColExpUsados.Exists(CStr(p_ExpC.IDExpediente)) Then
        m_ColExpUsados.Add CStr(p_ExpC.IDExpediente), p_ExpC.IDExpediente
    End If

    If p_IncluirDerivados Then
        If Not p_ExpC.Derivados Is Nothing Then
            For Each m_ID In p_ExpC.Derivados
                Set m_ExpCDerivado = p_ExpC.Derivados(m_ID)
                RellenarLinea = RellenarLinea(p_ColCampos, m_ColExpUsados, wbHoja, intFila, m_ExpCDerivado, p_IncluirDerivados, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                Set m_ExpCDerivado = Nothing
            Next
        End If
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Helper_ExpedientesExportExcel.RellenarLinea ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If

End Function

' ==============================================================================
' ConvertirATabla - convierte el rango A1:CurrentRegion a ListObject de Excel.
' ==============================================================================
Public Function ConvertirATabla( _
                                        p_Hoja As Excel.Worksheet, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_Rango As Excel.Range
    On Error GoTo errores
    Set m_Rango = p_Hoja.Range("A1").CurrentRegion
    p_Hoja.Application.CutCopyMode = False
    p_Hoja.ListObjects.Add(xlSrcRange, m_Rango, , xlYes).Name = "Tabla1"

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Helper_ExpedientesExportExcel.ConvertirATabla ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

' ==============================================================================
' AjustarCeldas - formato por defecto de las celdas de la hoja.
' ==============================================================================
Public Function AjustarCeldas( _
                                p_Hoja As Excel.Worksheet, _
                                Optional ByRef p_Error As String _
                                ) As String

    On Error GoTo errores


    With p_Hoja.Cells
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Helper_ExpedientesExportExcel.AjustarCeldas ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
