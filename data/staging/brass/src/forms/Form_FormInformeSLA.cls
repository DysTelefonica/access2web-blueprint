
Option Compare Database
Option Explicit

'===========================================================
' Form_FormInformeSLA
' Spec-010: Gestor de Informes SLA (solo eventos franqueados)
'===========================================================

'-----------------------------------------------------------
' Variables privadas del formulario
'-----------------------------------------------------------
Private m_strIDEventoSeleccionado As String
Private m_strFechaInicioFiltro As String
Private m_strFechaFinFiltro As String

Private Const LISTA_COL_COUNT As Integer = 6
Private Const LISTA_COL_WIDTHS As String = "3,591cm;2,501cm;4cm;3cm;2cm;3cm"

'-----------------------------------------------------------
' ObtenerEstadoFacturadoEvento
' Devuelve "Sí" si el evento está facturado, "No" en caso contrario
'-----------------------------------------------------------
Private Function ObtenerEstadoFacturadoEvento(ByVal p_IDEvento As String) As String
    Dim m_Total As Long

    On Error GoTo errores

    If Nz(p_IDEvento, "") = "" Then
        ObtenerEstadoFacturadoEvento = "No"
        Exit Function
    End If

    m_Total = DCount("*", "TbFacturaEventosInvolucrados", "IDEvento='" & Replace(p_IDEvento, "'", "''") & "'")

    If m_Total > 0 Then
        ObtenerEstadoFacturadoEvento = "Sí"
    Else
        ObtenerEstadoFacturadoEvento = "No"
    End If
    Exit Function

errores:
    ObtenerEstadoFacturadoEvento = "No"
End Function

'-----------------------------------------------------------
' FilaCabeceraLista
' Cabecera fija para ListaFiltrados (Value List)
'-----------------------------------------------------------
Private Function FilaCabeceraLista() As String
    FilaCabeceraLista = "IDEvento;Fecha Alta;BUI;Subsistema;Criticidad;F. Recepción"
End Function

'-----------------------------------------------------------
' InicializarListaFiltrados
' Define estructura fija y cabecera aunque no haya resultados
'-----------------------------------------------------------
Private Sub InicializarListaFiltrados()
    Me.ListaFiltrados.RowSourceType = "Value List"
    Me.ListaFiltrados.ColumnHeads = False
    Me.ListaFiltrados.ColumnCount = LISTA_COL_COUNT
    Me.ListaFiltrados.ColumnWidths = LISTA_COL_WIDTHS
    Me.ListaFiltrados.RowSource = FilaCabeceraLista()
End Sub

'-----------------------------------------------------------
' cmdObtgenerEventos_Click
' Valida fechas y carga ListaFiltrados con eventos franqueados
' que tengan al menos un campo SLA informado
'-----------------------------------------------------------
Private Sub cmdObtgenerEventos_Click()
    Dim strTextoError As String
    Dim strFechaInicio As String
    Dim strFechaFin As String
    Dim rcdEventos As DAO.Recordset
    Dim strSQL As String
    Dim strListaValores As String
    Dim strValorCelda As String
    Dim intCol As Integer
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    '----- Validación de fechas -----
    strFechaInicio = Nz(Me.txtFechaInicioInforme.Value, "")
    strFechaFin = Nz(Me.txtFechaFInInforme.Value, "")
    
    If strFechaInicio = "" Then
        strTextoError = "Debe introducir una fecha de inicio."
        Err.Raise 1000
    End If
    
    If strFechaFin = "" Then
        strTextoError = "Debe introducir una fecha fin."
        Err.Raise 1000
    End If
    
    If Not IsDate(strFechaInicio) Then
        strTextoError = "La fecha de inicio no es válida."
        Err.Raise 1000
    End If
    
    If Not IsDate(strFechaFin) Then
        strTextoError = "La fecha fin no es válida."
        Err.Raise 1000
    End If
    
    If CDate(strFechaInicio) > CDate(strFechaFin) Then
        strTextoError = "La fecha de inicio no puede ser mayor que la fecha fin."
        Err.Raise 1000
    End If
    
    ' Guardar fechas para exportar
    m_strFechaInicioFiltro = strFechaInicio
    m_strFechaFinFiltro = strFechaFin
    
    '----- Query de eventos franqueados con SLA (delegado a SLAReportService) -----
    Dim strErrorSQL As String
    strSQL = SLAReportService.ConstruirSQLEventosFranqueados(strFechaInicio, strFechaFin, strErrorSQL)
    If strErrorSQL <> "" Then
        strTextoError = "Error al construir consulta: " & strErrorSQL
        Err.Raise 1000
    End If
    Set rcdEventos = CurrentDb().OpenRecordset(strSQL)
    
    '----- Cargar ListaFiltrados -----
    strListaValores = FilaCabeceraLista()
    
    With rcdEventos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strListaValores = strListaValores & ";"
                ' Columnas: IDEvento | FechaAlta | BUI | Subsistema | Criticidad | FechaRecepcionNotificacion
                strValorCelda = Nz(.Fields("IDEVENTO"), "")
                strListaValores = strListaValores & strValorCelda & ";"
                
                strValorCelda = Format(Nz(.Fields("FECHAALTAEVENTO"), ""), "dd/mm/yyyy")
                strListaValores = strListaValores & strValorCelda & ";"
                
                strValorCelda = Nz(.Fields("BUI"), "")
                strListaValores = strListaValores & strValorCelda & ";"
                
                strValorCelda = Nz(.Fields("SUBSISTEMA"), "")
                strListaValores = strListaValores & strValorCelda & ";"
                
                strValorCelda = Nz(.Fields("CRITICIDAD"), "")
                strListaValores = strListaValores & strValorCelda & ";"
                
                ' FechaRecepcionNotificacion - clave SLA
                strValorCelda = Format(Nz(.Fields("FechaRecepcionNotificacion"), ""), "dd/mm/yyyy")
                If strValorCelda = "30/12/1899" Or strValorCelda = "" Then
                    strValorCelda = ""
                End If
                strListaValores = strListaValores & strValorCelda
                
                .MoveNext
            Loop
            
            Me.cmdExportarHTML.Enabled = True
        Else
            
            Me.cmdExportarHTML.Enabled = False
        End If
    End With
    
    rcdEventos.Close
    Set rcdEventos = Nothing
    
    ' Asignar a la lista (ColumnCount=6)
    Me.ListaFiltrados.RowSourceType = "Value List"
    Me.ListaFiltrados.RowSource = strListaValores
    
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al obtener eventos: " & Err.Number & vbNewLine & Err.Description
    End If
    DoCmd.Hourglass False
    VBA.DoEvents
    If strTextoError <> "" Then
        MsgBox strTextoError, vbCritical, "Error"
    End If
End Sub

'-----------------------------------------------------------
' cmdVerEvento_Click
' Abre el detalle del evento seleccionado en modo solo lectura
'-----------------------------------------------------------
Private Sub cmdVerEvento_Click()
    Dim strTextoError As String
    Dim strIDEvento As String
    Dim lngIndex As Long
    Dim strEstaFacturado As String
    Dim strFranqueado As String
    Dim m_ObjEvento As Evento
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    ' Obtener IDEvento seleccionado
    lngIndex = Me.ListaFiltrados.ListIndex
    If lngIndex < 0 Then
        strTextoError = "Debe seleccionar un evento de la lista."
        Err.Raise 1000
    End If

    If lngIndex = 0 Then
        strTextoError = "Debe seleccionar una fila de datos, no la cabecera."
        Err.Raise 1000
    End If
    
    ' La columna 0 contiene el IDEvento
    strIDEvento = Nz(Me.ListaFiltrados.Column(0, lngIndex), "")
    If strIDEvento = "" Then
        strTextoError = "No se pudo obtener el identificador del evento."
        Err.Raise 1000
    End If
    
    m_strIDEventoSeleccionado = strIDEvento

    Set m_ObjEvento = Constructor.getEvento(strIDEvento)
    If m_ObjEvento Is Nothing Then
        strTextoError = "No se pudo cargar el evento seleccionado."
        Err.Raise 1000
    End If

    strEstaFacturado = ObtenerEstadoFacturadoEvento(strIDEvento)
    strFranqueado = "Sí" ' ListaFiltrados solo muestra eventos franqueados

    If FormularioAbierto("FormEventoEdicion") Then
        DoCmd.Close acForm, "FormEventoEdicion", acSaveNo
    End If

    Set m_ObjEventoActivo = m_ObjEvento
    DoCmd.OpenForm "FormEventoEdicion", , , , , , strIDEvento & "|" & strEstaFacturado & "|" & strFranqueado
    
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al abrir el evento: " & Err.Number & vbNewLine & Err.Description
    End If
    DoCmd.Hourglass False
    VBA.DoEvents
    If strTextoError <> "" Then
        MsgBox strTextoError, vbCritical, "Error"
    End If
End Sub

'-----------------------------------------------------------
' cmdDetalleSLA_Click
' Abre FormEventoSLA para el IDEvento seleccionado en modo solo lectura
'-----------------------------------------------------------
Private Sub cmdDetalleSLA_Click()
    Dim strTextoError As String
    Dim strIDEvento As String
    Dim lngIndex As Long
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    ' Obtener IDEvento seleccionado
    lngIndex = Me.ListaFiltrados.ListIndex
    If lngIndex < 0 Then
        strTextoError = "Debe seleccionar un evento de la lista."
        Err.Raise 1000
    End If

    If lngIndex = 0 Then
        strTextoError = "Debe seleccionar una fila de datos, no la cabecera."
        Err.Raise 1000
    End If
    
    ' La columna 0 contiene el IDEvento
    strIDEvento = Nz(Me.ListaFiltrados.Column(0, lngIndex), "")
    If strIDEvento = "" Then
        strTextoError = "No se pudo obtener el identificador del evento."
        Err.Raise 1000
    End If
    
    ' Abrir formulario SLA con OpenArgs=IDEvento (read-only)
    DoCmd.OpenForm "FormEventoSLA", acNormal, , , acFormReadOnly, , strIDEvento
    
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al abrir el detalle SLA: " & Err.Number & vbNewLine & Err.Description
    End If
    DoCmd.Hourglass False
    VBA.DoEvents
    If strTextoError <> "" Then
        MsgBox strTextoError, vbCritical, "Error"
    End If
End Sub

'-----------------------------------------------------------
' cmdLimpiarEventos_Click
' Limpia fechas y lista de eventos
'-----------------------------------------------------------
Private Sub cmdLimpiarEventos_Click()
    On Error Resume Next
    
    Me.txtFechaInicioInforme.Value = Null
    Me.txtFechaFInInforme.Value = Null
    Call InicializarListaFiltrados
    
    m_strIDEventoSeleccionado = ""
    m_strFechaInicioFiltro = ""
    m_strFechaFinFiltro = ""
End Sub

'-----------------------------------------------------------
' cmdSalir_Click
' Cierra el formulario
'-----------------------------------------------------------
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

'-----------------------------------------------------------
' cmdExportarHTML_Click
''
'' FIX 3265: Añadida validación de recordset y campos antes de acceso.
'' Error 3265 "Item not found in collection" ocurre cuando se accede
'' a un campo que no existe en el recordset. Las causas raíz eran:
'' 1) Acceso a campos sin verificar que el recordset tiene datos
'' 2) Sin validación de existencia de campo antes de leer
''-----------------------------------------------------------
'Private Sub cmdExportarExcel_Click()
'    Dim strTextoError As String
'    Dim strFechaInicio As String
'    Dim strFechaFin As String
'    Dim strSQL As String
'    Dim rcdEventos As DAO.Recordset
'    Dim strRutaArchivo As String
'    Dim strNombreArchivo As String
'    Dim strSeparador As String
'    Dim strLinea As String
'    Dim objFSO As Object
'    Dim objArchivo As Object
'    Dim objExcel As Object
'    Dim objWorkbook As Object
'    Dim objHoja As Object
'    Dim intFila As Integer
'    Dim intCol As Integer
'    Dim strValorCelda As String
'    Dim m_CumpleTRES As Boolean, m_CumpleTRCM As Boolean, m_CumpleTRSS As Boolean, m_CumpleGlobal As Boolean
'    Dim m_Incidencia As Variant, m_ServicioAfectado As Variant
'    Dim fld As DAO.Field
'
'    On Error GoTo errores
'
'    VBA.DoEvents
'    DoCmd.Hourglass True
'    VBA.DoEvents
'
'    ' Usar fechas guardadas o las actuales del formulario
'    strFechaInicio = m_strFechaInicioFiltro
'    strFechaFin = m_strFechaFinFiltro
'
'    If strFechaInicio = "" Then
'        strFechaInicio = Nz(Me.txtFechaInicioInforme.Value, "")
'    End If
'    If strFechaFin = "" Then
'        strFechaFin = Nz(Me.txtFechaFInInforme.Value, "")
'    End If
'
'    If strFechaInicio = "" Or strFechaFin = "" Then
'        strTextoError = "No hay datos filtrados para exportar. Ejecute primero la búsqueda."
'        Err.Raise 1000
'    End If
'
'    ' Verificar que hay datos en la lista
'    If Me.ListaFiltrados.ListCount = 0 Then
'        strTextoError = "No hay datos en la lista para exportar."
'        Err.Raise 1000
'    End If
'
'    ' Construir nombre de archivo
'    strNombreArchivo = "Informe_SLA_" & Format(Now(), "yyyyMMdd_HHmmss") & ".xlsx"
'
'    ' Intentar usar directorio Documents del usuario
'    strSeparador = IIf(InStr(1, Environ("USERPROFILE"), "\") > 0, "\", "/")
'    strRutaArchivo = Environ("USERPROFILE") & strSeparador & "Documents" & strSeparador & strNombreArchivo
'
'    ' Query para obtener datos completos (delegado a SLAReportService)
'    Dim strErrorSQL As String
'    strSQL = SLAReportService.ConstruirSQLEventosFranqueados(strFechaInicio, strFechaFin, strErrorSQL)
'    If strErrorSQL <> "" Then
'        strTextoError = "Error al construir consulta: " & strErrorSQL
'        Err.Raise 1000
'    End If
'    Set rcdEventos = CurrentDb().OpenRecordset(strSQL)
'
'    ' FIX 3265: Verificar que el recordset se abrió correctamente y tiene campos
'    If rcdEventos Is Nothing Then
'        strTextoError = "Error al abrir el recordset de eventos."
'        Err.Raise 1000
'    End If
'
'    If rcdEventos.Fields.count = 0 Then
'        strTextoError = "El recordset de eventos está vacío (sin campos)."
'        rcdEventos.Close
'        Set rcdEventos = Nothing
'        Err.Raise 1000
'    End If
'
'    ' FIX 3265: Verificar campos requeridos para el informe SLA
'    ' Esto previene el error 3265 mostrando un mensaje claro si falta algún campo
'    Dim arrCamposRequeridos As Variant
'    Dim i As Integer
'    arrCamposRequeridos = Array("IDEVENTO", "FECHAALTAEVENTO", "BUI", "SUBSISTEMA", "CRITICIDAD", _
'        "FechaRecepcionNotificacion", "FechaInicioContactoCliente", "IncidenciaAveriaOReparacion", _
'        "TipoReparacion", "Urgente", "EventoConServicioAfectado", "FechaRestablecimientoServicio", _
'        "FechaInicioTiempoAdquisicion", "FechaFinTiempoAdquisicion")
'
'    For i = LBound(arrCamposRequeridos) To UBound(arrCamposRequeridos)
'        If Not ExisteCampoEnRecordset(rcdEventos, CStr(arrCamposRequeridos(i))) Then
'            strTextoError = "Campo requerido no encontrado en el recordset: " & arrCamposRequeridos(i) & _
'                ". Esto puede causar error 3265 durante la exportación."
'            rcdEventos.Close
'            Set rcdEventos = Nothing
'            Err.Raise 1000
'        End If
'    Next i
'
'    ' Crear objeto Excel y exportar
'    Set objExcel = CreateObject("Excel.Application")
'    objExcel.Visible = False
'    Set objWorkbook = objExcel.Workbooks.Add
'    Set objHoja = objWorkbook.Sheets(1)
'    objHoja.Name = "Informe SLA"
'
'    ' Encabezados
'    objHoja.Cells(1, 1).Value = "IDEVENTO"
'    objHoja.Cells(1, 2).Value = "Fecha Alta"
'    objHoja.Cells(1, 3).Value = "BUI"
'    objHoja.Cells(1, 4).Value = "Subsistema"
'    objHoja.Cells(1, 5).Value = "Criticidad"
'    objHoja.Cells(1, 6).Value = "Fecha Recepción Notif."
'    objHoja.Cells(1, 7).Value = "Fecha Inicio Contacto"
'    objHoja.Cells(1, 8).Value = "Incidencia/Avería"
'    objHoja.Cells(1, 9).Value = "Tipo Reparación"
'    objHoja.Cells(1, 10).Value = "Urgente"
'    objHoja.Cells(1, 11).Value = "Servicio Afectado"
'    objHoja.Cells(1, 12).Value = "Fecha Restablecimiento"
'    objHoja.Cells(1, 13).Value = "Cumple TRES"
'    objHoja.Cells(1, 14).Value = "Cumple TRCM"
'    objHoja.Cells(1, 15).Value = "Cumple TRSS"
'    objHoja.Cells(1, 16).Value = "Cumple Global"
'
'    ' Estilo cabecera
'    objHoja.Range("A1:P1").Font.Bold = True
'    objHoja.Range("A1:P1").Interior.Color = RGB(221, 235, 247)
'
'    ' Datos
'    intFila = 2
'    With rcdEventos
'        If Not .EOF Then
'            .MoveFirst
'            Do While Not .EOF
'                objHoja.Cells(intFila, 1).Value = Nz(.Fields("IDEVENTO"), "")
'
'                If IsDate(.Fields("FECHAALTAEVENTO")) Then
'                    objHoja.Cells(intFila, 2).Value = .Fields("FECHAALTAEVENTO")
'                    objHoja.Cells(intFila, 2).NumberFormat = "dd/mm/yyyy"
'                End If
'
'                objHoja.Cells(intFila, 3).Value = Nz(.Fields("BUI"), "")
'                objHoja.Cells(intFila, 4).Value = Nz(.Fields("SUBSISTEMA"), "")
'                objHoja.Cells(intFila, 5).Value = Nz(.Fields("CRITICIDAD"), "")
'
'                If IsDate(.Fields("FechaRecepcionNotificacion")) Then
'                    objHoja.Cells(intFila, 6).Value = .Fields("FechaRecepcionNotificacion")
'                    objHoja.Cells(intFila, 6).NumberFormat = "dd/mm/yyyy"
'                End If
'
'                If IsDate(.Fields("FechaInicioContactoCliente")) Then
'                    objHoja.Cells(intFila, 7).Value = .Fields("FechaInicioContactoCliente")
'                    objHoja.Cells(intFila, 7).NumberFormat = "dd/mm/yyyy"
'                End If
'
'                If Not IsNull(.Fields("IncidenciaAveriaOReparacion")) Then
'                    objHoja.Cells(intFila, 8).Value = IIf(.Fields("IncidenciaAveriaOReparacion") = True, "Sí", "No")
'                End If
'
'                objHoja.Cells(intFila, 9).Value = Nz(.Fields("TipoReparacion"), "")
'
'                If Not IsNull(.Fields("Urgente")) Then
'                    objHoja.Cells(intFila, 10).Value = IIf(.Fields("Urgente") = True, "Sí", "No")
'                End If
'
'                If Not IsNull(.Fields("EventoConServicioAfectado")) Then
'                    objHoja.Cells(intFila, 11).Value = IIf(.Fields("EventoConServicioAfectado") = True, "Sí", "No")
'                End If
'
'                If IsDate(.Fields("FechaRestablecimientoServicio")) Then
'                    objHoja.Cells(intFila, 12).Value = .Fields("FechaRestablecimientoServicio")
'                    objHoja.Cells(intFila, 12).NumberFormat = "dd/mm/yyyy"
'                End If
'
'                '-------------------------------
'                ' Evaluación de cumplimiento SLU
'                '-------------------------------
'                m_CumpleTRES = False
'                If IsDate(.Fields("FechaRecepcionNotificacion")) And IsDate(.Fields("FechaInicioContactoCliente")) Then
'                    m_CumpleTRES = (CDate(.Fields("FechaInicioContactoCliente")) >= CDate(.Fields("FechaRecepcionNotificacion")))
'                End If
'
'                m_CumpleTRCM = True
'                m_Incidencia = .Fields("IncidenciaAveriaOReparacion")
'                If Not IsNull(m_Incidencia) Then
'                    If m_Incidencia = True Then
'                        m_CumpleTRCM = IsDate(.Fields("FechaInicioTiempoAdquisicion")) And _
'                                      IsDate(.Fields("FechaFinTiempoAdquisicion")) And _
'                                      Nz(.Fields("TipoReparacion"), "") <> ""
'                        If m_CumpleTRCM Then
'                            m_CumpleTRCM = (CDate(.Fields("FechaFinTiempoAdquisicion")) >= CDate(.Fields("FechaInicioTiempoAdquisicion")))
'                        End If
'                    End If
'                End If
'
'                m_CumpleTRSS = True
'                m_ServicioAfectado = .Fields("EventoConServicioAfectado")
'                If Not IsNull(m_ServicioAfectado) Then
'                    If m_ServicioAfectado = True Then
'                        m_CumpleTRSS = IsDate(.Fields("FechaRestablecimientoServicio"))
'                        If m_CumpleTRSS And IsDate(.Fields("FechaRecepcionNotificacion")) Then
'                            m_CumpleTRSS = (CDate(.Fields("FechaRestablecimientoServicio")) >= CDate(.Fields("FechaRecepcionNotificacion")))
'                        End If
'                    End If
'                End If
'
'                m_CumpleGlobal = (m_CumpleTRES And m_CumpleTRCM And m_CumpleTRSS)
'
'                objHoja.Cells(intFila, 13).Value = IIf(m_CumpleTRES, "Cumple", "No cumple")
'                objHoja.Cells(intFila, 14).Value = IIf(m_CumpleTRCM, "Cumple", "No cumple")
'                objHoja.Cells(intFila, 15).Value = IIf(m_CumpleTRSS, "Cumple", "No cumple")
'                objHoja.Cells(intFila, 16).Value = IIf(m_CumpleGlobal, "Cumple", "No cumple")
'
'                ' Coloreo visual rápido
'                objHoja.Cells(intFila, 13).Interior.Color = IIf(m_CumpleTRES, RGB(198, 239, 206), RGB(255, 199, 206))
'                objHoja.Cells(intFila, 14).Interior.Color = IIf(m_CumpleTRCM, RGB(198, 239, 206), RGB(255, 199, 206))
'                objHoja.Cells(intFila, 15).Interior.Color = IIf(m_CumpleTRSS, RGB(198, 239, 206), RGB(255, 199, 206))
'                objHoja.Cells(intFila, 16).Interior.Color = IIf(m_CumpleGlobal, RGB(198, 239, 206), RGB(255, 199, 206))
'
'                intFila = intFila + 1
'                .MoveNext
'            Loop
'        End If
'    End With
'
'    rcdEventos.Close
'    Set rcdEventos = Nothing
'
'    ' Ajustar columnas
'    objHoja.Columns("A:P").AutoFit
'    objHoja.Range("A1:P1").AutoFilter
'
'    ' Guardar y cerrar
'    objWorkbook.SaveAs strRutaArchivo
'    objWorkbook.Close False
'    objExcel.Quit
'
'    Set objHoja = Nothing
'    Set objWorkbook = Nothing
'    Set objExcel = Nothing
'
'    MsgBox "Exportación completada." & vbNewLine & "Archivo: " & strRutaArchivo, vbInformation, "Éxito"
'
'    DoCmd.Hourglass False
'    VBA.DoEvents
'    Exit Sub
'
'errores:
'    If Err.Number <> 1000 Then
'        strTextoError = "Error al exportar a Excel: " & Err.Number & vbNewLine & Err.Description
'    End If
'
'    ' Limpiar objetos Excel
'    On Error Resume Next
'    If Not objWorkbook Is Nothing Then objWorkbook.Close False
'    If Not objExcel Is Nothing Then objExcel.Quit
'    Set objHoja = Nothing
'    Set objWorkbook = Nothing
'    Set objExcel = Nothing
'
'    DoCmd.Hourglass False
'    VBA.DoEvents
'
'    If strTextoError <> "" Then
'        MsgBox strTextoError, vbCritical, "Error"
'    End If
'End Sub

'-----------------------------------------------------------
' cmdExportarHTML_Click
' Genera informe HTML con los eventos franqueados del período
' y lo abre en el navegador
'-----------------------------------------------------------
Private Sub cmdExportarHTML_Click()
    Dim strTextoError As String
    Dim strFechaInicio As String
    Dim strFechaFin As String
    Dim strResultado As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    ' Usar fechas guardadas o las actuales del formulario
    strFechaInicio = m_strFechaInicioFiltro
    strFechaFin = m_strFechaFinFiltro
    
    If strFechaInicio = "" Then
        strFechaInicio = Nz(Me.txtFechaInicioInforme.Value, "")
    End If
    If strFechaFin = "" Then
        strFechaFin = Nz(Me.txtFechaFInInforme.Value, "")
    End If
    
    If strFechaInicio = "" Or strFechaFin = "" Then
        strTextoError = "No hay datos filtrados para exportar. Ejecute primero la búsqueda."
        Err.Raise 1000
    End If
    
    ' Verificar que hay datos en la lista
    If Me.ListaFiltrados.ListCount = 0 Then
        strTextoError = "No hay datos en la lista para exportar."
        Err.Raise 1000
    End If
    
    ' Generar HTML via servicio (delegado a SLAHTMLService)
    strResultado = SLAHTMLService.GenerarInformeHTML_Corporate(strFechaInicio, strFechaFin, strTextoError)
    
    If Left(strResultado, 5) = "#ERR|" Then
        strTextoError = Mid(strResultado, 6)
        Err.Raise 1000
    End If
    
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al generar HTML: " & Err.Number & vbNewLine & Err.Description
    End If
    
    DoCmd.Hourglass False
    VBA.DoEvents
    
    If strTextoError <> "" Then
        MsgBox strTextoError, vbCritical, "Error"
    End If
End Sub

'-----------------------------------------------------------
' Form_Load
' Inicialización del formulario
'-----------------------------------------------------------
Private Sub Form_Load()
    On Error Resume Next
    VBA.DoEvents
    Me.txtFechaInicioInforme.Value = Null
    Me.txtFechaFInInforme.Value = Null
    Call InicializarListaFiltrados
    m_strIDEventoSeleccionado = ""
    m_strFechaInicioFiltro = ""
    m_strFechaFinFiltro = ""
End Sub

'-----------------------------------------------------------
' ListaFiltrados_Click
' Al hacer click en la lista, guarda el IDEvento seleccionado
'-----------------------------------------------------------
Private Sub ListaFiltrados_Click()
    Dim lngIndex As Long
    
    On Error Resume Next
    lngIndex = Me.ListaFiltrados.ListIndex
    If lngIndex = 0 Then
        ' La fila 0 es cabecera visual, no debe poder seleccionarse
        Me.ListaFiltrados = Null
        Me.ListaFiltrados.ListIndex = -1
        m_strIDEventoSeleccionado = ""
    ElseIf lngIndex > 0 Then
        m_strIDEventoSeleccionado = Nz(Me.ListaFiltrados.Column(0, lngIndex), "")
        Me.cmdVerEvento.Enabled = True
        Me.cmdDetalleSLA.Enabled = True
    Else
        m_strIDEventoSeleccionado = ""
        Me.cmdVerEvento.Enabled = False
        Me.cmdDetalleSLA.Enabled = False
    End If
End Sub

'-----------------------------------------------------------
' ExisteCampoEnRecordset
' Verifica si un campo existe en el recordset (case-insensitive)
' Previene error 3265 "Item not found in collection"
'-----------------------------------------------------------
Private Function ExisteCampoEnRecordset(p_rcd As DAO.Recordset, p_strNombreCampo As String) As Boolean
    Dim fld As DAO.Field
    On Error Resume Next
    ExisteCampoEnRecordset = False
    Set fld = p_rcd.Fields(p_strNombreCampo)
    ' Si no hay error y el campo no es Nothing, existe
    If Err.Number = 0 And Not fld Is Nothing Then
        ExisteCampoEnRecordset = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

'-----------------------------------------------------------
' ObtenerValorCampo
' Obtiene valor de campo de forma segura, retorna Empty si no existe
' Previene error 3265 en acceso a campos
'-----------------------------------------------------------
Private Function ObtenerValorCampo(p_rcd As DAO.Recordset, p_strNombreCampo As String) As Variant
    On Error Resume Next
    ObtenerValorCampo = Empty
    If Not p_rcd Is Nothing Then
        If ExisteCampoEnRecordset(p_rcd, p_strNombreCampo) Then
            ObtenerValorCampo = p_rcd.Fields(p_strNombreCampo)
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function
