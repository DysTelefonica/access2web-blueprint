Option Compare Database
Option Explicit
Dim strIDGastoSeleccionado As String, strURLAnexo As String
Dim strTecnicoAnterior As String, strFechaInicialAnterior As String, strFechaFinalAnterior As String, strIDGastoAnterior As String
Public Sub LimpiarUltimos()
    strTecnicoAnterior = ""
    strFechaInicialAnterior = ""
    strFechaFinalAnterior = ""
    strIDGastoAnterior = ""
End Sub
Public Sub ComandoActualizarTodo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    ComandoLimpiarTecnico_Click
    ComandoLimpiarFechas_Click
    ComandoLimpiarGasto_Click
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ActualizarLista la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAlta_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '           strEstado="1"---->true or false
    '           strEstado="-1"--->false
    '-------------------------------------------------------------------
    If FormularioAbierto("FormGastoAlta") Then
        DoCmd.Close acForm, "FormGastoAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormGastoAlta"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un gasto se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAnexar_Click()
    Dim m_Gasto As Gasto
    Dim m_FormularioQueAbre As String
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDGastoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDGastoSeleccionado = "" Then
        strTextoError = "Se ha de seleccionar el gasto"
        Err.Raise 1000
    End If
    Set m_Gasto = Constructor.getGasto(strIDGastoSeleccionado, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    Set m_ObjEntidadParaAnexoActiva = m_Gasto
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , m_FormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un anexo se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoEditar_Click()
    Dim strTextoError As String, strFacturado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDGastoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    strFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    If Not IsNumeric(strIDGastoSeleccionado) Then
        strTextoError = "Seleccione el gasto que pretende editar"
        Err.Raise 1000
    End If
    If strFacturado = "" Then
        strFacturado = "No"
    Else
        strFacturado = "Sí"
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '           strEstado="1"---->true or false
    '           strEstado="-1"--->false
    '-------------------------------------------------------------------
    If FormularioAbierto("FormGastoEdicion") Then
        DoCmd.Close acForm, "FormGastoEdicion", acSaveNo
    End If
    'strArgumento=strIDGasto & "|" & strURLAnexo & "|" & strFacturado
    DoCmd.OpenForm "FormGastoEdicion", , , , , , strIDGastoSeleccionado & "|" & strURLAnexo & "|" & strFacturado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar el gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDGastoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDGastoSeleccionado = "" Then
        If strIDGastoSeleccionado = "" Then
            strTextoError = "Seleccione el gasto que pretende eliminar"
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If strURLAnexo = "" Then
        pregunta = MsgBox("¿Desea eliminar el gasto : " & Nz(Me.ListaFiltrados.Column(1), "") & _
                    " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Gasto")
    Else
        pregunta = MsgBox("¿Desea eliminar el gasto : " & Nz(Me.ListaFiltrados.Column(1), "") & _
                    " junto con el posible anexos implicado?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Gasto")
    End If
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    flag = Eliminar(strIDGastoSeleccionado, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    Me.ComandoAlta.SetFocus
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar el gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoExportarExcel_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_SQL = Me.ListaFiltrados.RowSource
    GenerarConsultasSQL m_SQL, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoLimpiarGasto_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboGastoBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboGastoBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComboGastoBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboGastoBusqueda, "")
    If strValor <> strIDGastoAnterior Then
        ActualizarLista
    End If
    strIDGastoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboTecnicoBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboTecnicoBusqueda, "")
    If strValor <> strTecnicoAnterior Then
        ActualizarLista
    End If
    strTecnicoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub FechaFinalBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.FechaFinalBusqueda, "")
    If strValor <> strFechaFinalAnterior Then
        ActualizarLista
    End If
    strFechaFinalAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub FechaInicialBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.FechaInicialBusqueda, "")
    If strValor <> strFechaInicialAnterior Then
        ActualizarLista
    End If
    strFechaInicialAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboGastoBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboGastoBusqueda, "")
    If strValor <> strIDGastoAnterior Then
        ActualizarLista
    End If
    strIDGastoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboTecnicoBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboTecnicoBusqueda, "")
    If strValor <> strTecnicoAnterior Then
        ActualizarLista
    End If
    strTecnicoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Canel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    Me.SetFocus
    VBA.DoEvents
    Me.Caption = Me.lblTitulo1.Caption
    ActualizarLista
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ImagenAnexo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    
    If strURLAnexo = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If InStr(1, strURLAnexo, ";") <> 0 Then
        ComandoAnexar_Click
        Exit Sub
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AbrirEnLocal = strURLAnexoLocal
    '       AbrirEnLocal = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = AbrirEnLocal(strURLAnexo, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Sub ListaFiltrados_Click()
    Dim m_Gasto As Gasto
    Dim fso As New FileSystemObject, strNombreArchivo As String, strFacturado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDGastoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDGastoSeleccionado = "" Then
        Me.ComandoEditar.Enabled = False
        Me.ComandoEliminar.Enabled = False
        Me.ComandoAnexar.Enabled = False
        Me.ImagenAnexo.Visible = False
        Me.lblFacturado.Visible = False
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoEditar.Enabled = True
    Me.ComandoAnexar.Enabled = True
    strFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    If strFacturado = "" Then
        strFacturado = "No"
        Me.ComandoEliminar.Enabled = True
        Me.lblFacturado.Visible = False
    Else
        strFacturado = "Sí"
        Me.ComandoEliminar.Enabled = False
        Me.lblFacturado.Visible = True
    End If
    Set m_Gasto = Constructor.getGasto(strIDGastoSeleccionado, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    strURLAnexo = m_Gasto.URLAnexos
    If strURLAnexo <> "" Then
        
        Me.ImagenAnexo.Visible = True
    Else
        Me.ImagenAnexo.Visible = False
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar hacer click sobre la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Filtrar( _
                            Optional strAlias As String, _
                            Optional strGasto As String, _
                            Optional strFechaInicial As String, _
                            Optional strFechaFinal As String _
                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    '       -me.ActualizarLista
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    Dim strParteWhere As String, strSQLParteInicial As String
    Dim strFechaInicialParaConsulta As String, strFechaFinalParaConsulta As String, strSQLFinal As String
    Dim strParteWhereGasto As String, strParteWhereAliasTecnico As String, strParteWhereFechas As String, strTextoError As String
    On Error GoTo errores
    strSQLParteInicial = "SELECT TbGastos.IDGasto AS ID, TbGastos.TipoGasto as Gasto, " & _
                    "IIf(IsNumeric([TbGastos].[ImporteUnitario]) And IsNumeric([TbGastos].[NumeroUnidades]),Format(CDbl([TbGastos].[ImporteUnitario])*CDbl([TbGastos].[NumeroUnidades]),'#,##0.00 €'),Null) AS [Importe Total], " & _
                    "TbGastos.FechaImputacionGasto AS Fecha, TbTecnicos.NOMBRE AS Técnico, TbFacturaGastosInvolucrados.IDFactura , TbAnexos.NombreArchivo " & _
                    "FROM ((TbGastos LEFT JOIN TbFacturaGastosInvolucrados ON " & _
                    "TbGastos.IDGasto = TbFacturaGastosInvolucrados.IDGasto) " & _
                    "LEFT JOIN TbTecnicos ON TbGastos.ALIAS = TbTecnicos.ALIAS) " & _
                    "LEFT JOIN TbAnexos ON TbGastos.IDGasto = TbAnexos.IDGasto "
    If strAlias = "" And strGasto = "" And strFechaInicial = "" And strFechaFinal = "" Then
        strParteWhere = "WHERE (((TbGastos.IDGasto)=-1));"
        strSQLFinal = strSQLParteInicial & strParteWhere
        Filtrar = strSQLFinal
        Exit Function
    End If
    If Not IsDate(strFechaInicial) Then
        strFechaInicialParaConsulta = "01/01/1900"
    Else
        strFechaInicialParaConsulta = Format(strFechaInicial, "mm/dd/yyyy")
    End If
    If Not IsDate(strFechaFinal) Then
        strFechaFinalParaConsulta = "01/01/2100"
    Else
        strFechaFinalParaConsulta = Format(strFechaFinal, "mm/dd/yyyy")
    End If
    If strGasto = "" Then
        strParteWhereGasto = "((TbGastos.TipoGasto) Like '*' Or (TbGastos.TipoGasto) Is Null)"
    Else
        strParteWhereGasto = "((TbGastos.TipoGasto) = '" & strGasto & "')"
    End If
    If strAlias = "" Then
        strParteWhereAliasTecnico = "((TbGastos.ALIAS) Like '*' Or (TbGastos.ALIAS) Is Null)"
    Else
        strParteWhereAliasTecnico = "((TbGastos.ALIAS) = '" & strAlias & "')"
    End If
    strParteWhereFechas = "((TbGastos.FechaImputacionGasto) Between #" & strFechaInicialParaConsulta & "# And #" & strFechaFinalParaConsulta & "#)"
    strParteWhere = "WHERE (" & _
                            strParteWhereGasto & _
                            " AND " & strParteWhereAliasTecnico & _
                            " AND " & strParteWhereFechas & _
                            ") ORDER BY  TbGastos.FechaImputacionGasto DESC;"
    strSQLFinal = strSQLParteInicial & strParteWhere
    Filtrar = strSQLFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Filtrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Filtrar = "#ERR" & "|" & strTextoError
End Function
Public Sub ActualizarLista()
    Dim intElementos As Integer, lst As ListBox, VarItem As Variant
    Dim strIDGastoSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDGastoSeleccionado = ""
    Set lst = Me.ListaFiltrados
    strIDGastoSeleccionadoActualmente = Nz(lst.Column(0), "")
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoAnexar.Enabled = False
    Me.ImagenAnexo.Visible = False
    Me.lblFacturado.Visible = False
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    flag = Filtrar( _
                    Nz(Me.ComboTecnicoBusqueda, ""), Nz(Me.ComboGastoBusqueda, ""), _
                    Nz(Me.FechaInicialBusqueda, ""), Nz(Me.FechaFinalBusqueda, ""))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Filtrar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_SQL = flag
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 0 Then
        Me.ComandoExportarExcel.Enabled = False
    Else
        If intElementos = 1 Then
            strIDGastoSeleccionadoActualmente = lst.Column(0, 1)
            lst.Selected(1) = True
            ListaFiltrados_Click
        Else
            For i = 1 To lst.ListCount - 1
                If lst.Selected(i) = True Then
                    strIDGastoSeleccionadoActualmente = lst.Column(0, i)
                    ListaFiltrados_Click
                    Exit For
                End If
            Next
        End If
        Me.ComandoExportarExcel.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ActualizarLista la lista de Actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoLimpiarTecnico_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboTecnicoBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboTecnicoBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarFechas_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.FechaInicialBusqueda, "") = "" And Nz(Me.FechaFinalBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.FechaInicialBusqueda = Null
    Me.FechaFinalBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Public Function Eliminar( _
                            p_IDGasto As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    
    Dim m_Gasto As Gasto
    On Error GoTo errores
    
    If Not IsNumeric(p_IDGasto) Then
        p_Error = "No se ha introducido el gasto que se desea eliminar"
        Err.Raise 1000
    End If
    Set m_Gasto = Constructor.getGasto(p_IDGasto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Gasto Is Nothing Then
        p_Error = "Gasto no registrado"
        Err.Raise 1000
    End If
    m_Gasto.Eliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Eliminar = "Gasto eliminado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

