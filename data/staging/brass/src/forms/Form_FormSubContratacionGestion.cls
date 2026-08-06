Option Compare Database
Option Explicit
Dim strIDSubContratacionSeleccionada As String, strURLAnexo As String
Dim strEmpresaBusquedaAnterior As String, strFechaFinalBusquedaAnterior As String, strIDBusquedaAnterior As String, strFormularioQueAbre As String
Private Sub ComandoActualizarTodo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    ComandoLimpiarEmpresa_Click
    ComandoLimpiarFechas_Click
    ComandoLimpiarID_Click
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar actualizar las listas se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Sub LimpiarUltimos()
    strEmpresaBusquedaAnterior = ""
    strFechaFinalBusquedaAnterior = ""
    strIDBusquedaAnterior = ""
End Sub
Public Sub ActualizarLista()
    Dim intElementos As Integer, lbl As Label, lst As ListBox, VarItem As Variant, strIDSubContratacionSeleccionadaActualmente As String, _
        blnEncontrada As Boolean, i As Integer, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDSubContratacionSeleccionada = ""
    Me.ComandoEditarSubContratacion.Enabled = False
    Me.ComandoEliminarSubContratacion.Enabled = False
    Me.ComandoExportarExcelSubContrataciones.Enabled = False
    Me.ComandoAnexar.Enabled = False
    Me.ImagenAnexo.Visible = False
    strIDSubContratacionSeleccionadaActualmente = Nz(Me.ListaFiltrados.Column(0), "")
    '------------------------------------------------------
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '------------------------------------------------------
    flag = Filtrar(Nz(Me.ComboEmpresaBusqueda, ""), Nz(Me.FechaInicialBusqueda, ""), _
                    Nz(Me.FechaFinalBusqueda, ""), Nz(Me.ComboIDSubContratacion, ""))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Filtrar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_SQL = flag
    Set lst = Me.ListaFiltrados
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    
    If intElementos = 0 Then
        Me.ComandoExportarExcelSubContrataciones.Enabled = False
    ElseIf intElementos = 1 Then
        strIDSubContratacionSeleccionadaActualmente = lst.Column(0, 1)
        lst.Selected(1) = True
        ListaFiltrados_Click
        Me.ComandoExportarExcelSubContrataciones.Enabled = True
    Else
        Me.ComandoExportarExcelSubContrataciones.Enabled = True
        For i = 1 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                strIDSubContratacionSeleccionadaActualmente = lst.Column(0, i)
                ListaFiltrados_Click
                Exit For
            End If
        Next
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar actualizar la lista  se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAltaSubContratacion_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '------------------------------------------
    If FormularioAbierto("FormSubContratacionAlta") Then
        DoCmd.Close acForm, "FormSubContratacionAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormSubContratacionAlta"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta una subcontratación se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAnexar_Click()
    Dim m_SubContratacion As Subcontratacion
    Dim strEscritura As String, strEstaFacturado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strEstaFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    strIDSubContratacionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDSubContratacionSeleccionada = "" Then
        strTextoError = "Se ha de seleccionar la subcontratación primero"
        Err.Raise 1000
    End If
    If strEstaFacturado <> "" Then
        strEstaFacturado = "Sí"
        strEscritura = "No"
    Else
        strEstaFacturado = "No"
        strEscritura = "Sí"
    End If
    Set m_SubContratacion = Constructor.getSubContratacion(strIDSubContratacionSeleccionada, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    strFormularioQueAbre = ""
    Set m_ObjEntidadParaAnexoActiva = m_SubContratacion
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Filtrar( _
                            Optional strEmpresa As String, _
                            Optional strFechaInicial As String, _
                            Optional strFechaFinal As String, _
                            Optional strIDSubContratacion As String _
                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 30/09/2019
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    Dim strParteWhere As String, strSQLParteInicial As String, strFechaInicialParaConsulta As String, strFechaFinalParaConsulta As String, _
        strParteWhereIDSubContratacion As String, strParteWhereEmpresas As String, strParteWhereFechas As String, strSQLFinal As String, _
        strTextoError As String
    On Error GoTo errores
    strSQLParteInicial = "SELECT TbSubcontrataciones.IDSubcontratacion AS ID, TbSubcontrataciones.Empresa, TbSubcontrataciones.FechaAlta,TbSubcontrataciones.FechaFin, " & _
                            "TbSubcontrataciones.Importe , TbFacturaSubcontratacionesInvolucradas.IDFactura " & _
                            "FROM TbSubcontrataciones LEFT JOIN TbFacturaSubcontratacionesInvolucradas ON " & _
                            "TbSubcontrataciones.IDSubcontratacion = TbFacturaSubcontratacionesInvolucradas.IDSubcontratacion "
    If strEmpresa = "" And strFechaInicial = "" And strFechaFinal = "" And strIDSubContratacion = "" Then
        strParteWhereIDSubContratacion = "((TbSubcontrataciones.IDSubcontratacion) = -1)"
        strParteWhere = "WHERE (" & strParteWhereIDSubContratacion & ") ORDER BY  TbSubcontrataciones.FechaAlta DESC;"
        strSQLFinal = strSQLParteInicial & strParteWhere
        Filtrar = strSQLFinal
        Exit Function
    End If
    If Not IsDate(strFechaInicial) Then
        strFechaInicial = "01/01/1900"
    End If
    If Not IsDate(strFechaFinal) Then
        strFechaFinal = "01/01/2100"
    End If
    strFechaInicialParaConsulta = Format(strFechaInicial, "mm/dd/yyyy")
    strFechaFinalParaConsulta = Format(strFechaFinal, "mm/dd/yyyy")
    
    If strIDSubContratacion = "" Then
        strParteWhereIDSubContratacion = "((TbSubcontrataciones.IDSubcontratacion) Like '*' Or (TbSubcontrataciones.IDSubcontratacion) Is Null)"
    Else
        strParteWhereIDSubContratacion = "((TbSubcontrataciones.IDSubcontratacion) = " & strIDSubContratacion & ")"
    End If
    If strEmpresa = "" Then
        strParteWhereEmpresas = "((TbSubcontrataciones.Empresa) Like '*' Or (TbSubcontrataciones.Empresa) Is Null)"
    Else
        strParteWhereEmpresas = "((TbSubcontrataciones.Empresa) = '" & strEmpresa & "')"
    End If
    strParteWhereFechas = "((TbSubcontrataciones.FechaAlta) Between #" & strFechaInicialParaConsulta & "# And #" & strFechaFinalParaConsulta & "#)"
    strParteWhere = "WHERE (" & _
                            strParteWhereIDSubContratacion & _
                            " AND " & strParteWhereEmpresas & _
                            " AND " & strParteWhereFechas & _
                            ") ORDER BY  TbSubcontrataciones.FechaAlta DESC;"
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
Private Sub ComandoEditarSubContratacion_Click()
    Dim strEscritura As String, strEstaFacturado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDSubContratacionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    strEstaFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    If Not IsNumeric(strIDSubContratacionSeleccionada) Then
        strTextoError = "Seleccione la Subcontratacion que pretende editar"
        Err.Raise 1000
    End If
    If strEstaFacturado <> "" Then
        strEstaFacturado = "Sí"
        strEscritura = "No"
    Else
        strEstaFacturado = "No"
        strEscritura = "Sí"
    End If
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '------------------------------------------
    If FormularioAbierto("FormSubContratacionEdicion") Then
        DoCmd.Close acForm, "FormSubContratacionEdicion", acSaveNo
    End If
    'strArgumento=strIDActividad & "|" & strEstaFacturado & "|" & strEscritura
    DoCmd.OpenForm "FormSubContratacionEdicion", , , , , , strIDSubContratacionSeleccionada & "|" & strEstaFacturado & "|" & strEscritura
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar la Subcontratacion se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminarSubContratacion_Click()
    Dim strEstaFacturado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDSubContratacionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDSubContratacionSeleccionada = "" Then
        strTextoError = "Seleccione la Subcontratacion que pretende eliminar"
        Err.Raise 1000
    End If
    strEstaFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    If strEstaFacturado <> "" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strEstaFacturado = "Sí" Then
        strTextoError = "No se puede eliminar un elemento que está facturado"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '
    '           strEstado="1"---->true or false
    '           strEstado="-1"--->false
    '-------------------------------------------------------------------
    pregunta = MsgBox("¿Desea eliminar la Subcontratacion con ID: " & strIDSubContratacionSeleccionada & _
                " junto con los anexos implicados?", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Subcontratacion")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Eliminar(strIDSubContratacionSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SubContratacion.Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDSubContratacionSeleccionada = ""
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar la Subcontratacion se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Eliminar( _
                            strIDSubContratacion As String, _
                            Optional strEstaFacturada As String _
                            ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/06/13
    '   -Modificaciones:

    '   -Funcionamiento:
    '       -va a mirar que no esté ya facturada
    '       -eliminará el registro y además su anexo así como de las lista del no para R
    '   -Llamada desde
    '       -Form_FormSubContratacionGestion.ComandoEliminarSubContratacion_Click
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim m_SubContratacion As Subcontratacion
    Dim rcdDatos As DAO.Recordset, fso As New FileSystemObject
    Dim VarItem As Variant, strURLAnexo As String, strCadenaAnexos As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDSubContratacion) Then
        strTextoError = "No se ha introducido el IDSubcontratación que se desea eliminar"
        Err.Raise 1000
    End If
    If strEstaFacturada <> "Sí" And strEstaFacturada <> "No" Then
        '------------------------------------------------------
        '   -Devuelve:
        '       EstaFacturado = strEstaFacturado
        '       EstaFacturado = "#ERR" & "|" & strTextoError
        '------------------------------------------------------
        flag = EstaFacturado("SUB", strIDSubContratacion)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFacturado ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEstaFacturada = flag
    End If
    If strEstaFacturada = "Sí" Then
        strTextoError = "No se puede eliminar una subcontratación que forme parte de alguna de las facturas"
        Err.Raise 1000
    End If
    Set m_SubContratacion = Constructor.getSubContratacion(strIDSubContratacion, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    strCadenaAnexos = m_SubContratacion.URLAnexos
    If strCadenaAnexos <> "" Then
        If InStr(1, strCadenaAnexos, ";") <> 0 Then
            dato = Split(strCadenaAnexos, ";")
            For Each VarItem In dato
                strURLAnexo = CStr(VarItem)
                '----------------------------------------------
                '   -Devuelve:
                '       FicheroAbierto = true or false
                '       FicheroAbierto =False-->Error
                '----------------------------------------------
                If FicheroAbierto(strURLAnexo) Then
                    strTextoError = "El anexo llamado " & fso.GetFileName(strURLAnexo) & " asociado está abierto"
                    Err.Raise 1000
                End If
            Next
        Else
            If fso.FileExists(strCadenaAnexos) Then
                '----------------------------------------------
                '   -Devuelve:
                '       FicheroAbierto = true or false
                '       FicheroAbierto =False-->Error
                '----------------------------------------------
                If FicheroAbierto(strCadenaAnexos) Then
                    strTextoError = "El anexo llamado " & fso.GetFileName(strCadenaAnexos) & " asociado está abierto"
                    Err.Raise 1000
                End If
            End If
        End If
    End If
    m_SQL = "SELECT TbSubcontrataciones.* " & _
                "FROM TbSubcontrataciones " & _
                "WHERE (((TbSubcontrataciones.IDSubcontratacion)=" & strIDSubContratacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No está registrada la subcontratación que pretende eliminar"
            Err.Raise 1000
        End If
       .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SubContratacion.EliminarAnexos strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    
    Eliminar = "Subcontratación eliminada correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método SubContratacion.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Eliminar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoExportarExcelSubContrataciones_Click()
    Dim strSQLTotal As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strSQLTotal = Me.ListaFiltrados.RowSource
    '------------------------------------------
    '   -Devuelve:
    '       GenerarConsultas = Descriptivo
    '       GenerarConsultas = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = GenerarConsultas("FiltroDeSubContrataciones", Me.hWnd, strSQLTotal, , "No")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método GenerarConsultas ha devuelto el error: " & vbNewLine & dato(1)
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
Private Sub ComandoLimpiarEmpresa_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboEmpresaBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEmpresaBusqueda = Null
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
Private Sub ComandoLimpiarID_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboIDSubContratacion, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboIDSubContratacion = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComboEmpresaBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEmpresaBusqueda, "")
    If strValor <> strEmpresaBusquedaAnterior Then
        ActualizarLista
    End If
    strEmpresaBusquedaAnterior = strValor
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
Private Sub ComboEmpresaBusqueda_Enter()
    On Error Resume Next
    Me.ComboEmpresaBusqueda.Requery
End Sub
Private Sub ComboEmpresaBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEmpresaBusqueda, "")
    If strValor <> strEmpresaBusquedaAnterior Then
        ActualizarLista
    End If
    strEmpresaBusquedaAnterior = strValor
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
Public Sub ComboIDSubContratacion_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboIDSubContratacion, "")
    If strValor <> strIDBusquedaAnterior Then
        ActualizarLista
    End If
    strIDBusquedaAnterior = strValor
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
Private Sub ComboIDSubContratacion_Enter()
    On Error Resume Next
    Me.ComboIDSubContratacion.Requery
End Sub
Private Sub ComboIDSubContratacion_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboIDSubContratacion, "")
    If strValor <> strIDBusquedaAnterior Then
        ActualizarLista
    End If
    strIDBusquedaAnterior = strValor
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
    If strValor <> strFechaFinalBusquedaAnterior Then
        ActualizarLista
    End If
    strFechaFinalBusquedaAnterior = strValor
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
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
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
Private Sub ListaFiltrados_Click()
    Dim m_SubContratacion As Subcontratacion
    Dim fso As New FileSystemObject, strEstaFacturado As String, strNombreArchivo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComandoAltaSubContratacion.Enabled = True
    Me.ComandoEditarSubContratacion.Enabled = False
    Me.ComandoEliminarSubContratacion.Enabled = False
    Me.ComandoAnexar.Enabled = False
    strIDSubContratacionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    strEstaFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    If strIDSubContratacionSeleccionada = "" Then
        If Not fso Is Nothing Then
            Set fso = Nothing
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComandoEditarSubContratacion.Enabled = True
    Me.ComandoEliminarSubContratacion.Enabled = True
    Me.ComandoAnexar.Enabled = True
    Set m_SubContratacion = Constructor.getSubContratacion(strIDSubContratacionSeleccionada, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    strURLAnexo = m_SubContratacion.URLAnexos
    If strURLAnexo = "" Then
        Me.ImagenAnexo.Visible = False
    Else
        Me.ImagenAnexo.Visible = True
    End If
    If strEstaFacturado <> "" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strEstaFacturado = "Sí" Then
        Me.ComandoEliminarSubContratacion.Enabled = False
        
    Else
        Me.ComandoEliminarSubContratacion.Enabled = True
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





