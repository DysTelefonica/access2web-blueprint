Option Compare Database
Option Explicit
Dim strIDGasto As String, strURLAnexo As String, strFacturado As String
Dim strTipoGastoAlInicio As String, strImporteUnitarioAlInicio As String, strNumeroUnidadesAlInicio As String, _
    strFechaImputacionGastoAlInicio As String, strAliasAlInicio As String, strObservacionesAlInicio As String
Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  05/09/2014
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '
    '           strEstado="1"----> True Or False
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    Dim strTipoGasto As String, strImporteUnitario As String, strNumeroUnidades As String, strFechaImputacionGasto As String, strAlias As String, _
        strObservaciones As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strTipoGasto = Nz(Me.ComboTipoGasto, "")
    strImporteUnitario = Nz(Me.ImporteUnitario, "")
    strNumeroUnidades = Nz(Me.NUnidades, "")
    strFechaImputacionGasto = Nz(Me.FechaImputacion, "")
    strAlias = Nz(Me.ComboAlias, "")
    strObservaciones = Nz(Me.Observaciones, "")
    If IsDate(strFechaImputacionGasto) Then
        strFechaImputacionGasto = Format(strFechaImputacionGasto, "dd/mm/yyyy")
    End If
    If IsNumeric(strImporteUnitario) Then
        strImporteUnitario = CStr(CDbl(strImporteUnitario))
    End If
    If strTipoGasto <> strTipoGastoAlInicio Or strImporteUnitario <> strImporteUnitarioAlInicio Or strNumeroUnidades <> strNumeroUnidadesAlInicio Or _
        strFechaImputacionGasto <> strFechaImputacionGastoAlInicio Or strAlias <> strAliasAlInicio Or strObservaciones <> strObservacionesAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Function
errores:
    DoCmd.Hourglass False
    HaHabidoCambios = True
End Function
Private Sub cmdSalir_Click()
    Dim strTextoError As String, strNombreFormulario As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    If HaHabidoCambios() And strFacturado = "No" Then
        pregunta = MsgBox("¿Desea guardar los cambios producidos?", vbExclamation + vbYesNo + vbDefaultButton1, "Guardar Cambios")
        If pregunta = vbYes Then
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            '-------------------------------------------------------------------
            '   -Devuelve:
            '       Registrar = Descriptivo
            '       Registrar = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = Registrar()
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            '-------------------------------------------------------------------
            '       FormularioAbierto = true or false
            '       FormularioAbierto =False-->Error
            '--------------------------------------------
            If FormularioAbierto("FormGastoGestion") Then
                Forms("FormGastoGestion").Controls("ListaFiltrados").Requery
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
        End If
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Function Registrar() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strTipoGasto As String, strImporteUnitario As String, strNumeroUnidades As String, strFechaImputacionGasto As String, strAlias As String, _
        strObservaciones As String, strTextoError As String
    On Error GoTo errores
    If strIDGasto = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    strTipoGasto = Nz(Me.ComboTipoGasto, "#")
    strImporteUnitario = Nz(Me.ImporteUnitario, "#")
    strNumeroUnidades = Nz(Me.NUnidades, "#")
    strFechaImputacionGasto = Nz(Me.FechaImputacion, "#")
    strAlias = Nz(Me.ComboAlias, "#")
    strObservaciones = Nz(Me.Observaciones, "#")
    '-----------------------------------------------
    '       EditarGasto = Descriptivo
    '       EditarGasto = "#ERR" & "|" &strTextoError
    '-----------------------------------------------
    flag = EditarGasto(strIDGasto, strTipoGasto, strImporteUnitario, strNumeroUnidades, _
                        strFechaImputacionGasto, strAlias, strObservaciones)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EditarGasto ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComboTipoGasto_Change()
    Dim rcdDatos As DAO.Recordset, strImporteUnitarioDeTabla As String, strTipoGasto As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strTipoGasto = Nz(Me.ComboTipoGasto, "")
    m_SQL = "SELECT TbGastosImportePorTipo.ImporteUnitario " & _
            "FROM TbGastosImportePorTipo " & _
            "WHERE (((TbGastosImportePorTipo.TipoGasto)='" & strTipoGasto & "') AND ((TbGastosImportePorTipo.FechaFinalImporte) Is Null));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strImporteUnitarioDeTabla = Nz(.Fields("ImporteUnitario"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If IsNumeric(strImporteUnitarioDeTabla) Then
        Me.ImporteUnitario = strImporteUnitarioDeTabla
    Else
        Me.ImporteUnitario = Null
    End If
    CalculoTotal
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboTipoGasto_Exit(Cancel As Integer)
    Dim rcdDatos As DAO.Recordset, strImporteUnitarioDeTabla As String, strTipoGasto As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strTipoGasto = Nz(Me.ComboTipoGasto, "")
    m_SQL = "SELECT TbGastosImportePorTipo.ImporteUnitario " & _
            "FROM TbGastosImportePorTipo " & _
            "WHERE (((TbGastosImportePorTipo.TipoGasto)='" & strTipoGasto & "') AND ((TbGastosImportePorTipo.FechaFinalImporte) Is Null));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strImporteUnitarioDeTabla = Nz(.Fields("ImporteUnitario"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If IsNumeric(strImporteUnitarioDeTabla) Then
        Me.ImporteUnitario = strImporteUnitarioDeTabla
    Else
        Me.ImporteUnitario = Null
    End If
    CalculoTotal
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    If Not HaHabidoCambios() Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("No ha habido cambios que guardar", vbCritical, "Sin cambios")
        Exit Sub
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Registrar = strEstaFacturado
    '       Registrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '--------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------------
    If FormularioAbierto("FormGastoGestion") Then
        Forms("FormGastoGestion").Controls("ListaFiltrados").Requery
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar el gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strIDGasto & "|" & strURLAnexo & "|" & strFacturado
    strArgumento = Nz(Me.OpenArgs, "")
    If InStr(1, strArgumento, "|") = 0 Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 2 Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    strIDGasto = dato(0)
    strURLAnexo = dato(1)
    strFacturado = dato(2)
    If Not IsNumeric(strIDGasto) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = Descriptivo
    '       EstablecerDatos = "#ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
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
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
End Sub
Private Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, ctl As Control, strTextoError As String
    On Error GoTo errores
    If strIDGasto = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbGastos.* " & _
                "FROM TbGastos " & _
                "WHERE (((TbGastos.IDGasto)=" & strIDGasto & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strTipoGastoAlInicio = Nz(.Fields("TipoGasto"), "")
        strImporteUnitarioAlInicio = Nz(.Fields("ImporteUnitario"), "")
        strNumeroUnidadesAlInicio = Nz(.Fields("NumeroUnidades"), "")
        strFechaImputacionGastoAlInicio = Nz(.Fields("FechaImputacionGasto"), "")
        strAliasAlInicio = Nz(.Fields("ALIAS"), "")
        strObservacionesAlInicio = Nz(.Fields("Observaciones"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strTipoGastoAlInicio <> "" Then
        Me.ComboTipoGasto = strTipoGastoAlInicio
    End If
    If strImporteUnitarioAlInicio <> "" Then
        Me.ImporteUnitario = Format(strImporteUnitarioAlInicio, "#,##0.00 €")
    End If
    If strNumeroUnidadesAlInicio <> "" Then
        Me.NUnidades = strNumeroUnidadesAlInicio
    End If
    If IsDate(strFechaImputacionGastoAlInicio) Then
        Me.FechaImputacion = Format(strFechaImputacionGastoAlInicio, "dd/mm/yyyy")
    End If
    If strAliasAlInicio <> "" Then
        Me.ComboAlias = strAliasAlInicio
    End If
    If strObservacionesAlInicio <> "" Then
        Me.Observaciones = strObservacionesAlInicio
    End If
    If IsNumeric(strImporteUnitarioAlInicio) And IsNumeric(strNumeroUnidadesAlInicio) Then
        Me.ImporteTotal = Format(CDbl(strImporteUnitarioAlInicio) * CDbl(strNumeroUnidadesAlInicio), "#,##0.00 €")
    End If
    If strFacturado = "Sí" Then
        Me.ComandoRegistrar.Enabled = False
        Me.lblFacturado.Visible = True
    Else
        Me.ComandoRegistrar.Enabled = True
        Me.lblFacturado.Visible = False
    End If
    If strURLAnexo <> "" Then
        Me.ImagenAnexo.Visible = True
    Else
        Me.ImagenAnexo.Visible = False
    End If
    RellenarComboTecnicos Me.ComboAlias, EnumSino.Sí, , strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatos = "Datos del gasto establecidos correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstablecerDatos = "#ERR" & "|" & strTextoError
End Function
Private Sub ImagenAnexo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
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
Private Sub CalculoTotal()
    Dim dblImporteUnitario As Double, dblNUnidades As Double, dblImporteTotal As Double, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If IsNumeric(Nz(Me.ImporteUnitario, "")) And IsNumeric(Nz(Me.NUnidades, "")) Then
        dblImporteUnitario = CDbl(Me.ImporteUnitario)
        dblNUnidades = CDbl(Me.NUnidades)
        dblImporteTotal = dblImporteUnitario * dblNUnidades
    End If
    If dblImporteTotal = 0 Then
        Me.ImporteTotal = Null
    Else
        Me.ImporteTotal = Format(dblImporteTotal, "#,##0.00 €")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar calcular el total se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ImporteUnitario_Exit(Cancel As Integer)
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    CalculoTotal
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub NUnidades_Exit(Cancel As Integer)
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    CalculoTotal
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Public Function EditarGasto( _
                                strIDGasto As String, _
                                Optional strTipoGasto As String, _
                                Optional strImporteUnitario As String, _
                                Optional strNumeroUnidades As String, _
                                Optional strFechaImputacionGasto As String, _
                                Optional strAlias As String, _
                                Optional strObservaciones As String _
                                ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 05/09/2014
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -1) Ver si se ha facturado ya si lo está no se puede cambiar strImporteUnitario,strNumeroUnidades,strFechaImputacionGasto
    '   -Llamada desde
    
    '   -Devuelve:
    
    '       EditarGasto = Descriptivo
    '       EditarGasto = "#ERR" & "|" &strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strID As String, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strIDGasto) Then
        strTextoError = "No se ha introducido un IDGasto correcto"
        Err.Raise 1000
    End If
    If strTipoGasto = "" Or strTipoGasto = "#" Then
        strTextoError = "Se ha de introducir el tipo de gasto"
        Err.Raise 1000
    End If
    If Not IsNumeric(strImporteUnitario) Then
        strTextoError = "Se ha de introducir el importe unitario"
        Err.Raise 1000
    End If
    If Not IsNumeric(strNumeroUnidades) Then
        strTextoError = "Se ha de introducir el número de unidades"
        Err.Raise 1000
    End If
     If Not IsDate(strFechaImputacionGasto) Then
        strTextoError = "Se ha de introducir la fecha de imputación del gasto"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbGastos.* " & _
            "FROM TbGastos " & _
            "WHERE (((TbGastos.IDGasto)=" & strIDGasto & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El Gasto: " & strIDGasto & " no estaba registrado en la base de datos"
            Err.Raise 1000
        End If
        .Edit
            .Fields("TipoGasto") = strTipoGasto
            .Fields("ImporteUnitario") = CDbl(strImporteUnitario)
            .Fields("NumeroUnidades") = CDbl(strNumeroUnidades)
            .Fields("FechaImputacionGasto") = strFechaImputacionGasto
            If strAlias <> "" And strAlias <> "#" Then
                .Fields("ALIAS") = strAlias
            Else
                If strAlias = "#" Then
                    .Fields("ALIAS") = Null
                End If
            End If
            If strObservaciones <> "" And strObservaciones <> "#" Then
                .Fields("Observaciones") = strObservaciones
            Else
                 If strObservaciones = "#" Then
                    .Fields("Observaciones") = Null
                End If
            End If
        .Update
    End With
    EditarGasto = "Editada correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Gasto.EditarGasto ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EditarGasto = "#ERR" & "|" & strTextoError
End Function



