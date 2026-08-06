Option Compare Database
Option Explicit
Dim strIDSubContratacion As String, strURLAnexo As String, strEscritura As String, strEstaFacturado As String
Dim strEmpresaAlInicio As String, strFechaInicialAlInicio As String, strDescripcionAlInicio As String, strFechaFinAlInicio As String, _
    strImporteAlInicio As String, strFormularioQueAbre As String
Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  13/06/2013
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
    Dim strEmpresa As String, strFechaInicial As String, strDESCRIPCION As String, strFechaFin As String, strImporte As String
    On Error GoTo errores
    strEmpresa = Nz(Me.ComboEmpresa, "")
    strFechaInicial = Nz(Me.FechaAlta, "")
    strDESCRIPCION = Nz(Me.Descripcion, "")
    strFechaFin = Nz(Me.FechaFin, "")
    strImporte = Nz(Me.Importe, "")
    If IsDate(strFechaInicial) Then
        strFechaInicial = Format(strFechaInicial, "dd/mm/yyyy")
    End If
    If IsDate(strFechaFin) Then
        strFechaFin = Format(strFechaFin, "dd/mm/yyyy")
    End If
    If IsDate(strImporte) Then
        strImporte = CStr(CDbl(strImporte))
    End If
    If strEmpresa <> strEmpresaAlInicio Or strFechaInicial <> strFechaInicialAlInicio Or strDESCRIPCION <> strDescripcionAlInicio Or _
        strFechaFin <> strFechaFinAlInicio Or strImporte <> strImporteAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Private Sub cmdSalir_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    If HaHabidoCambios() And strEscritura = "Sí" Then
        pregunta = MsgBox("¿Desea guardar los cambios producidos?", vbExclamation + vbYesNo + vbDefaultButton1, "Guardar Cambios")
        If pregunta = vbYes Then
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            '-------------------------------------------------------------------
            '   -Devuelve:
            '       Registrar = strEstado & "|" & strTextoOk & "|" & strTextoError
            '           strEstado="1"---->strTextoOk=Descriptivo
            '           strEstado="-1"--->strTextoError
            '-------------------------------------------------------------------
            flag = Registrar()
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Registrar ha devuelto un error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
        End If
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAdjuntar_Click()
    Dim m_SubContratacion As Subcontratacion
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strIDSubContratacion = "" Then
        strTextoError = "No se ha podido determinar La IDSubcontratación"
        Err.Raise 1000
    End If
    If HaHabidoCambios() Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       Registrar = strEstado & "|" & strTextoOk & "|" & strTextoError
        '           strEstado="1"---->strTextoOk=Descriptivo
        '           strEstado="-1"--->strTextoError
        '-------------------------------------------------------------------
        flag = Registrar()
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Registrar ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    Set m_SubContratacion = Constructor.getSubContratacion(strIDSubContratacion, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    strFormularioQueAbre = Me.Name
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
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Function Registrar() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '   -llamada por:
    
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strEmpresa As String, strFechaInicial As String, strDESCRIPCION As String, strFechaFin As String, strImporte As String, strTextoError As String
    On Error GoTo errores
    If strIDSubContratacion = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    strEmpresa = Nz(Me.ComboEmpresa, "")
    strFechaInicial = Nz(Me.FechaAlta, "")
    strDESCRIPCION = Nz(Me.Descripcion, "#")
    strFechaFin = Nz(Me.FechaFin, "#")
    strImporte = Nz(Me.Importe, "")
    '-----------------------------------------------
    '   -Devuelve:
    '       editar = Descriptivo
    '       editar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = editar(strIDSubContratacion, strEmpresa, strFechaInicial, strDESCRIPCION, strImporte, strFechaFin)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SubContratacion.editar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not HaHabidoCambios() Then
        strTextoError = "No hay cambios que registrar"
        Err.Raise 1000
    End If
    '-----------------------------------------------------
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Edición correcta", vbInformation, "Edición")
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strIDSubContratacion & "|" & strEstaFacturado & "|" & strEscritura
    strArgumento = Nz(Me.OpenArgs, "")
    If strArgumento = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 2 Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    strIDSubContratacion = Nz(dato(0), "")
    strEstaFacturado = Nz(dato(1), "")
    strEscritura = Nz(dato(2), "")
    If strIDSubContratacion = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        strEstaFacturado = "Sí"
    End If
    If strEscritura <> "Sí" And strEscritura <> "No" Then
        strEscritura = "No"
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerDatos ha devuelto un error: " & vbNewLine & dato(1)
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos() As String
    Dim m_SubContratacion As Subcontratacion
    Dim rcdDatos As DAO.Recordset, ctl As Control, strTextoError As String
    On Error GoTo errores
    If strIDSubContratacion = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbSubcontrataciones.* " & _
                "FROM TbSubcontrataciones " & _
                "WHERE (((TbSubcontrataciones.IDSubcontratacion)=" & strIDSubContratacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strEmpresaAlInicio = Nz(.Fields("Empresa"), "")
        strFechaInicialAlInicio = Nz(.Fields("FechaAlta"), "")
        strDescripcionAlInicio = Nz(.Fields("Descripcion"), "")
        strFechaFinAlInicio = Nz(.Fields("FechaFin"), "")
        strImporteAlInicio = Nz(.Fields("Importe"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEmpresaAlInicio <> "" Then
        Me.ComboEmpresa = strEmpresaAlInicio
    End If
    If strFechaInicialAlInicio <> "" Then
        If IsDate(strFechaInicialAlInicio) Then
            strFechaInicialAlInicio = Format(strFechaInicialAlInicio, "dd/mm/yyyy")
        End If
        Me.FechaAlta = strFechaInicialAlInicio
    End If
    If strDescripcionAlInicio <> "" Then
        Me.Descripcion = strDescripcionAlInicio
    End If
    If IsDate(strFechaFinAlInicio) Then
        If IsDate(strFechaFinAlInicio) Then
            strFechaFinAlInicio = Format(strFechaFinAlInicio, "dd/mm/yyyy")
        End If
        Me.FechaFin = strFechaFinAlInicio
    End If
    If IsNumeric(strImporteAlInicio) Then
        Me.Importe = Format(strImporteAlInicio, "#,##0.00 €")
    End If
    If strEstaFacturado = "No" Then
        Me.lblTitulo1.Caption = "EDICIÓN DE SUBCONTRATACIÓN: " & Format(strIDSubContratacion, "000")
       For Each ctl In Me.Controls
            If ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox Then
                ctl.BackColor = lngColorCampoEditable
                ctl.Enabled = True
                ctl.Locked = False
            End If
        Next
        Me.ComandoRegistrar.Enabled = True
        Me.ComandoAdjuntar.Enabled = True
    Else
        Me.lblTitulo1.Caption = "DETALLE DE SUBCONTRATACIÓN: " & Format(strIDSubContratacion, "000")
        Me.cmdSalir.SetFocus
        For Each ctl In Me.Controls
            If ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox Then
                ctl.BackColor = lngColorCampoNoEditable
                ctl.Enabled = True
                ctl.Locked = True
            End If
        Next
        Me.ComandoRegistrar.Enabled = False
        Me.ComandoAdjuntar.Enabled = False
    End If
    If strEstaFacturado = "Sí" Then
        Me.lblFacturada.Visible = True
    Else
        Me.lblFacturada.Visible = False
    End If
    
    Set m_SubContratacion = Constructor.getSubContratacion(strIDSubContratacion, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    strURLAnexo = m_SubContratacion.URLAnexos
    If strURLAnexo = "" Then
        Me.ImagenAnexo.Visible = False
    Else
        Me.ImagenAnexo.Visible = True
    End If
    EstablecerDatos = "Datos de la subcontratación establecidos correctamente"
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
    DoCmd.Hourglass True
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
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function editar( _
                        strIDSubContratacion As String, _
                        Optional strEmpresa As String, _
                        Optional strFechaInicial As String, _
                        Optional strDESCRIPCION As String, _
                        Optional strImporte As String, _
                        Optional strFechaFin As String _
                        ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 13/06/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -1) Ver si se ha facturado ya si lo está no se puede cambiar strFechaInicial,strImporte,strFechaFin
    '   -Llamada desde
    '       -Form_FormSubContratacionEdicion.Editar
    '   -Devuelve:
    '       editar = Descriptivo
    '       editar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strID As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDSubContratacion) Then
        strTextoError = "No se ha introducido un IDSubcontratación correcto"
        Err.Raise 1000
    End If
    If strEmpresa = "" Or strEmpresa = "#" Then
        strTextoError = "Se ha de introducir el nombre de la empresa subcontratada"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaInicial) Then
        strTextoError = "Se ha de introducir una fecha inicial válida"
        Err.Raise 1000
    End If
    If Not IsNumeric(strImporte) Then
        strTextoError = "Se ha de introducir un importe válido"
        Err.Raise 1000
    End If
    If strFechaFin <> "" And strFechaFin <> "#" Then
        If CDate(strFechaFin) < CDate(strFechaInicial) Then
            strTextoError = "Si se introduce una fecha de fin ha de ser mayor o igual que la inicial"
            Err.Raise 1000
        End If
    End If
    If strEstaFacturado = "Sí" Then
        strTextoError = "Ya está facturada, por lo que no se permiten cambios"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbSubcontrataciones.* " & _
            "FROM TbSubcontrataciones " & _
            "WHERE (((TbSubcontrataciones.IDSubcontratacion)=" & strIDSubContratacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "La subcontratación: " & strIDSubContratacion & " no estaba registrada en la base de datos"
            Err.Raise 1000
        End If
        .Edit
            If IsDate(strFechaInicial) Then
                .Fields("FechaAlta") = strFechaInicial
            End If
            If strEmpresa <> "" Then
                .Fields("Empresa") = strEmpresa
            End If
            If strDESCRIPCION <> "" And strDESCRIPCION <> "#" Then
                .Fields("Descripcion") = strDESCRIPCION
            ElseIf strDESCRIPCION = "#" Then
                .Fields("Descripcion") = Null
            End If
            If IsDate(strFechaFin) Then
                 .Fields("FechaFin") = strFechaFin
            Else
                If strFechaFin = "#" Then
                    .Fields("FechaFin") = Null
                End If
            End If
            If IsNumeric(strImporte) Then
                .Fields("Importe") = strImporte
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    editar = "Editada correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Editar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    editar = "#ERR" & "|" & strTextoError
End Function




