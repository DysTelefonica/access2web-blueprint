Option Compare Database
Option Explicit
Private m_Error As String
Private m_CalibracionSeleccionada As EquipoMedidaCalibracion
Private WithEvents m_FormCalibracion As Form_FormEquipoMedidaCalibracion
Public Function Filtrar( _
                        Optional ByRef p_Error As String _
                        ) As String
       
       
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_objCalibracion As EquipoMedidaCalibracion
    Dim lst As ListBox
    
    On Error GoTo errores
    
    Set m_CalibracionSeleccionada = Nothing
    Me.ComandoAlta.Enabled = True
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoExportarExcel.Enabled = False
    
    
    Set lst = Me.ListaFiltrados
    lst.RowSource = "ID;Fecha Inicial;Fecha Final"
    Set m_Col = m_ObjEquipoMedidaActivo.Calibraciones
    p_Error = m_ObjEquipoMedidaActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_objCalibracion = m_Col(m_ID)
        With m_objCalibracion
            lst.AddItem .IDCalibracion & ";" & .FechaCalibracion & ";" & .FechaFinCalibracion
        End With
        Set m_objCalibracion = Nothing
    Next
    Me.ComandoExportarExcel.Enabled = True
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaFiltrados_Click
        Exit Function
    Else
       ListaFiltrados_Click
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function


Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub ComandoAlta_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjCalibracionEquipoMedidaActiva = Nothing
    If FormularioAbierto("FormEquipoMedidaCalibracion") Then
        DoCmd.Close acForm, "FormEquipoMedidaCalibracion", acSaveNo
    End If
    DoCmd.OpenForm "FormEquipoMedidaCalibracion"
    Set m_FormCalibracion = Forms("FormEquipoMedidaCalibracion")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEditar_Click()

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_CalibracionSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_CalibracionSeleccionada = Constructor.getEquipoMedidaCalibracion(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_CalibracionSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar la calibración de la lista primero"
            Err.Raise 1000
        End If
    End If
    
    Set m_ObjCalibracionEquipoMedidaActiva = m_CalibracionSeleccionada
    If FormularioAbierto("FormEquipoMedidaCalibracion") Then
        DoCmd.Close acForm, "FormEquipoMedidaCalibracion", acSaveNo
    End If
    DoCmd.OpenForm "FormEquipoMedidaCalibracion"
    Set m_FormCalibracion = Forms("FormEquipoMedidaCalibracion")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_CalibracionSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_CalibracionSeleccionada = Constructor.getEquipoMedidaCalibracion(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_CalibracionSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar la calibración de la lista primero"
            Err.Raise 1000
        End If
    End If
    m_CalibracionSeleccionada.Eliminar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_CalibracionSeleccionada = Nothing
    
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoExportarExcel_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    flag = GenerarConsultas("Lista de Calibraciones", Me.hWnd, , Me.ListaFiltrados, "No")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método GenerarConsultas ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoExportarExcel_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If m_ObjEquipoMedidaActivo Is Nothing Then
        m_Error = "No se conoce el equipo"
        Err.Raise 1000
    End If
    Me.Caption = Me.lblTitulo1.Caption
    Me.Nombre = m_ObjEquipoMedidaActivo.NombreParaCombo
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub ListaFiltrados_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.ComandoAlta.Enabled = True
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_CalibracionSeleccionada Is Nothing Then
        Set m_CalibracionSeleccionada = Constructor.getEquipoMedidaCalibracion(Me.ListaFiltrados.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaFiltrados.Column(0) <> m_CalibracionSeleccionada.IDCalibracion Then
            Set m_CalibracionSeleccionada = Constructor.getEquipoMedidaCalibracion(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_CalibracionSeleccionada Is Nothing Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    Me.ComandoEditar.Enabled = True
    Me.ComandoEliminar.Enabled = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub m_FormCalibracion_Alta()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjEquipoMedidaActivo.Calibraciones = Nothing
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormCalibracion_Alta se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormCalibracion_Edicion()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjEquipoMedidaActivo.Calibraciones = Nothing
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormCalibracion_Edicion se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub