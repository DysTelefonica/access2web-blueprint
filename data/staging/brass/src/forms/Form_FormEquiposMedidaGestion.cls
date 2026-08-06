Option Compare Database
Option Explicit
Private m_Error As String
Private m_EquipoSeleccionado As EquipoMedida
Private WithEvents m_FormEquipo As Form_FormEquipoMedida
Public Event EquipoSeleccionado(m_Equipo As EquipoMedida)

Public Function Filtrar( _
                        Optional ByRef p_Error As String _
                        ) As String
       
       
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjEquipo As EquipoMedida
    Dim lst As ListBox
    
    On Error GoTo errores
    
    Set m_EquipoSeleccionado = Nothing
    Me.ComandoAltaEquipo.Enabled = True
    Me.ComandoEditarEquipo.Enabled = False
    Me.ComandoEliminarEquipo.Enabled = False
    Me.ComandoCalibraciones.Enabled = False
    
    Me.ComandoExportarExcel.Enabled = False
    Set lst = Me.ListaFiltrados
    lst.RowSource = "ID;Equipo;Marca;Modelo;NS"
    Set m_Col = Constructor.getEquiposMedida(Nz(Me.Nombre, ""), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjEquipo = m_Col(m_ID)
        With m_ObjEquipo
            lst.AddItem .IDEquipoMedida & ";" & .Nombre & ";" & .Marca & ";" & .Modelo & ";" & .NS
        End With
        Set m_ObjEquipo = Nothing
    Next
    Me.ComandoExportarExcel.Enabled = True
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaFiltrados_Click
        Exit Function
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

Private Sub ComandoAltaEquipo_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjEquipoMedidaActivo = Nothing
    If FormularioAbierto("FormEquipoMedida") Then
        DoCmd.Close acForm, "FormEquipoMedida", acSaveNo
    End If
    DoCmd.OpenForm "FormEquipoMedida"
    Set m_FormEquipo = Forms("FormEquipoMedida")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaEquipo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoCalibraciones_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_EquipoSeleccionado Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_EquipoSeleccionado = Constructor.getEquipo(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_EquipoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar el equipo de la lista primero"
            Err.Raise 1000
        End If
    End If
    
    Set m_ObjEquipoMedidaActivo = m_EquipoSeleccionado
    If FormularioAbierto("FormEquipoMedidaCalibracionesGestion") Then
        DoCmd.Close acForm, "FormEquipoMedidaCalibracionesGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormEquipoMedidaCalibracionesGestion"
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCalibraciones_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEditarEquipo_Click()

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_EquipoSeleccionado Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_EquipoSeleccionado = Constructor.getEquipo(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_EquipoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar el equipo de la lista primero"
            Err.Raise 1000
        End If
    End If
    
    Set m_ObjEquipoMedidaActivo = m_EquipoSeleccionado
    If FormularioAbierto("FormEquipoMedida") Then
        DoCmd.Close acForm, "FormEquipoMedida", acSaveNo
    End If
    DoCmd.OpenForm "FormEquipoMedida"
    Set m_FormEquipo = Forms("FormEquipoMedida")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditarEquipo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarEquipo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_EquipoSeleccionado Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_EquipoSeleccionado = Constructor.getEquipo(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_EquipoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar el equipo de la lista primero"
            Err.Raise 1000
        End If
    End If
    m_EquipoSeleccionado.Eliminar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjEquipoMedidaActivo = Nothing
    Set m_EquipoSeleccionado = Nothing
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
        m_Error = "Al ComandoEditarEquipo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
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
    
    flag = GenerarConsultas("Lista de equipos de medida", Me.hWnd, , Me.ListaFiltrados, "No")
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

Private Sub ComandoLimpiarNombre_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.Nombre = Empty
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
        m_Error = "Al ComandoLimpiarNombre_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoSeleccionarEquipo_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_EquipoSeleccionado Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_EquipoSeleccionado = Constructor.getEquipo(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_EquipoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar el equipo de la lista primero"
            Err.Raise 1000
        End If
    End If
    If m_EquipoSeleccionado.Calibrado = EnumSino.No Then
        m_Error = m_EquipoSeleccionado.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "El equipo seleccionado no está activo. Ya sea porque tiene fecha de fin de uso o por estar fuera de calibración"
        Err.Raise 1000
    Else
        m_Error = m_EquipoSeleccionado.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    RaiseEvent EquipoSeleccionado(m_EquipoSeleccionado)
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSeleccionarEquipo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
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
    Me.Caption = Me.lblTitulo1.Caption
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
    Dim m_EstadoCalibracion As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Me.ComandoAltaEquipo.Enabled = True
    Me.ComandoEditarEquipo.Enabled = False
    Me.ComandoEliminarEquipo.Enabled = False
    Me.ComandoCalibraciones.Enabled = False
    
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        DoCmd.Hourglass False
        Me.lblEstadoCalibracion.Visible = False
        Exit Sub
    End If
    If m_EquipoSeleccionado Is Nothing Then
        Set m_EquipoSeleccionado = Constructor.getEquipoMedida(Me.ListaFiltrados.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Me.lblEstadoCalibracion.Visible = True
        Me.lblEstadoCalibracion.Caption = m_EquipoSeleccionado.EstadoCalibracion
    Else
        If Me.ListaFiltrados.Column(0) <> m_EquipoSeleccionado.IDEquipoMedida Then
            Set m_EquipoSeleccionado = Constructor.getEquipoMedida(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            Me.lblEstadoCalibracion.Visible = True
            Me.lblEstadoCalibracion.Caption = m_EquipoSeleccionado.EstadoCalibracion
        End If
    End If
    If m_EquipoSeleccionado Is Nothing Then
        Me.lblEstadoCalibracion.Visible = False
        DoCmd.Hourglass False
        Exit Sub
    End If
    Me.ComandoEditarEquipo.Enabled = True
    Me.ComandoEliminarEquipo.Enabled = True
    Me.ComandoCalibraciones.Enabled = True
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

Private Sub m_FormEquipo_EquipoAlta()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
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
        m_Error = "Al m_FormEquipo_EquipoAlta se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormEquipo_EquipoModificado()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
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
        m_Error = "Al m_FormEquipo_EquipoModificado se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Nombre_Exit(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
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
        m_Error = "Al Nombre_Exit se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub