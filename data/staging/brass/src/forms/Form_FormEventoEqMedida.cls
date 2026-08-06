Option Compare Database
Option Explicit


Private m_EquipoEnCombo As EquipoMedida
Private m_EquipoSeleccionado As EquipoMedida
Private m_Error As String
Public Event Alta()
Public Event Eliminado()

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAlta_Click()
    
    Dim m_EquipoEnCombo As EquipoMedida
    Dim lst As ListBox
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    Set m_EquipoEnCombo = Constructor.getEquipoMedida(Nz(Me.ComboEquipoDeMedida.Column(0), ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EquipoEnCombo Is Nothing Then
        m_Error = "Primero seleccione un equipo de medida"
        Err.Raise 1000
    End If
    If m_ColEquiposMedidaPorEventoActivos Is Nothing Then
        Set m_ColEquiposMedidaPorEventoActivos = New Scripting.Dictionary
        m_ColEquiposMedidaPorEventoActivos.CompareMode = TextCompare
    End If
    If Not m_ColEquiposMedidaPorEventoActivos.Exists(m_EquipoEnCombo.IDEquipoMedida) Then
        m_ColEquiposMedidaPorEventoActivos.Add m_EquipoEnCombo.IDEquipoMedida, m_EquipoEnCombo
        RaiseEvent Alta
    End If
    Me.ComboEquipoDeMedida = Null
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.SetFocus
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
    
    Dim m_EquipoEnLista As EquipoMedida
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    Set m_EquipoEnLista = Constructor.getEquipoMedida(Nz(Me.ListaEqMedida.Column(0), ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EquipoEnLista Is Nothing Then
        m_Error = "Primero seleccione un equipo de medida"
        Err.Raise 1000
    End If
    If Not m_ColEquiposMedidaPorEventoActivos Is Nothing Then
        If m_ColEquiposMedidaPorEventoActivos.Exists(m_EquipoEnLista.IDEquipoMedida) Then
            m_ColEquiposMedidaPorEventoActivos.Remove m_EquipoEnLista.IDEquipoMedida
            If m_ColEquiposMedidaPorEventoActivos.count = 0 Then
                Set m_ColEquiposMedidaPorEventoActivos = Nothing
            End If
            RaiseEvent Eliminado
        End If
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_EquipoSeleccionado = Nothing
    Me.ComandoEliminar.Enabled = False
    Me.SetFocus
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboEquipoDeMedida_AfterUpdate()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Nz(Me.ComboEquipoDeMedida, "") = "" Then
        Set m_EquipoEnCombo = Nothing
        Me.ComandoAlta.Enabled = False
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_EquipoEnCombo Is Nothing Then
        Set m_EquipoEnCombo = Constructor.getEquipoMedida(Me.ComboEquipoDeMedida.Column(0))
    Else
        If m_EquipoEnCombo.IDEquipoMedida <> Me.ComboEquipoDeMedida.Column(0) Then
            Set m_EquipoEnCombo = Constructor.getEquipoMedida(Me.ComboEquipoDeMedida.Column(0))
        End If
    End If
    If m_EquipoEnCombo Is Nothing Then
         Me.ComandoAlta.Enabled = False
    Else
        If m_ObjEventoActivo Is Nothing Then
            Me.ComandoAlta.Enabled = True
        Else
            If m_ObjEventoActivo.FechaFacturado <> "" Or m_ObjEventoActivo.Franqueado = True Then
                Me.ComandoAlta.Enabled = False
            Else
                Me.ComandoAlta.Enabled = True
            End If
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If m_EquipoEnCombo.Calibrado = EnumSino.No Then
        pregunta = MsgBox("Equipo fuera de calibración", vbExclamation, "Equipo fuera de calibración")
    End If
    If Me.ComandoAlta.Enabled = True Then
    
        ComandoAlta_Click
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboEquipoDeMedida_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub Form_Load()
    
    Dim cmb As ComboBox
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AjustarTamaño Me
    Set m_ColEquiposMedidaPorEventoActivos = Nothing
    
    
    RellenarListaEquiposMedida Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = Me.lblTitulo1.Caption
    
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
   
    Dim m_Col As Scripting.Dictionary
    Dim m_Equipo As EquipoMedida
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    Me.ComandoEliminar.Enabled = False
    Me.ComandoAlta.Enabled = False
    Set lst = Me.ListaEqMedida
    lst.RowSource = "ID;Equipo Medida;Calibración"
    Set m_Col = m_ColEquiposMedidaPorEventoActivos
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Equipo = m_Col(m_ID)
        lst.AddItem m_Equipo.IDEquipoMedida & ";" & m_Equipo.NombreParaCombo & ";" & m_Equipo.EstadoCalibracion
        Set m_Equipo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub ListaEqMedida_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Nz(Me.ListaEqMedida, "") = "" Then
        Set m_EquipoSeleccionado = Nothing
        Me.ComandoEliminar.Enabled = False
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_EquipoSeleccionado Is Nothing Then
        Set m_EquipoSeleccionado = Constructor.getEquipoMedida(Me.ListaEqMedida.Column(0))
    Else
        If m_EquipoSeleccionado.IDEquipoMedida <> Me.ListaEqMedida.Column(0) Then
            Set m_EquipoSeleccionado = Constructor.getEquipoMedida(Me.ListaEqMedida.Column(0))
        End If
    End If
    If m_EquipoSeleccionado Is Nothing Then
         Me.ComandoEliminar.Enabled = False
    Else
        If m_ObjEventoActivo Is Nothing Then
            Me.ComandoEliminar.Enabled = True
        Else
            If m_ObjEventoActivo.FechaFacturado <> "" Or m_ObjEventoActivo.Franqueado = True Then
                Me.ComandoEliminar.Enabled = False
            Else
                Me.ComandoEliminar.Enabled = True
            End If
        End If
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaEqMedida_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub