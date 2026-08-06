Option Compare Database
Option Explicit
Private m_Error As String

Private m_EquipoEnCombo As EquipoMedida

Private m_EventoEquipoSeleccionado As EventoEquipoMedida
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
    m_ObjEventoActivo.RegistrarEquipoMedida m_EquipoEnCombo.IDEquipoMedida, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent Alta
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
    If m_ObjEventoActivo.Modificable = EnumSino.No Then
        m_Error = "El evento no es modificable"
        Err.Raise 1000
    End If
    If m_EventoEquipoSeleccionado Is Nothing Then
        Set m_EventoEquipoSeleccionado = Constructor.getEquipoMedidaCalibracion(Nz(Me.ListaEventoEqMedida.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_EventoEquipoSeleccionado Is Nothing Then
            m_Error = "No se ha seleccionado un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea desvincular el equipo de medida con el evento?", vbExclamation + vbYesNo + vbDefaultButton2, "Desvincular equipo de medida")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjEventoActivo.EliminarEquipoMedida m_EventoEquipoSeleccionado.IDEventoEquipoMedida, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent Eliminado
    Me.ComboEquipoDeMedida.SetFocus
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
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
    If m_ObjEventoActivo.Modificable = EnumSino.No Then
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
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.ComandoAlta.Enabled = False
    End If
    Me.ComandoAlta.Enabled = True
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If m_EquipoEnCombo.Calibrado = EnumSino.No Then
        pregunta = MsgBox("Equipo fuera de calibración", vbExclamation, "Equipo fuera de calibración")
    End If
    ComandoAlta_Click
    
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
   
    If m_ObjEventoActivo Is Nothing Then
        m_Error = "No hay evento activo"
        Err.Raise 1000
    End If
    Set cmb = Me.ComboEquipoDeMedida
    cmb.RowSource = ""
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
    Dim m_EventoEquipo As EventoEquipoMedida
    Dim m_EstadoCalibracion As String
    
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    Me.ComandoEliminar.Enabled = False
    Me.ComandoAlta.Enabled = False
    Set lst = Me.ListaEventoEqMedida
    lst.RowSource = "ID;Equipo Medida;Calibración"
    
    
    
    Set m_Col = m_ObjEventoActivo.ColEquiposMedida
    p_Error = m_ObjEventoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
         For Each m_ID In m_Col
            Set m_EventoEquipo = m_Col(m_ID)
            If IsNumeric(m_EventoEquipo.IDCalibracion) Then
                m_EstadoCalibracion = "Calibrado"
            Else
                m_EstadoCalibracion = "Sin Calibrar"
            End If
            
            lst.AddItem m_EventoEquipo.IDEventoEquipoMedida & ";" & m_EventoEquipo.EquipoMedida.NombreParaCombo & ";" & m_EstadoCalibracion
            Set m_EventoEquipo = Nothing
        Next
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub ListaEventoEqMedida_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Me.ComandoEliminar.Enabled = False
    If Nz(Me.ListaEventoEqMedida, "") = "" Then
        Set m_EventoEquipoSeleccionado = Nothing
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjEventoActivo.Modificable = EnumSino.No Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_EventoEquipoSeleccionado Is Nothing Then
        Set m_EventoEquipoSeleccionado = Constructor.getEventoEquipoMedida(Me.ListaEventoEqMedida.Column(0))
    Else
        If m_EventoEquipoSeleccionado.IDEventoEquipoMedida <> Me.ListaEventoEqMedida.Column(0) Then
            Set m_EventoEquipoSeleccionado = Constructor.getEventoEquipoMedida(Me.ListaEventoEqMedida.Column(0))
        End If
    End If
    If Not m_EventoEquipoSeleccionado Is Nothing Then
         Me.ComandoEliminar.Enabled = True
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaEventoEqMedida_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub