Option Compare Database
Option Explicit

Private m_CalibracionAlInicio As EquipoMedidaCalibracion
Private m_Error As String
Public Event Alta()
Public Event Edicion()

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoRegistrar_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Not HaHabidoCambios(m_Error) Then
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No ha habido cambios"
        Err.Raise 1000
    End If
    RellenarObjeto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_CalibracionAlInicio Is Nothing Then
        m_ObjCalibracionEquipoMedidaActiva.ParaAlta = EnumSino.No
    Else
        m_ObjCalibracionEquipoMedidaActiva.ParaAlta = EnumSino.Sí
    End If
    m_ObjCalibracionEquipoMedidaActiva.Registrar m_CalibracionAlInicio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CalibracionAlInicio Is Nothing Then
        RaiseEvent Alta
    Else
        RaiseEvent Edicion
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If m_ObjEquipoMedidaActivo Is Nothing Then
        m_Error = "No se ha indicado el equipo de medida"
        Err.Raise 1000
    End If
    
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
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_Titulo As String
    Dim m_Valor As String
    
    On Error GoTo errores
    
    If m_ObjEquipoMedidaActivo Is Nothing Then
        p_Error = "No se ha indicado el equipo de medida"
        Err.Raise 1000
    End If
    
    With m_ObjEquipoMedidaActivo
        Me.Nombre = .Nombre & " (" & .Marca & ")"
        Me.Modelo = .Modelo
        Me.NS = .NS
        Me.PN = .PN
    End With
    If m_ObjCalibracionEquipoMedidaActiva Is Nothing Then
        m_Titulo = "ALTA DE CALIBRACIÓN"
        Me.lblTitulo1.Caption = m_Titulo
        Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
        Me.Caption = Me.lblTitulo1.Caption
        Exit Function
    End If
    m_Titulo = "EDICIÓN DE CALIBRACIÓN"
    Me.lblTitulo1.Caption = m_Titulo
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    Set m_CalibracionAlInicio = Constructor.getEquipoMedidaCalibracion(m_ObjCalibracionEquipoMedidaActiva.IDCalibracion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CalibracionAlInicio Is Nothing Then
        p_Error = "No parece estar registrado en la base de datos"
        Err.Raise 1000
    End If
    With m_CalibracionAlInicio
        If IsDate(.FechaCalibracion) Then
            Me.FechaCalibracion = .FechaCalibracion
        End If
        If IsDate(.FechaFinCalibracion) Then
            Me.FechaFinCalibracion = .FechaFinCalibracion
        End If
    End With
    
    
    EstablecerDatos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function RellenarObjeto(Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim m_Valor As String
    On Error GoTo errores
    
    
    If m_ObjCalibracionEquipoMedidaActiva Is Nothing Then
        Set m_ObjCalibracionEquipoMedidaActiva = New EquipoMedidaCalibracion
    End If
    With m_ObjCalibracionEquipoMedidaActiva
        .FechaCalibracion = Nz(Me.FechaCalibracion, "")
        .FechaFinCalibracion = Nz(Me.FechaFinCalibracion, "")
        .IDEquipoMedida = m_ObjEquipoMedidaActivo.IDEquipoMedida
    End With
        
    
    RellenarObjeto = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjeto ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    On Error GoTo errores
    
    If m_CalibracionAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_CalibracionAlInicio
        If .FechaCalibracion <> Nz(Me.FechaCalibracion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaFinCalibracion <> Nz(Me.FechaFinCalibracion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function