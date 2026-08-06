Option Compare Database
Option Explicit

Private m_ObjLibranzaAlInicio As LIbranza
Private m_Error As String
Public Event Alta()
Public Event Editado()

Private Sub ComandoRegistrar_Click()
    
    
    Dim m_FechaInicial As String
    Dim m_FechaFinal As String
    Dim m_Texto As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    If m_ObjLibranzaActiva Is Nothing Then
        Set m_ObjLibranzaActiva = New LIbranza
    End If
    
    m_FechaInicial = Nz(Me.FechaInicial, "")
    m_FechaFinal = Nz(Me.FechaFinal, "")
    
    With m_ObjLibranzaActiva
        Set .Tecnico = m_ObjTecnicoActivo
        m_Error = .Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        .Alias = m_ObjTecnicoActivo.Alias
        .MOTIVO = Nz(Me.ComboMotivo, "")
        .Observaciones = Nz(Me.Observaciones, "")
        
    End With
    m_Texto = m_ObjTecnicoActivo.AltaAusencias(m_ObjLibranzaActiva, m_ObjLibranzaAlInicio, m_FechaInicial, m_FechaFinal, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent Alta
    pregunta = MsgBox(m_Texto, vbInformation, "Alta por lotes")
    
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
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Load()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AjustarTamaño Me

    
    If m_ObjTecnicoActivo Is Nothing Then
        If m_ObjLibranzaActiva Is Nothing Then
            m_Error = "No se conoce el técnico"
            Err.Raise 1000
        End If
        Set m_ObjTecnicoActivo = m_ObjLibranzaActiva.Tecnico
        m_Error = m_ObjLibranzaActiva.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
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
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
    
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    RellenarCombo p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjLibranzaActiva Is Nothing Then
        Me.lblTitulo1.Caption = "Alta de Libranzas para " & m_ObjTecnicoActivo.Alias
        Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
        Me.Caption = Me.lblTitulo1.Caption
        Exit Function
    End If
    Set m_ObjLibranzaAlInicio = Constructor.getLibranza(m_ObjLibranzaActiva.IDAusencia, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblTitulo1.Caption = "Edición de Libranza para " & m_ObjTecnicoActivo.Alias
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    With m_ObjLibranzaAlInicio
        If IsDate(.FechaLibranza) Then
            Me.FechaInicial = .FechaLibranza
            Me.FechaFinal = .FechaLibranza
        End If
        If .MOTIVO <> "" Then
            Me.ComboMotivo = .MOTIVO
        End If
        If .Observaciones <> "" Then
            Me.Observaciones = .Observaciones
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As EnumSino
    
    On Error GoTo errores
    If m_ObjLibranzaAlInicio Is Nothing Then
        HaHabidoCambios = EnumSino.Sí
        Exit Function
    End If
    
    With m_ObjLibranzaAlInicio
        If Nz(Me.FechaFinal, "") <> Nz(Me.FechaInicial, "") Then
            HaHabidoCambios = EnumSino.Sí
            Exit Function
        End If
        If Nz(Me.FechaInicial, "") <> .FechaLibranza Then
            HaHabidoCambios = EnumSino.Sí
            Exit Function
        End If
        If Nz(Me.ComboMotivo, "") <> .MOTIVO Then
            HaHabidoCambios = EnumSino.Sí
            Exit Function
        End If
        If Nz(Me.Observaciones, "") <> .Observaciones Then
            HaHabidoCambios = EnumSino.Sí
            Exit Function
        End If
        
    End With
    HaHabidoCambios = EnumSino.No
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function RellenarCombo(Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim cmb As ComboBox
    On Error GoTo errores
    
    Set cmb = Me.ComboMotivo
    cmb.RowSource = ""
    m_SQL = "SELECT DISTINCT TbTecnicosAusencias.Motivo FROM TbTecnicosAusencias; "
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            cmb.AddItem Nz(.Fields("Motivo"), "")
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCombo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function