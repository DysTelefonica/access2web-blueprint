Option Compare Database
Option Explicit

Public Event Alta(m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio)
Private m_Error As String





Private Sub ComandoRegistrar_Click()
    
    
    Dim m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio
    Dim m_ObjTipoTecnico As TIPOTECNICO
    Dim blnTipoDadoDeAlta As Boolean
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Nz(Me.Tipo, "") = "" Then
        m_Error = "Falta tipo técnico"
        Err.Raise 1000
    End If
    If Not IsDate(Nz(Me.FechaInicial, "")) Then
        m_Error = "Hay que indicar la fecha inicial"
        Err.Raise 1000
    End If
    If Not IsNumeric(Nz(Me.PrecioHoraLab, "")) Then
        m_Error = "Hay que indicar el precio hora laborable"
        Err.Raise 1000
    End If
    If Not IsNumeric(Nz(Me.PrecioHoraExtra, "")) Then
        m_Error = "Hay que indicar el precio hora Extra"
        Err.Raise 1000
    End If
    Set m_ObjTipoTecnico = New TIPOTECNICO
    m_ObjTipoTecnico.ParaAlta = EnumSino.Sí
    m_ObjTipoTecnico.Tipo = Me.Tipo
    m_ObjTipoTecnico.Registrar , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    blnTipoDadoDeAlta = True
    Set m_ObjTipoTecnicoPrecio = New TipoTecnicoPrecio
    With m_ObjTipoTecnicoPrecio
        .ParaAlta = EnumSino.Sí
        Set .TIPOTECNICOObj = m_ObjTipoTecnico
        .TIPOTECNICO = Me.Tipo
        .FechaInicial = Me.FechaInicial
        .PrecioHoraLab = Me.PrecioHoraLab
        .PrecioHoraExtra = Me.PrecioHoraExtra
        .Registrar , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    RaiseEvent Alta(m_ObjTipoTecnicoPrecio)
    
    
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If blnTipoDadoDeAlta Then
        m_ObjTipoTecnico.Borrar m_Error
    End If
    
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
    
    AjustarTamaño Me

    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
