Option Compare Database
Option Explicit

Private m_Error As String
Private m_ObjEventoAlInicio As Evento

Public Event Franqueado()

Private Sub ComandoFranquear_Click()
    
    On Error GoTo errores
    m_Error = ""
    If Nz(Me.CausaFin, "") = "" Then
        m_Error = "Ha de indicar alguna causa del franqueo"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjEventoActivo.Franquear Me.CausaFin, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent Franqueado
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoFranquear_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    Dim m_CausasCalculadas  As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    m_Error = ""
    
    If m_ObjEventoActivo Is Nothing Then
        m_Error = "No hay un evento activo"
        Err.Raise 1000
    End If
    Set m_ObjEventoActivo = Constructor.getEvento(m_ObjEventoActivo.IDEVENTO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjEventoActivo Is Nothing Then
        m_Error = "No hay un evento activo"
        Err.Raise 1000
    End If
    If m_ObjEventoActivo.Franqueado = True Then
        Me.ComandoFranquear.Enabled = False
        Me.AllowEdits = False
        If m_ObjEventoActivo.CausaFin <> "" Then
            Me.CausaFin = m_ObjEventoActivo.CausaFin
        End If
        Me.ComandoFranquear.Caption = "Quitar Franqueo"
    Else
        Me.ComandoFranquear.Enabled = True
        Me.AllowEdits = True
        m_CausasCalculadas = m_ObjEventoActivo.CausaFinCalculada
        m_Error = m_ObjEventoActivo.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_CausasCalculadas <> "" Then
            Me.CausaFin = m_CausasCalculadas
        End If
        Me.ComandoFranquear.Caption = "Franquear"
    End If
        
    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    Me.Repaint
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

Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub