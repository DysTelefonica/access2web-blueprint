 Option Compare Database
Option Explicit

Public Event Alta()
Public Event Editado()

Private m_Error As String

Private Sub cmdSalir_Click()
     On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If m_ObjMotivoActivo Is Nothing Then
       Exit Function
       
    End If
    With m_ObjMotivoActivo
        If .MotivoHPS <> "" Then
            Me.MotivoHPS = .MotivoHPS
        End If
        
    End With
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Sub ComandoGrabar_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Nz(Me.MotivoHPS, "") = "" Then
        m_Error = "Introduzca el motivo"
        Err.Raise 1000
    End If
    If m_EnOficina = EnumSiNo.no Then
        m_Error = "Sólo puede hacerse en la oficina"
        Err.Raise 1000
    End If
    If m_ObjMotivoActivo Is Nothing Then
        Set m_ObjMotivoActivo = New MotivoHPS
        m_ObjMotivoActivo.MotivoHPS = Nz(Me.MotivoHPS, "")
        m_ObjMotivoActivo.AltaMotivo m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjEntorno.ColMotivos = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("Registrado correctamente", vbInformation, "Registro")
        RaiseEvent Alta
        DoCmd.Close acForm, Me.Name, acSaveNo
        Exit Sub
        
    End If
    If m_ObjMotivoActivo.MotivoHPS = Nz(Me.MotivoHPS, "") Then
        m_Error = "Sin cambios que guardar"
        Err.Raise 1000
    End If
    m_ObjMotivoActivo.EditarMotivo m_ObjMotivoActivo.MotivoHPS, Me.MotivoHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Registrado correctamente", vbInformation, "Registro")
    RaiseEvent Editado
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Ajustar Me
    Me.Caption = Me.lblTitulo.Caption
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
