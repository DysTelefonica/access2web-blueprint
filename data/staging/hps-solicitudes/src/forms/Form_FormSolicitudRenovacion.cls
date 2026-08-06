Option Compare Database
Option Explicit


Private m_Tipo As String
Private m_Error As String
Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    DoCmd.Hourglass True
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "El método ComandoAyudaFormulario_Click ha devuelto el error: " & Err.Numbr & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerLogs_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    InformeDelRegistro p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=m_Error
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
        m_Error = "Al ComandoVerLogs_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

'
Public Sub Form_Load()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me

    Me.AllowEdits = False
    Me.Caption = m_TituloFormulario
    m_Tipo = Nz(Me.OpenArgs(), "")
    If m_Tipo <> "Renovación" And m_Tipo <> "Elevación Grado" Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    If m_ObjUsuarioHPSActivo Is Nothing Then
        m_Error = "No se ha seleccionado previamente el usuario HPS"
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.NavDatos.NavigationTargetName = "FormSolicitudRenovacionDatos"
    Me.FrmDetalle.SourceObject = "FormSolicitudRenovacionDatos"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
   
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
     If Not m_ObjSolicitudActiva Is Nothing Then
        If m_ObjSolicitudActiva.IDSolicitud = "" Then
            Me.lblModificacion.Visible = False
            Me.lblTitulo.Caption = "ALTA DE SOLICITUD DE HPS"
            Me.NavFechas.Enabled = False
        Else
            Set m_ObjSolicitudAlInicio = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjSolicitudAlInicio Is Nothing Then
                p_Error = "El formulario se ha abierto con parámetros insuficientes"
                Err.Raise 1000
            End If
            
            Me.NavFechas.Enabled = True
        End If
        
       
        Me.lblModificacion.Visible = True
        Me.lblModificacion.Caption = m_ObjSolicitudActiva.UltimaModificacionTexto
        
    Else
         Me.lblModificacion.Visible = False
         Set m_ObjSolicitudActiva = New solicitud
         m_ObjSolicitudActiva.TIPO = m_Tipo
         m_ObjSolicitudActiva.IDUsuarioHPS = m_ObjUsuarioHPSActivo.id
         
         
         Me.NavFechas.Enabled = False
    End If
    Me.lblTitulo.Caption = "RENOVACIÓN PARA " & UCase(m_ObjUsuarioHPSActivo.NombreCompleto)
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function


Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub