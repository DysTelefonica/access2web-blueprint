Option Compare Database
Option Explicit

Public Event Alta()
Public Event Editado()

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
    m_Error = ""
    DoCmd.Hourglass True
    InformeDelRegistro p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "El método ComandoAyudaFormulario_Click ha devuelto el error: " & Err.Numbr & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
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
        Me.lblTitulo.Caption = "SOLICITUD DE " & UCase(m_Tipo) & " PARA " & UCase(m_ObjSolicitudActiva.NombreCompleto)
    Else
        
         Me.lblModificacion.Visible = False
         Set m_ObjSolicitudActiva = New solicitud
         m_ObjSolicitudActiva.TIPO = m_Tipo
         
         Me.NavFechas.Enabled = False
        Me.lblTitulo.Caption = "SOLICITUD DE " & UCase(m_Tipo)
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
'
Public Sub Form_Load()



    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me

    'Me.AllowEdits = False
    Me.Caption = m_TituloFormulario
    m_Tipo = Nz(Me.OpenArgs(), "")
    If m_Tipo <> "Alta" And m_Tipo <> "Traspaso" Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.NavDatos.NavigationTargetName = "FormSolicitudesAltaTraspasoDatos"
    Me.FrmDetalle.SourceObject = "FormSolicitudesAltaTraspasoDatos"
    If m_Tipo = "Alta" Then
        Me.NavFechas.NavigationTargetName = "FormSolicitudesAltaFechas"
    Else
        Me.NavFechas.NavigationTargetName = "FormSolicitudesTraspasoFechas"
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    'Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    'Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
