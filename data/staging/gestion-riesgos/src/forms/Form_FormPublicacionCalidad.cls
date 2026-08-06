Option Compare Database
Option Explicit
Private m_Error As String



Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    
    Dim m_EnUTE As String
    Dim m_NecesitaEvidendiasSum As EnumSiNo
    Dim m_Particula As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
'    Debug.Print Me.InsideHeight
'    Debug.Print Me.InsideWidth
    Me.InsideHeight = 7185
    Me.InsideWidth = 13215
    Set m_ObjAnexoEvicenciaUTE = Nothing
    blnPermitidoEscribir = False
    m_URLInforme = ""
    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "No hay ninguna edición activa"
        Err.Raise 1000
    End If
    With m_ObjEdicionActiva
        If .EsActivo = EnumSiNo.Sí Then
            blnEdicionActiva = True
            m_Particula = "(edición activa)"
        Else
            blnEdicionActiva = False
            m_Particula = "(edición no activa)"
        End If
        Me.lblTitulo.Caption = .Proyecto.Proyecto & " Ed. " & .Edicion & " " & m_Particula
        
        If .EsActivo = EnumSiNo.No Then
            blnPermitidoEscribir = False
        Else
            If .UsuarioConectadoAutorizado = EnumSiNo.Sí Then
                blnPermitidoEscribir = True
            Else
                blnPermitidoEscribir = False
            End If
        End If
        m_EnUTE = .Proyecto.EnUTE
        'Set .ColSuministradores = Nothing
        m_NecesitaEvidendiasSum = .EvidenciasSuministradoresRequeridas
        If m_EnUTE = "Sí" Then
            Me.NavUTE.Enabled = True
            Set m_ObjAnexoEvicenciaUTE = .AnexoEvidenciaUTE
        Else
            Me.NavUTE.Enabled = False
        End If
        If m_NecesitaEvidendiasSum = EnumSiNo.Sí Then
            Me.NavAcuerdoSuministradores.Enabled = True
        Else
            Me.NavAcuerdoSuministradores.Enabled = False
        End If
    
    End With
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    
    On Error GoTo errores
    
    
        
    Me.NavPublicacion.NavigationTargetName = "FormPublicacionCalidadPublicar"
    Me.FormDetalle.SourceObject = "FormPublicacionCalidadPublicar"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Establecer datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function
