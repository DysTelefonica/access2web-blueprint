Option Compare Database
Option Explicit

Private WithEvents m_FormConf As Form_FormConfiguracion
Private WithEvents m_FormTareas As Form_FormTareasTramitadorPendientes
Private m_NombreTramitadorAlInicio As String
Private m_Error As String

Private Sub cmdSalir_Click()

    On Error GoTo errores
    m_Error = ""

    pregunta = MsgBox("¿Desea realmente salir de la aplicación", vbExclamation + vbYesNo, "Salir de la aplicación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    DoCmd.Quit acQuitSaveAll
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al cmdSalir_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
'
Private Sub ComandoActualizar_Click()

    On Error GoTo errores
  

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    ' --- NUEVO: FORZAR RECÁLCULO ---
    ' Como estamos en el botón actualizar manual, obligamos a consultar la BD de nuevo
    m_ObjEntorno.CalcularTotalesPendientes m_ObjUsuarioConectado, EsTecnico
    ' -------------------------------
    
    Ajustar Me
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AvanceCerrar
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
   
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
   DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
'
Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    CorreoAlAdministrador "Prueba de Error"
    Me.AllowEdits = False
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoConfiguracion_Click()

    On Error GoTo errores


    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If FormularioAbierto("FormConfiguracion") Then
        DoCmd.Close acForm, "FormConfiguracion", acSaveNo
    End If
    DoCmd.OpenForm "FormConfiguracion"
    If FormularioAbierto("FormConfiguracion") Then
        Set m_FormConf = Forms("FormConfiguracion")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoConfiguracion_Click se ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEntidades_Click()
    On Error GoTo errores


    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If Not FormularioAbierto("FormEntidades") Then
        DoCmd.OpenForm "FormEntidades"
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEntidades_Click se ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoGestionCorreos_Click()
    On Error GoTo errores


    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If Not FormularioAbierto("FormCorreosGestion") Then
        DoCmd.OpenForm "FormCorreosGestion"
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGestionCorreos_Click se ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoGestionSolicitudes_Click()
    On Error GoTo errores


    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If Not FormularioAbierto("FormSolicitudesGestion") Then
        DoCmd.OpenForm "FormSolicitudesGestion"
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGestionSolicitudes_Click se ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoTareasPendientes_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If FormularioAbierto("FormTareasTramitadorPendientes") Then
        DoCmd.Close acForm, "FormTareasTramitadorPendientes", acSaveNo
    End If
    DoCmd.OpenForm "FormTareasTramitadorPendientes"
    If FormularioAbierto("FormTareasTramitadorPendientes") Then
        Set m_FormTareas = Forms("FormTareasTramitadorPendientes")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoTareasPendientes_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String

   
    On Error GoTo errores
    
    
    
    Me.Caption = m_TituloFormulario
    Me.lblVersion.Caption = "Versión: " & m_ObjEntorno.VersionAplicacion
    Me.lblConfCorreosAutomaticos.Caption = m_ObjEntorno.TextoConfiguracionCorreosAutomaticos
    EstablecerIndicadores p_Error
    If p_Error <> "" Then Err.Raise 1000
    

    Me.Caption = m_TituloFormulario

    Me.ComandoTareasPendientes.Enabled = True
    Me.ComandoConfiguracion.Enabled = True
    Me.ComandoGestionSolicitudes.Enabled = True
    
    If m_FaltaAlgunParametroEnConfiguracion = True Then
        Me.lblParametros.Caption = "Parámetros INCOMPLETOS"
        Me.lblParametros.ForeColor = 683236
    Else
        Me.lblParametros.Caption = "Parámetros COMPLETOS"
        Me.lblParametros.ForeColor = 16737792
    End If
'    If Application.TempVars("ActivadoCorreoAutomatico") = "No" Then
'        Me.lblConfCorreosAutomaticos.Visible = True
'        Me.lblConfCorreosAutomaticos.Caption = "Aún no implementado correos automáticos"
'        Me.ComandoGestionCorreos.Enabled = False
'    Else
        Me.lblConfCorreosAutomaticos.Visible = False
        Me.lblEstado.Visible = False
        Me.ComandoGestionCorreos.Enabled = True
    'End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    Ajustar Me
    
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AvanceCerrar
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    If Application.TempVars("EnDesarrollo") <> "Sí" Then
        DoCmd.Quit acQuitSaveAll
    End If
    Me.AllowEdits = True
End Sub

Public Function EstablecerIndicadores( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_NTotalesTotal As Long
    Dim m_NTotalesUC As Long
    
    On Error GoTo errores
    
    ' -------------------------------------------------------------------------
    ' LECTURA INMEDIATA DE MEMORIA (Calculado previamente en Splash o al Actualizar)
    ' -------------------------------------------------------------------------
    
    m_NTotalesTotal = m_ObjEntorno.TotalPendientesGlobal
    m_NTotalesUC = m_ObjEntorno.TotalPendientesUsuario
    
    ' -------------------------------------------------------------------------
    ' PINTADO DE LA UI
    ' -------------------------------------------------------------------------
    
    If EsTecnico = EnumSiNo.Sí Then
        Me.lblTareasPendientes.Caption = "Pendientes (0/" & m_NTotalesTotal & ")"
        If m_NTotalesTotal > 0 Then
            Me.lblTareasPendientes.ForeColor = 683236  ' Rojo
        Else
            Me.lblTareasPendientes.ForeColor = 16737792 ' Verde/Azul
        End If
    Else
        Me.lblTareasPendientes.Caption = "Pendientes (" & m_NTotalesUC & "/" & m_NTotalesTotal & ")"
        
        If m_NTotalesTotal > 0 Then
            Me.lblTareasPendientes.ForeColor = 683236  ' Rojo
        Else
            Me.lblTareasPendientes.ForeColor = 16737792 ' Verde/Azul
        End If
    End If
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerIndicadores ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Sub m_FormConf_Modificado()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    If m_FaltaAlgunParametroEnConfiguracion = True Then
        Me.lblParametros.Caption = "Parámetros INCOMPLETOS"
        Me.lblParametros.ForeColor = 2366701
    Else
        Me.lblParametros.Caption = "Parámetros COMPLETOS"
        Me.lblParametros.ForeColor = -2147483602
    End If
    Me.lblConfCorreosAutomaticos.Caption = m_ObjEntorno.TextoConfiguracionCorreosAutomaticos
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al m_FormConf_Modificado se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormTareas_SolicitudModificada()
    EstablecerIndicadores p_Error:=m_Error
End Sub
