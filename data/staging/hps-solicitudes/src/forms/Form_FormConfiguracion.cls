Option Compare Database
Option Explicit

Public Event Modificado()
Private m_ConfAlInicio As Configuracion
Private m_Conf As Configuracion

Private m_Error As String


Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
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

Private Sub ComandoRegistrar_Click()
    
    Dim m_HaHabidoCambios As Boolean
    Dim m_ConfOp As ConfiguracionOperaciones
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_HaHabidoCambios Then
        m_Error = "No hay cambios que guardar"
        Err.Raise 1000
    End If
    RellenarObjeto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjEntorno.Configuracion = Nothing
    Set m_ConfOp = New ConfiguracionOperaciones
    With m_ConfOp
        .Registrar m_Conf, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Me.AllowEdits = True
    
    RaiseEvent Modificado
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Function ActivarCamposCorreo( _
                                        p_Activar As Boolean, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim ctl As Control
    
    On Error GoTo errores
    For Each ctl In Me.Controls
        If InStr(1, Nz(ctl.Tag, ""), "CORREO") <> 0 Then
            ctl.Enabled = p_Activar
        End If
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActivarCamposCorreo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Sub ComandoURLPlantillasExcel_Click()
     On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Not fso.FolderExists(m_ObjEntorno.URLDirectorioPlantillasEXCELConVersion) Then
        m_Error = "El directorio no es accesible en estos momentos"
        Err.Raise 1000
    End If
    Ejecutar Me.hwnd, "open", m_ObjEntorno.URLDirectorioPlantillasEXCELConVersion, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoURLPlantillasExcel_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
   
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoURLPlantillasHTML_Click()
    
    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Not fso.FolderExists(m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida) Then
        m_Error = "El directorio no es accesible en estos momentos"
        Err.Raise 1000
    End If
    Ejecutar Me.hwnd, "open", m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoURLPlantillasHTML_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
   
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub CorreosAutomaticos_AfterUpdate()
    Dim blnActivar As Boolean
    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If Nz(Me.CorreosAutomaticos, "") = "Sí" Then
        blnActivar = True
    Else
        blnActivar = False
    End If
    ActivarCamposCorreo blnActivar
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
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()

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
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String

   
    On Error GoTo errores
    Me.Caption = m_TituloFormulario
    Set m_ConfAlInicio = constructor.getConfiguracion(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ConfAlInicio Is Nothing Then
        Me.AutocancelacionMARGA = "Sí"
        Me.AutocancelacionPreMARGA = "Sí"
        ActivarCamposCorreo False
        Exit Function
    End If
    With m_ConfAlInicio
        If .CorreosAutomaticos <> "" Then
            Me.CorreosAutomaticos = .CorreosAutomaticos
            ActivarCamposCorreo True
        Else
            ActivarCamposCorreo False
        End If
        If .BuzonSeguridad <> "" Then
            Me.BuzonSeguridad = .BuzonSeguridad
        End If
        If .DiasParaRecordatorioExcel1 <> "" Then
            Me.DiasParaRecordatorioExcel1 = .DiasParaRecordatorioExcel1
        End If
        If .EmailDirectorSeguridad <> "" Then
            Me.EmailDirectorSeguridad = .EmailDirectorSeguridad
        End If
        If .DiasParaRecordatorioExcel2 <> "" Then
            Me.DiasParaRecordatorioExcel2 = .DiasParaRecordatorioExcel2
        End If
        
        If .DiasParaRecordatorioRellenoMarga1 <> "" Then
            Me.DiasParaRecordatorioRellenoMarga1 = .DiasParaRecordatorioRellenoMarga1
        End If
        If .DiasParaRecordatorioRellenoMarga2 <> "" Then
            Me.DiasParaRecordatorioRellenoMarga2 = .DiasParaRecordatorioRellenoMarga2
        End If
        If .DiasCancelacionMARGA <> "" Then
            Me.DiasCancelacionMARGA = .DiasCancelacionMARGA
        End If
        If .AutocancelacionMARGA <> "" Then
            Me.AutocancelacionMARGA = .AutocancelacionMARGA
        End If
        If .DiasCancelacionPreMARGA <> "" Then
            Me.DiasCancelacionPreMARGA = .DiasCancelacionPreMARGA
        End If
        If .AutocancelacionPreMARGA <> "" Then
            Me.AutocancelacionPreMARGA = .AutocancelacionPreMARGA
        End If
        If .VersionPlantillasHTML <> "" Then
            Me.VersionPlantillasHTML = .VersionPlantillasHTML
        End If
        If .VersionPlantillasEXCEL <> "" Then
            Me.VersionPlantillasEXCEL = .VersionPlantillasEXCEL
        End If
        
    End With
    If Application.TempVars("ActivadoCorreoAutomatico") = "No" Then
        Me.lblConfCorreosAutomaticos.Visible = True
        Me.lblConfCorreosAutomaticos.Caption = "Aún no implementado correos automáticos"
        
    Else
        Me.lblConfCorreosAutomaticos.Visible = False
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
Private Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    On Error GoTo errores
    
    If m_ConfAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_ConfAlInicio
        If .CorreosAutomaticos <> Nz(Me.CorreosAutomaticos, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .BuzonSeguridad <> Nz(Me.BuzonSeguridad, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .DiasParaRecordatorioExcel1 <> Nz(Me.DiasParaRecordatorioExcel1, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .EmailDirectorSeguridad <> Nz(Me.EmailDirectorSeguridad, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .DiasParaRecordatorioExcel2 <> Nz(Me.DiasParaRecordatorioExcel2, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        If .DiasParaRecordatorioRellenoMarga1 <> Nz(Me.DiasParaRecordatorioRellenoMarga1, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .DiasParaRecordatorioRellenoMarga2 <> Nz(Me.DiasParaRecordatorioRellenoMarga2, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .AutocancelacionPreMARGA <> Nz(Me.AutocancelacionPreMARGA, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .AutocancelacionMARGA <> Nz(Me.AutocancelacionMARGA, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .DiasCancelacionPreMARGA <> Nz(Me.DiasCancelacionPreMARGA, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        If .DiasCancelacionMARGA <> Nz(Me.DiasCancelacionMARGA, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .VersionPlantillasHTML <> Nz(Me.VersionPlantillasHTML, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .VersionPlantillasEXCEL <> Nz(Me.VersionPlantillasEXCEL, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HahabidoCambios ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function RellenarObjeto(Optional ByRef p_Error As String) As String

   
    On Error GoTo errores
    
    Set m_Conf = New Configuracion
    
    With m_Conf
        If Not m_ConfAlInicio Is Nothing Then
            m_Conf.id = m_ConfAlInicio.id
        End If
        .CorreosAutomaticos = Nz(Me.CorreosAutomaticos, "")
        .BuzonSeguridad = Nz(Me.BuzonSeguridad, "")
        .DiasParaRecordatorioExcel1 = Nz(Me.DiasParaRecordatorioExcel1, "")
        .EmailDirectorSeguridad = Nz(Me.EmailDirectorSeguridad, "")
        .DiasParaRecordatorioExcel2 = Nz(Me.DiasParaRecordatorioExcel2, "")
        
        .DiasParaRecordatorioRellenoMarga1 = Nz(Me.DiasParaRecordatorioRellenoMarga1, "")
        .DiasParaRecordatorioRellenoMarga2 = Nz(Me.DiasParaRecordatorioRellenoMarga2, "")
        .AutocancelacionPreMARGA = Nz(Me.AutocancelacionPreMARGA, "")
        .AutocancelacionMARGA = Nz(Me.AutocancelacionMARGA, "")
        .DiasCancelacionPreMARGA = Nz(Me.DiasCancelacionPreMARGA, "")
        .DiasCancelacionMARGA = Nz(Me.DiasCancelacionMARGA, "")
        .VersionPlantillasHTML = Nz(Me.VersionPlantillasHTML, "")
        .VersionPlantillasEXCEL = Nz(Me.VersionPlantillasEXCEL, "")
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjeto ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function


