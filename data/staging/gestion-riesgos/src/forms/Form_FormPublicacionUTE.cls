Option Compare Database
Option Explicit
Private m_Error As String

Private Sub ComandoEvidenciaDeAcuerdo_Click()
    Dim m_URL As String
    Dim m_Anexo As Anexo
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
     m_Error = ""
    If m_ObjEdicionActiva.Proyecto.EnUTE = "No" Then
        m_Error = "No es un proyecto en UTE"
        Err.Raise 1000
    End If
    
    m_URL = Seleccionar(True, "Seleccione el archivo a anexar", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Set m_Anexo = New Anexo
    With m_Anexo
        .IDEdicion = m_ObjEdicionActiva.IDEdicion
        .EvidenciaUTE = "Sí"
        .EvidenciaSuministrador = "No"
        .Titulo = "ED " & m_ObjEdicionActiva.Edicion & " Acuerdo tratamiento riesgos con miembros de UTE"
        .Registrar m_ObjEdicionActiva, m_URL, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_ObjAnexoEvicenciaUTE = Constructor.getAnexo(m_Anexo.IDAnexo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjEdicionActiva.AnexoEvidenciaUTE = m_Anexo
    InvalidarPublicabilidadPorCambioEvidenciaEdicion m_ObjEdicionActiva.IDEdicion, m_Error
    If m_Error <> "" Then
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
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEvidenciaDeAcuerdo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    AvanceCerrar
    
End Sub

Private Sub ComandoEvidenciaDeAcuerdoQuitar_Click()
    On Error GoTo errores

    
    m_Error = ""
    If m_ObjAnexoEvicenciaUTE Is Nothing Then
        m_Error = "No parece tener registrado un anexo de evidencia de acuerdo con la UTE"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea eliminar el anexo de la evidencia de acuerdo con la UTE?", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar anexo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjAnexoEvicenciaUTE.EliminarAnexo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    InvalidarPublicabilidadPorCambioEvidenciaEdicion m_ObjEdicionActiva.IDEdicion, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjEdicionActiva = Constructor.getEdicion(m_ObjEdicionActiva.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjAnexoEvicenciaUTE = m_ObjEdicionActiva.AnexoEvidenciaUTE
    m_URLInforme = ""
    
    
    Me.ComandoEvidenciaDeAcuerdo.SetFocus
    EstablecerDatos
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEvidenciaDeAcuerdoQuitar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me

    
    
    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "No hay ninguna edición activa"
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
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub


Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    On Error Resume Next
    If blnPermitidoEscribir = True Then
        Me.ComandoEvidenciaDeAcuerdo.Enabled = True
    Else
        Me.ComandoEvidenciaDeAcuerdo.Enabled = False
        
    End If
    If m_ObjAnexoEvicenciaUTE Is Nothing Then
        Me.ImagenAnexo.Visible = False
        Me.ComandoEvidenciaDeAcuerdoQuitar.Visible = False
    Else
        Me.ImagenAnexo.Visible = True
        If blnPermitidoEscribir = True Then
            Me.ComandoEvidenciaDeAcuerdoQuitar.Visible = True
            Me.ComandoEvidenciaDeAcuerdoQuitar.Enabled = True
        End If
    End If
End Function



Private Sub ImagenAnexo_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not fso.FileExists(m_ObjAnexoEvicenciaUTE.URLAnexo) Then
        m_Error = "El archivo no es alcanzable en estos momentos: " & vbNewLine & m_ObjAnexoEvicenciaUTE.URLAnexo
        Err.Raise 1000
    End If
    If FicheroAbierto(m_ObjAnexoEvicenciaUTE.URLAnexo) Then
        m_Error = "Ya tiene el fichero abierto"
        Err.Raise 1000
    End If
    AbrirEnLocal m_ObjAnexoEvicenciaUTE.URLAnexo, Me.hWnd, m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenAnexo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
