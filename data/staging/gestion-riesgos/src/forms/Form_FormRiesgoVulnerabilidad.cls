Option Compare Database
Option Explicit
Private m_Error As String



Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
   
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
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_valor As String
    Dim i As Long
    Dim blnPermitidoEditar As Boolean
'    On Error GoTo errores
    On Error Resume Next
    Me.AllowEdits = True
    
    With m_ObjRiesgoActivo
        If .Vulnerabilidad <> "" Then
            Set lst = Me.ListaVulnerabilidad
            For i = 0 To lst.ListCount - 1
                m_valor = Nz(lst.Column(0, i), "")
                If m_valor = .Vulnerabilidad Then
                    lst.Selected(i) = True
                    Exit For
                End If
            Next
            EstablecerControlCombo Me.ListaVulnerabilidad, EnumSiNo.Sí
        Else
            EstablecerControlCombo Me.ListaVulnerabilidad, EnumSiNo.No
        End If
        If m_ObjRiesgoActivo.Edicion.EsActivo = EnumSiNo.No Then
            blnPermitidoEditar = False
        Else
            If m_ObjRiesgoActivo.Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
                blnPermitidoEditar = True
            
            Else
                blnPermitidoEditar = True
            End If
        End If
    End With


End Function

Private Sub ListaVulnerabilidad_Click()
    
    On Error Resume Next
    If Me.AllowEdits = True Then
        
        If Nz(Me.ListaVulnerabilidad.Column(0), "") = "" Then
            EstablecerControlCombo Me.ListaVulnerabilidad, EnumSiNo.No
        Else
            EstablecerControlCombo Me.ListaVulnerabilidad, EnumSiNo.Sí
        End If
        EstablecerColorPestañasRiesgo
    End If
    
End Sub