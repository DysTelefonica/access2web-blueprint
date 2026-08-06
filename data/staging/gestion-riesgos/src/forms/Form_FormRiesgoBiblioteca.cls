Option Compare Database
Option Explicit
Private m_Error As String
Private m_ObjRiesgoBibliotecaAlIncio As RiesgoBiblioteca

Public Event AltaRiesgoBiblioteca(m_objRiesgoDeBiblioteca As RiesgoBiblioteca)
Public Event EdicionRiesgoBiblioteca(m_objRiesgoDeBiblioteca As RiesgoBiblioteca)


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

Private Sub ComandoGrabar_Click()
    
    Dim m_HaHabidoCambios As Boolean
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = False Then
        m_Error = "No ha habido cambios que guardar"
        Err.Raise 1000
    End If
    RellenarObjetoConDatosFormulario m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjRiesgoBibliotecaActivo.Registrar m_ObjRiesgoBibliotecaAlIncio, EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjRiesgoBibliotecaAlIncio Is Nothing Then
        RaiseEvent AltaRiesgoBiblioteca(m_ObjRiesgoBibliotecaActivo)
    Else
        RaiseEvent EdicionRiesgoBiblioteca(m_ObjRiesgoBibliotecaActivo)
    End If
    If m_ObjRiesgoBibliotecaAlIncio Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("¿Desea dar de alta algún riesgo tipo máSí", vbExclamation + vbYesNo + vbDefaultButton2, "Alta riesgo tipo")
        If pregunta <> vbYes Then
            DoCmd.Close acForm, Me.Name, acSaveNo
            Exit Sub
        End If
       
        Me.Familia = Null
        Me.Descripcion = Null
        Me.Activo = Null
        Set m_ObjRiesgoBibliotecaActivo = Nothing
        Exit Sub
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub Form_Load()
    
    On Error GoTo errores
    
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    Dim m_Titulo As String
    
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        Me.AllowEdits = False
        Me.ComandoGrabar.Enabled = False
    Else
        
        Me.AllowEdits = True
        Me.ComandoGrabar.Enabled = True
    End If
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjRiesgoBibliotecaActivo Is Nothing Then
        m_Titulo = "Alta de riesgo de biblioteca"
    Else
        If Me.AllowEdits = False Then
            m_Titulo = "Detalle de riesgo de biblioteca"
        Else
            m_Titulo = "Edición de riesgo de biblioteca"
        End If
        Set m_ObjRiesgoBibliotecaAlIncio = Constructor.getRiesgoBiblioteca(m_ObjRiesgoBibliotecaActivo.IDRiesgoTipo, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjRiesgoBibliotecaAlIncio.RiesgoParaRetipificar = EnumSiNo.Sí Then
            Me.AllowEdits = False
            Me.ComandoGrabar.Enabled = False
        End If
    End If
    Me.lblTitulo1.Caption = m_Titulo
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    
    If m_ObjRiesgoBibliotecaAlIncio Is Nothing Then
        Exit Function
    End If
    
    With m_ObjRiesgoBibliotecaAlIncio
        
        If .Familia <> "" Then
            Me.Familia = .Familia
        End If
        If .Activo <> "" Then
            Me.Activo = .Activo
        End If
        If .Descripcion <> "" Then
            Me.Descripcion = .Descripcion
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    On Error GoTo errores
    If m_ObjRiesgoBibliotecaAlIncio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_ObjRiesgoBibliotecaAlIncio
        
        If .Familia <> Nz(Me.Familia, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Activo <> Nz(Me.Activo, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Descripcion <> Nz(Me.Descripcion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Function RellenarObjetoConDatosFormulario( _
                                                Optional ByRef p_Error As String _
                                                ) As String
   
    
    
    On Error GoTo errores
    
    If m_ObjRiesgoBibliotecaActivo Is Nothing Then
        Set m_ObjRiesgoBibliotecaActivo = New RiesgoBiblioteca
    End If
    With m_ObjRiesgoBibliotecaActivo
        .Familia = Nz(Me.Familia, "")
        .Activo = Nz(Me.Activo, "")
        .Descripcion = Nz(Me.Descripcion, "")
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjetoConDatosFormulario ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim cmb As ComboBox
    Dim m_Familia As Variant
    
    
    On Error GoTo errores
    
    Set cmb = Me.Familia
    cmb.RowSource = ""
    If Not m_ObjEntorno.ColRiesgosBibliotecasFamilias Is Nothing Then
        
        For Each m_Familia In m_ObjEntorno.ColRiesgosBibliotecasFamilias
            cmb.AddItem m_Familia
            
        Next
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
