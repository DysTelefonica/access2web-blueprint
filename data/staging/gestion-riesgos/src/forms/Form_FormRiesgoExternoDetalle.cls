Option Compare Database
Option Explicit
Private m_Error As String

Private m_CodRiesgoBiblioteca As String
Private m_CausaRaiz As String
Private m_RiesgoExtAlinicio As RiesgoExterno
Public Event Registrado()
Private WithEvents m_FormRiesgosBibliotecaGestion As Form_FormRiesgosBibliotecaGestion


Private Function Registrar(Optional ByRef p_Error As String) As String
    
    Dim m_ObjCorreo As correo
    Dim m_FechaRef As Date
    Dim m_LineaAlFinal As String
    On Error GoTo errores
    
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    With m_ObjRiesgoExtActivo
        
        .Trasladar = Nz(Me.Trasladar, "")
        .Descripcion = Nz(Me.Descripcion, "")
        If m_ObjRiesgoExtActivo.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            .CausaRaiz = Nz(Me.CausaRaiz, "")
        End If
        .MotivoNoIntegrado = Nz(Me.MotivoNoIntegrado, "")
        If Nz(Me.MotivoNoIntegrado, "") <> "" And Not IsDate(.FechaMotivo) Then
            .FechaMotivo = CStr(m_FechaRef)
        End If
        .CodRiesgoBiblioteca = Nz(Me.CodRiesgoBiblioteca, "")
        .Registrar m_RiesgoExtAlinicio, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí And _
            .Edicion.Proyecto.ParaInformeAvisos <> "No" Then
            If .RiesgoParaRetipificar = EnumSiNo.Sí Then
                Set m_ObjRiesgoActivo = .riesgo
                If m_ObjRiesgoActivo.CodRiesgoBiblioteca <> m_CodRiesgoBiblioteca Or _
                    m_ObjRiesgoActivo.CausaRaiz <> m_CausaRaiz Then
                    EnviarCorreoRiesgoRequiereRetipificacion m_ObjRiesgoActivo, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    
                End If
            End If
           
        End If
    End With
    
    
    
    
    RaiseEvent Registrado
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.description
    End If
End Function

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
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGrabar_Click()
    Dim m_HaHabidoCambios As Boolean
    On Error GoTo errores
    
'    If blnPermitidoEscribir = False Then
'        Exit Sub
'    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = False Then
        m_Error = "Sin cambios que guardar"
        Err.Raise 1000
    End If
    
    Registrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    
    On Error GoTo errores
    m_Error = ""
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSalir_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSeleccionarRiesgoDeBiblioteca_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        DoCmd.Close acForm, "FormRiesgosBibliotecaGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormRiesgosBibliotecaGestion"
    If FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        Set m_FormRiesgosBibliotecaGestion = Forms("FormRiesgosBibliotecaGestion")
        With m_FormRiesgosBibliotecaGestion
            .ComandoElegir.Visible = True
            .ComandoElegir.Enabled = True
            
        End With
    End If
    
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSeleccionarRiesgoDeBiblioteca_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim blnPermitidoEditar As Boolean
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
   
    Me.Caption = Me.lblTitulo1.Caption
    
    If m_ObjRiesgoExtActivo Is Nothing Then
        m_Error = "No se ha podido determinar el riesgo externo activo para editar"
        Err.Raise 1000
    End If
    
    Set m_RiesgoExtAlinicio = Constructor.getRiesgoExterno(m_ObjRiesgoExtActivo.IDRiesgoExt, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoExtAlinicio.Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
        blnPermitidoEditar = True
    Else
        blnPermitidoEditar = False
    End If
    
    If m_RiesgoExtAlinicio.Edicion.EsActivo = EnumSiNo.Sí And blnPermitidoEditar = True Then
        Me.AllowEdits = True
        Me.ComandoGrabar.Enabled = True
        If m_RiesgoExtAlinicio.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            Me.ComandoSeleccionarRiesgoDeBiblioteca.Enabled = True
            
            Me.Descripcion.Enabled = False
            Me.Descripcion.Locked = False
            Me.CausaRaiz.Enabled = True
            Me.CausaRaiz.Locked = False
        ElseIf m_RiesgoExtAlinicio.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            Me.ComandoSeleccionarRiesgoDeBiblioteca.Enabled = False
            
            Me.Descripcion.Enabled = True
            Me.Descripcion.Locked = False
            Me.CausaRaiz.Enabled = False
            Me.CausaRaiz.Locked = False
        Else
            m_Error = "No se sabe si la gestión de riesgos requiere riesgos de la biblioteca"
            Err.Raise 1000
        End If
    Else
        Me.ComandoGrabar.Enabled = False
        Me.AllowEdits = False
        blnPermitidoEscribir = False
        Me.ComandoSeleccionarRiesgoDeBiblioteca.Enabled = False
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
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    If m_RiesgoExtAlinicio Is Nothing Then
        Exit Function
    End If
    With m_RiesgoExtAlinicio
        Me.CodRiesgo = .CodRiesgo
        
        Me.Trasladar = .Trasladar
        Me.CodRiesgoBiblioteca = .CodRiesgoBiblioteca
        Me.Descripcion = .Descripcion
        Me.MotivoNoIntegrado = .MotivoNoIntegrado
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            Me.CausaRaiz = .CausaRaiz
        End If
        m_CodRiesgoBiblioteca = .CodRiesgoBiblioteca
        m_CausaRaiz = .CausaRaiz
    End With
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & Err.description
    End If
                  
End Function

Private Sub Trasladar_AfterUpdate()
    On Error Resume Next
    If blnPermitidoEscribir = False Then
        Exit Sub
    End If
    If Nz(Me.Trasladar, "") = "Sí" Then
        Me.MotivoNoIntegrado.Enabled = False
        Me.MotivoNoIntegrado.Locked = True
        Me.MotivoNoIntegrado.value = Null
    Else
        
        Me.MotivoNoIntegrado.Enabled = True
        Me.MotivoNoIntegrado.Locked = False
    End If
    
End Sub

Private Sub Trasladar_Exit(Cancel As Integer)
    On Error Resume Next
    If blnPermitidoEscribir = False Then
        Exit Sub
    End If
    If Nz(Me.Trasladar, "") = "Sí" Then
        Me.MotivoNoIntegrado.Enabled = False
        Me.MotivoNoIntegrado.Locked = True
        Me.MotivoNoIntegrado.value = Null
    Else
        
        Me.MotivoNoIntegrado.Enabled = True
        Me.MotivoNoIntegrado.Locked = False
    End If
End Sub


Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean

    On Error GoTo errores
    If m_RiesgoExtAlinicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_RiesgoExtAlinicio
        If .Trasladar <> Nz(Me.Trasladar, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Descripcion <> Nz(Me.Descripcion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .MotivoNoIntegrado <> Nz(Me.MotivoNoIntegrado, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .CodRiesgoBiblioteca <> Nz(Me.CodRiesgoBiblioteca, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            If .CausaRaiz <> Nz(Me.CausaRaiz, "") Then
                HaHabidoCambios = True
                Exit Function
            End If
        End If
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & Err.description
    End If
End Function

Private Sub m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido(m_ObjRiesgoBiblioteca As RiesgoBiblioteca)
    On Error GoTo errores
     VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    With m_ObjRiesgoBiblioteca
        Me.CodRiesgoBiblioteca = .codigo
        
        If .RiesgoParaRetipificar = EnumSiNo.No Then
            Me.Descripcion = .Descripcion
            Me.Descripcion.Locked = False
            Me.Descripcion.Enabled = True
            Me.CausaRaiz.Locked = False
            Me.CausaRaiz.Enabled = True
        Else
            Me.Descripcion = Null
            Me.Descripcion.Locked = False
            Me.Descripcion.Enabled = False
            Me.CausaRaiz.Locked = False
            Me.CausaRaiz.Enabled = False
        End If
    End With
     VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
