Attribute VB_Name = "Form_FormUsuario02DatosObservacion"
Option Compare Database
Option Explicit

Private m_Error As String


Private m_ObjObservacionAlInicio As Observacion
Public Event AltaRegistro(p_Tipo As String)
Public Event EdicionRegistro(p_Tipo As String)



Private Sub CmdSalir_Click()
    
    Dim m_HaHabidoCambios As Boolean
    
    On Error GoTo errores
    m_Error = ""
    
    
    If m_ObjObservacionActiva.ParaAlta = EnumSiNo.No Then
        m_HaHabidoCambios = HaHabidoCambios(m_ObjObservacionAlInicio, Me, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_HaHabidoCambios Then
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            pregunta = MsgBox("¿Desea guardar los cambios antes de cerrar el formulario?", vbExclamation + vbYesNo + vbDefaultButton1, "Guardar Cambios")
            If pregunta = vbYes Then
                m_ObjObservacionActiva.Registrar m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                
                RaiseEvent EdicionRegistro(m_ObjObservacionActiva.TIPO)
            End If
        End If
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar salir se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoGrabar_Click()
    
    Dim m_HaHabidoCambios As Boolean
    
    On Error GoTo errores
    
    m_Error = ""
    RellenarObjetoConDatos m_ObjObservacionActiva, Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    If m_ObjObservacionAlInicio Is Nothing Then
        m_ObjObservacionActiva.ParaAlta = EnumSiNo.Sí
        m_ObjObservacionActiva.ID = m_ObjUsuarioActivo.ID
        m_ObjObservacionActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If Me.TIPO = "General" Then
            Set m_ObjUsuarioActivo.ColObservaciones = Nothing
        ElseIf Me.TIPO = "Curso" Then
            Set m_ObjUsuarioActivo.ColObservacionesCurso = Nothing
        End If
        
        RaiseEvent AltaRegistro(m_ObjObservacionActiva.TIPO)
    Else
        m_HaHabidoCambios = HaHabidoCambios(m_ObjObservacionAlInicio, Me, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_HaHabidoCambios Then
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_Error = "No hay cambios que guardar"
            Err.Raise 1000
        End If
        m_ObjObservacionActiva.ParaAlta = EnumSiNo.No
        m_ObjObservacionActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
        RaiseEvent EdicionRegistro(m_ObjObservacionActiva.TIPO)
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    Dim m_Titulo As String
    Dim m_Argumento As String
    On Error GoTo errores
    
    m_Error = ""
    AjustarTamaño Me
   
    m_Argumento = Nz(Me.OpenArgs(), "")
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "No se conoce el usuario activo"
        Err.Raise 1000
    End If
    If m_ObjObservacionActiva Is Nothing Then
        If Me.AllowEdits = False Then
            m_Error = "El Sistema está en modo Sólo lectura"
            Err.Raise 1000
        End If
        If m_Argumento <> "General" And m_Argumento <> "Curso" Then
            m_Error = "Formulario abierto con parámetros insuficientes"
            Err.Raise 1000
        End If
        
        Set m_ObjObservacionActiva = New Observacion
        m_ObjObservacionActiva.ID = m_ObjUsuarioActivo.ID
        m_ObjObservacionActiva.ParaAlta = EnumSiNo.Sí
        Me.TIPO = m_Argumento
        Me.Fecha = Date
        Me.TIPO.Enabled = False
        Me.Observacion.SetFocus
        m_Titulo = "ALTA DE OBSERVACIÓN " & UCase(m_Argumento) & ": " & m_ObjUsuarioActivo.NombreCompleto
        Me.ComandoGrabar.Enabled = True
    Else
        If Me.AllowEdits = True Then
            m_Titulo = "EDICIÓN DE OBSERVACIÓN " & UCase(m_Argumento) & ": " & m_ObjUsuarioActivo.NombreCompleto
            Me.ComandoGrabar.Enabled = True
        Else
            m_Titulo = "DETALLE DE OBSERVACIÓN" & UCase(m_Argumento) & ": " & m_ObjUsuarioActivo.NombreCompleto
            Me.ComandoGrabar.Enabled = False
        End If
        Set m_ObjObservacionAlInicio = getObservacion(m_ObjObservacionActiva.IDObservacion, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
        DatosDelObjetoAlFormulario m_ObjObservacionAlInicio, Me, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Me.Observacion.SetFocus
    End If
    Me.lblTitulo.Caption = m_Titulo
    Me.Caption = m_Titulo
    CopiarEstiloBoton Me.ComandoGrabar, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.cmdSalir, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.EncabezadoDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor
    Me.PieDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
