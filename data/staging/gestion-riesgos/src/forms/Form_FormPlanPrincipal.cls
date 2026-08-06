Option Compare Database
Option Explicit
Private m_EsMitigacion As String
Private m_Error As String
Public Event PlanNuevo(p_Plan As Object)

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
        
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_EsMitigacion = "Sí" Then
        Set m_ObjPMActivo = New pm
        m_ObjPMActivo.IDRiesgo = m_ObjRiesgoActivo.IDRiesgo
        m_ObjPMActivo.DisparadorDelPlan = Nz(Me.DisparadorDelPlan, "")
            
        m_ObjPMActivo.Registrar , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        RaiseEvent PlanNuevo(m_ObjPMActivo)
    Else
        Set m_ObjPCActivo = New pc
        m_ObjPCActivo.IDRiesgo = m_ObjRiesgoActivo.IDRiesgo
        m_ObjPCActivo.DisparadorDelPlan = Nz(Me.DisparadorDelPlan, "")
            
        m_ObjPCActivo.Registrar , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        RaiseEvent PlanNuevo(m_ObjPCActivo)
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
     If m_EsMitigacion = "Sí" Then
        Set m_ObjPMActivo = getPM(m_ObjPMActivo.IDMitigacion, m_Error)
    Else
        Set m_ObjPCActivo = getPC(m_ObjPCActivo.IDContingencia, m_Error)
    End If
   
   
End Sub
Private Sub ComandoSalir_Click()
    
    
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub DisparadorDelPlan_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
    
End Sub

Private Sub DisparadorDelPlan_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub

Private Sub Form_Load()

    
    Dim m_Titulo As String
    Dim m_CodRiesgo As String
    Dim m_Cod As String
    Dim blnPermitidoEditar As Boolean
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    m_EsMitigacion = Nz(Me.openArgs(), "")
    If m_EsMitigacion <> "Sí" And m_EsMitigacion <> "No" Then
        m_Error = "No se ha abierto el formulario con los parámetros insuficientes"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo Is Nothing Then
        m_Error = "No hay ningún riesgo activo"
        Err.Raise 1000
    End If
    Set m_Edicion = m_ObjRiesgoActivo.Edicion
    With m_ObjRiesgoActivo
        Me.Descripcion = .Descripcion
        Me.CausaRaiz = .CausaRaiz
        m_CodRiesgo = .CodigoRiesgo
        If m_EsMitigacion = "Sí" Then
            m_Titulo = "ALTA DE PLAN DE MITIGACIÓN PARA " & m_CodRiesgo
            Me.lblDisparador.Caption = "Denominación"
        Else
            m_Titulo = "ALTA DE plan de contingencia PARA " & m_CodRiesgo
            Me.lblDisparador.Caption = "Disparador"
        End If
    End With
    
    Me.lblTitulo1.Caption = m_Titulo
    
     Me.Caption = m_ObjEntorno.TituloUsuarioConectado
     If m_Edicion.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        If m_Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
            blnPermitidoEditar = True
        
        Else
            blnPermitidoEditar = True
        End If
    End If
    
    If blnPermitidoEditar = True Then
        Me.ComandoGrabar.Enabled = True
        Me.AllowEdits = True
    Else
        Me.ComandoGrabar.Enabled = False
        Me.AllowEdits = False
    End If
     
     
    RellenarBordeCamposObligatorios Me
    
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