Option Compare Database
Option Explicit

Private m_EsMitigacion As String
Private WithEvents m_FormAccion As Form_FormPlanAcciones

Private m_ObjPMAlInicio As pm
Private m_ObjPCAlInicio As pc
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_Error As String
' ==========================================
' MSG BOX HELPER (per access-vba-tdd -0 - no interactive UI in tests)
' p_PromptResult: -1 (default) = show MsgBox; 0 = skip MsgBox (test mode)
' ==========================================
Private Function MostrarMsgBoxSiNoEsTest(ByVal p_Message As String, _
                                       ByVal p_Style As VbMsgBoxStyle, _
                                       ByVal p_Title As String, _
                                       Optional ByVal p_PromptResult As Long = -1) As Long
    If p_PromptResult = 0 Then
        MostrarMsgBoxSiNoEsTest = vbOK
    Else
        MostrarMsgBoxSiNoEsTest = MsgBox(p_Message, p_Style, p_Title)
    End If
End Function

Private Sub ComandoAltaAccion_Click()
    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormPlanAcciones") Then
        DoCmd.Close acForm, "FormPlanAcciones", acSaveNo
    End If
    Set m_ObjPMAccionActiva = Nothing
    Set m_ObjPCAccionActiva = Nothing
    DoCmd.OpenForm "FormPlanAcciones", , , , , , m_EsMitigacion
    If FormularioAbierto("FormPlanAcciones") Then
        Set m_FormAccion = Forms("FormPlanAcciones")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaAccion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminar_Click()
    Dim m_NodoActual As MSComctlLib.Node
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMActivo Is Nothing Then
            m_Error = "No se sabe el plan de mitigación"
            Err.Raise 1000
        End If
    Else
        If m_ObjPCActivo Is Nothing Then
            m_Error = "No se sabe el plan de contingencia"
            Err.Raise 1000
        End If
    End If
    
    If modFormCoordinationHelper.Coord_GetPermitidoEditar() = False Then
        m_Error = "Usuario no autorizado a esa operación"
        Err.Raise 1000
    End If
    Set m_NodoActual = modFormCoordinationHelper.Coord_ArbolSelectedItem()
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("¿Desea realmente borrar el plan seleccionado?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de un plan")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_EsMitigacion = "Sí" Then
        m_ObjPMActivo.Borrar m_Error
        Set m_ObjRiesgoActivo.ColPMs = Nothing
    Else
        m_ObjPCActivo.Borrar m_Error
        Set m_ObjRiesgoActivo.ColPCs = Nothing
    End If
    
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("Plan borrado con éxito", vbInformation, "Plan borrado")
    ' HR3c closure: post-delete refresh patron (parent select + NodeClick + Remove)
    ' encapsulado en modFormCoordinationHelper.Coord_ArbolEliminarNodoConRefresco.
    modFormCoordinationHelper.Coord_ArbolEliminarNodoConRefresco m_NodoActual.key, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    


    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGrabar_Click()
       
    Dim m_HaHabidoCambios As Boolean
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If modFormCoordinationHelper.Coord_GetPermitidoEditar() = False Then
        Exit Sub
    End If
    If m_EsMitigacion = "Sí" Then
        m_ObjPMActivo.DisparadorDelPlan = Nz(Me.DisparadorDelPlan, "")
    Else
        m_ObjPCActivo.DisparadorDelPlan = Nz(Me.DisparadorDelPlan, "")
    End If
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = False Then
        m_Error = "No hay nada que grabar"
        Err.Raise 1000
    End If
    If m_EsMitigacion = "Sí" Then
        m_ObjPMActivo.Registrar m_ObjPMAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    Else
        
        m_ObjPCActivo.Registrar m_ObjPCAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End If
    modFormCoordinationHelper.Coord_RefrescarArbolRiesgosScope "plan", m_Error
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
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
   
     If m_EsMitigacion = "Sí" Then
        Set m_ObjPMActivo = getPM(m_ObjPMActivo.IDMitigacion, m_Error)
    Else
        Set m_ObjPCActivo = getPC(m_ObjPCActivo.IDContingencia, m_Error)
    End If
   
   
End Sub



Private Sub DisparadorDelPlan_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub DisparadorDelPlan_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
End Sub

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
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    On Error GoTo errores
    p_Error = ""

    Call EstablecerDatos_CargarPermisos(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_AplicarEstado(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_RenderizarControles(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub EstablecerDatos_CargarPermisos(ByRef p_Error As String)
    On Error GoTo errores
    If modFormCoordinationHelper.Coord_GetEsMitigacion() = "" Then
        p_Error = "No se sabe si es plan de mitigación o de contingencia"
        Err.Raise 1000
    End If
    If modFormCoordinationHelper.Coord_GetEsMitigacion() = CStr(EnumSiNo.Sí) Then
        m_EsMitigacion = "Sí"
    Else
        m_EsMitigacion = "No"
    End If
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMActivo Is Nothing Then
            p_Error = "Este formulario es para editar"
            Err.Raise 1000
        End If
        Set m_ObjPMAlInicio = getPM(m_ObjPMActivo.IDMitigacion, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        Me.lblDisparador.Caption = "Denominación"
    Else
        If m_ObjPCActivo Is Nothing Then
            p_Error = "Este formulario es para editar"
            Err.Raise 1000
        End If
        Set m_ObjPCAlInicio = getPC(m_ObjPCActivo.IDContingencia, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        Me.lblDisparador.Caption = "Disparador"
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.Description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    If modFormCoordinationHelper.Coord_GetPermitidoEditar() = True Then
        Me.ComandoEliminar.Enabled = True
        Me.ComandoAltaAccion.Enabled = True
        Me.ComandoGrabar.Enabled = True
        Me.AllowEdits = True
    Else
        Me.ComandoEliminar.Enabled = False
        Me.ComandoAltaAccion.Enabled = False
        Me.ComandoGrabar.Enabled = False
        Me.AllowEdits = False
    End If
    If m_EsMitigacion = "Sí" Then
        With m_ObjPMAlInicio
            Me.lblTitulo.Caption = "Plan de mitigación " & " " & .CodMitigacion & " " & "riesgo " & .riesgo.CodigoRiesgo
            Me.DisparadorDelPlan = .DisparadorDelPlan
            Me.EstadoCompletoPlanEstado = .EstadoPlanParaEtiqueta
            Me.EstadoCompletoFechaActivacion = .FechaActivacionParaEtiqueta
            Me.EstadoCompletoFechaDesactivacion = .FechaDesactivacionParaEtiqueta
        End With
    Else
        With m_ObjPCAlInicio
            Me.lblTitulo.Caption = "Plan de contingencia " & " " & .CodContingencia & " " & "riesgo " & .riesgo.CodigoRiesgo
            Me.DisparadorDelPlan = .DisparadorDelPlan
            Me.EstadoCompletoPlanEstado = .EstadoPlanParaEtiqueta
            Me.EstadoCompletoFechaActivacion = .FechaActivacionParaEtiqueta
            Me.EstadoCompletoFechaDesactivacion = .FechaDesactivacionParaEtiqueta
        End With
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.Description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    ' HR3 closure: leer nodo + ver descripcion via modFormCoordinationHelper
    ' (no acceso directo a otro form .cls / Forms("X").Controls).
    If FormularioAbierto("FormRiesgosGestion") Then
        Set m_NodoSeleccionado = modFormCoordinationHelper.Coord_GetSelectedNode()
        m_VerDescripcion = modFormCoordinationHelper.Coord_GetComboVerDescripcion()
    End If
    RellenarBordeCamposObligatorios Me
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.Description
End Sub

Private Function HaHabidoCambios( _
                                    Optional ByRef p_Error As String _
                                    ) As Boolean
    Dim detector As RiesgoPlanActionChangeDetector

    On Error GoTo errores
    Set detector = New RiesgoPlanActionChangeDetector

    If m_EsMitigacion = "Sí" Then
        HaHabidoCambios = detector.HasPMChanges(m_ObjPMAlInicio, m_ObjPMActivo, p_Error)
    Else
        HaHabidoCambios = detector.HasPCChanges(m_ObjPCAlInicio, m_ObjPCActivo, p_Error)
    End If
    If p_Error <> "" Then
        Exit Function
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Sub m_FormAccion_AccionNueva(m_PMAccion As Object)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    If Left(m_NodoSeleccionado.key, 2) = "PM" Then
        Me.EstadoCompletoPlanEstado = m_ObjPMActivo.EstadoPlanParaEtiqueta
        Me.EstadoCompletoFechaActivacion = m_ObjPMActivo.FechaActivacionParaEtiqueta
        Me.EstadoCompletoFechaDesactivacion = m_ObjPMActivo.FechaDesactivacionParaEtiqueta
        
        
    ElseIf Left(m_NodoSeleccionado.key, 2) = "PC" Then
        Me.EstadoCompletoPlanEstado = m_ObjPCActivo.EstadoPlanParaEtiqueta
        Me.EstadoCompletoFechaActivacion = m_ObjPCActivo.FechaActivacionParaEtiqueta
        Me.EstadoCompletoFechaDesactivacion = m_ObjPCActivo.FechaDesactivacionParaEtiqueta
        
    End If
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    modFormCoordinationHelper.Coord_CargarArbolAccion m_NodoSeleccionado, m_PMAccion, m_Error
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
        m_Error = "Al m_FormAccion_AccionNueva se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
