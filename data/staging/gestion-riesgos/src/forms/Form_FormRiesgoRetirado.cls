Option Compare Database
Option Explicit

Private blnRiesgoActivo As Boolean
Private m_EstadoRiesgo As EnumRiesgoEstado
Private blnPermitidoEditar As Boolean
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



Private Sub ComandoAceptar_Click()
    Dim m_Fecha As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_Fecha = Nz(Me.FechaRetirado, "")
    If Not IsDate(m_Fecha) Then
        m_Error = "Se ha de indicar la fecha"
        Err.Raise 1000
    End If
    With m_ObjRiesgoActivo
        
        If Me.ComandoAceptar.Caption = "Aceptar" Then
            .RetiroAprobar m_Fecha, m_EstadoRiesgo, m_Error
        Else
            .RetiroAprobacionQuitar m_EstadoRiesgo, m_Error
        End If
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    Set m_ObjRiesgoActivo = GetCachedRiesgo(p_IDRiesgo:=m_ObjRiesgoActivo.IDRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    EstablecerDatosCalidad m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Call modFormCoordinationHelper.Coord_RefreshRiesgoGestionRiesgo(m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Call modFormCoordinationHelper.Coord_RefreshRiesgo(m_Error)
    If m_Error <> "" Then Err.Raise 1000
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "calidad", m_Error
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
        m_Error = "Al ComandoAceptar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRechazar_Click()
     Dim m_Fecha As String
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_Fecha = Nz(Me.FechaRechazoRetiroPorCalidad, "")
    If Not IsDate(m_Fecha) Then
        m_Error = "Se ha de indicar la fecha de rechazo"
        Err.Raise 1000
    End If
    With m_ObjRiesgoActivo
       
        If Me.ComandoRechazar.Caption = "Rechazar" Then
            .RetiroRechazo m_EstadoRiesgo, m_Fecha, m_Error
        Else
            .RetiroRechazoQuitar m_EstadoRiesgo, m_Error
        End If
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    Set m_ObjRiesgoActivo = GetCachedRiesgo(p_IDRiesgo:=m_ObjRiesgoActivo.IDRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    EstablecerDatosCalidad m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Call modFormCoordinationHelper.Coord_RefreshRiesgoGestionRiesgo(m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Call modFormCoordinationHelper.Coord_RefreshRiesgo(m_Error)
    If m_Error <> "" Then Err.Raise 1000
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "calidad", m_Error
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
        m_Error = "Al ComandoRechazar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRetiroRegistrar_Click()
    
    Dim m_NombreIcono As String
    Dim m_Fecha As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
   
    If Nz(Me.JustificacionRetiroRiesgo, "") = "" Then
        m_Error = "Se ha de indicar alguna justificación"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo.IDRiesgo = "" Then
        m_Error = "Primero ha de registrar el resto de campos del Riesgo"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Materializado Then
        pregunta = MostrarMsgBoxSiNoEsTest("El riesgo actualmente está materializado." & vbNewLine & _
                    "¿Desea proponer la retirada del riesgo?", vbExclamation + vbYesNo + vbDefaultButton2, "Propuesta de retirada")
        If pregunta <> vbYes Then
            Exit Sub
        End If
        m_ObjRiesgoActivo.MaterializacionQuitarRegistrar m_ObjRiesgoActivo.EstadoEnum, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_Fecha = Nz(Me.FechaRetirado, "")
    If Not IsDate(m_Fecha) Then
        m_Error = "Primero ha de indicar la fecha de retirada"
        Err.Raise 1000
    End If
    m_ObjRiesgoActivo.RetiroRegistrar Me.JustificacionRetiroRiesgo, m_Fecha, m_ObjRiesgoActivo.EstadoEnum, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    Set m_ObjRiesgoAlInicio = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
    If IsDate(m_ObjRiesgoAlInicio.FechaJustificacionRetiroRiesgo) Then
        Me.FechaJustificacionRetiroRiesgo = m_ObjRiesgoAlInicio.FechaJustificacionRetiroRiesgo
        Me.ComandoEliminar.Visible = True
        Me.ComandoEliminar.Enabled = True
    End If
    ActualizarEtiquetaEstadoRiesgo
    ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_NodoSeleccionado, p_VerDescripcion:=m_VerDescripcion, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "retiro", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatosCalidad
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("Ha sido registrado la propuesta de Retiro del Riesgo y ha salido un correo electrónico a Calidad para su aprobación", vbInformation, "Riesgo propuesto para Retirar")
    
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRetiroRegistrar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminar_Click()
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim m_Estado As EnumRiesgoEstado
    On Error GoTo errores
    
    
    m_Error = ""
    
    pregunta = MostrarMsgBoxSiNoEsTest("¿Desea quitar la retirada del Riesgo?", vbExclamation + vbYesNo + vbDefaultButton2, "Quitar retirada")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjRiesgoActivo.IDRiesgo = "" Then
        m_Error = "Primero ha de registrar el resto de campos del Riesgo"
        Err.Raise 1000
    End If
    m_Estado = m_ObjRiesgoActivo.ESTADOCalculado
    m_RiesgoEnRetirada = RiesgoEnRetirada(m_Estado)
    If m_RiesgoEnRetirada = EnumSiNo.No Then
        m_Error = "El retiro del riesgo no es sin justificar, sin visar o rechazado"
        Err.Raise 1000
    End If
    
    
    
    m_ObjRiesgoActivo.RetiroRegistrarQuitar m_Estado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    Set m_ObjRiesgoAlInicio = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
    ActualizarEtiquetaEstadoRiesgo
    ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_NodoSeleccionado, p_VerDescripcion:=m_VerDescripcion, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "retiro", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    EstablecerDatos
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("Ha sido registrado la habilitación del riesgo", vbInformation, "Riesgo habilitado")
    
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

Private Sub ComandoFRetirado_Click()
    Dim m_Fecha As String
    Dim m_RiesgoEnAceptacion As EnumSiNo
    On Error GoTo errores
    m_Error = ""
    If blnPermitidoEditar = False Then
       Exit Sub
    End If
    m_RiesgoEnAceptacion = RiesgoEnAceptacion(m_ObjRiesgoActivo.EstadoEnum)
    If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
        Me.FechaRetirado = Null
        m_Error = "El riesgo está aceptado o en fase de aceptación"
        Err.Raise 1000
    End If
    m_Fecha = Nz(InputBox("Introduzca la fecha de retiro del riesgo", "Fecha retirado", Date), "")
    If Not IsDate(m_Fecha) Then
        Exit Sub
    End If
    
    If Not m_ObjRiesgoAlInicio Is Nothing Then
        If m_ObjRiesgoAlInicio.FechaRetirado = m_Fecha Then
            Exit Sub
        End If
        HacerVisiblesCamposOcultos
        RellenarCamposOcultos m_ObjRiesgoAlInicio
        If blnPermitidoEditar = True Then
            If m_ObjRiesgoAlInicio.EstadoEnum = EnumRiesgoEstado.Retirado Then
                Me.JustificacionRetiroRiesgo.Enabled = False
                Me.ComandoRetiroRegistrar.Enabled = False
                Me.ComandoEliminar.Enabled = True
                Me.ComandoEliminar.Visible = True
            Else
                Me.JustificacionRetiroRiesgo.Enabled = True
                Me.ComandoRetiroRegistrar.Enabled = True
                Me.ComandoEliminar.Visible = False
                
            End If
        Else
            Me.JustificacionRetiroRiesgo.Enabled = False
            Me.ComandoRetiroRegistrar.Enabled = False
            Me.ComandoEliminar.Enabled = False
            Me.ComandoEliminar.Visible = False
        End If
        Me.FechaRetirado = m_Fecha
        Exit Sub
    End If
    Me.FechaRetirado = m_Fecha
    HacerVisiblesCamposOcultos
    Me.JustificacionRetiroRiesgo.Enabled = True
    Me.ComandoRetiroRegistrar.Enabled = True
    Me.ComandoEliminar.Visible = False
    Me.ComandoEliminar.Enabled = False
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoFRetirado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
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
    EstablecerDatosCalidad m_Error
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
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""

    If m_ObjRiesgoAlInicio Is Nothing Then
        HacerInVisiblesCamposOcultos
        Me.AllowEdits = False
        Exit Function
    End If

    Call EstablecerDatos_CargarPermisos(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_AplicarEstado(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_RenderizarControles(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
    End If
End Function

Private Sub EstablecerDatos_CargarPermisos(ByRef p_Error As String)
    On Error GoTo errores
    Dim m_Edicion As Edicion
    ' HR3e (2026-07-01): eliminar comentario dead de cross-form call.
    Set m_NodoSeleccionado = modFormCoordinationHelper.Coord_GetSelectedNode(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    m_VerDescripcion = modFormCoordinationHelper.Coord_GetComboVerDescripcion(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Set m_Edicion = m_ObjRiesgoActivo.Edicion
    If m_Edicion.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        blnPermitidoEditar = True
    End If
    If m_ObjRiesgoActivo.EsActivo = EnumSiNo.Sí Then
        blnRiesgoActivo = True
    Else
        blnRiesgoActivo = False
    End If
    If blnPermitidoEditar = True Then
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Retirado Then
        Me.FechaRetirado = m_ObjRiesgoAlInicio.FechaRetirado
        Me.ComandoFRetirado.Enabled = False
        If blnPermitidoEditar = True Then
            Me.ComandoEliminar.Enabled = True
        End If
        HacerVisiblesCamposOcultos
        RellenarCamposOcultos m_ObjRiesgoAlInicio
        Exit Sub
    End If
    If IsDate(m_ObjRiesgoAlInicio.FechaRetirado) Then
        Me.FechaRetirado = m_ObjRiesgoAlInicio.FechaRetirado
        HacerVisiblesCamposOcultos
        RellenarCamposOcultos m_ObjRiesgoAlInicio
        If blnPermitidoEditar = True Then
            Me.JustificacionRetiroRiesgo.Enabled = True
            Me.ComandoRetiroRegistrar.Enabled = True
            Me.ComandoEliminar.Enabled = True
        Else
            Me.ComandoFRetirado.Enabled = False
            Me.JustificacionRetiroRiesgo.Enabled = False
            Me.ComandoRetiroRegistrar.Enabled = False
            Me.ComandoEliminar.Enabled = False
            Me.ComandoEliminar.Visible = False
        End If
        Exit Sub
    End If
    HacerInVisiblesCamposOcultos
    Me.FechaRetirado = Null
    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    ' Renderizado de controles adicionales (placeholder para futuras extensiones)
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.description
End Sub

Private Function EstablecerDatosCalidad(Optional ByRef p_Error As String) As String
    

    Dim m_Builder As RiesgoEstadoCalidadViewStateBuilder
    Dim m_State As RiesgoEstadoCalidadViewState
    
    On Error GoTo errores
    
    m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum
    Set m_Builder = New RiesgoEstadoCalidadViewStateBuilder
    Set m_State = m_Builder.BuildRetiradaCalidad( _
        p_EstadoRiesgo:=m_EstadoRiesgo, _
        p_EsTecnico:=EsTecnico, _
        p_RiesgoActivo:=blnRiesgoActivo, _
        p_HayRiesgoInicial:=(Not m_ObjRiesgoAlInicio Is Nothing), _
        p_FechaRetiradoRegistrada:=RiesgoInicialTieneFechaRetirado(), _
        p_Error:=p_Error)
    AplicarEstadoBotonesCalidad m_State
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatosCalidad ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Function RiesgoInicialTieneFechaRetirado() As Boolean
    If m_ObjRiesgoAlInicio Is Nothing Then
        RiesgoInicialTieneFechaRetirado = False
        Exit Function
    End If

    RiesgoInicialTieneFechaRetirado = IsDate(m_ObjRiesgoAlInicio.FechaRetirado)
End Function

Private Sub AplicarEstadoBotonesCalidad(ByVal p_State As RiesgoEstadoCalidadViewState)
    Me.ComandoAceptar.Visible = p_State.ComandoAceptarVisible
    Me.ComandoRechazar.Visible = p_State.ComandoRechazarVisible

    If Not p_State.HasButtonDecision Then
        Exit Sub
    End If

    Me.ComandoAceptar.Caption = p_State.ComandoAceptarCaption
    Me.ComandoRechazar.Caption = p_State.ComandoRechazarCaption
    Me.ComandoAceptar.Enabled = p_State.ComandoAceptarEnabled
    Me.ComandoRechazar.Enabled = p_State.ComandoRechazarEnabled
End Sub
Private Sub HacerVisiblesCamposOcultos()

    On Error Resume Next
    Me.lblJustificacionRetiroRiesgo.Visible = True
    Me.ComandoRetiroRegistrar.Visible = True
    Me.ComandoEliminar.Visible = True
    
    Me.JustificacionRetiroRiesgo.Visible = True
    Me.lblFechaJustificacionRetiroRiesgo.Visible = True
    Me.lblFechaAprobacionRetiroPorCalidad.Visible = True
    Me.lblFechaRechazoRetiroPorCalidad.Visible = True
    Me.FechaJustificacionRetiroRiesgo.Visible = True
    Me.FechaAprobacionRetiroPorCalidad.Visible = True
    Me.FechaRechazoRetiroPorCalidad.Visible = True
End Sub
Private Sub HacerInVisiblesCamposOcultos()

    On Error Resume Next
    Me.lblJustificacionRetiroRiesgo.Visible = False
    Me.ComandoRetiroRegistrar.Visible = False
   
    Me.ComandoEliminar.Visible = False
    
    Me.JustificacionRetiroRiesgo.Visible = False
    Me.lblFechaJustificacionRetiroRiesgo.Visible = False
    Me.lblFechaAprobacionRetiroPorCalidad.Visible = False
    Me.lblFechaRechazoRetiroPorCalidad.Visible = False
    Me.FechaJustificacionRetiroRiesgo.Visible = False
    Me.FechaAprobacionRetiroPorCalidad.Visible = False
    Me.FechaRechazoRetiroPorCalidad.Visible = False
End Sub

Private Sub RellenarCamposOcultos(Optional p_Riesgo As riesgo)

    On Error Resume Next
    With p_Riesgo
        If .JustificacionRetiroRiesgo <> "" Then
            Me.JustificacionRetiroRiesgo = .JustificacionRetiroRiesgo
        End If
        If IsDate(.FechaJustificacionRetiroRiesgo) Then
            Me.FechaJustificacionRetiroRiesgo = .FechaJustificacionRetiroRiesgo
        End If
        If IsDate(.FechaAprobacionRetiroPorCalidad) Then
            Me.FechaAprobacionRetiroPorCalidad = .FechaAprobacionRetiroPorCalidad
        End If
        If IsDate(.FechaRechazoRetiroPorCalidad) Then
            Me.FechaRechazoRetiroPorCalidad = .FechaRechazoRetiroPorCalidad
        End If
    End With
    
End Sub

