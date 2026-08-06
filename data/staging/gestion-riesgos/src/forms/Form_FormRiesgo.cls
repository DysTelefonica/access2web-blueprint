Option Compare Database
Option Explicit
Public m_ObjColEmpresas As Scripting.Dictionary
Public m_ObjColNombres As Scripting.Dictionary

Public Event RiesgoNuevo()
Public Event RiesgoEditado()

Public Event cambioEnPriorizacion(p_Priorizacion As String)
Public m_PermitidoEditar As EnumSiNo
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_Error As String



Private Sub ComandoActualizar_Click()
    Dim m_Form As Form
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    RefrescarDetalleRiesgoManualBridge p_Error:=m_Error
    If m_Error <> "" Then Err.Raise 1000

    Set m_Form = Me.FrmDetalle.Form
    ' HR3e (2026-07-01): reemplazar las 6 llamadas cross-form
    ' Form_FormRiesgoX.EstablecerDatos por el dispatcher
    ' modFormCoordinationHelper.Coord_RefrescarSubformDetalleRiesgo.
    Call modFormCoordinationHelper.Coord_RefrescarSubformDetalleRiesgo( _
        p_SubformName:=m_Form.Name, _
        p_Error:=m_Error)
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
        m_Error = "Al ComandoActualizar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Function RefrescarDetalleRiesgoManualBridge(Optional ByRef p_Error As String) As String
    ' Thin UI wrapper: delega la logica de refresco al helper
    ' modRiesgoDetalleRefrescoHelper. El form solo valida que haya
    ' un riesgo activo y devuelve el IDRiesgo (compatibilidad con
    ' call site en ComandoActualizar_Click).
    Dim m_IdRiesgo As String

    On Error GoTo errores

    p_Error = ""
    RefrescarDetalleRiesgoManualBridge = ""

    If m_ObjRiesgoActivo Is Nothing Then
        p_Error = "No hay riesgo activo para refrescar"
        Err.Raise 1000
    End If

    m_IdRiesgo = CStr(m_ObjRiesgoActivo.IDRiesgo)

    RefrescarDetalleRiesgoManualBridge = modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual( _
        m_IdRiesgo, m_ObjRiesgoActivo, m_ObjRiesgoAlInicio, , p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método RefrescarDetalleRiesgoManualBridge ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RefrescarDetalleRiesgoManualBridge", p_Error
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


Private Sub ComandoSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoVerEdicion_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    'Set m_ObjEdicionActiva = m_ObjRiesgoActivo.Edicion
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.OpenForm "FormRiesgosGestion"
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerEdicion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    'Me.ControlDeNavegación0.Tabs(0).Enabled = False
    Me.InsideHeight = 9060
    Me.InsideWidth = 13980
    
    
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
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
' ----------------------------------------------------------------------------
' Public Sub Refrescar (HR4 UI adapter)
'   Hard rule 4 fix: el proxy de modFormCoordinationHelper llama este Sub
'   (UI adapter tiny, delega a Private EstablecerDatos).
' ----------------------------------------------------------------------------
Public Sub Refrescar(Optional ByRef p_Error As String)
    EstablecerDatos p_Error
End Sub

' ----------------------------------------------------------------------------
' Private Function EstablecerDatos (HR4 fix — antes era Public)
'   Logica de negocio extraida a modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular.
'   Esta funcion queda como thin coordinator: pure compute via helper +
'   UI wiring local (Me.* controls) + persistencia del estado del form.
' ----------------------------------------------------------------------------
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String

    Dim m_Resultado As RiesgoEstablecerDatos_Resultado
    Dim m_ObjRiesgoAlInicio_Local As Object
    Dim m_ObjEdicionActiva_Local As Object
    Dim m_DbForm As DAO.Database
    Dim m_CalcError As String
    Dim m_EstadoRiesgo As EnumRiesgoEstado

    On Error GoTo errores

    ' Snapshot editable state del form
    Set m_ObjEdicionActiva_Local = m_ObjEdicionActiva

    ' --- Pure compute via helper (parent: p_EsGemelo=False) ---
    Set m_DbForm = getdb(m_CalcError)
    If m_CalcError <> "" Then
        p_Error = m_CalcError
        Err.Raise 1000
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_ObjRiesgoActivo, m_ObjEdicionActiva_Local, m_EsAlta, _
        m_PermitidoEditar, False, _
        m_Resultado, m_ObjRiesgoAlInicio_Local, _
        m_DbForm, m_CalcError
    If m_CalcError <> "" Then
        p_Error = m_CalcError
        Err.Raise 1000
    End If

    ' --- Persistir estado del compute ---
    Set m_ObjEdicionActiva = m_ObjEdicionActiva_Local
    m_PermitidoEditar = m_Resultado.PermitidoEditar
    Set m_ObjRiesgoAlInicio = m_ObjRiesgoAlInicio_Local

    ' --- Alta path: crear nuevo riesgo (mutacion especifica del form) ---
    If m_EsAlta = EnumSiNo.Sí Then
        Set m_ObjRiesgoActivo = New riesgo
        m_ObjRiesgoActivo.IDEdicion = m_ObjEdicionActiva.IDEdicion
        Set m_ObjRiesgoAlInicio = Nothing
    End If

    ' Compute m_EstadoRiesgo local (form-level state)
    m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum

    ' --- UI wiring (form-specific) ---
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado

    If m_PermitidoEditar = EnumSiNo.Sí Then
        Me.AllowEdits = True
        Me.ComandoGrabar.Enabled = True
    Else
        Me.AllowEdits = False
        Me.ComandoGrabar.Enabled = False
    End If

    Me.lblEstadoRiesgo.Visible = m_Resultado.LblEstadoRiesgoVisible
    Me.NavGeneral.Enabled = (m_EsAlta = EnumSiNo.Sí)
    Me.NavRetirado.Enabled = m_Resultado.NavRetiradoEnabled
    Me.NavMaterializado.Enabled = m_Resultado.NavMaterializadoEnabled
    Me.lblTitulo.Caption = m_Resultado.Titulo
    Me.NavGeneral.NavigationTargetName = m_Resultado.NavGeneralTarget
    blnEdicionActiva = m_Resultado.blnEdicionActiva

    ' --- Post-UI hooks (form-specific) ---
    ActualizarEtiquetaEstadoRiesgo , p_Error
    If p_Error <> "" Then Err.Raise 1000

    EstablecerColorPestañasRiesgo p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000

    ' HR3 closure: leer nodo + ver descripcion via modFormCoordinationHelper
    ' (no acceso directo a otro form .cls / Forms("X").Controls).
    If FormularioAbierto("FormRiesgosGestion") Then
        Set m_NodoSeleccionado = modFormCoordinationHelper.Coord_GetSelectedNode()
        m_VerDescripcion = modFormCoordinationHelper.Coord_GetComboVerDescripcion()
    End If

    Me.FrmDetalle.SourceObject = Me.NavGeneral.NavigationTargetName

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

    
Private Sub ComandoGrabar_Click()
    
    Dim m_HaHabidoCambios As Boolean
    Dim m_EstadoRiesgo As String
    Dim m_NombreIcono As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    RellenarObjeto
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_HaHabidoCambios = False Then
        m_Error = "Ningún cambio que guardar"
        Err.Raise 1000
    End If
    
    With m_ObjRiesgoActivo
        '.Estado = m_ObjRiesgoActivo.EstadoCalculadoTexto
        If m_EsAlta = EnumSiNo.Sí Then
            .IDEdicion = m_ObjEdicionActiva.IDEdicion
        End If
        
        .Registrar EnumSiNo.Sí, m_ObjRiesgoAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_EstadoRiesgo = .EstadoEnum
        ActualizarEtiquetaEstadoRiesgo
    End With
    If m_EstadoRiesgo = EnumRiesgoEstado.Incompleto Then
        Me.NavRetirado.Enabled = False
        Me.NavMaterializado.Enabled = False
    Else
        Me.NavRetirado.Enabled = True
        Me.NavMaterializado.Enabled = True
    End If
    
    If m_ObjRiesgoAlInicio Is Nothing Then
       RaiseEvent RiesgoNuevo
        
    Else

        ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_NodoSeleccionado, p_VerDescripcion:=m_VerDescripcion, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
       
        
        ' HR3 closure: el subform FormRiesgosGestionRiesgo se refresca
        ' via modFormCoordinationHelper (no llamada directa a otro form .cls).
        ' Coord_RefreshRiesgoGestionRiesgo ya verifica que el subform sea
        ' el adecuado (Coord no toca FormDetalle.SourceObject: el caller
        ' valida, Coord delega a modRiesgoEstablecerDatosHelper).
        If FormularioAbierto("FormRiesgosGestion") Then
            If Forms("FormRiesgosGestion").FormDetalle.SourceObject = "FormRiesgosGestionRiesgo" Then
                modFormCoordinationHelper.Coord_RefreshRiesgoGestionRiesgo
            End If

        End If
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    AvanceCerrar
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



Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    ' Issue #38: pure comparison logic delegated to RiesgoChangeDetector.
    ' Form keeps only the form-specific guards (Nothing snapshot = first load).
    Dim m_Detector As RiesgoChangeDetector
    Dim m_DetectorError As String

    On Error GoTo errores

    If m_ObjRiesgoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If

    Set m_Detector = New RiesgoChangeDetector
    HaHabidoCambios = m_Detector.HasChanges(m_ObjRiesgoAlInicio, m_ObjRiesgoActivo, m_DetectorError)
    If m_DetectorError <> vbNullString Then
        p_Error = "HaHabidoCambios: " & m_DetectorError
    End If
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Private Sub NavMaterializado_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
   
End Sub

Private Sub NavMitigacion_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
   

End Sub

Public Function RellenarObjeto(Optional ByRef p_Error As String) As String
    Dim m_Form As Form
    On Error GoTo errores
    Set m_Form = Me.FrmDetalle.Form
    With m_ObjRiesgoActivo
        If Me.FrmDetalle.SourceObject = "FormRiesgoDefinicion" Then
            .Origen = Nz(m_Form.Origen, "")
            .Priorizacion = Nz(m_Form.Priorizacion, "")
            .Descripcion = Nz(m_Form.Descripcion, "")
            .CausaRaiz = Nz(m_Form.CausaRaiz, "")
            .FechaDetectado = Nz(m_Form.FechaDetectado, "")
            .EntidadDetecta = Nz(m_Form.EntidadDetecta, "")
            .DetectadoPor = Nz(m_Form.DetectadoPor, "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoDefinicionNoBiblioteca" Then
            .Origen = Nz(m_Form.Origen, "")
            .Priorizacion = Nz(m_Form.Priorizacion, "")
            .Descripcion = Nz(m_Form.Descripcion, "")
            .FechaDetectado = Nz(m_Form.FechaDetectado, "")
            .EntidadDetecta = Nz(m_Form.EntidadDetecta, "")
            .DetectadoPor = Nz(m_Form.DetectadoPor, "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoPlazoCosteCalidad" Then
            .Plazo = Nz(m_Form.ListaPlazo.Column(0), "")
            .Calidad = Nz(m_Form.ListaCalidad.Column(0), "")
            .Coste = Nz(m_Form.ListaCoste.Column(0), "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoVulnerabilidad" Then
            .Vulnerabilidad = Nz(m_Form.ListaVulnerabilidad.Column(0), "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoMitigacion" Then
            .Mitigacion = Nz(m_Form.ListaMitigacion.Column(0), "")
            If m_Form.JustificacionAceptacionRiesgo.Visible Then
                .JustificacionAceptacionRiesgo = Nz(m_Form.JustificacionAceptacionRiesgo, "")
            Else
                .JustificacionAceptacionRiesgo = ""
            End If
            If m_Form.FechaJustificacionAceptacionRiesgo.Visible Then
                .FechaJustificacionAceptacionRiesgo = Nz(m_Form.FechaJustificacionAceptacionRiesgo, "")
            Else
                .FechaJustificacionAceptacionRiesgo = ""
            End If
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoRetirado" Then
            .FechaRetirado = Nz(m_Form.FechaRetirado, "")
            .JustificacionRetiroRiesgo = Nz(m_Form.JustificacionRetiroRiesgo, "")
            .FechaJustificacionRetiroRiesgo = Nz(m_Form.FechaJustificacionRetiroRiesgo, "")
            .FechaRechazoRetiroPorCalidad = Nz(m_Form.FechaRechazoRetiroPorCalidad, "")
            
       
        End If
        
    End With
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjeto ha devuelto el error: " & Err.description
    End If
End Function



Private Sub NavPlazoCosteCalidad_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
   
End Sub

Private Sub NavRetirado_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
    
End Sub

Private Sub NavVulnerabilidad_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
    
End Sub
