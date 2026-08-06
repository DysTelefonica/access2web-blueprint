Option Compare Database
Option Explicit

Private m_URLInforme As String
Private WithEvents m_FormRiesgo As Form_FormRiesgo
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

Private Sub ComandoAbrirProyecto_Click()
    On Error GoTo errores
    
    
    
    m_Error = ""
    pregunta = MostrarMsgBoxSiNoEsTest("¿Desea realmente quitar la última publicación y convertir esa edición e la edición activa?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Abrir gestión de riesgos")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_ObjProyectoActivo.Abrir m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyectoActivo = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.cmdSalir.SetFocus
    ' HR3e (2026-07-01): reemplazar acceso directo a
    ' Form_FormRiesgosGestion.ComandoActualizarContador_Click por helper
    ' modFormCoordinationHelper.
    Call modFormCoordinationHelper.Coord_ComandoActualizarContadorRiesgosGestion(m_Error)
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
        m_Error = "Al ComandoAbrirProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoAltaRiesgo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    If Form_FormRiesgosGestion.blnPermitidoEditar = False Then
        m_Error = "No está permitido dar de alta más riesgos"
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormRiesgo") Then
        DoCmd.Close acForm, "FormRiesgo", acSaveNo
    End If
    Set m_ObjRiesgoActivo = Nothing
    Set m_ObjRiesgoAlInicio = Nothing
    m_EsAlta = EnumSiNo.Sí
    DoCmd.OpenForm "FormRiesgo"
    If FormularioAbierto("FormRiesgo") Then
        Set m_FormRiesgo = Forms("FormRiesgo")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaRiesgo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Private Sub ComandoGenerarInforme_Click()
    Dim url As String
    Dim alcance As EnumControlCambiosAlcance
    Dim valorAlcance As String
    
    
    On Error GoTo errores
    
    If m_ObjEdicionActiva.TodosLosRiesgosPriorizados = EnumSiNo.No Then
        pregunta = MostrarMsgBoxSiNoEsTest("Para poder generar el informe ha de tener todos sus riesgos priorizados", vbExclamation, "Error")
        Exit Sub
    End If
    
   
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    valorAlcance = modFormInteractionHelper.FormInteraction_ObtenerTextoControl(Me, "cboControlCambiosAlcance", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    alcance = ControlCambios_ResolverAlcanceDesdeSelector(valorAlcance, EnumControlCambiosAlcanceResumen3)

    url = GenerarInformePublicacion( _
            p_Edicion:=m_ObjEdicionActiva, _
            p_FechaCierre:="", _
            p_FechaPublicacion:="", _
            p_Error:=m_Error, _
            p_ControlCambiosAlcance:=alcance)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' NOTA: La función GenerarInformeEdicionHTML ya tiene internamente un 'Ejecutar p_hWnd, "open"...'
    ' Si notas que el documento se abre dos veces seguidas, simplemente comenta o borra la siguiente línea:
    Ejecutar Me.hWnd, "open", url, "", "", 1
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    AvanceCerrar
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarInforme_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoPublicacion_Click()
    
    Dim m_NombreFormulario As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If EsTecnico = EnumSiNo.Sí Then
        m_NombreFormulario = "FormPublicacionTecnico"
    Else
        m_NombreFormulario = "FormPublicacionCalidad"
    End If
    If FormularioAbierto(m_NombreFormulario) Then
        DoCmd.Close acForm, m_NombreFormulario, acSaveNo
    End If
    DoCmd.OpenForm m_NombreFormulario
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminarEdicion_Click()
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjEdicionUltima As Edicion
    Dim m_IDProyecto As String
    Dim m_ObjProyecto As Proyecto
    Dim m_ObjEdicionAnterior As Edicion
    Dim m_ObjDocumento As Documento
    Dim m_IdEdicionBorrada As String
    Dim m_DeleteDecisionService As GestionRiesgosEdicionDeleteDecisionService
    Dim m_ExisteArchivoInformeAnterior As Boolean
    Dim m_InformeAnteriorAbierto As Boolean
    Dim m_HayEdicionActiva As Boolean
    Dim m_EsPrimeraEdicion As EnumSiNo
        
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_DeleteDecisionService = New GestionRiesgosEdicionDeleteDecisionService
    m_HayEdicionActiva = (Not (m_ObjEdicionActiva Is Nothing))
    If m_HayEdicionActiva Then
        m_EsPrimeraEdicion = m_ObjEdicionActiva.EsPrimeraEdicion
    End If

    If Not m_DeleteDecisionService.PuedeIniciarBorrado( _
            m_HayEdicionActiva, _
            m_EsPrimeraEdicion, _
            EsTecnico, _
            m_Error) Then
        Err.Raise 1000
    End If

    'Set m_ObjEdicionActiva = Constructor.getEdicion(m_ObjEdicionActiva.IDEdicion)
    m_IDProyecto = m_ObjEdicionActiva.IDProyecto
    
    Set m_Col = m_ObjEdicionActiva.colCambiosConEdicionAnterior
    m_Error = m_ObjEdicionActiva.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_DeleteDecisionService.DebeConfirmarCambiosConEdicionAnterior(Not (m_Col Is Nothing)) Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MostrarMsgBoxSiNoEsTest("Se han producido cambios con la edición anterior" & vbNewLine & _
                        "¿Desea seguir eliminando la edición?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar edición")
        If pregunta <> vbYes Then
            Exit Sub
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("¿Desea realmente borrar completamente la edición seleccionada?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar edición")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not m_ObjEdicionActiva.EdicionAnterior Is Nothing Then
        
        Set m_ObjDocumento = m_ObjEdicionActiva.EdicionAnterior.DocumentoObj
        If Not m_ObjDocumento Is Nothing Then
            m_ExisteArchivoInformeAnterior = fso.FileExists(m_ObjDocumento.URLAdjunto)
            If m_ExisteArchivoInformeAnterior Then
                m_InformeAnteriorAbierto = FicheroAbierto(m_ObjDocumento.URLAdjunto)
            End If
            If m_DeleteDecisionService.InformeAnteriorBloqueaBorrado( _
                    True, _
                    m_ExisteArchivoInformeAnterior, _
                    m_InformeAnteriorAbierto, _
                    m_Error) Then
                Err.Raise 1000
            End If
        End If
    End If
    
    
    m_IdEdicionBorrada = m_ObjEdicionActiva.IDEdicion
    m_ObjEdicionActiva.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    InvalidarCacheEdicion m_IdEdicionBorrada
    EstablecerTareasCalidad EnumSiNo.No, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_ObjProyectoActivo = GetCachedProyecto(m_IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjProyectoActivo.EstablecerFechaMaxPublicacion m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjEdicionActiva.EstablecerFechaMaxPublicacion m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyectoActivo = GetCachedProyecto(m_IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Form_FormRiesgosGestion.m_EdicionAlAbrir = Nothing
    modFormCoordinationHelper.Coord_RefreshRiesgosGestion m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    ' HR4-sister-b (HR3d follow-up #6 - 2026-07-01):
    '   Era cross-form call a Form_FormProyectosGestion.ActualizarProyecto
    '   + .ListaFiltrados_Click. Hard rule 3: form .cls MUST NOT call
    '   other form .cls. Reemplazado por modProyectosGestionHelper.
    If FormularioAbierto("FormProyectosGestion") Then
        modProyectosGestionHelper.ProyectosGestion_ActualizarProyecto_ActualizarLinea _
            Forms("FormProyectosGestion"), m_ObjProyecto, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        modProyectosGestionHelper.ProyectosGestion_ListaFiltrados_Click_Run _
            Forms("FormProyectosGestion"), m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    'DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarEdicion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRiesgoPriorizacion_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If FormularioAbierto("FormRiesgosEstablecerPrioridades") Then
        DoCmd.Close acForm, "FormRiesgosEstablecerPrioridades", acSaveNo
    End If
    DoCmd.OpenForm "FormRiesgosEstablecerPrioridades"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRiesgoPriorizacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    

    Me.AllowEdits = True
    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "No hay ninguna edición activa"
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If

End Sub


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""

    Me.ComandoPublicacion.Enabled = True
    Me.ComandoEliminarEdicion.Enabled = False
    Me.ComandoAltaRiesgo.Enabled = False
    Me.ComandoRiesgoPriorizacion.Enabled = False
    Me.ComandoGenerarInforme.Enabled = False
    Me.ComandoVerAnexos.Enabled = True
    Me.ComandoAbrirProyecto.Visible = False

    Call EstablecerDatos_CargarPermisos(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_AplicarEstado(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_RenderizarControles(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub EstablecerDatos_CargarPermisos(ByRef p_Error As String)
    On Error GoTo errores
    Dim m_FechaVisible As String
    Dim m_Proyecto As Proyecto

    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No se conoce la edición activa"
        Err.Raise 1000
    End If
    Set m_Proyecto = GetCachedProyecto(m_ObjEdicionActiva.IDProyecto, p_Error)
    If p_Error <> "" Then Exit Sub

    ' Punto 08 / Fase A — spec Calidad 2026-07-08:
    '   con FechaPublicacion -> label "Fecha Edición" + value FechaPublicacion
    '   sin FechaPublicacion  -> label "Fecha publicación anterior" + value
    '     FechaEdicion system (que == FechaPublicacionEdicionAnterior para
    '     ediciones subsiguientes o == project start date para la primera)
    '
    ' Decisión de modelado: la fecha de creación de la edición activa
    ' (FechaEdicion system) coincide con la fecha de publicación de la
    ' edición anterior. Por tanto el value a mostrar es el mismo en
    ' ambos casos (FechaEdicion system == FechaPublicacionEdicionAnterior).
    If IsDate(m_ObjEdicionActiva.FechaPublicacion) Then
        m_FechaVisible = Format(m_ObjEdicionActiva.FechaPublicacion, "dd/mm/yyyy")
        ' Label dinámico: "Fecha Edición" (con publicación)
        Me.Controls("Etiqueta76").Caption = "Fecha Edición"
    Else
        m_FechaVisible = Format(m_ObjEdicionActiva.FechaEdicion, "dd/mm/yyyy")
        ' Label dinámico: "Fecha publicación anterior" (sin publicación)
        Me.Controls("Etiqueta76").Caption = "Fecha publicación anterior"
    End If
    Me.FechaEdicion = m_FechaVisible

    Me.CodigoDocumento = m_Proyecto.CodigoDocumento
    Me.Edicion = m_ObjEdicionActiva.Edicion
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.Description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    Dim m_FechaProxEdicion As String

    With m_ObjEdicionActiva
        If .Proyecto.FechaCierre <> "" Then
            Me.lblTitulo.Caption = .Proyecto.NombreParaNodo & " ed. " & .Edicion & " (No activa cerrada)"
        Else
            Me.lblTitulo.Caption = .Proyecto.NombreParaNodo & " ed. " & .Edicion
        End If
        Me.FechaPrevistaCierre = .Proyecto.FechaPrevistaCierreCalculada
        If .EsActivo = EnumSiNo.Sí Then
            m_FechaProxEdicion = .FechaMaxProximaPublicacionCalculado
            If IsDate(m_FechaProxEdicion) Then
                Me.FechaMaxProximaPublicacion = Format(m_FechaProxEdicion, "dd/mm/yyyy")
            End If
            If m_FechaProxEdicion <> .FechaMaxProximaPublicacion Then
                .FechaMaxProximaPublicacionGrabar
            End If
        Else
            Me.FechaMaxProximaPublicacion = "N/A"
        End If
        If Form_FormRiesgosGestion.blnPermitidoEditar = True Then
            If EsTecnico <> EnumSiNo.Sí Then
                If m_ObjEdicionActiva.EsPrimeraEdicion = EnumSiNo.No Then
                    Me.ComandoEliminarEdicion.Enabled = True
                End If
                If IsDate(.Proyecto.FechaCierre) Then
                    Me.ComandoAbrirProyecto.Visible = True
                End If
            End If
            Me.ComandoAltaRiesgo.Enabled = True
            Me.ComandoRiesgoPriorizacion.Enabled = True
        End If
        Me.Elaborado = .Elaborado
        Me.Revisado = .Revisado
        Me.Aprobado = .Aprobado
        m_URLInforme = .URLArchivoInforme
        p_Error = .Error
        If p_Error <> "" Then Exit Sub
        If m_URLInforme <> "" Then
            Me.ImagenInforme.Visible = True
            Me.lblInformeGenerado.Visible = True
        Else
            Me.ImagenInforme.Visible = False
            Me.lblInformeGenerado.Visible = False
        End If
    End With
    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.Description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    EstablecerlblRechazado Me
    Me.ComandoGenerarInforme.Enabled = True
    If EsTecnico <> EnumSiNo.Sí Then
        If CInt(m_ObjEdicionActiva.IDEdicion) = CInt(m_ObjEdicionActiva.Proyecto.EdicionUltima.IDEdicion) Then
            If m_ObjEdicionActiva.Proyecto.FechaCierre <> "" Then
                Me.ComandoAbrirProyecto.Visible = True
            End If
        End If
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.Description
End Sub

Private Sub ImagenInforme_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not fso.FileExists(m_URLInforme) Then
        m_Error = "El archivo del informe no es alcanzable en estos momentos: " & vbNewLine & m_URLInforme
        Err.Raise 1000
    End If
    If FicheroAbierto(m_URLInforme, m_Error) Then
        m_Error = "Ya tiene el fichero abierto"
        Err.Raise 1000
    End If
    AbrirEnLocal m_URLInforme, Me.hWnd, m_Error
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
        m_Error = "Al ImagenInforme_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub m_FormRiesgo_RiesgoNuevo()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not Form_FormRiesgosGestion.m_ColRiesgosAplicados.Exists(CStr(m_ObjRiesgoActivo.IDRiesgo)) Then
        Form_FormRiesgosGestion.m_ColRiesgosAplicados.Add CStr(m_ObjRiesgoActivo.IDRiesgo), m_ObjRiesgoActivo
    End If
    
    Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
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
        m_Error = "Al m_FormRiesgo_RiesgoNuevo se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

' ----------------------------------------------------------------------------
' B2 / Punto 15 — Abre popup de anexos de TODA la edicion (cache-first).
'   Mismo patron que Form_FormRiesgosGestion.ComandoAnexos_Click pero
'   hardcoded a "Edicion" para no depender del openArgs del nav form.
'   Muestra TODOS los anexos (directos + de riesgos hijos) via
'   modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached.
' ----------------------------------------------------------------------------
Private Sub ComandoVerAnexos_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    DoCmd.OpenForm "FormAnexos", openArgs:="Edición"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerAnexos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

