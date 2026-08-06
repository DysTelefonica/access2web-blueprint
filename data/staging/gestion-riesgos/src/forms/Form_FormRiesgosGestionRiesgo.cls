Attribute VB_Name = "Form_FormRiesgosGestionRiesgo"
Option Compare Database
Option Explicit
Public Event RiesgoAceptado()
Public Event RiesgoRetirado()
Public Event RiesgoMaterializado()

Private m_URLInformeRiesgo As String
Private WithEvents m_FormPlan As Form_FormPlanPrincipal


' Constantes de diseño (en Twips: 1cm = 567 twips)
' Ajustados para que no solapen con la nueva sección de Histórico
Private Const ALTO_DESCRIPCION_SIN_CRAIZ As Long = 2000 ' ~3.5cm (reducido para dar espacio)
Private Const ALTO_DESCRIPCION_CON_CRAIZ As Long = 1137 ' ~2.0cm

' Cache del histórico de estados (Punto 18 / B3 — descubierto durante implementación):
' getEstadosDiferentesHastaEdicion hace 2N queries (1 por edición para el riesgo
' + 1 por edición para materializaciones). m_ObjHistoricoCache evita recalcular en
' cada render del histórico. Se rellena en Form_Load y se borra en Form_Close.
Private m_ObjHistoricoCache As Scripting.Dictionary

Private m_Error As String

' ==========================================
' MSG BOX HELPER (per access-vba-tdd §0 — no interactive UI in tests)
' p_PromptResult: -1 (default) = show MsgBox; 0 = skip MsgBox (test mode)
' ==========================================
Private Function MostrarMsgBoxSiNoEsTest(ByVal p_Message As String, _
                                       ByVal p_Style As VbMsgBoxStyle, _
                                       ByVal p_Title As String, _
                                       Optional ByVal p_PromptResult As Long = -1) As Long
    If p_PromptResult = 0 Then
        ' Test mode: skip MsgBox. Return vbOK as default.
        MostrarMsgBoxSiNoEsTest = vbOK
    Else
        MostrarMsgBoxSiNoEsTest = MsgBox(p_Message, p_Style, p_Title)
    End If
End Function

' ==========================================
' EVENTOS DE BOTONES
' ==========================================

Private Sub ComandoAltaPC_Click()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    If m_ObjRiesgoActivo.FichaRiesgoCompleta = EnumSiNo.No Then
        DoCmd.Hourglass False
        pregunta = MostrarMsgBoxSiNoEsTest("No se puede editar nin dar de alta un plan de mitigación de un riesgo al que " & _
                "le falta completar los datos", vbExclamation, "Riesgo incompleto")
        
        If Me.ComandoDetalle.Enabled Then
            ComandoDetalle_Click
            Exit Sub
        End If
    End If
    If FormularioAbierto("FormPlanPrincipal") Then DoCmd.Close acForm, "FormPlanPrincipal", acSaveNo
    DoCmd.OpenForm "FormPlanPrincipal", , , , , , "No"
    If FormularioAbierto("FormPlanPrincipal") Then Set m_FormPlan = Forms("FormPlanPrincipal")
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoAltaPC_Click", Err
End Sub

Private Sub ComandoAltaPM_Click()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    If m_ObjRiesgoActivo.FichaRiesgoCompleta = EnumSiNo.No Then
        DoCmd.Hourglass False
        pregunta = MostrarMsgBoxSiNoEsTest("No se puede editar nin dar de alta un plan de contingencia de un riesgo al que " & _
                "le falta completar los datos", vbExclamation, "Riesgo incompleto")
        
        If Me.ComandoDetalle.Enabled Then
            ComandoDetalle_Click
            Exit Sub
        End If
    End If
    If FormularioAbierto("FormPlanPrincipal") Then DoCmd.Close acForm, "FormPlanPrincipal", acSaveNo
    DoCmd.OpenForm "FormPlanPrincipal", , , , , , "Sí"
    If FormularioAbierto("FormPlanPrincipal") Then Set m_FormPlan = Forms("FormPlanPrincipal")
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoAltaPM_Click", Err
End Sub

Private Sub ComandoDetalle_Click()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormRiesgo") Then DoCmd.Close acForm, "FormRiesgo", acSaveNo
    m_EsAlta = EnumSiNo.No
    Set m_ObjRiesgoAlInicio = Nothing
    m_EsAlta = EnumSiNo.No
    DoCmd.OpenForm "FormRiesgo"
    Forms("FormRiesgo").ComandoVerEdicion.Visible = False
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoDetalle_Click", Err
End Sub

Private Sub ComandoEliminar_Click()
    Dim m_IdRiesgo As String

    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""

    If m_ObjRiesgoActivo Is Nothing Then Err.Raise 1000, , "No se sabe el riesgo activo"
    If m_ObjRiesgoActivo.Borrable = EnumSiNo.No Then Err.Raise 1000, , "El riesgo no permite ser borrado"
    If m_ObjRiesgoActivo.Edicion.UsuarioConectadoAutorizado = EnumSiNo.No Then Err.Raise 1000, , "Usuario no autorizado"

    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    If MostrarMsgBoxSiNoEsTest("¿Desea realmente borrar el riesgo seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de un riesgo") <> vbYes Then Exit Sub

    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_IdRiesgo = m_ObjRiesgoActivo.IDRiesgo
    m_ObjRiesgoActivo.Borrar m_Error
    If m_Error <> "" Then Err.Raise 1000
    InvalidarCacheRiesgo m_IdRiesgo

    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents

    ' HR3c closure: post-delete refresh via modFormCoordinationHelper
    ' (no acceso directo a Form_FormRiesgosGestion desde este form .cls).
    ' m_NodoPadre (antes L123) era dead-code: Set m_NodoPadre = ... sin uso.
    modFormCoordinationHelper.Coord_ColRiesgosAplicados_RemoveIfExists CStr(m_IdRiesgo), m_Error
    If m_Error <> "" Then Err.Raise 1000
    modFormCoordinationHelper.Coord_CargarArbol m_Error
    If m_Error <> "" Then Err.Raise 1000
    modFormCoordinationHelper.Coord_SeleccionarNodo m_ObjEdicionActiva, m_Error
    If m_Error <> "" Then Err.Raise 1000
    MostrarMsgBoxSiNoEsTest "Riesgo borrado con éxito", vbInformation, "Riesgo borrado"
    Exit Sub
errores:
    GestionarError "ComandoEliminar_Click", Err
End Sub

Private Sub ComandoVerInformeRiesgo_Click()
    Dim m_URLInforme As String
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    m_URLInforme = GenerarInformeRiesgoHTML(p_IDRiesgo:=m_ObjRiesgoActivo.IDRiesgo, p_hWnd:=Application.hWndAccessApp, p_Error:=m_Error)
    If m_Error <> "" Then Err.Raise 1000
    AvanceCerrar
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoVerInformeRiesgo_Click", Err
End Sub

' ==========================================
' CARGA DEL FORMULARIO
' ==========================================

Private Sub Form_Load()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    If m_ObjRiesgoActivo Is Nothing Then
        VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
        Exit Sub
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "Form_Load", Err
End Sub

' ----------------------------------------------------------------------------
' Form_Close (B3 / Punto 18 — cache cleanup)
'   Libera m_ObjHistoricoCache cuando se cierra el form. El cache vive
'   solo lo que vive el form (Form_Load → Form_Close).
' ----------------------------------------------------------------------------
Private Sub Form_Close()
    On Error Resume Next
    Set m_ObjHistoricoCache = Nothing
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
'   El gemelo mantiene su estructura de 3 sub-rutinas extraidas en Block 1 WI-5.
'   Esta funcion queda como thin coordinator (UI wiring + llamadas a subs).
' ----------------------------------------------------------------------------
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.ComandoDetalle.Enabled = True
    Me.ComandoEliminar.Visible = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoAltaPC.Enabled = False
    Me.ComandoAltaPM.Enabled = False

    ' Sub-rutinas extraídas (Block 1 WI-5 — thin form)
    Call EstablecerDatos_CargarRiesgo(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_AplicarEstado(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_RenderizarControles(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

' =============================================================================
' Sub-rutina 1: cargar el riesgo en los controles del form
'   - Título + permisos + campos básicos + Causa Raíz + fechas básicas
'   - Lee m_ObjRiesgoActivo directamente (cache via GetCachedRiesgo upstream)
' =============================================================================
Private Sub EstablecerDatos_CargarRiesgo(ByRef p_Error As String)
    Dim m_RiesgoAltoOMuyAltoTexto As String
    Dim blnPermitidoEditar As Boolean
    Dim m_Edicion As Edicion

    On Error GoTo errores

    With m_ObjRiesgoActivo
        ' Título
        m_RiesgoAltoOMuyAltoTexto = .RiesgoAltoOMuyAltoTexto
        If m_RiesgoAltoOMuyAltoTexto = "Sí" Then
            Me.lblTitulo.Caption = "Riesgo " & .CodigoRiesgo & " RIESGO ALTO O MUY ALTO****"
            Me.lblTitulo.ForeColor = RGB(255, 0, 0)
        Else
            Me.lblTitulo.Caption = "Riesgo " & .CodigoRiesgo
            Me.lblTitulo.ForeColor = RGB(0, 102, 255)
        End If

        ' Permisos
        Set m_Edicion = m_ObjRiesgoActivo.Edicion
        If m_Edicion.EsActivo = EnumSiNo.No Then
            blnPermitidoEditar = False
        Else
            blnPermitidoEditar = True
        End If
        If blnPermitidoEditar = True Then
            Me.ComandoAltaPM.Enabled = True
            Me.ComandoAltaPC.Enabled = True
            If .Borrable = EnumSiNo.Sí Then
                Me.ComandoEliminar.Visible = True
                Me.ComandoEliminar.Enabled = True
            End If
        End If

        ' Campos básicos
        Me.CodigoRiesgo = .CodigoRiesgo
        If .Priorizacion <> "" Then Me.Priorizacion = .Priorizacion
        If .Origen <> "" Then Me.Origen = .Origen
        If .Descripcion <> "" Then Me.Descripcion = .Descripcion

        ' Causa Raíz
        If m_ObjProyectoActivo.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            Me.lblCRaiz.Visible = False
            Me.CausaRaiz.Visible = False
            Me.Descripcion.Height = ALTO_DESCRIPCION_SIN_CRAIZ
            Me.Descripcion.FontSize = 11
        Else
            Me.lblCRaiz.Visible = True
            Me.CausaRaiz.Visible = True
            Me.Descripcion.Height = ALTO_DESCRIPCION_CON_CRAIZ
            If .CausaRaiz <> "" Then Me.CausaRaiz = .CausaRaiz
            Me.Descripcion.FontSize = 9
            Me.CausaRaiz.FontSize = 9
        End If
    End With

    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarRiesgo: " & Err.Description
End Sub

' =============================================================================
' Sub-rutina 2: aplicar el estado del riesgo (Estado, fechas, publicabilidad)
'   - Lee m_ObjRiesgoActivo directamente
'   - Podría delegar a modEstadoRiesgosActivo.ObtenerEstado pero el form
'     ya tiene la instancia cacheada: redundante serializar a JSON
' =============================================================================
Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    Dim m_ErroresRiesgo As String

    On Error GoTo errores

    With m_ObjRiesgoActivo
        ' Resto de campos (fechas, impacto, estado)
        If .EntidadDetecta <> "" Then Me.EntidadDetecta = .EntidadDetecta
        If .DetectadoPor <> "" Then Me.DetectadoPor = .DetectadoPor
        If .FechaDetectado <> "" Then Me.FechaDetectado = .FechaDetectado

        If .FechaMaterializado <> "" Then
            Me.FechaMaterializado = .FechaMaterializado
            Me.FechaMaterializado.ForeColor = m_ColorNOOK
        Else
            Me.FechaMaterializado.ForeColor = m_ColorOK
        End If

        If .FechaRetirado <> "" Then
            Me.FechaRetirado = .FechaRetirado
            Me.FechaRetirado.ForeColor = m_ColorNOOK
        Else
            Me.FechaRetirado.ForeColor = m_ColorOK
        End If

        If .ImpactoGlobal <> "" Then Me.ImpactoGlobal = .ImpactoGlobal
        Me.Estado = .ESTADOCalculadoTexto
        Me.FechaEstado = .FechaEstado
        Me.RiesgoAltoOMuyAlto = .RiesgoAltoOMuyAltoTexto

        m_ErroresRiesgo = .ErroresRiesgoTexto
        p_Error = .Error
        If p_Error <> "" Then Err.Raise 1000
    End With

    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.Description
End Sub

' =============================================================================
' Sub-rutina 3: renderizar controles complejos (publicabilidad, histórico)
'   - ConstruirDatosPublicabilidadRiesgo, EvaluarPublicabilidadRiesgo
'   - getEstadosDiferentesHastaEdicion para histórico
' =============================================================================
Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    Dim m_DatosPublicabilidad As tPublicabilidadRiesgoDatos
    Dim m_ChecksPublicabilidad As Scripting.Dictionary
    Dim m_VeredictoPublicabilidad As EnumPublicabilidadVeredicto
    Dim m_Edicion As Edicion
    Dim m_ObjColRiesgosEstados As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Resultado As String
    Dim dato As Variant
    Dim sFechaEstado As String, sEstadoRiesgo As String

    On Error GoTo errores

    Set m_Edicion = m_ObjRiesgoActivo.Edicion

    ' Informe / Carencias
    If m_Edicion.EsActivo = EnumSiNo.Sí Then
        Me.ComandoVerInformeRiesgo.Visible = True
        If ConstruirDatosPublicabilidadRiesgo(m_ObjRiesgoActivo, m_DatosPublicabilidad, , p_Error) = EnumSiNo.No Then Err.Raise 1000
        Dim m_Publicable As EnumSiNo
        m_Publicable = EvaluarPublicabilidadRiesgo(m_DatosPublicabilidad, m_ChecksPublicabilidad, m_VeredictoPublicabilidad, p_Error)
        If p_Error <> "" Then Err.Raise 1000

        If m_VeredictoPublicabilidad = EnumPublicabilidadVeredicto.NoPublicable Then
            Me.ComandoVerInformeRiesgo.ForeColor = 2366701
        Else
            Me.ComandoVerInformeRiesgo.ForeColor = 16737792
        End If
    Else
        Me.ComandoVerInformeRiesgo.Visible = False
    End If

    ' Histórico de estados
    ' B3 / Punto 18 + Puntos 19-21 (issue #100):
    '   - Cache a nivel de form (m_ObjHistoricoCache) para no recalcular 2N queries
    '     en cada render del histórico.
    '   - Orden cronológico inverso (más reciente arriba): lo hace el HELPER
    '     (getEstadosDiferentesHastaEdicion ordena descendente post-bucle),
    '     el form itera en orden natural del Dictionary (LBound -> UBound).
    '   - Columna NC por-edición: "Si" si el riesgo en esa edición tiene
    '     materializaciones pendientes de decidir NC, "-" en caso contrario.
    '     El NC viaja dentro de la tupla del helper (dato(2)), NO se calcula
    '     aquí (antes se computaba global con m_ObjRiesgoActivo, eso era el
    '     bug del issue #100 — el NC debe ser por-edición, no global).
    Dim m_Keys As Variant
    Dim i As Long
    Dim sNCEdicion As String

    On Error Resume Next ' Tolerar ausencia del control lstEstadosHistoricos
    Me.lstEstadosHistoricos.RowSource = ""
    Me.lstEstadosHistoricos.AddItem "Estado;Fecha;NC"

    ' 1. Cache: si está vacío, calcular; si está lleno, reusar.
    If m_ObjHistoricoCache Is Nothing Then
        On Error GoTo errores
        Set m_ObjHistoricoCache = getEstadosDiferentesHastaEdicion( _
                                        m_Edicion, _
                                        m_ObjRiesgoActivo.CodigoRiesgo, _
                                        CStr(m_Edicion.FechaPublicacion), _
                                        "", _
                                        p_Error)
        If p_Error <> "" Then Err.Raise 1000
    End If
    Set m_ObjColRiesgosEstados = m_ObjHistoricoCache
    On Error Resume Next

    ' 2. Iterar en orden natural (el helper ya devuelve orden descendente).
    If Not m_ObjColRiesgosEstados Is Nothing Then
        m_Keys = m_ObjColRiesgosEstados.Keys
        For i = LBound(m_Keys) To UBound(m_Keys)
            m_ID = m_Keys(i)
            m_Resultado = m_ObjColRiesgosEstados(m_ID)
            dato = Split(m_Resultado, "|")
            sEstadoRiesgo = dato(0)

            If UBound(dato) >= 1 Then
                sFechaEstado = dato(1)
                If IsDate(sFechaEstado) Then
                    sFechaEstado = Format(sFechaEstado, "dd/mm/yyyy")
                End If
            Else
                sFechaEstado = ""
            End If

            ' NC por-edición (issue #100): dato(2) trae "Si" o "-".
            ' Tolerar entradas con formato antiguo (2 campos) default a "-".
            If UBound(dato) >= 2 Then
                sNCEdicion = dato(2)
                If sNCEdicion = "" Then sNCEdicion = "-"
            Else
                sNCEdicion = "-"
            End If

            Me.lstEstadosHistoricos.AddItem sEstadoRiesgo & ";" & sFechaEstado & ";" & sNCEdicion
        Next i
    End If
    On Error GoTo errores

    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.Description
End Sub

Private Sub GestionarError(Metodo As String, objErr As ErrObject)
    DoCmd.Hourglass False
    If objErr.Number <> 1000 Then
        m_Error = "Al " & Metodo & " se ha producido el error n: " & objErr.Number & vbNewLine & "Detalle: " & objErr.Description
        CorreoAlAdministrador m_Error
        MostrarMsgBoxSiNoEsTest m_Error, vbCritical, "Error"
    Else
        MostrarMsgBoxSiNoEsTest m_Error, vbExclamation, "Advertencia"
    End If
End Sub

' =============================================================================
' B2 / Punto 15 — Abre popup de anexos del riesgo actual (cache-first).
'   Mismo patron que Form_FormRiesgosGestionEdicion.ComandoVerAnexos_Click
'   pero hardcoded a "Riesgo" para no depender del openArgs del nav form.
'   Muestra los anexos directos del riesgo via
'   modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached.
' -----------------------------------------------------------------------------
Private Sub ComandoVerAnexos_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    DoCmd.OpenForm "FormAnexos", openArgs:="Riesgo"

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

Private Sub m_FormPlan_PlanNuevo(p_Plan As Object)
    Dim m_NodoRiesgo As Object

    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""

    ' HR3 closure: leer nodo + invocar CargarArbolPM/PC via
    ' modFormCoordinationHelper (no acceso directo a otro form .cls).
    Set m_NodoRiesgo = modFormCoordinationHelper.Coord_GetSelectedNode()

    If TypeOf p_Plan Is pm Then
        modFormCoordinationHelper.Coord_CargarArbolPM _
            P_NodoRiesgo:=m_NodoRiesgo, _
            p_Riesgo:=m_ObjRiesgoActivo, _
            p_PM:=p_Plan, _
            p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
            p_Refrescando:=EnumSiNo.No, _
            p_Error:=m_Error
    Else
        modFormCoordinationHelper.Coord_CargarArbolPC _
            P_NodoRiesgo:=m_NodoRiesgo, _
            p_Riesgo:=m_ObjRiesgoActivo, _
            p_PC:=p_Plan, _
            p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
            p_Refrescando:=EnumSiNo.No, _
            p_Error:=m_Error
    End If
    If m_Error <> "" Then Err.Raise 1000
    modFormCoordinationHelper.Coord_SeleccionarNodo p_Plan, m_Error
    If m_Error <> "" Then Err.Raise 1000
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "m_FormPlan_PlanNuevo", Err
End Sub
