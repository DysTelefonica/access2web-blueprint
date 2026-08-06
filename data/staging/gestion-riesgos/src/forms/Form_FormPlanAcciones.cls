Attribute VB_Name = "Form_FormPlanAcciones"
Option Compare Database
Option Explicit
Private m_EsMitigacion As String
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_NodoPlan As MSComctlLib.Node
Public Event AccionNueva(m_PMAccion As Object)
Private Const ALTO_DESCRIPCION_SIN_CRAIZ As Long = 2712
Private Const ALTO_DESCRIPCION_CON_CRAIZ As Long = 1017
' Cache-first del issue #90: m_FechaBaseProyecto se resuelve UNA sola vez en
' Form_Load (vía modEdicionFechasPresenter.ObtenerFechaBaseProyecto) y se
' pasa a modAccionFechasHelper.ValidarFechasAccion desde ComandoGrabar_Click.
' Variant para soportar Null (proyecto sin ediciones: constraint (b) no aplica).
Private m_FechaBaseProyecto As Variant
Private m_Error As String



Private Sub Accion_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub

Private Sub Accion_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub

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
    Dim m_Plan As Object
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    ' issue #90 — Validación cronológica de fechas (constraints (a) y (b))
    ' antes de .Registrar. Cache-first: m_FechaBaseProyecto se resolvió en
    ' Form_Load UNA sola vez (1 query SQL). Validamos ANTES de asignar al
    ' PMAccion/PCAccion para evitar registros con constraints violadas.
    Dim m_ErrValidacion As String
    Dim m_ResultadoValidacion As Boolean
    Dim m_FechaBaseTxt As String
    Dim m_FechaInicioVal As String
    Dim m_FechaFinPrevistaVal As String
    If Not IsNull(m_FechaBaseProyecto) Then
        m_FechaBaseTxt = Format$(CDate(m_FechaBaseProyecto), "dd/mm/yyyy")
    Else
        m_FechaBaseTxt = ""
    End If
    m_FechaInicioVal = Nz(Me.Controls("FechaInicio").Value, "")
    m_FechaFinPrevistaVal = Nz(Me.Controls("FechaFinPrevista").Value, "")

    m_ResultadoValidacion = modAccionFechasHelper.ValidarFechasAccion( _
        m_FechaInicioVal, _
        m_FechaFinPrevistaVal, _
        m_FechaBaseTxt, _
        m_ErrValidacion)
    If Not m_ResultadoValidacion Then
        m_Error = IIf(Len(m_ErrValidacion) > 0, m_ErrValidacion, "Validación de fechas fallida")
        Err.Raise 1000
    End If

    If m_EsMitigacion = "Sí" Then
        Set m_ObjPMAccionActiva = New PMAccion
        With m_ObjPMAccionActiva
            .IDMitigacion = m_ObjPMActivo.IDMitigacion
            .Accion = Nz(Me.Controls("Accion").Value, "")
            .ResponsableAccion = Nz(Me.Controls("ResponsableAccion").Value, "")
            .FechaInicio = Nz(Me.Controls("FechaInicio").Value, "")
            .FechaFinPrevista = Nz(Me.Controls("FechaFinPrevista").Value, "")
            .FechaFinReal = Nz(Me.Controls("FechaFinReal").Value, "")
            .Registrar , m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            RaiseEvent AccionNueva(m_ObjPMAccionActiva)
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Set m_ObjRiesgoActivo.ColPMs = Nothing
            modFormCoordinationHelper.Coord_RefrescarArbolRiesgosScope "plan", m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            pregunta = MsgBox("¿Desea dar de alta otra acción", vbExclamation + vbYesNo + vbDefaultButton1, "Nueva acción")
            If pregunta <> vbYes Then
                DoCmd.Close acForm, Me.Name, acSaveNo
                Exit Sub
            End If
            Set m_ObjPMAccionActiva = Nothing
            Me.Controls("Accion").Value = Null
            Me.Controls("ResponsableAccion").Value = Null
            Me.Controls("FechaInicio").Value = Null
            Me.Controls("FechaFinPrevista").Value = Null
            Me.Controls("FechaFinReal").Value = Null

            Me.Controls("Accion").SetFocus
            Exit Sub

        End With


    Else
        Set m_ObjPCAccionActiva = New PCAccion
        With m_ObjPCAccionActiva

            .IDContingencia = m_ObjPCActivo.IDContingencia
            .Accion = Nz(Me.Controls("Accion").Value, "")
            .ResponsableAccion = Nz(Me.Controls("ResponsableAccion").Value, "")
            .FechaInicio = Nz(Me.Controls("FechaInicio").Value, "")
            .FechaFinPrevista = Nz(Me.Controls("FechaFinPrevista").Value, "")
            .FechaFinReal = Nz(Me.Controls("FechaFinReal").Value, "")
            .Registrar , m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            RaiseEvent AccionNueva(m_ObjPCAccionActiva)
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Set m_ObjRiesgoActivo.ColPCs = Nothing
            modFormCoordinationHelper.Coord_RefrescarArbolRiesgosScope "plan", m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            pregunta = MsgBox("¿Desea dar de alta otra acción", vbExclamation + vbYesNo + vbDefaultButton1, "Nueva acción")
            If pregunta <> vbYes Then
                DoCmd.Close acForm, Me.Name, acSaveNo
                Exit Sub
            End If
            Set m_ObjPCAccionActiva = Nothing
            Me.Controls("Accion").Value = Null
            Me.Controls("ResponsableAccion").Value = Null
            Me.Controls("FechaInicio").Value = Null
            Me.Controls("FechaFinPrevista").Value = Null
            Me.Controls("FechaFinReal").Value = Null

            Me.Controls("Accion").SetFocus
            Exit Sub

        End With
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
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If

End Sub
Private Sub ComandoSalir_Click()
    
    
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub FechaFinPrevista_BeforeUpdate(Cancel As Integer)
    EstablecerColorFechaFinPrevista Me
        
    
End Sub

Private Sub FechaFinPrevista_Change()
    EstablecerColorFechaFinPrevista Me
End Sub

Private Sub FechaFinPrevista_Exit(Cancel As Integer)
    
    EstablecerColorFechaFinPrevista Me
End Sub

Private Sub FechaFinReal_BeforeUpdate(Cancel As Integer)
    EstablecerColorFechaFinReal Me
End Sub

Private Sub FechaFinReal_Change()
    EstablecerColorFechaFinReal Me
End Sub

Private Sub FechaFinReal_Exit(Cancel As Integer)
    EstablecerColorFechaFinReal Me
End Sub

Private Sub FechaInicio_BeforeUpdate(Cancel As Integer)
    EstablecerColorFechaInicio Me
End Sub

Private Sub FechaInicio_Change()
    EstablecerColorFechaInicio Me
End Sub

Private Sub FechaInicio_Exit(Cancel As Integer)
    EstablecerColorFechaInicio Me
End Sub

Private Sub Form_Load()


    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    m_EsMitigacion = Nz(Me.openArgs(), "")

    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    ' Cache-first del issue #90 — resolver UNA sola vez FechaBaseProyecto
    ' (= MIN(FechaEdicion) por IDProyecto). Se guarda en m_FechaBaseProyecto
    ' y se pasa a modAccionFechasHelper.ValidarFechasAccion desde
    ' ComandoGrabar_Click antes de .Registrar.
    m_FechaBaseProyecto = Null
    Dim m_ObjRiesgoParaCache As riesgo
    If m_EsMitigacion = "Sí" Then
        If Not m_ObjPMActivo Is Nothing Then
            Set m_ObjRiesgoParaCache = m_ObjPMActivo.riesgo
        End If
    Else
        If Not m_ObjPCActivo Is Nothing Then
            Set m_ObjRiesgoParaCache = m_ObjPCActivo.riesgo
        End If
    End If
    If Not m_ObjRiesgoParaCache Is Nothing Then
        If Not m_ObjRiesgoParaCache.Edicion Is Nothing Then
            If Not m_ObjRiesgoParaCache.Edicion.Proyecto Is Nothing Then
                Dim m_ErrCachePlan As String
                Dim m_IDProyectoPlan As String
                m_IDProyectoPlan = CStr(m_ObjRiesgoParaCache.Edicion.Proyecto.IDProyecto)
                If Len(Trim$(m_IDProyectoPlan)) > 0 Then
                    m_FechaBaseProyecto = modEdicionFechasPresenter.ObtenerFechaBaseProyecto( _
                        m_IDProyectoPlan, Nothing, m_ErrCachePlan)
                End If
            End If
        End If
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
   
    Dim m_DescRiesgo As String
    Dim m_Riesgo As riesgo
    Dim m_CausaRaiz As String
    Dim m_Titulo As String
    Dim m_RowSource As String
    Dim blnPermitidoEditar As Boolean
    Dim m_Edicion As Edicion
    Dim m_EnumTipoNodo As EnumTipoNodo
    On Error GoTo errores
    
    p_Error = ""
    
    Me.Controls("ComandoGrabar").Enabled = True
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMActivo Is Nothing Then
            p_Error = "No se conoce el plan de mitigación"
            Err.Raise 1000
        End If
        Set m_Edicion = m_ObjPMActivo.riesgo.Edicion
        Set m_Riesgo = m_ObjPMActivo.riesgo
        m_Titulo = "ALTA ACCIÓN PARA EL PLAN " & m_ObjPMActivo.CodMitigacion
        Me.Controls("lblDisparador").Caption = "Denominación"
        Me.Controls("DisparadorDelPlan").Value = m_ObjPMActivo.DisparadorDelPlan
    Else
        If m_ObjPCActivo Is Nothing Then
            p_Error = "No se conoce el plan de contingencia"
            Err.Raise 1000
        End If
        Set m_Edicion = m_ObjPCActivo.riesgo.Edicion
        Set m_Riesgo = m_ObjPCActivo.riesgo
        m_Titulo = "ALTA ACCIÓN PARA EL PLAN " & m_ObjPCActivo.CodContingencia
        Me.Controls("lblDisparador").Caption = "Disparador"
        Me.Controls("DisparadorDelPlan").Value = m_ObjPCActivo.DisparadorDelPlan
    End If
    If m_Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
        m_DescRiesgo = m_Riesgo.Descripcion
        m_CausaRaiz = m_Riesgo.CausaRaiz
    Else
        m_DescRiesgo = m_Riesgo.Descripcion
    End If
    Me.Controls("lblTitulo1").Caption = m_Titulo
    If m_Edicion.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        If m_Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
            blnPermitidoEditar = True
        
        Else
            blnPermitidoEditar = True
        End If
        ' HR3 closure: leer nodo + ver descripcion via modFormCoordinationHelper
        ' (no acceso directo a otro form .cls / Forms("X").Controls).
        If FormularioAbierto("FormRiesgosGestion") Then
            Set m_NodoSeleccionado = modFormCoordinationHelper.Coord_GetSelectedNode()
            m_VerDescripcion = modFormCoordinationHelper.Coord_GetComboVerDescripcion()
            m_EnumTipoNodo = getTipoNodo(m_NodoSeleccionado)
            If m_EnumTipoNodo = EnumTipoNodo.pm Then
                Set m_NodoPlan = m_NodoSeleccionado
            ElseIf m_EnumTipoNodo = EnumTipoNodo.pc Then
                Set m_NodoPlan = m_NodoSeleccionado
            ElseIf m_EnumTipoNodo = EnumTipoNodo.PMA Then
                Set m_NodoPlan = m_NodoSeleccionado.Parent
            ElseIf m_EnumTipoNodo = EnumTipoNodo.PCA Then
                Set m_NodoPlan = m_NodoSeleccionado.Parent
            End If
        End If
    End If
    
    If blnPermitidoEditar = True Then
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
    End If
    
    If m_Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
        Me.Controls("lblCRaiz").Visible = False
        Me.Controls("CausaRaiz").Visible = False
        Me.Controls("Descripcion").Height = ALTO_DESCRIPCION_SIN_CRAIZ
        Me.Controls("Descripcion").FontSize = 11
    Else
        Me.Controls("lblCRaiz").Visible = True
        Me.Controls("CausaRaiz").Visible = True
        Me.Controls("Descripcion").Height = ALTO_DESCRIPCION_CON_CRAIZ
        If m_Riesgo.CausaRaiz <> "" Then
            Me.Controls("CausaRaiz").Value = m_Riesgo.CausaRaiz
        End If
        Me.Controls("Descripcion").FontSize = 9
        Me.Controls("CausaRaiz").FontSize = 9
    End If
    
    Me.Controls("Descripcion").Value = m_DescRiesgo
    
    m_RowSource = m_Riesgo.Edicion.Proyecto.CadenaNombreAutorizados
    m_Error = m_Riesgo.Edicion.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Me.Controls("ResponsableAccion").RowSource = m_RowSource
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Sub ResponsableAccion_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub

Private Sub ResponsableAccion_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub
