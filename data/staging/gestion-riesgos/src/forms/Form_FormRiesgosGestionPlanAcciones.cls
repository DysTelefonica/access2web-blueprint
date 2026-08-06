Attribute VB_Name = "Form_FormRiesgosGestionPlanAcciones"
Option Compare Database
Option Explicit

Private blnPermitidoEditar As Boolean
Private m_ObjPMAccionAlInicio As PMAccion
Private m_ObjPCAccionAlInicio As PCAccion
Private m_EsMitigacion As String
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_NodoPlan As MSComctlLib.Node
Private m_NodoAccion As MSComctlLib.Node

Private m_Riesgo As riesgo
' Cache-first del issue #90: m_FechaBaseProyecto se resuelve UNA sola vez en
' Form_Load (vía modEdicionFechasPresenter.ObtenerFechaBaseProyecto) y se
' pasa a modAccionFechasHelper.ValidarFechasAccion desde los BeforeUpdate
' de FechaInicio y FechaFinPrevista. Variant para soportar Null (proyecto
' sin ediciones: constraint (b) no aplica).
Private m_FechaBaseProyecto As Variant
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

Private Sub ComandoEliminar_Click()
    Dim m_Plan As Object

    
    Dim m_NombreIcono As String
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMAccionActiva Is Nothing Then
            m_Error = "No se sabe la acción plan de mitigación"
            Err.Raise 1000
        End If
        Set m_Plan = m_ObjPMAccionActiva.Mitigacion
    Else
        If m_ObjPCAccionActiva Is Nothing Then
            m_Error = "No se sabe la acción del plan de contingencia"
            Err.Raise 1000
        End If
        Set m_Plan = m_ObjPCAccionActiva.Contingencia
    End If
    
    If blnPermitidoEditar = False Then
        m_Error = "Usuario no autorizado a esa operación"
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("¿Desea realmente borrar la acción seleccionada?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de una acción")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_EsMitigacion = "Sí" Then
        m_ObjPMAccionActiva.Borrar m_Error
        
    Else
        m_ObjPCAccionActiva.Borrar m_Error
        
    End If
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    modFormCoordinationHelper.Coord_RefrescarArbolRiesgosScope "action-delete", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("Acción borrada con éxito", vbInformation, "Acción borrada")
    
    
    
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
    Dim m_Plan As Object
    Dim m_NombreIcono As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If blnPermitidoEditar = False Then
        Exit Sub
    End If
    If m_EsMitigacion = "Sí" Then
        With m_ObjPMAccionActiva
            
            .Accion = Nz(Me.Controls("Accion").Value, "")
            .ResponsableAccion = Nz(Me.Controls("ResponsableAccion").Value, "")
            .FechaInicio = Nz(Me.Controls("FechaInicio").Value, "")
            .FechaFinPrevista = Nz(Me.Controls("FechaFinPrevista").Value, "")
            .FechaFinReal = Nz(Me.Controls("FechaFinReal").Value, "")
        End With
        Set m_Plan = m_ObjPMAccionActiva.Mitigacion
    Else
        With m_ObjPCAccionActiva
            
            .Accion = Nz(Me.Controls("Accion").Value, "")
            .ResponsableAccion = Nz(Me.Controls("ResponsableAccion").Value, "")
            .FechaInicio = Nz(Me.Controls("FechaInicio").Value, "")
            .FechaFinPrevista = Nz(Me.Controls("FechaFinPrevista").Value, "")
            .FechaFinReal = Nz(Me.Controls("FechaFinReal").Value, "")
        End With
        Set m_Plan = m_ObjPCAccionActiva.Contingencia
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
        m_ObjPMAccionActiva.Registrar m_ObjPMAccionAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjPMAccionAlInicio = Constructor.getPMAccion(m_ObjPMAccionAlInicio.IDAccionMitigacion)
        Set m_ObjRiesgoActivo.ColPMs = Nothing
    Else
        m_ObjPCAccionActiva.Registrar m_ObjPCAccionAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjPCAccionAlInicio = Constructor.getPCAccion(m_ObjPCAccionAlInicio.IDAccionContingencia)
        Set m_ObjRiesgoActivo.ColPCs = Nothing
    End If
    
    
    EstablecerDatos

    modFormCoordinationHelper.Coord_RefrescarArbolRiesgosScope "action", m_Error
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
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub





Private Sub FechaFinReal_Change()
    EstablecerColorFechaFinReal Me
End Sub

Private Sub FechaInicio_Change()
    EstablecerColorFechaInicio Me
End Sub

Private Sub Form_Load()

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If modFormCoordinationHelper.Coord_GetEsMitigacion() = "" Then
        m_Error = "No se sabe si es plan de mitigación o de contingencia"
        Err.Raise 1000
    End If
    If modFormCoordinationHelper.Coord_GetEsMitigacion() = CStr(EnumSiNo.Sí) Then
        m_EsMitigacion = "Sí"
    Else
        m_EsMitigacion = "No"
    End If
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMAccionActiva Is Nothing Then
            m_Error = "No se conoce el plan de mitigación de la acción"
            Err.Raise 1000
        End If
        Set m_ObjPMAccionAlInicio = getPMAccion(m_ObjPMAccionActiva.IDAccionMitigacion, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjPMAccionAlInicio Is Nothing Then
            m_Error = "Es una edición y no se han podido cargar los datos"
            Err.Raise 1000
        End If
        
    Else
        If m_ObjPCAccionActiva Is Nothing Then
            m_Error = "No se conoce el plan de contingencia de la acción"
            Err.Raise 1000
        End If
        Set m_ObjPCAccionAlInicio = getPCAccion(m_ObjPCAccionActiva.IDAccionContingencia, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjPCAccionAlInicio Is Nothing Then
            m_Error = "Es una edición y no se han podido cargar los datos"
            Err.Raise 1000
        End If
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If blnPermitidoEditar = True Then
        Me.Controls("ComandoGrabar").Enabled = True
        Me.Controls("ComandoEliminar").Enabled = True
        
        Me.AllowEdits = True
    Else
        Me.Controls("ComandoGrabar").Enabled = False
        Me.Controls("ComandoEliminar").Enabled = False
        Me.AllowEdits = False
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

    Me.Controls("ComandoEliminar").Enabled = False
    Me.Controls("ComandoGrabar").Enabled = False
    Me.Controls("FechaFinPrevista").ForeColor = m_ColorOK

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
    Dim m_RowSource As String
    m_RowSource = m_ObjEdicionActiva.Proyecto.CadenaNombreAutorizados
    ' HR3 closure: leer nodo + ver descripcion via modFormCoordinationHelper
    ' (no acceso directo a otro form .cls / Forms("X").Controls).
    If FormularioAbierto("FormRiesgosGestion") Then
        Set m_NodoSeleccionado = modFormCoordinationHelper.Coord_GetSelectedNode()
        m_VerDescripcion = modFormCoordinationHelper.Coord_GetComboVerDescripcion()
    End If
    If m_NodoSeleccionado Is Nothing Then
        p_Error = "No se conoce el nodo seleccionado"
        Err.Raise 1000
    End If
    Dim m_ObjSel As Object
    Set m_ObjSel = modFormCoordinationHelper.Coord_GetObjSeleccionado()
    If TypeOf m_ObjSel Is pm Or TypeOf m_ObjSel Is pc Then
        Set m_NodoPlan = m_NodoSeleccionado
        Set m_NodoAccion = Nothing
    ElseIf TypeOf m_ObjSel Is PMAccion Or TypeOf m_ObjSel Is PCAccion Then
        Set m_NodoAccion = m_NodoSeleccionado
        Set m_NodoPlan = m_NodoAccion.Parent
    Else
        p_Error = "Nodo seleccionado desconocido"
        Err.Raise 1000
    End If
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        blnPermitidoEditar = True
    End If
    Me.Controls("ResponsableAccion").RowSource = m_RowSource
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.Description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    If m_EsMitigacion = "Sí" Then
        With m_ObjPMAccionAlInicio
            Me.Controls("lblTitulo").Caption = "Accción " & .CodAccion & " " & " " & "plan de mitigación " & " " & .Mitigacion.CodMitigacion
            Me.Controls("Accion").Value = .Accion
            If .ResponsableAccion <> "" Then Me.Controls("ResponsableAccion").Value = .ResponsableAccion
            If IsDate(.FechaInicio) Then Me.Controls("FechaInicio").Value = .FechaInicio
            If IsDate(.FechaFinPrevista) Then Me.Controls("FechaFinPrevista").Value = .FechaFinPrevista
            If m_ObjPMAccionAlInicio.EsActivo = No Then
                Me.Controls("FechaFinPrevista").ForeColor = m_ColorOK
            Else
                If IsDate(.FechaFinPrevista) Then
                    Me.Controls("FechaFinPrevista").Value = .FechaFinPrevista
                    If .NecesitaReplanificacion = EnumSiNo.Sí Then
                        Me.Controls("FechaFinPrevista").ForeColor = m_ColorNOOK
                    Else
                        Me.Controls("FechaFinPrevista").ForeColor = m_ColorOK
                    End If
                End If
            End If
            If IsDate(.FechaFinReal) Then Me.Controls("FechaFinReal").Value = .FechaFinReal
            Set m_Riesgo = .Mitigacion.riesgo
        End With
    Else
        With m_ObjPCAccionAlInicio
            Me.Controls("lblTitulo").Caption = "Accción " & .CodAccion & " " & " " & "plan de contingencia " & " " & .Contingencia.CodContingencia
            Me.Controls("Accion").Value = .Accion
            If .ResponsableAccion <> "" Then Me.Controls("ResponsableAccion").Value = .ResponsableAccion
            If IsDate(.FechaInicio) Then Me.Controls("FechaInicio").Value = .FechaInicio
            If IsDate(.FechaFinPrevista) Then
                Me.Controls("FechaFinPrevista").Value = .FechaFinPrevista
                If .NecesitaReplanificacion = EnumSiNo.Sí Then
                    Me.Controls("FechaFinPrevista").ForeColor = m_ColorNOOK
                Else
                    Me.Controls("FechaFinPrevista").ForeColor = m_ColorOK
                End If
            End If
            If IsDate(.FechaFinReal) Then Me.Controls("FechaFinReal").Value = .FechaFinReal
            Set m_Riesgo = .Contingencia.riesgo
        End With
    End If

    ' Cache-first del issue #90 — resolver UNA sola vez FechaBaseProyecto
    ' (= MIN(FechaEdicion) por IDProyecto). Se guarda en m_FechaBaseProyecto
    ' (Variant) y se pasa a modAccionFechasHelper.ValidarFechasAccion desde
    ' los BeforeUpdate. Si m_Riesgo no expone Edicion o Proyecto, se deja
    ' Null (constraint (b) no aplicará, comportamiento defensivo).
    m_FechaBaseProyecto = Null
    If Not m_Riesgo Is Nothing Then
        If Not m_Riesgo.Edicion Is Nothing Then
            If Not m_Riesgo.Edicion.Proyecto Is Nothing Then
                Dim m_ErrCache As String
                Dim m_IDProyecto As String
                m_IDProyecto = CStr(m_Riesgo.Edicion.Proyecto.IDProyecto)
                If Len(Trim$(m_IDProyecto)) > 0 Then
                    m_FechaBaseProyecto = modEdicionFechasPresenter.ObtenerFechaBaseProyecto( _
                        m_IDProyecto, Nothing, m_ErrCache)
                    ' Si m_ErrCache está poblado, dejamos Null y la constraint
                    ' (b) se omite (skip defensivo — el form no rompe por un
                    ' error de cache).
                End If
            End If
        End If
    End If

    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.Description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    RellenarBordeCamposObligatorios Me
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.Description
End Sub

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    Dim detector As RiesgoPlanActionChangeDetector
    
    On Error GoTo errores
    Set detector = New RiesgoPlanActionChangeDetector
    
    If m_EsMitigacion = "Sí" Then
        HaHabidoCambios = detector.HasPMActionChanges(m_ObjPMAccionAlInicio, m_ObjPMAccionActiva, p_Error)
    Else
        HaHabidoCambios = detector.HasPCActionChanges(m_ObjPCAccionAlInicio, m_ObjPCAccionActiva, p_Error)
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
Private Sub Accion_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub Accion_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub FechaFinPrevista_BeforeUpdate(Cancel As Integer)

   Dim m_ErrValidacion As String
   Dim m_Resultado As Boolean
   Dim m_FechaBaseTxt As String
   Dim m_FechaInicioVal As String
   Dim m_FechaFinPrevistaVal As String

   ' issue #90 — constraint (a): FechaFinPrevista >= FechaInicio.
   ' constraint (b): FechaInicio >= FechaBaseProyecto (cache en m_FechaBaseProyecto).
   ' El helper modAccionFechasHelper.ValidarFechasAccion valida ambas en una
   ' sola llamada. Si falla, mostramos MsgBox y cancelamos el cambio del campo.
   If Not IsNull(m_FechaBaseProyecto) Then
      m_FechaBaseTxt = Format$(CDate(m_FechaBaseProyecto), "dd/mm/yyyy")
   Else
      m_FechaBaseTxt = ""
   End If
   m_FechaInicioVal = Nz(Me.Controls("FechaInicio").Value, "")
   m_FechaFinPrevistaVal = Nz(Me.Controls("FechaFinPrevista").Value, "")

   m_Resultado = modAccionFechasHelper.ValidarFechasAccion( _
       m_FechaInicioVal, _
       m_FechaFinPrevistaVal, _
       m_FechaBaseTxt, _
       m_ErrValidacion)

   If Not m_Resultado Then
      If Len(m_ErrValidacion) > 0 Then
         pregunta = MostrarMsgBoxSiNoEsTest(m_ErrValidacion, vbExclamation, "Validación de fechas")
      End If
      Cancel = True
      Exit Sub
   End If

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

Private Sub FechaFinReal_Exit(Cancel As Integer)
    EstablecerColorFechaFinReal Me
End Sub

Private Sub FechaInicio_BeforeUpdate(Cancel As Integer)

   Dim m_ErrValidacion As String
   Dim m_Resultado As Boolean
   Dim m_FechaBaseTxt As String
   Dim m_FechaInicioVal As String
   Dim m_FechaFinPrevistaVal As String

   ' issue #90 — Misma validación que FechaFinPrevista_BeforeUpdate.
   ' Llamamos al helper con los valores ACTUALES de los controles (antes
   ' de que el cambio se persista). El cache m_FechaBaseProyecto fue
   ' resuelto en Form_Load/EstablecerDatos (cache-first: 1 sola query).
   If Not IsNull(m_FechaBaseProyecto) Then
      m_FechaBaseTxt = Format$(CDate(m_FechaBaseProyecto), "dd/mm/yyyy")
   Else
      m_FechaBaseTxt = ""
   End If
   m_FechaInicioVal = Nz(Me.Controls("FechaInicio").Value, "")
   m_FechaFinPrevistaVal = Nz(Me.Controls("FechaFinPrevista").Value, "")

   m_Resultado = modAccionFechasHelper.ValidarFechasAccion( _
       m_FechaInicioVal, _
       m_FechaFinPrevistaVal, _
       m_FechaBaseTxt, _
       m_ErrValidacion)

   If Not m_Resultado Then
      If Len(m_ErrValidacion) > 0 Then
         pregunta = MostrarMsgBoxSiNoEsTest(m_ErrValidacion, vbExclamation, "Validación de fechas")
      End If
      Cancel = True
      Exit Sub
   End If

   EstablecerColorFechaInicio Me
End Sub

Private Sub FechaInicio_Exit(Cancel As Integer)
    EstablecerColorFechaInicio Me
End Sub
