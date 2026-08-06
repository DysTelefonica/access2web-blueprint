Option Compare Database
Option Explicit

Private blnPermitidoEditar As Boolean
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_RiesgoMatSeleccionado As RiesgoMaterializacion
Private m_NC As nc
Private m_Error As String
Private Const FORM_DIALOGO_MATERIALIZACION As String = "FormMaterializacionPlanContingencia"
Private Const TEMP_MAT_ACEPTADA As String = "MaterializacionPlanContingenciaAceptada"
Private Const TEMP_MAT_FECHA As String = "MaterializacionPlanContingenciaFecha"
Private Const TEMP_MAT_PLAN As String = "MaterializacionPlanContingenciaPlan"
Private Const OPENARGS_EDIT_MATERIALIZACION As String = "EDIT"
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

Private Sub ComandoFechaMaterializado_Click()
    
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim m_FechaMaterializado As String
    Dim m_IDPlanContingencia As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    With m_ObjRiesgoActivo
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(m_ObjRiesgoActivo.EstadoEnum)
        If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
            m_Error = "El riesgo está en camino de ser aceptado"
            Err.Raise 1000
        End If
        m_RiesgoEnRetirada = RiesgoEnRetirada(.EstadoEnum)
        If m_RiesgoEnRetirada = EnumSiNo.Sí Then
            m_Error = "El riesgo está en camino de ser retirado"
            Err.Raise 1000
        End If
        If .EstadoEnum = EnumRiesgoEstado.Materializado Then
            m_Error = "El riesgo aparece como materializado"
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MostrarMsgBoxSiNoEsTest("Al materializarse un riesgo, recuerde que ha de ejecutar el plan de contingencia y " & _
                "publicar la edición para que quede constancia", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Materialización")
        If pregunta <> vbYes Then
            
            Exit Sub
        End If
        If Not SolicitarDatosMaterializacion(m_FechaMaterializado, m_IDPlanContingencia, m_Error) Then
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Sub
        End If
        If Not .ValidarFechaMaterializacionPermitida(m_FechaMaterializado, m_Error) Then
            Err.Raise 1000
        End If
         VBA.DoEvents
        DoCmd.Hourglass True
        VBA.DoEvents
        Me.FechaMaterializado = m_FechaMaterializado
        
        m_ObjRiesgoActivo.FechaMaterializado = Me.FechaMaterializado
        .MaterializacionRegistrar Me.FechaMaterializado, m_Error, m_IDPlanContingencia
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjRiesgoAlInicio = GetCachedRiesgo(.idRiesgo)
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
        If m_Error <> "" Then
           Err.Raise 1000
        End If
        
    End With
    ActualizarEtiquetaEstadoRiesgo
    ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_NodoSeleccionado, p_VerDescripcion:=m_VerDescripcion, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "materializacion", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatos
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoFechaMaterializado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
   
End Sub

Private Function SolicitarDatosMaterializacion( _
    ByRef p_FechaMaterializado As String, _
    ByRef p_IDPlanContingencia As String, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo errores

    p_Error = ""
    p_FechaMaterializado = ""
    p_IDPlanContingencia = ""

    LimpiarTempVarsMaterializacion
    DoCmd.OpenForm FORM_DIALOGO_MATERIALIZACION, , , , , acDialog, _
        m_ObjRiesgoActivo.idRiesgo

    If TempVarValueMaterializacion(TEMP_MAT_ACEPTADA) <> "Sí" Then
        GoTo salida
    End If

    p_FechaMaterializado = TempVarValueMaterializacion(TEMP_MAT_FECHA)
    p_IDPlanContingencia = TempVarValueMaterializacion(TEMP_MAT_PLAN)
    SolicitarDatosMaterializacion = True

salida:
    LimpiarTempVarsMaterializacion
    Exit Function

errores:
    p_Error = "El método SolicitarDatosMaterializacion ha devuelto el error: " & _
        Err.description
    Resume salida
End Function

Private Sub ListasMaterializaciones_DblClick(Cancel As Integer)
    EditarMaterializacionSeleccionada
End Sub

Private Sub EditarMaterializacionSeleccionada()
    Dim idMaterializacion As String
    Dim fechaMaterializacion As String
    Dim idPlanContingencia As String
    Dim openArgs As String
    Dim editError As String

    On Error GoTo errores

    m_Error = ""
    If Not blnPermitidoEditar Then Exit Sub
    If Me.ListasMaterializaciones.ListIndex < 0 Then Exit Sub

    idMaterializacion = Trim$(Nz(Me.ListasMaterializaciones.Column(0), ""))
    fechaMaterializacion = Trim$(Nz(Me.ListasMaterializaciones.Column(2), ""))
    idPlanContingencia = Trim$(Nz(Me.ListasMaterializaciones.Column(6), ""))
    If idMaterializacion = "" Then Exit Sub

    If EsMaterializacionVigenteEditable( _
        m_ObjRiesgoActivo.idRiesgo, _
        idMaterializacion, _
        editError) <> EnumSiNo.Sí Then
        m_Error = editError
        If m_Error = "" Then
            m_Error = "Solo se puede editar la materialización vigente del riesgo"
        End If
        Err.Raise 1000
    End If

    openArgs = OPENARGS_EDIT_MATERIALIZACION & "|" & _
        m_ObjRiesgoActivo.idRiesgo & "|" & _
        idMaterializacion & "|" & _
        fechaMaterializacion & "|" & _
        idPlanContingencia

    LimpiarTempVarsMaterializacion
    DoCmd.OpenForm FORM_DIALOGO_MATERIALIZACION, , , , , acDialog, openArgs
    If TempVarValueMaterializacion(TEMP_MAT_ACEPTADA) = "Sí" Then
        EstablecerDatos m_Error
        If m_Error <> "" Then Err.Raise 1000
    End If
    LimpiarTempVarsMaterializacion
    Exit Sub

errores:
    LimpiarTempVarsMaterializacion
    If Err.Number <> 1000 Then
        m_Error = "Al editar la materialización se ha producido el error: " & _
            Err.description
    End If
    MsgBox m_Error, vbExclamation, "Materialización"
End Sub

Private Function TempVarValueMaterializacion(ByVal p_Nombre As String) As String
    On Error GoTo salida

    TempVarValueMaterializacion = Nz(TempVars(p_Nombre).value, "")

salida:
End Function

Private Sub LimpiarTempVarsMaterializacion()
    On Error Resume Next
    TempVars.Remove TEMP_MAT_ACEPTADA
    TempVars.Remove TEMP_MAT_FECHA
    TempVars.Remove TEMP_MAT_PLAN
    On Error GoTo 0
End Sub


Private Sub ComandoQuitarMaterializado_Click()
    
    On Error GoTo errores
    m_Error = ""
    pregunta = MostrarMsgBoxSiNoEsTest("¿Desea habilitar el riesgo?", vbExclamation + vbYesNo + vbDefaultButton2, "Habilitar el riesgo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    With m_ObjRiesgoActivo
         .MaterializacionQuitarRegistrar .EstadoEnum, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjRiesgoAlInicio = GetCachedRiesgo(.idRiesgo)
        Me.ComandoFechaMaterializado.SetFocus
        Me.ComandoQuitarMaterializado.Visible = False
        If IsDate(.FechaMaterializado) Then
            Me.FechaMaterializado = .FechaMaterializado
        Else
            Me.FechaMaterializado = Null
        End If
        
    End With
    ActualizarEtiquetaEstadoRiesgo EnumSiNo.Sí
    EstablecerDatos
    ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_NodoSeleccionado, p_VerDescripcion:=m_VerDescripcion, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "materializacion", m_Error
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
        m_Error = "Al ComandoQuitarMaterializado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoVerEnTarea_Click()
   
    On Error GoTo errores

    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If FormularioAbierto("FormCalidadRiesgoMaterializaciones") Then
        DoCmd.Close acForm, "FormCalidadRiesgoMaterializaciones", acSaveNo
    End If
    DoCmd.OpenForm "FormCalidadRiesgoMaterializaciones"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents


    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerEnTarea_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
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
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""

    Me.FechaMaterializado.Enabled = False

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
    Dim m_EditabilitySvc As MaterializacionPlanEligibilityService

    EstablecerLista p_Error
    If p_Error <> "" Then Exit Sub

    Set m_Edicion = m_ObjRiesgoActivo.Edicion
    Set m_EditabilitySvc = New MaterializacionPlanEligibilityService
    blnPermitidoEditar = m_EditabilitySvc.IsMaterializationEditable( _
        m_Edicion.EsActivo, _
        m_Edicion.Proyecto.UsuarioAutorizado)
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    If blnPermitidoEditar = True Then
        Set m_NodoSeleccionado = modFormCoordinationHelper.Coord_GetSelectedNode(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        m_VerDescripcion = modFormCoordinationHelper.Coord_GetComboVerDescripcion(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        Me.ComandoFechaMaterializado.Enabled = True
        If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Materializado Then
            Me.ComandoQuitarMaterializado.Visible = True
            Me.ComandoQuitarMaterializado.Enabled = True
            Me.FechaMaterializado = m_ObjRiesgoActivo.FechaMaterializado
        Else
            Me.ComandoQuitarMaterializado.Visible = False
            Me.FechaMaterializado = Null
        End If
    Else
        Me.ComandoFechaMaterializado.Enabled = False
        Me.ComandoQuitarMaterializado.Visible = False
        Me.FechaMaterializado = m_ObjRiesgoActivo.FechaMaterializado
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    ConfigurarListaMaterializacionesEditable
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.description
End Sub

Private Sub ConfigurarListaMaterializacionesEditable()
    With Me.ListasMaterializaciones
        .Enabled = blnPermitidoEditar
        .Locked = Not blnPermitidoEditar
        .TabStop = blnPermitidoEditar
    End With
End Sub

Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ID As Variant
    Dim m_RiesgoMat As RiesgoMaterializacion
    Dim m_Decidido As String
    Dim m_Edicion As String
    Dim m_ResumenPlan As String
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    Me.ComandoVerEnTarea.Visible = False
    Set lst = Me.ListasMaterializaciones
    lst.RowSource = "ID;Edición;FMaterializacion;Mat.;Decidido;Plan;IDPlan"
    
    Set m_ObjRiesgoActivo.colMaterializaciones = Nothing
    Set m_Col = m_ObjRiesgoActivo.colMaterializaciones
    p_Error = m_ObjRiesgoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    If EsTecnico = EnumSiNo.No Then
        Me.ComandoVerEnTarea.Visible = True
    End If
    For Each m_ID In m_Col
        Set m_RiesgoMat = m_Col(m_ID)
        If m_RiesgoMat.Decidido = EnumSiNo.Sí Then
            m_Decidido = "Sí"
        Else
            m_Decidido = "No"
        End If
        m_Edicion = m_RiesgoMat.Edicion.Edicion
        m_ResumenPlan = ResumenPlanMaterializacionLista(m_RiesgoMat.idPlanContingencia)
        
        lst.AddItem m_RiesgoMat.ID & ";" & m_Edicion & ";" & _
            m_RiesgoMat.Fecha & ";" & m_RiesgoMat.EsMaterializacion & ";" & _
            m_Decidido & ";" & EscapeMaterializacionListValue(m_ResumenPlan) & ";" & _
            m_RiesgoMat.idPlanContingencia
        Set m_RiesgoMat = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.description
    End If

End Function

Private Function ResumenPlanMaterializacionLista( _
    ByVal p_IDPlanContingencia As String _
) As String
    Dim resumenError As String

    If Trim$(Nz(p_IDPlanContingencia, "")) = "" Then Exit Function
    ResumenPlanMaterializacionLista = _
        ResumenPlanContingenciaMaterializacion(p_IDPlanContingencia, resumenError)
End Function

Private Function EscapeMaterializacionListValue(ByVal p_Value As String) As String
    EscapeMaterializacionListValue = Replace(Replace(Replace(Nz(p_Value, ""), _
        ";", ","), vbCrLf, " "), vbLf, " ")
End Function
