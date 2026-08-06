Option Compare Database
Option Explicit

Private WithEvents m_FormExpedientes As Form_FormExpedientesBusqueda
Private m_UltimoValorJuridica As String
Private m_Expediente As Expediente
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


Private Sub Aprobado_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.Aprobado = Nz(Me.Aprobado, "")
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub ComandoActualizar_Click()
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_Expediente = Constructor.getExpediente(p_IDExpediente:=Nz(Me.idExpediente, ""), p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarDatosExpediente m_Expediente, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarBordeCamposObligatorios Me
    Me.Elaborado.SetFocus
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoIrABusquedaExpedientes_Click()
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormExpedientesBusqueda") Then
        DoCmd.Close acForm, "FormExpedientesBusqueda", acSaveNo
    End If
    DoCmd.OpenForm "FormExpedientesBusqueda"
    If FormularioAbierto("FormExpedientesBusqueda") Then
        ' HR3d: WithEvents ref (legitimo). m_FormExpedientes_PalabraClave
        ' y .Filtrar ahora pasan por Coord_ExpedientesBusqueda_Filtrar.
        Set m_FormExpedientes = Forms("FormExpedientesBusqueda")
        If Nz(Me.PalabraClave, "") <> "" Then
            Call modFormCoordinationHelper.Coord_ExpedientesBusqueda_Filtrar( _
                p_PalabraClave:=Nz(Me.PalabraClave, ""), _
                p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIrABusquedaExpedientes_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
    AvanceCerrar
    
End Sub

Private Sub Elaborado_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.Elaborado = Nz(Me.Elaborado, "")
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    RellenarDatosAlObjeto m_Error
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
        m_Error = "Al Form_Unload se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
    AvanceCerrar
   
End Sub

Private Sub m_FormExpedientes_Seleccionado(m_Exp As Expediente)
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    RellenarDatosExpediente m_Exp
    If Nz(Me.idExpediente, "") = "" Then
        Me.ComandoActualizar.Enabled = False
    Else
        Me.ComandoActualizar.Enabled = True
    End If
    Me.PalabraClave = Null
    RellenarBordeCamposObligatorios Me
    Me.Elaborado.SetFocus
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormExpedientes_Seleccionado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
    AvanceCerrar
   
End Sub

Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Me.ComandoBuscarExpediente.Enabled = False Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        Me.ComandoBuscarExpediente.SetFocus
        Me.PalabraClave.SetFocus
        'Debug.Print Me.Pass
        ' Código que deseas ejecutar cuando se pulse Enter
        ComandoBuscarExpediente_Click
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    End If
End Sub


Private Sub ParaInformeAvisos_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.ParaInformeAvisos = Nz(Me.ParaInformeAvisos, "")
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub Revisado_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.Revisado = Nz(Me.Revisado, "")
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
   ' Debug.Print Me.IDExpediente.BackColor
    AvanceCerrar
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
    AvanceCerrar
    
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub




Private Sub ComandoBuscarExpediente_Click()
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Expediente As Expediente
    
    Dim m_PalabraClave As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    m_PalabraClave = Nz(Me.PalabraClave, "")
    If m_PalabraClave = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    Set m_Col = Constructor.getExpedientesBusqueda(p_PalabraClave:=m_PalabraClave, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Me.idExpediente = Null
        Me.Juridica = Null
        Me.CodExp = Null
        Me.Nemotecnico = Null
        Me.FechaInicioContrato = Null
        Me.FechaFinContrato = Null
        Me.OrganoContratacion = Null
        Me.ResponsableCalidad = Null
        Me.CodigoDocumento = Null
        Me.CorreoRAC = Null
        If Nz(Me.idExpediente, "") = "" Then
            Me.ComandoActualizar.Enabled = False
        Else
            Me.ComandoActualizar.Enabled = True
        End If
        RellenarBordeCamposObligatorios Me
        ComandoIrABusquedaExpedientes_Click
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    If m_Col.Count = 1 Then
        For Each m_ID In m_Col
            Set m_Expediente = m_Col(m_ID)
            RellenarDatosExpediente m_Expediente, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If Nz(Me.idExpediente, "") = "" Then
                Me.ComandoActualizar.Enabled = False
            Else
                Me.ComandoActualizar.Enabled = True
            End If
        Next
        Me.PalabraClave = Null
        RellenarBordeCamposObligatorios Me
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If FormularioAbierto("FormExpedientesBusqueda") Then
        DoCmd.Close acForm, "FormExpedientesBusqueda", acSaveNo
    End If
    DoCmd.OpenForm "FormExpedientesBusqueda"
    If FormularioAbierto("FormExpedientesBusqueda") Then
        ' HR3d: WithEvents ref (legitimo) + keyword/filter via Coord.
        Set m_FormExpedientes = Forms("FormExpedientesBusqueda")
        Call modFormCoordinationHelper.Coord_ExpedientesBusqueda_Filtrar( _
            p_PalabraClave:=Nz(Me.PalabraClave, ""), _
            p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBuscarExpediente_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

' HR4-sister (HR3d follow-up #9, HIGH-PRIORITY - 2026-07-01):
'   Era Public Function. Ahora Private. Body es thin adapter que delega
'   al helper modGestionRiesgosDatosGeneralesHelper. Sin external callers
'   en codegraph blast radius (1 internal caller in m_FormExpedientes_Seleccionado
'   y 2 en ComandoActualizar_Click/ComandoBuscarExpediente). Hard rule 4.
Private Function RellenarDatosExpediente(Optional p_Expediente As Expediente, Optional ByRef p_Error As String) As String

    On Error GoTo errores

    ' HR4-sister: delega al helper. El helper hace el render a los Me.* controls.
    modGestionRiesgosDatosGeneralesHelper.DatosGenerales_RellenarDatosExpediente _
        Me, p_Expediente, m_ObjProyectoAlInicio, p_Error

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo RellenarDatosExpediente ha devuelto el error: " & vbNewLine & Err.Description

    End If
End Function
' HR4-sister (HR3d follow-up #9, #5 - 2026-07-01):
'   Era Public Function. Ahora Private. Body es thin adapter que delega
'   al helper. 1 internal caller in Form_Unload. Hard rule 4.
Private Function RellenarDatosAlObjeto(Optional ByRef p_Error As String) As String


    On Error GoTo errores

    ' HR4-sister: delega al helper. El helper lee Me.* y escribe al proyecto.
    modGestionRiesgosDatosGeneralesHelper.DatosGenerales_RellenarDatosAlObjeto _
        Me, m_ObjProyectoActivo, p_Error

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo RellenarDatosAlObjeto ha devuelto el error: " & vbNewLine & Err.Description

    End If
End Function

' HR4-sister (HR3d follow-up #9, HIGH-PRIORITY - 2026-07-01):
'   Era Public Function. Ahora Private. Body es thin UI adapter que
'   delega a 3 Private Subs (EstablecerDatos_CargarPermisos,
'   _AplicarEstado, _RenderizarControles) que ya existian.
'   1 internal caller in Form_Load. Hard rule 4: no external callers.
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""

    Me.PalabraClave.Enabled = True
    Me.PalabraClave.Locked = False
    Me.ComandoIrABusquedaExpedientes.Enabled = True
    Me.ComandoBuscarExpediente.Enabled = True

    Call EstablecerDatos_CargarPermisos(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_AplicarEstado(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_RenderizarControles(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub EstablecerDatos_CargarPermisos(ByRef p_Error As String)
    On Error GoTo errores
    If m_EsAlta = EnumSiNo.No Then
        If m_ObjProyectoAlInicio Is Nothing Then
            Set m_ObjProyectoAlInicio = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, p_Error)
            If p_Error <> "" Then Err.Raise 1000
            If m_ObjProyectoAlInicio Is Nothing Then
                p_Error = "No se conoce la gestión de riesgos"
                Err.Raise 1000
            End If
        End If
    End If
    If Not m_ObjProyectoAlInicio Is Nothing Then
        If Not m_ObjProyectoAlInicio.EdicionActiva Is Nothing Then
            If Not m_ObjProyectoAlInicio.EdicionActiva.FechaPublicacion <> "" Then
                Me.PalabraClave.Enabled = False
                Me.PalabraClave.Locked = True
                Me.ComandoIrABusquedaExpedientes.Enabled = False
                Me.ComandoBuscarExpediente.Enabled = False
            End If
        End If
    End If
    If Nz(Me.idExpediente, "") = "" Then
        Me.ComandoActualizar.Enabled = False
    Else
        Me.ComandoActualizar.Enabled = True
    End If
    If m_EsAlta = Empty Then
        p_Error = "No se sabe si es para alta o no"
        Err.Raise 1000
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    Me.Elaborado = ""
    Me.Revisado = ""
    Me.Aprobado = ""
    EstablecerCombos p_Error
    If p_Error <> "" Then Exit Sub
    ' HR3d: cross-form read Form_FormGestionRiesgos.blnPermitidoEditar ->
    '   Coord_GetPermitidoEditar (reuso, entry existente).
    If modFormCoordinationHelper.Coord_GetPermitidoEditar(p_Error) = False Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    If m_ObjProyectoActivo Is Nothing Then
        RellenarBordeCamposObligatorios Me
        Exit Sub
    End If
    With m_ObjProyectoActivo
        RellenarDatosExpediente .Expediente, p_Error
        If p_Error <> "" Then Exit Sub
        If IsDate(.FechaCierre) Then Me.FechaCierre = .FechaCierre
        If .ParaInformeAvisos = "" Then .ParaInformeAvisos = "Sí"
        Me.ParaInformeAvisos = .ParaInformeAvisos
        If .Elaborado <> "" Then Me.Elaborado = .Elaborado Else Me.Elaborado = Null
        If .Revisado <> "" Then Me.Revisado = .Revisado Else Me.Revisado = Null
        If .Aprobado <> "" Then Me.Aprobado = .Aprobado Else Me.Aprobado = Null
    End With
    RellenarBordeCamposObligatorios Me
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.description
End Sub

Private Function EstablecerCombos(Optional ByRef p_Error As String) As String
    
    Dim m_UsuarioRed As Variant
    Dim m_ObjUsuario As Usuario
    Dim m_ColUsuariosTecnicos As Scripting.Dictionary
   
    Dim m_SoloActivos As EnumSiNo
    On Error GoTo errores
    
    
    
    
    
    Me.ParaInformeAvisos.RowSource = "Sí;No"
    Me.ParaInformeAvisos.Requery
    
   
    
    
    Me.Elaborado.RowSource = ""
    Me.Elaborado.Requery
    Me.Revisado.RowSource = ""
    Me.Revisado.Requery
    Me.Aprobado.RowSource = ""
    Me.Aprobado.Requery
    If m_ObjProyectoAlInicio Is Nothing Then
        m_SoloActivos = EnumSiNo.Sí
    Else
        m_SoloActivos = EnumSiNo.No
    End If
    
    EstablecerComboTecnicos Me.Elaborado, m_SoloActivos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerComboTecnicos Me.Revisado, m_SoloActivos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerComboTecnicos Me.Aprobado, m_SoloActivos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.description
        
    End If
    
    
End Function



Private Function EstablecerComboTecnicos( _
                                            ByRef cmb As ComboBox, _
                                            Optional ByRef p_SoloActivos As EnumSiNo = EnumSiNo.Sí, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_objColUsuarios As Scripting.Dictionary
    Dim m_UsuarioRed  As Variant
    Dim m_ObjUsuario As Usuario
    
    On Error GoTo errores
    cmb.RowSource = ""
    If p_SoloActivos = Empty Then
        p_SoloActivos = EnumSiNo.Sí
    End If
    If p_SoloActivos = EnumSiNo.Sí Then
        Set m_objColUsuarios = m_ObjEntorno.ColUsuariosTecnicosActivos
    Else
        Set m_objColUsuarios = m_ObjEntorno.ColUsuariosTecnicos
    End If
    If m_objColUsuarios Is Nothing Then
        Exit Function
    End If
    
    For Each m_UsuarioRed In m_objColUsuarios
        Set m_ObjUsuario = m_objColUsuarios(m_UsuarioRed)
        cmb.AddItem m_ObjUsuario.Nombre
        Set m_ObjUsuario = Nothing
siguiente:
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboTecnicos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
