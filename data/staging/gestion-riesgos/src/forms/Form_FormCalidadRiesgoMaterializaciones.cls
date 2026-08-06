Attribute VB_Name = "Form_FormCalidadRiesgoMaterializaciones"
Option Compare Database
Option Explicit
Private WithEvents m_FormNC As Form_FormRiesgoNC
Private m_Error As String
Private m_NC As NC
Private blnCalculadoNC As Boolean
Private m_RiesgoMatSeleccionado As RiesgoMaterializacion
Private m_EstadoRiesgo As EnumRiesgoEstado

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




Private Sub ComandoNoParaNC_Click()
    Dim m_ID As String
    On Error GoTo errores

    m_Error = ""
    If m_RiesgoMatSeleccionado Is Nothing Then
        Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_RiesgoMatSeleccionado.id <> Nz(Me.ListasMaterializaciones.Column(0), "") Then
            Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_RiesgoMatSeleccionado Is Nothing Then
        m_Error = "Ha de seleccionar un elemento de la lista"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea no generar ni vincular esta materialización con una NC ?", vbExclamation + vbYesNo + vbDefaultButton2, "No para NC")

    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_RiesgoMatSeleccionado.RegistrarParaNONC , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Not m_RiesgoMatSeleccionado Is Nothing Then
        m_ID = m_RiesgoMatSeleccionado.id
    End If
    Set m_ObjRiesgoActivo.ColMaterializaciones = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ID = "" Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    SeleccionarDeLista m_ID, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
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
        m_Error = "Al ComandoNoParaNC_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRevocarDecision_Click()
    
    Dim m_ID As String
    On Error GoTo errores

    m_Error = ""
    If m_RiesgoMatSeleccionado Is Nothing Then
        Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_RiesgoMatSeleccionado.id <> Nz(Me.ListasMaterializaciones.Column(0), "") Then
            Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_RiesgoMatSeleccionado Is Nothing Then
        m_Error = "Ha de seleccionar un elemento de la lista superior"
        Err.Raise 1000
    End If
    
    
    pregunta = MsgBox("¿Desea dejarlo para decidirlo en otro momento ?", vbExclamation + vbYesNo + vbDefaultButton2, "Revocar decisión")

    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents


    m_RiesgoMatSeleccionado.RegistrarPorDecidir m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Not m_RiesgoMatSeleccionado Is Nothing Then
        m_ID = m_RiesgoMatSeleccionado.id
    End If
    Set m_ObjRiesgoActivo.ColMaterializaciones = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ID = "" Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    SeleccionarDeLista m_ID, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
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
        m_Error = "Al ComandoRevocarDecision_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoSalir_Click()
     DoCmd.Close acForm, Me.Name
End Sub

Private Sub ComandoVerPlanesContingencia_Click()
    
    Dim m_Riesgo As riesgo
    On Error GoTo errores

    
    m_Error = ""
    If m_RiesgoMatSeleccionado Is Nothing Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.OpenForm "FormRiesgosGestion", openArgs:=m_RiesgoMatSeleccionado.Edicion.IDEdicion
    Set m_Riesgo = Constructor.getRiesgo(p_IDEdicion:=m_RiesgoMatSeleccionado.Edicion.IDEdicion, p_CodigoRiesgo:=m_RiesgoMatSeleccionado.riesgo.CodigoRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Riesgo Is Nothing Then
        m_Error = "No se ha encontrado el Riesgo"
        Err.Raise 1000
    End If
    ' HR3e (2026-07-01): reemplazar acceso directo a Form_FormRiesgosGestion
    ' .SeleccionarNodo por helper modFormCoordinationHelper.
    Call modFormCoordinationHelper.Coord_SeleccionarNodo( _
        p_Objeto:=m_Riesgo, _
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
        m_Error = "Al ComandoVerPlanesContingencia_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    AjustarTamaño Me
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
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub



Private Function VisibilizarDesvisibilizarGrupoNC(p_Activar As EnumSiNo, Optional ByRef p_Error As String) As String

    If p_Activar = EnumSiNo.Sí Then
        Me.DescripcionDecisionMaterializado.Visible = True
        
       
        Me.ComandoVincularNC.Visible = True
        Me.ComboNC.Visible = True
        Me.ComandoNoParaNC.Visible = True
    Else
       
        Me.DescripcionDecisionMaterializado.Visible = False
        
        Me.ComandoVincularNC.Visible = False
        Me.ComboNC.Visible = False
        Me.ComandoNoParaNC.Visible = False
        Me.ComandoRevocarDecision.Visible = True
    End If
End Function
Private Function EstablecerBotoneraGrupoNC(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    
    
    
    Me.DescripcionDecisionMaterializado.Enabled = False
    
    Me.ComandoVincularNC.Enabled = False
    Me.ComboNC.Enabled = False
    Me.ComandoNoParaNC.Enabled = False
    Me.ComandoRevocarDecision.Visible = False
   

    
    
    If m_RiesgoMatSeleccionado Is Nothing Then
        Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_RiesgoMatSeleccionado.id <> Nz(Me.ListasMaterializaciones.Column(0), "") Then
            Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_RiesgoMatSeleccionado Is Nothing Then
        Exit Function
    End If
    'Me.FechaMaterializado.SetFocus
    VisibilizarDesvisibilizarGrupoNC EnumSiNo.Sí
    With m_RiesgoMatSeleccionado
        If .EsMaterializacionCalcuado <> EnumSiNo.Sí Then
            Exit Function
        End If
        If Not IsDate(.FechaDecison) Then
            Me.DescripcionDecisionMaterializado.value = "Decisión por tomar"
            
            Me.ComandoVincularNC.Enabled = True
            Me.ComandoNoParaNC.Enabled = True
            Me.ComboNC.Enabled = True
            Me.ComandoRevocarDecision.Visible = False
            Exit Function
        End If
        If .ParaNCCalculado = EnumSiNo.No Then
            Me.DescripcionDecisionMaterializado.value = "Decisión tomada. No se genera NC"
            Me.ComandoRevocarDecision.Visible = True
            Me.ComandoRevocarDecision.Enabled = True
            Exit Function
        End If
        Set m_NC = .NC
        If Not m_NC Is Nothing Then
            Me.DescripcionDecisionMaterializado.value = "Decisión tomada. Se vinculó a la NC: " & m_NC.CodigoNoConformidad
            Me.ComandoRevocarDecision.Visible = True
            Me.ComandoRevocarDecision.Enabled = True
            Exit Function
        End If
    End With
    
    
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerBotoneraGrupoNC ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum
    
    VisibilizarDesvisibilizarGrupoNC EnumSiNo.No
    Me.ListasMaterializaciones.RowSource = "ID;Edición;FMaterializacion;Mat.;Decidido"
    
    EstablecerComboNC m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerBotoneraGrupoNC m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerLista p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ID As Variant
    Dim m_RiesgoMat As RiesgoMaterializacion
    Dim m_pteDecidir As String
    Dim m_Edicion As String
    Dim m_NC As String
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    
    Set lst = Me.ListasMaterializaciones
    lst.RowSource = "ID;Edición;FMaterializacion;Mat.;pte.decidir NC;NC"
    
    Set m_ObjRiesgoActivo.ColMaterializaciones = Nothing
    Set m_Col = m_ObjRiesgoActivo.ColMaterializaciones
    p_Error = m_ObjRiesgoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_RiesgoMat = m_Col(m_ID)
        If m_RiesgoMat.Decidido = EnumSiNo.Sí Then
            m_pteDecidir = "No"
        Else
            m_pteDecidir = "Sí"
        End If
        If m_RiesgoMat.NC Is Nothing Then
            m_NC = "--"
        Else
            m_NC = m_RiesgoMat.NC.CodigoNoConformidad
        End If
        m_Edicion = m_RiesgoMat.Edicion.Edicion
        
        lst.AddItem m_RiesgoMat.id & ";" & m_Edicion & ";" & m_RiesgoMat.Fecha & ";" & m_RiesgoMat.EsMaterializacion & ";" & _
                    m_pteDecidir & ";" & m_NC
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
Private Sub ComandoVincularNC_Click()
    Dim m_NCSeleccionada As NC
    Dim m_ID As String
    On Error GoTo errores

    m_Error = ""
    If m_RiesgoMatSeleccionado Is Nothing Then
        Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_RiesgoMatSeleccionado.id <> Nz(Me.ListasMaterializaciones.Column(0), "") Then
            Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Nz(Me.ListasMaterializaciones.Column(0), ""), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_RiesgoMatSeleccionado Is Nothing Then
        m_Error = "Ha de seleccionar un elemento de la lista superior"
        Err.Raise 1000
    End If
    
    If Nz(Me.ComboNC, "") = "" Then
        m_Error = "Ha de seleccionar una NC del cuadro desplegable"
        Err.Raise 1000
    End If
    Set m_NCSeleccionada = Constructor.getNC(Me.ComboNC.Column(0), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_NCSeleccionada Is Nothing Then
        m_Error = "Ha de seleccionar una NC del cuadro desplegable"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea vincularlo a la NC seleccionada ?", vbExclamation + vbYesNo + vbDefaultButton2, "Vincular a NC")

    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    

    m_RiesgoMatSeleccionado.VincularNC m_NCSeleccionada, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Not m_RiesgoMatSeleccionado Is Nothing Then
        m_ID = m_RiesgoMatSeleccionado.id
    End If
    Set m_ObjRiesgoActivo.ColMaterializaciones = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ID = "" Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    SeleccionarDeLista m_ID, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
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
        m_Error = "Al ComandoVincularNC_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ListasMaterializaciones_Click()
    
    On Error GoTo errores

    m_Error = ""
    DoCmd.Hourglass True
    Me.DescripcionDecisionMaterializado = Null
    Me.ComandoRevocarDecision.Enabled = False
    
    Me.ComandoVincularNC.Enabled = False
    Me.ComandoNoParaNC.Enabled = False
    Me.ComboNC.Enabled = False
    Me.ComandoVerPlanesContingencia.Enabled = False
    If Nz(Me.ListasMaterializaciones.Column(0), "") = "" Then
        Set m_RiesgoMatSeleccionado = Nothing
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_RiesgoMatSeleccionado Is Nothing Then
        Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Me.ListasMaterializaciones.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_RiesgoMatSeleccionado.id <> Me.ListasMaterializaciones.Column(0) Then
            Set m_RiesgoMatSeleccionado = Constructor.getRiesgoMaterializado(Me.ListasMaterializaciones.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_RiesgoMatSeleccionado Is Nothing Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    Me.ComandoVerPlanesContingencia.Enabled = True
    EstablecerBotoneraGrupoNC m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListasMaterializaciones_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub m_FormNC_NCRegistrada()
    
    Dim m_ID As String
    
    On Error GoTo errores

    m_Error = ""
    DoCmd.Hourglass True
    If Not m_RiesgoMatSeleccionado Is Nothing Then
        m_ID = m_RiesgoMatSeleccionado.id
    End If
    Set m_ObjRiesgoActivo.ColMaterializaciones = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ID = "" Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    SeleccionarDeLista m_ID, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
       Err.Raise 1000
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormNC_NCRegistrada se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Function EstablecerComboNC(Optional ByRef p_Error As String) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_NC As NC
    Dim m_ID As Variant
    Dim cmb As ComboBox
    
    On Error GoTo errores

    p_Error = ""
    If m_ObjRiesgoAlInicio Is Nothing Then
        Exit Function
    End If
    If Me.ListasMaterializaciones.Visible = False Then
        Exit Function
    End If
    
    Set cmb = Me.ComboNC
    cmb.RowSource = "ID;COD;Descripción"
    Set m_Col = Constructor.getNCs(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_NC = m_Col(m_ID)
        With m_NC
            cmb.AddItem .IDNoConformidad & ";" & .CodigoNoConformidad & ";" & .Descripcion
        End With
        Set m_NC = Nothing
    Next
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboNC ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function SeleccionarDeLista(p_Id As String, Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    On Error GoTo errores
    
    
    If p_Id = "" Then
        Exit Function
    End If
    Set lst = Me.ListasMaterializaciones
    If lst.ListCount < 2 Then
        Exit Function
    End If
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = p_Id Then
            lst.Selected(i) = True
            ListasMaterializaciones_Click
            Exit Function
        End If
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarDeLista ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
