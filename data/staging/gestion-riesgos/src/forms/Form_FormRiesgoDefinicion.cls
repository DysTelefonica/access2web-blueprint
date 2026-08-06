Option Compare Database
Option Explicit

Private m_FamiliaUltima As String
Private m_PalabraClaveUltima As String
Private m_ObjRiesgoBibliotecaSeleccionado As RiesgoBiblioteca
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

Private Sub CausaRaiz_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    
    
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenarBordeCamposObligatorios Me
        EstablecerColorPestañasRiesgo
    End If
End Sub

Private Sub ComandoLimpiarFamilia_Click()
    On Error Resume Next
    Me.ComboFamilia = Null
     m_FamiliaUltima = ""
    EstablecerListaRiesgosBiblioteca
    
End Sub

Private Sub ComandoLimpiarPalabraClave_Click()
    On Error Resume Next
    Me.PalabraClave = Null
    m_PalabraClaveUltima = ""
   
    EstablecerListaRiesgosBiblioteca
    
End Sub

Private Sub ComboFamilia_BeforeUpdate(Cancel As Integer)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Nz(Me.ComboFamilia, "") = m_FamiliaUltima Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    EstablecerListaRiesgosBiblioteca m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_FamiliaUltima = Nz(Me.ComboFamilia, "")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComboFamilia_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Descripcion_Exit(Cancel As Integer)
    On Error Resume Next
    
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenarBordeCamposObligatorios Me
        EstablecerColorPestañasRiesgo
    End If
End Sub

Private Sub DetectadoPor_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenarBordeCamposObligatorios Me
        EstablecerColorPestañasRiesgo
    End If
End Sub



Private Sub EntidadDetecta_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    RellenarComboDetectadoPor Nz(Me.EntidadDetecta, "")
   
   If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenarBordeCamposObligatorios Me
        EstablecerColorPestañasRiesgo
    End If
End Sub


Private Sub FechaDetectado_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    
   
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenarBordeCamposObligatorios Me
        EstablecerColorPestañasRiesgo
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
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub ListaFiltrados_Click()
    
    Dim lst As ListBox
    Dim i As Long
    Dim m_VarItem As Variant
    On Error GoTo errores
    
    
    m_Error = ""
    Set lst = Me.ListaFiltrados
    If Me.CausaRaiz.Visible = False Then
        Exit Sub
    End If
    

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Me.ListaFiltrados.Column(0), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjRiesgoBibliotecaSeleccionado.IDRiesgoTipo <> Me.ListaFiltrados.Column(0) Then
            Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Me.ListaFiltrados.Column(0), , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    With m_ObjRiesgoBibliotecaSeleccionado
        'Me.CodRiesgoBiblioteca = .Codigo
        If .RiesgoParaRetipificar = EnumSiNo.Sí Then
            Me.TextoPendienteRetipificacion.Visible = True
            Me.TextoPendienteRetipificacion.value = "Pendiente de Calidad"
            Me.Descripcion.Enabled = True
        Else
            Me.TextoPendienteRetipificacion.Visible = False
             Me.Descripcion.Enabled = False
        End If
        Me.Descripcion = .Descripcion
        If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
            RellenarBordeCamposObligatorios Me
        End If
        m_ObjRiesgoActivo.CodRiesgoBiblioteca = .codigo
    End With
    
    
    m_ObjRiesgoActivo.Descripcion = Nz(Me.Descripcion, "")
    
       
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Function EstablecerListaRiesgosBiblioteca( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim lst As ListBox
    Dim m_ID As Variant
    Dim m_ObjRiesgoBiblioteca As RiesgoBiblioteca
    Dim m_Col As Scripting.Dictionary
    Dim m_Descripcion As String
    On Error GoTo errores
    
    p_Error = ""
    Set lst = Me.ListaFiltrados
    lst.RowSource = ""
    Set m_Col = Constructor.getRiesgosBiblioteca(Nz(Me.ComboFamilia, ""), Nz(Me.PalabraClave, ""), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjRiesgoBiblioteca = m_Col(m_ID)
        With m_ObjRiesgoBiblioteca
            m_Descripcion = Replace(.Descripcion, ";", ":")
            If Len(m_Descripcion) > 45 Then
                m_Descripcion = Left(m_Descripcion, 45) & " ..."
            
            
            End If
            lst.AddItem .IDRiesgoTipo & ";" & m_Descripcion
            Set m_ObjRiesgoBiblioteca = Nothing
        End With
    Next
    If Me.ListaFiltrados.Enabled = True Then
        If lst.ListCount = 1 Then
            lst.Selected(0) = True
            ListaFiltrados_Click
        End If
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerListaRiesgosBiblioteca ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function SeleccionarValorEnLIsta(p_Valor As String, Optional p_HaciendoClick As EnumSiNo = EnumSiNo.Sí, Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    On Error GoTo errores
    
    If p_Valor = "" Then
        Exit Function
    End If
    Set lst = Me.ListaFiltrados
    For i = 0 To lst.ListCount - 1
        If Nz(lst.Column(0, i), "") = p_Valor Then
            lst.Selected(i) = True
            If p_HaciendoClick = EnumSiNo.Sí Then
                ListaFiltrados_Click
            End If
            
            Exit For
        End If
    Next
    
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarValorEnLIsta ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    On Error GoTo errores
    p_Error = ""

    Me.Descripcion.Enabled = False
    Me.CausaRaiz.Enabled = False
    Me.Origen.Enabled = False
    Me.ListaFiltrados.Enabled = False
    Me.ComboFamilia.Enabled = False
    Me.ComandoLimpiarFamilia.Enabled = False
    Me.PalabraClave.Enabled = False
    Me.ComandoLimpiarPalabraClave.Enabled = False

    Call EstablecerDatos_CargarPermisos(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_AplicarEstado(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_RenderizarControles(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Sub EstablecerDatos_CargarPermisos(ByRef p_Error As String)
    On Error GoTo errores
    Dim m_RiesgoBiblioteca As RiesgoBiblioteca
    Dim m_PriorizacionCorrecta As EnumSiNo

    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No se puede determinar la edición activa"
        Err.Raise 1000
    End If
    If Form_FormRiesgo.m_PermitidoEditar = Empty Then
        With m_ObjEdicionActiva
            If .EsActivo = EnumSiNo.No Then
               Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.No
            Else
                If .Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
                    Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.Sí
                Else
                    Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.No
                End If
            End If
        End With
    End If
    If m_EsAlta = EnumSiNo.Sí Then
        If Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.No Then
            m_Error = "Es un alta y el usuario no tiene permitida esa operación"
            Err.Raise 1000
        End If
        If m_ObjRiesgoActivo Is Nothing Then
            Set m_ObjRiesgoActivo = New riesgo
            m_ObjRiesgoActivo.IDEdicion = m_ObjEdicionActiva.IDEdicion
            Set m_ObjRiesgoAlInicio = Nothing
        End If
    Else
        If m_ObjRiesgoActivo Is Nothing Then
            m_Error = "Es una edición y no hay datos"
            Err.Raise 1000
        End If
        Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo, , , m_Error)
        If m_Error <> "" Then Err.Raise 1000
        If m_ObjRiesgoAlInicio Is Nothing Then
            m_Error = "Es una edición y no hay datos"
            Err.Raise 1000
        End If
    End If
    If m_EsAlta = EnumSiNo.No Then
        If m_ObjRiesgoActivo Is Nothing Then
            p_Error = "No es un alta y no se sabe qué riesgo es"
            Err.Raise 1000
        End If
        If m_ObjRiesgoAlInicio Is Nothing Then
            Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo, , , p_Error)
            If p_Error <> "" Then Err.Raise 1000
            If m_ObjRiesgoAlInicio Is Nothing Then
                p_Error = "No es un alta y no se sabe qué riesgo es"
                Err.Raise 1000
            End If
        End If
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.Description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    Dim m_PriorizacionCorrecta As EnumSiNo
    Dim m_RiesgoBiblioteca As RiesgoBiblioteca

    If Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.Sí Then
        Me.CausaRaiz.Enabled = True
        Me.Origen.Enabled = True
        Me.CausaRaiz.Enabled = True
        Me.ListaFiltrados.Enabled = True
        Me.ComboFamilia.Enabled = True
        Me.ComandoLimpiarFamilia.Enabled = True
        Me.PalabraClave.Enabled = True
        Me.ComandoLimpiarPalabraClave.Enabled = True
    End If
    With m_ObjRiesgoActivo
        If .Descripcion <> "" Then Me.Descripcion = .Descripcion
        If .CausaRaiz <> "" Then Me.CausaRaiz = .CausaRaiz
        If IsDate(.FechaDetectado) Then Me.FechaDetectado = .FechaDetectado
        If .EntidadDetecta <> "" Then Me.EntidadDetecta = .EntidadDetecta
        If .DetectadoPor <> "" Then Me.DetectadoPor = .DetectadoPor
        If .Origen <> "" Then Me.Origen = .Origen
        If IsNumeric(.Priorizacion) Then Me.Priorizacion = .Priorizacion
        If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
            If .IDRiesgo = "" Then
                m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, Nz(Me.Priorizacion, ""))
            Else
                m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, Nz(Me.Priorizacion, ""), .IDRiesgo)
            End If
        End If
    End With
    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.Description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    Dim m_RiesgoBiblioteca As RiesgoBiblioteca
    Dim m_PriorizacionCorrecta As EnumSiNo

    EstablecerComboOrigen Me.Origen, p_Error
    If p_Error <> "" Then Exit Sub
    EstablecerListaRiesgosBiblioteca p_Error
    If p_Error <> "" Then Exit Sub
    EstablecerComboFamilia Me.ComboFamilia, p_Error
    If p_Error <> "" Then Exit Sub
    Me.lblCausaRaiz.Visible = True
    Me.CausaRaiz.Visible = True
    Me.Descripcion.Enabled = False
    With m_ObjRiesgoActivo
        If .CodRiesgoBiblioteca <> "" Then
            If m_RiesgoBiblioteca Is Nothing Then
                Set m_RiesgoBiblioteca = Constructor.getRiesgoBiblioteca(, .CodRiesgoBiblioteca, p_Error)
                If p_Error <> "" Then Exit Sub
            Else
                If m_RiesgoBiblioteca.codigo <> .CodRiesgoBiblioteca Then
                    Set m_RiesgoBiblioteca = Constructor.getRiesgoBiblioteca(, .CodRiesgoBiblioteca, p_Error)
                    If p_Error <> "" Then Exit Sub
                End If
            End If
            If Not m_RiesgoBiblioteca Is Nothing Then
                SeleccionarValorEnLIsta m_RiesgoBiblioteca.IDRiesgoTipo, EnumSiNo.No
                If m_RiesgoBiblioteca.RiesgoParaRetipificar = EnumSiNo.Sí Then
                    Me.TextoPendienteRetipificacion.Visible = True
                    Me.TextoPendienteRetipificacion.value = "Pendiente de Calidad"
                    Me.Descripcion.Enabled = True
                    Me.Descripcion.Locked = False
                    Me.Descripcion.SetFocus
                Else
                    Me.TextoPendienteRetipificacion.Visible = False
                End If
            End If
        End If
    End With
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenarBordeCamposObligatorios Me
        If m_PriorizacionCorrecta = EnumSiNo.Sí Then
            EstablecerControlCombo Me.Priorizacion, EnumSiNo.Sí
        Else
            EstablecerControlCombo Me.Priorizacion, EnumSiNo.No
        End If
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.Description
End Sub

Private Function RellenarComboDetectadoPor(Optional p_EntidadDetecta As String, Optional ByRef p_Error As String) As String

    Dim cmb As ComboBox
    Dim ColSuministradores As Scripting.Dictionary
    Dim m_Suministrador As Suministrador
    Dim m_ID As Variant
    Dim m_NombreAMostrar As String
    
    
    Dim m_ValorActual As String
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_objColRiesgos As Scripting.Dictionary
    Dim m_RowSource As String
    On Error GoTo errores
    
    Set cmb = Me.DetectadoPor
    m_ValorActual = Nz(cmb.value, "")
    cmb.value = Null
    cmb.RowSource = ""
    
    If p_EntidadDetecta <> "Persona" And p_EntidadDetecta <> "Empresa" Then
        Exit Function
    End If
    If p_EntidadDetecta = "Persona" Then
        
        m_RowSource = m_ObjEdicionActiva.Proyecto.CadenaNombreAutorizados
        p_Error = m_ObjEdicionActiva.Proyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        cmb.RowSource = m_RowSource

        If m_ValorActual <> "" Then
            cmb.value = m_ValorActual
            
        End If
        
        Exit Function
    End If
    
    Set ColSuministradores = m_ObjRiesgoActivo.Edicion.ColSuministradores
    If ColSuministradores Is Nothing Then
        Exit Function
    End If
    For Each m_ID In ColSuministradores
        Set m_Suministrador = ColSuministradores(m_ID)
        If m_Suministrador.Nemotecnico = "" Then
            m_NombreAMostrar = m_Suministrador.Nombre
        Else
            m_NombreAMostrar = m_Suministrador.Nemotecnico
        End If
        cmb.AddItem m_NombreAMostrar
        Set m_Suministrador = Nothing
    Next
    cmb.value = m_ValorActual
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboDetectadoPor ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Sub Origen_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        RellenarBordeCamposObligatorios Me
        EstablecerColorPestañasRiesgo
    End If
End Sub


Private Sub PalabraClave_Exit(Cancel As Integer)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Nz(Me.PalabraClave, "") = m_PalabraClaveUltima Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    EstablecerListaRiesgosBiblioteca m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_PalabraClaveUltima = Nz(Me.PalabraClave, "")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al PalabraClave_Exit se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub Priorizacion_Change()
    Dim m_PriorizacionCorrecta As EnumSiNo
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Me.FechaDetectado.SetFocus
    Me.Priorizacion.SetFocus
   
    If m_EsAlta = EnumSiNo.Sí Then
    
        m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, _
                                    Nz(Me.Priorizacion, ""))
    Else
        If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
            m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, _
                                    Nz(Me.Priorizacion, ""), m_ObjRiesgoActivo.IDRiesgo)
        End If
        
    End If
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        EstablecerColorPestañasRiesgo p_PriorizacionCorrecta:=m_PriorizacionCorrecta
    End If
    If Nz(Me.Priorizacion, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        'm_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjRiesgoActivo.IDEdicion, Me.Priorizacion, m_ObjRiesgoActivo.IDRiesgo)
        If m_PriorizacionCorrecta = EnumSiNo.No Then
            If Not m_ObjRiesgoAlInicio Is Nothing Then
                If IsNumeric(m_ObjRiesgoAlInicio.Priorizacion) Then
                    Me.Priorizacion = m_ObjRiesgoAlInicio.Priorizacion
                    m_ObjRiesgoActivo.Priorizacion = m_ObjRiesgoAlInicio.Priorizacion
                    
                    EstablecerColorPestañasRiesgo
                End If
            Else
                
                Me.Priorizacion = Null
            End If
            
            
            m_Error = "Esa priorización no es correcta. Puede que sea superior a los riesgos no retirados, o no sea un número válido, o esté repetida"
            Err.Raise 1000
        Else
           
            EstablecerControlCombo Me.Priorizacion, EnumSiNo.Sí
        End If
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Priorizacion_Change se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
