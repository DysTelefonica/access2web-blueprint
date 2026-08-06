Option Compare Database
Option Explicit

Private m_ObjRiesgoExtSeleccionado As RiesgoExterno
Private WithEvents m_FormRiesgosBibliotecaGestion As Form_FormRiesgosBibliotecaGestion
Private WithEvents m_FormRiesgosExt As Form_FormRiesgoExternoDetalle
Private blnPermitidoEscribir As Boolean
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





Public Sub ComandoActualizar_Click()
    
    Dim lst As ListBox
    Dim i As Integer
    Dim m_IDRiesgoExtSeleccionado As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set lst = Me.ListaRiesgosOferta
    lst.RowSource = "IDRiesgoExt;Código;descripción;Se Trasl.;Cód Riesgo"
    m_IDRiesgoExtSeleccionado = Nz(lst.Column(0), "")
    If m_ObjProyectoAlInicio Is Nothing Then
        m_Error = "No se conoce la gestión de riesgos"
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
   
    
    If m_IDRiesgoExtSeleccionado <> "" Then
        If lst.ListCount > 2 Then
            For i = 1 To lst.ListCount - 1
                If Nz(lst.Column(0, i), "") = m_IDRiesgoExtSeleccionado Then
                    lst.Selected(i) = True
                    ListaRiesgosOferta_Click
                    Exit For
                End If
            Next
        End If
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
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoAltaRiesgoOferta_Click()
    
    Dim lst As ListBox
    Dim m_NombreUsuarioCalidad As String
    Dim m_ObjUsuario As Usuario
    Dim m_UsuarioRegistra As String
    Dim m_CodigoRiesgo As String
    Dim m_ObjCorreo As correo
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    Set lst = Me.ListaRiesgosOferta
    
    m_UsuarioRegistra = m_ObjUsuarioConectado.UsuarioRed
    
    Set m_ObjRiesgoExtActivo = New RiesgoExterno
    With m_ObjRiesgoExtActivo
        .IDEdicion = m_ObjProyectoActivo.EdicionUltima.IDEdicion
        .Descripcion = Nz(Me.Descripcion, "")
        .FechaDetectado = Nz(Me.FechaDetectado, "")
        .Trasladar = Nz(Me.ComboTrasladar, "")
        .MotivoNoIntegrado = Nz(Me.MotivoNoIntegrado, "")
        If Nz(Me.MotivoNoIntegrado, "") <> "" Then
            .FechaMotivo = CStr(m_FechaRef)
        End If
        .Origen = "Oferta"
        .UsuarioRegistra = m_UsuarioRegistra
        .CodRiesgoBiblioteca = Nz(Me.CodRiesgoBiblioteca, "")
        .RequiereRiesgoDeBibliotecaCalculado = m_ObjProyectoActivo.RequiereRiesgoDeBibliotecaCalculado
        If m_ObjProyectoActivo.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            .CausaRaiz = Nz(Me.CausaRaiz, "")
        End If
        
        .Registrar , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If Not .riesgo Is Nothing Then
            m_CodigoRiesgo = .riesgo.codigoRiesgo
        End If
        
        
        lst.AddItem .IDRiesgoExt & ";" & .CodRiesgo & ";" & .Descripcion & ";" & .Trasladar & ";" & m_CodigoRiesgo
    End With
    Set m_ObjRiesgoExtActivo = Nothing
    Me.Descripcion = Null
    Me.FechaDetectado = Null
    Me.MotivoNoIntegrado = Null
    Me.ComboTrasladar.SetFocus
    ListaRiesgosOferta_Click
    Me.CodRiesgoBiblioteca = Null
    Me.Descripcion = Null
    Me.CausaRaiz = Null
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaRiesgoOferta_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoDetalle_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjRiesgoExtSeleccionado Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    Set m_ObjRiesgoExtActivo = m_ObjRiesgoExtSeleccionado
    If m_ObjRiesgoExtActivo.Origen <> "Pedido" Then
        If FormularioAbierto("FormRiesgoExternoDetalle") Then
            DoCmd.Close acForm, "FormRiesgoExternoDetalle", acSaveNo
        End If
        DoCmd.OpenForm "FormRiesgoExternoDetalle"
        If FormularioAbierto("FormRiesgoExternoDetalle") Then
            Set m_FormRiesgosExt = Forms("FormRiesgoExternoDetalle")
        End If
        
    Else
        If FormularioAbierto("FormRiesgoExternoPedido") Then
            DoCmd.Close acForm, "FormRiesgoExternoPedido", acSaveNo
        End If
        
        DoCmd.OpenForm "FormRiesgoExternoPedido"
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoDetalle_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminarRiesgoOferta_Click()
    
    Dim m_ObjRiesgo As riesgo
    Dim lst As ListBox
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjRiesgoExtSeleccionado Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    
    Set lst = Me.ListaRiesgosOferta
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MostrarMsgBoxSiNoEsTest("¿Desea realmente borrar el Riesgo externo seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de riesgo externo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_ObjRiesgo = m_ObjRiesgoExtSeleccionado.riesgo
    m_Error = m_ObjRiesgoExtSeleccionado.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjRiesgo Is Nothing Then
        m_ObjRiesgo.ValidacionEliminar m_MotivoNoOK
        
        If m_MotivoNoOK <> "" Then
            m_Error = "No se podría borrar el riesgo asociado por: " & vbNewLine & m_MotivoNoOK
            Err.Raise 1000
        End If
    End If
    m_ObjRiesgoExtSeleccionado.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjRiesgo Is Nothing Then
        InvalidarCacheRiesgo m_ObjRiesgo.idRiesgo
    End If
    Set m_ObjRiesgoExtSeleccionado = Nothing
    Set m_ObjRiesgoExtActivo = Nothing
    lst.RemoveItem (lst.Column(0))
     ListaRiesgosOferta_Click
    If Not m_ObjRiesgo Is Nothing Then
        Set m_ObjRiesgo = GetCachedRiesgo(p_IDRiesgo:=m_ObjRiesgo.idRiesgo, p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If Not m_ObjRiesgo Is Nothing Then
        
        If FormularioAbierto("FormRiesgo") Then
            DoCmd.Close acForm, "FormRiesgo", acSaveNo
        End If
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarRiesgoOferta_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub ComandoSeleccionarRiesgoDeBiblioteca_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        DoCmd.Close acForm, "FormRiesgosBibliotecaGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormRiesgosBibliotecaGestion", , , , , , "FormGestionRiesgos"
    If FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        Set m_FormRiesgosBibliotecaGestion = Forms("FormRiesgosBibliotecaGestion")
        With m_FormRiesgosBibliotecaGestion
            .ComandoElegir.Visible = True
            .ComandoAlta.Visible = False
            .ComandoEliminar.Visible = False
            .ComandoVerSiUsado.Visible = False
            .Activo = "Sí"
            .Activo.Enabled = False
            .ComandoLimpiarActivo.Enabled = False
            .ComandoNoExisteRiesgo.Visible = False
        End With
        ' HR3e (2026-07-01): reemplazar acceso directo a
        ' Form_FormRiesgosBibliotecaGestion.Filtrar por helper
        ' modFormCoordinationHelper.
        Call modFormCoordinationHelper.Coord_FiltrarRiesgosBibliotecaGestion(m_Error)
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
        m_Error = "Al ComandoSeleccionarRiesgoDeBiblioteca_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVerRiesgoTrasladado_Click()
    
    Dim lst As ListBox
    Dim m_Form As Form
    Dim i As Integer
    Dim m_IDRiesgoEnLista As String
    Dim m_ObjRiesgoDelExt As riesgo
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjRiesgoExtSeleccionado Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    If m_ObjRiesgoExtSeleccionado.riesgo Is Nothing Then
        m_Error = "El Riesgo Externo seleccionado no tiene un Riesgo Asociado"
        Err.Raise 1000
    End If
    Set m_ObjRiesgoDelExt = m_ObjRiesgoExtSeleccionado.riesgo
    If Not FormularioAbierto("FormRiesgosGestion") Then
        
        DoCmd.OpenForm "FormRiesgosGestion"
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRiesgoTrasladado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComboTrasladar_Change()
    On Error Resume Next
    If Nz(Me.ComboTrasladar, "") = "No" Then
        Me.MotivoNoIntegrado.Enabled = True
    Else
        Me.MotivoNoIntegrado.Enabled = False
        Me.MotivoNoIntegrado = Null
    End If
End Sub

Private Sub ComboTrasladar_Exit(Cancel As Integer)
    On Error Resume Next
    If Nz(Me.ComboTrasladar, "") = "No" Then
        Me.MotivoNoIntegrado.Enabled = True
    Else
        Me.MotivoNoIntegrado.Enabled = False
        Me.MotivoNoIntegrado = Null
    End If
End Sub

Private Sub Descripcion_Exit(Cancel As Integer)
    On Error Resume Next
    If Nz(Me.ComboTrasladar, "") = "No" Then
        Me.MotivoNoIntegrado.Enabled = True
    Else
        Me.MotivoNoIntegrado.Enabled = False
        Me.MotivoNoIntegrado = Null
    End If
End Sub

Private Sub FechaDetectado_Exit(Cancel As Integer)
    On Error Resume Next
    If Nz(Me.ComboTrasladar, "") = "No" Then
        Me.MotivoNoIntegrado.Enabled = True
    Else
        Me.MotivoNoIntegrado.Enabled = False
        Me.MotivoNoIntegrado = Null
    End If
End Sub
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""

    Me.ComandoDetalle.Enabled = False
    Me.ComandoEliminarRiesgoOferta.Enabled = False
    Me.ComandoVerRiesgoTrasladado.Enabled = False
    Me.ComandoAltaRiesgoOferta.Enabled = False
    Me.ComandoActualizar.Enabled = True
    Me.ComandoSeleccionarRiesgoDeBiblioteca.Enabled = False
    Me.Descripcion.Enabled = True
    Me.CausaRaiz.Enabled = True
    Me.AllowEdits = True

    Call EstablecerDatos_CargarPermisos(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_AplicarEstado(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Call EstablecerDatos_RenderizarControles(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Private Sub EstablecerDatos_CargarPermisos(ByRef p_Error As String)
    On Error GoTo errores
    If m_ObjProyectoActivo Is Nothing Or m_ObjProyectoAlInicio Is Nothing Then
        p_Error = "No se sabe qué proyecto es"
        Err.Raise 1000
    End If
    If m_ObjProyectoActivo.EsActivo = EnumSiNo.Sí Then
        If EsTecnico = EnumSiNo.Sí Then
            blnPermitidoEscribir = False
        Else
            If m_ObjProyectoActivo.EdicionActiva Is Nothing Then
                blnPermitidoEscribir = False
            Else
                If m_ObjProyectoActivo.EdicionActiva.EsPrimeraEdicion = EnumSiNo.Sí Then
                    blnPermitidoEscribir = True
                Else
                    blnPermitidoEscribir = False
                End If
            End If
        End If
    Else
        blnPermitidoEscribir = False
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_CargarPermisos: " & Err.description
End Sub

Private Sub EstablecerDatos_AplicarEstado(ByRef p_Error As String)
    On Error GoTo errores
    If blnPermitidoEscribir = True Then
        Me.ComandoAltaRiesgoOferta.Enabled = True
        If m_ObjProyectoActivo.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            Me.ComandoSeleccionarRiesgoDeBiblioteca.Enabled = True
            Me.CausaRaiz.Enabled = True
            Me.Descripcion.Enabled = False
        ElseIf m_ObjProyectoActivo.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            Me.ComandoSeleccionarRiesgoDeBiblioteca.Enabled = False
            Me.CausaRaiz.Enabled = False
            Me.Descripcion.Enabled = True
        Else
            p_Error = "No se sabe si la gestión de riesgos requiere riesgos del catálogo"
            Err.Raise 1000
        End If
        Me.FechaDetectado.Enabled = True
        Me.ComboTrasladar.Enabled = True
        Me.MotivoNoIntegrado.Enabled = True
        Me.ComandoAltaRiesgoOferta.Enabled = True
    Else
        Me.ComandoSeleccionarRiesgoDeBiblioteca.Enabled = False
        Me.Descripcion.Enabled = False
        Me.CausaRaiz.Enabled = False
        Me.FechaDetectado.Enabled = False
        Me.ComboTrasladar.Enabled = False
        Me.MotivoNoIntegrado.Enabled = False
        Me.ComandoAltaRiesgoOferta.Enabled = False
    End If
    Exit Sub
errores:
    p_Error = "EstablecerDatos_AplicarEstado: " & Err.description
End Sub

Private Sub EstablecerDatos_RenderizarControles(ByRef p_Error As String)
    On Error GoTo errores
    Dim lst As ListBox
    Dim m_ObjRiesgoExterno As RiesgoExterno
    Dim m_IDRiesgoExt As Variant
    Dim m_CodRiesgo As String
    Dim m_Col As Scripting.Dictionary

    Set m_ObjEdicionActiva = m_ObjProyectoActivo.EdicionPrimera
    p_Error = m_ObjProyectoActivo.Error
    If p_Error <> "" Then Exit Sub
    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No se ha podido determinar la edición activa"
        Exit Sub
    End If
    Set lst = Me.ListaRiesgosOferta
    lst.RowSource = "IDRiesgoExt;Código;descripción;Se Trasl.;Cód Riesgo"
    Set m_ObjEdicionActiva.ColRiesgosExternos = Nothing
    Set m_Col = m_ObjEdicionActiva.ColRiesgosExternos
    If m_Col Is Nothing Then
        Exit Sub
    End If
    For Each m_IDRiesgoExt In m_Col
        Set m_ObjRiesgoExterno = m_Col(m_IDRiesgoExt)
        With m_ObjRiesgoExterno
            If Not .riesgo Is Nothing Then
                m_CodRiesgo = .riesgo.codigoRiesgo
            Else
                m_CodRiesgo = ""
            End If
            lst.AddItem .IDRiesgoExt & ";" & .CodRiesgo & ";" & .Descripcion & ";" & .Trasladar & ";" & m_CodRiesgo
        End With
        Set m_ObjRiesgoExterno = Nothing
SiguienteRiesgoExt:
    Next
    lst.Selected(1) = True
    ListaRiesgosOferta_Click
    Exit Sub
errores:
    p_Error = "EstablecerDatos_RenderizarControles: " & Err.description
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub ListaRiesgosOferta_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.ComandoDetalle.Enabled = False
    Me.ComandoEliminarRiesgoOferta.Enabled = False
    Me.ComandoVerRiesgoTrasladado.Visible = False
   
     
    If Nz(Me.ListaRiesgosOferta.Column(0), "") = "" Then
        Set m_ObjRiesgoExtSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjRiesgoExtSeleccionado Is Nothing Then
        Set m_ObjRiesgoExtSeleccionado = getRiesgoExterno(Me.ListaRiesgosOferta.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaRiesgosOferta.Column(0) <> m_ObjRiesgoExtSeleccionado.IDRiesgoExt Then
            Set m_ObjRiesgoExtSeleccionado = getRiesgoExterno(Me.ListaRiesgosOferta.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjRiesgoExtSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Set m_ObjEdicionActiva = m_ObjRiesgoExtSeleccionado.Edicion
    m_Error = m_ObjRiesgoExtSeleccionado.Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    Me.ComandoDetalle.Enabled = True
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        If blnPermitidoEscribir = True Then
            Me.ComandoEliminarRiesgoOferta.Enabled = True
        Else
            Me.ComandoEliminarRiesgoOferta.Enabled = False
        End If
    Else
        Me.ComandoEliminarRiesgoOferta.Enabled = False
    End If
    If m_ObjRiesgoExtSeleccionado.riesgo Is Nothing Then
        Me.ComandoVerRiesgoTrasladado.Visible = False
    Else
        Me.ComandoVerRiesgoTrasladado.Visible = True
        Me.ComandoVerRiesgoTrasladado.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaRiesgosOferta_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ListaRiesgosOferta_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo True
    If Me.ListaRiesgosOferta.ItemsSelected.Count > 0 Then
        If Me.ComandoDetalle.Enabled Then
            ComandoDetalle_Click
        End If
    End If
    Application.Echo False
End Sub

Public Function ActualizarRiesgoExt( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim lst As ListBox
    Dim m_LineaActual As String
    Dim m_LineaFinal As String
    Dim i As Integer
    Dim m_CodRiesgo As String
    Dim m_ObjRiesgo As riesgo
    On Error GoTo errores
    
    Set m_ObjRiesgoExtSeleccionado = m_ObjRiesgoExtActivo
    If m_ObjRiesgoExtSeleccionado Is Nothing Then
        Exit Function
    End If
    Set lst = Me.ListaRiesgosOferta
    For i = 1 To lst.ListCount - 1
        If Nz(lst.Column(0, i), "") = m_ObjRiesgoExtSeleccionado.IDRiesgoExt Then
            m_LineaActual = Nz(lst.Column(0, i), "") & ";" & Nz(lst.Column(1, i), "") & ";" & Nz(lst.Column(2, i), "") & ";" & _
                            Nz(lst.Column(3, i), "") & ";" & Nz(lst.Column(4, i), "")
            lst.Selected(i) = True
            Exit For
        End If
    Next
    
    
    With m_ObjRiesgoExtSeleccionado
        Set m_ObjRiesgo = .riesgo
        If Not m_ObjRiesgo Is Nothing Then
            m_CodRiesgo = m_ObjRiesgo.codigoRiesgo
        End If
        If m_CodRiesgo <> "" Then
            m_LineaFinal = .IDRiesgoExt & ";" & .CodRiesgo & ";" & .Descripcion & ";" & _
                            .Trasladar & ";" & m_CodRiesgo
        Else
            m_LineaFinal = .IDRiesgoExt & ";" & .CodRiesgo & ";" & .Descripcion & ";" & _
                            .Trasladar & ";"
        End If
    End With
    If m_LineaActual <> m_LineaFinal Then
        lst.RowSource = Replace(lst.RowSource, m_LineaActual, m_LineaFinal)
        lst.Requery
        lst.Selected(i) = True
        ListaRiesgosOferta_Click
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarRiesgoExt ha producido el error: " & vbNewLine & Err.description
    End If
End Function
Private Sub m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido(m_ObjRiesgoBiblioteca As RiesgoBiblioteca)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    With m_ObjRiesgoBiblioteca
        Me.CodRiesgoBiblioteca = .codigo
        If .RiesgoParaRetipificar = EnumSiNo.No Then
            Me.Descripcion = .Descripcion
            Me.Descripcion.Locked = False
            Me.Descripcion.Enabled = True
        Else
            Me.Descripcion = Null
            Me.Descripcion.Locked = False
            Me.Descripcion.Enabled = False
        End If
        
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub m_FormRiesgosExt_Registrado()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    'Set m_ObjRiesgoExtSeleccionado = m_ObjRiesgoActivo
    ComandoActualizar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormRiesgosExt_Registrado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
