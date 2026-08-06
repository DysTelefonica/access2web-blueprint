Option Compare Database
Option Explicit
Private m_Error As String
Private m_ObjRiesgoBibliotecaSeleccionado As RiesgoBiblioteca
Private blnPermitidoEscribir As Boolean
Public Event AltaRiesgoBiblioteca(m_objRiesgoDeBiblioteca As RiesgoBiblioteca)
Public Event BajaRiesgoBiblioteca()
Public Event RiesgoDeBibliotecaElegido(m_ObjRiesgoBiblioteca As RiesgoBiblioteca)
Private WithEvents m_FormRiesgoBiblioteca As Form_FormRiesgoBiblioteca

Private m_FamiliaUltima As String
Private m_PalabraClaveUltima As String
Private m_ActivoUltimo As String
Private m_AbiertoDesdeFormProyectos As Boolean


Private Sub Activo_BeforeUpdate(Cancel As Integer)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Nz(Me.Activo, "") = m_ActivoUltimo Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ActivoUltimo = Nz(Me.Activo, "")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Activo_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoActualizarRiesgosBiblioteca_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Filtrar m_Error
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
        m_Error = "Al ComandoActualizarRiesgosBiblioteca_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
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

Private Sub ComandoElegir_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Nz(Me.ListaRiesgos.Column(0), ""), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    
    If m_ObjRiesgoBibliotecaSeleccionado.Activo <> "Sí" Then
        m_Error = "Para poder elegir un riesgo de la biblioteca, éste ha de estar activo por parte de Calidad"
        Err.Raise 1000
    End If
    RaiseEvent RiesgoDeBibliotecaElegido(m_ObjRiesgoBibliotecaSeleccionado)
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoElegir_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoLimpiarActivo_Click()
     On Error Resume Next
    Me.Activo = Null
    m_ActivoUltimo = ""
    Filtrar
End Sub

Private Sub ComandoLimpiarFamilia_Click()
     On Error Resume Next
    Me.ComboFamilia = Null
    m_FamiliaUltima = ""
    Filtrar
End Sub

Private Sub ComandoLimpiarPalabraClave_Click()
    On Error Resume Next
    Me.PalabraClave = Null
    m_PalabraClaveUltima = ""
    Filtrar
End Sub

Public Sub ComandoAlta_Click()
    
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjRiesgoBibliotecaActivo = Nothing
    If FormularioAbierto("FormRiesgoBiblioteca") Then
        DoCmd.Close acForm, "FormRiesgoBiblioteca", acSaveNo
    End If
    DoCmd.OpenForm "FormRiesgoBiblioteca"
    If FormularioAbierto("FormRiesgoBiblioteca") Then
        Set m_FormRiesgoBiblioteca = Forms("FormRiesgoBiblioteca")
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Private Sub ComandoEditar_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Nz(Me.ListaRiesgos.Column(0), ""), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    
    If FormularioAbierto("FormRiesgoBiblioteca") Then
        DoCmd.Close acForm, "FormRiesgoBiblioteca", acSaveNo
    End If
    Set m_ObjRiesgoBibliotecaActivo = m_ObjRiesgoBibliotecaSeleccionado
    DoCmd.OpenForm "FormRiesgoBiblioteca"
    If FormularioAbierto("FormRiesgoBiblioteca") Then
        Set m_FormRiesgoBiblioteca = Forms("FormRiesgoBiblioteca")
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminar_Click()
   
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Nz(Me.ListaRiesgos.Column(0), ""), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    Else
        If m_ObjRiesgoBibliotecaSeleccionado.IDRiesgoTipo <> Nz(Me.ListaRiesgos.Column(0), "") Then
            Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Nz(Me.ListaRiesgos.Column(0), ""), , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        m_Error = "Se ha de seleccionar un elemento de la lista"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente borrar el riesgo del catálogo?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrar riesgo de Biblioteca")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjRiesgoBibliotecaSeleccionado.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjRiesgoBibliotecaSeleccionado = Nothing
    Me.ListaRiesgos.RemoveItem (Me.ListaRiesgos.Column(0))
    Me.ComandoEditar.Enabled = False
    Me.ComandoSalir.SetFocus
    Me.ComandoEliminar.Enabled = False
    RaiseEvent BajaRiesgoBiblioteca
    If Me.ListaRiesgos.ListCount > 1 Then
        Me.ListaRiesgos.Selected(2) = True
        ListaRiesgos_Click
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoNoExisteRiesgo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    RaiseEvent RiesgoDeBibliotecaElegido(m_ObjEntorno.RiesgoBibliotecaNoExiste)
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNoExisteRiesgo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoVerSiUsado_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Nz(Me.ListaRiesgos.Column(0), ""), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If m_ObjRiesgoBibliotecaSeleccionado.RiesgoParaRetipificar = EnumSiNo.Sí Then
        pregunta = MsgBox("Es un tipo de riesgo especial. Es necesario que permanezca en la biblioteca", vbExclamation, "Riesgo reservado")
        Exit Sub
    End If
    If m_ObjRiesgoBibliotecaSeleccionado.Usado = EnumSiNo.Sí Then
        pregunta = MsgBox("Ha sido usado en algún riesgo", vbExclamation, "Usado")
    Else
        pregunta = MsgBox("No ha sido usado en ningún riesgo", vbInformation, "No Usado")
    End If
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerSiUsado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
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
    Filtrar m_Error
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
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub Form_Open(Cancel As Integer)
    
    
    Dim m_Titulo As String
   
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    If EsTecnico = EnumSiNo.Sí Then
        blnPermitidoEscribir = False
    Else
        blnPermitidoEscribir = True
    End If
    If blnPermitidoEscribir = True Then
        Me.ComandoAlta.Enabled = True
        Me.ComandoEliminar.Visible = True
        Me.ComandoVerSiUsado.Visible = True
    Else
        Me.ComandoAlta.Enabled = False
        Me.ComandoEliminar.Visible = False
        Me.ComandoVerSiUsado.Visible = False
    End If
    Me.ComandoEditar.Enabled = False
    
    m_Titulo = "LISTA DE RIESGOS DE LA BIBLIOTECA"
    Me.lblTitulo1.Caption = m_Titulo
   
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    EstablecerComboFamilia Me.ComboFamilia, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ComandoActualizarRiesgosBiblioteca.SetFocus
    If Nz(Me.openArgs, "") <> "" Then
        If Nz(Me.openArgs(), "") = "FormGestionRiesgos" Then
            m_AbiertoDesdeFormProyectos = True
        Else
            m_AbiertoDesdeFormProyectos = False
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    Filtrar m_Error
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
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Public Function Filtrar( _
                        Optional ByRef p_Error As String _
                        ) As String

    Dim lst As ListBox
    Dim m_ID As Variant
    Dim m_ObjRiesgoBiblioteca As RiesgoBiblioteca
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    
    p_Error = ""
    Set lst = Me.ListaRiesgos
    lst.RowSource = "IDRiesgoTipo;Familia;Descripción;Activo"
    Set m_Col = Constructor.getRiesgosBiblioteca(Nz(Me.ComboFamilia, ""), Nz(Me.PalabraClave, ""), Nz(Me.Activo, ""), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjRiesgoBiblioteca = m_Col(m_ID)
        With m_ObjRiesgoBiblioteca
            lst.AddItem .IDRiesgoTipo & ";" & .Familia & ";" & _
                        Replace(.Descripcion, ";", ":") & ";" & .Activo
            Set m_ObjRiesgoBiblioteca = Nothing
        End With
    Next
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaRiesgos_Click
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub ListaRiesgos_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.DescripcionRiesgo = Null
    Me.ComandoVerSiUsado.Enabled = False
    If Me.ComandoElegir.Visible = True Then
        Me.ComandoElegir.Enabled = False
    End If
    If Me.ComandoEliminar.Visible = True Then
        Me.ComandoEliminar.Enabled = False
    End If
    
    Me.ComandoEditar.Enabled = False
    If Nz(Me.ListaRiesgos.Column(0), "") = "" Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjRiesgoBibliotecaSeleccionado Is Nothing Then
        Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Me.ListaRiesgos.Column(0), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjRiesgoBibliotecaSeleccionado.IDRiesgoTipo <> Me.ListaRiesgos.Column(0) Then
            Set m_ObjRiesgoBibliotecaSeleccionado = Constructor.getRiesgoBiblioteca(Me.ListaRiesgos.Column(0), , m_Error)
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
    Me.DescripcionRiesgo = m_ObjRiesgoBibliotecaSeleccionado.Descripcion
    If Me.ComandoVerSiUsado.Visible = True Then
        Me.ComandoVerSiUsado.Enabled = True
    End If
    If Me.ComandoElegir.Visible = True Then
        If m_ObjRiesgoBibliotecaSeleccionado.Activo = "Sí" Then
            If m_AbiertoDesdeFormProyectos = True Then
                If m_ObjRiesgoBibliotecaSeleccionado.RiesgoParaRetipificar = EnumSiNo.Sí Then
                    Me.ComandoElegir.Enabled = False
                Else
                    Me.ComandoElegir.Enabled = True
                End If
            Else
                Me.ComandoElegir.Enabled = True
            End If
            
        Else
            Me.ComandoElegir.Enabled = False
        End If
        
    End If
    
    Me.ComandoEditar.Enabled = True
    If Me.ComandoEliminar.Visible = True Then
        If EsTecnico <> EnumSiNo.Sí Then
            If m_ObjRiesgoBibliotecaSeleccionado.RiesgoParaRetipificar <> EnumSiNo.Sí Then
                Me.ComandoEliminar.Enabled = True
            Else
                Me.ComandoEliminar.Enabled = False
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
        m_Error = "Al ListaRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Public Function ActualizarRiesgoBiblioteca( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim lst As ListBox
    Dim m_LineaInicial As String
    Dim m_LineaFinal As String
    Dim i As Integer
    
    On Error GoTo errores
    
    Set lst = Me.ListaRiesgos
    If Nz(lst.Column(0), "") = "" Then
        Exit Function
    End If
    'IDRiesgoTipo;Tipo;Código;Descripción
    For i = 0 To 4
        If m_LineaInicial = "" Then
            m_LineaInicial = Nz(lst.Column(i), "")
        Else
            m_LineaInicial = m_LineaInicial & ";" & Nz(lst.Column(i), "")
        End If
    Next
    If m_ObjRiesgoBibliotecaActivo Is Nothing Then
        Set m_ObjRiesgoBibliotecaActivo = Constructor.getRiesgoBiblioteca(lst.Column(0), , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjRiesgoBibliotecaActivo Is Nothing Then
            lst.RemoveItem (lst.ListIndex)
            Set m_ObjRiesgoBibliotecaSeleccionado = Nothing
            Exit Function
        End If
    Else
        If lst.Column(0) <> m_ObjRiesgoBibliotecaActivo.IDRiesgoTipo Then
            Set m_ObjRiesgoBibliotecaActivo = Constructor.getRiesgoBiblioteca(lst.Column(0), , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjRiesgoBibliotecaActivo Is Nothing Then
                lst.RemoveItem (lst.ListIndex)
                Set m_ObjRiesgoBibliotecaSeleccionado = Nothing
                Exit Function
            End If
        End If
    End If
    Set m_ObjRiesgoBibliotecaSeleccionado = m_ObjRiesgoBibliotecaActivo
    With m_ObjRiesgoBibliotecaSeleccionado
        m_LineaFinal = .IDRiesgoTipo & ";" & .Familia & ";" & .codigo & ";" & _
                            Replace(.Descripcion, ";", ":") & ";" & .Activo
        
       
    End With
    If m_LineaInicial <> m_LineaFinal Then
        lst.RowSource = Replace(lst.RowSource, m_LineaInicial, m_LineaFinal)
        lst.Requery
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarRiesgoBiblioteca ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub ListaRiesgos_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If Me.ComandoElegir.Visible = True Then
        If Me.ComandoElegir.Enabled = True Then
            ComandoElegir_Click
        End If
    Else
        If Me.ComandoEditar.Enabled = True Then
            ComandoEditar_Click
        End If
    End If
    
    Application.Echo True
    
End Sub

Private Sub m_FormRiesgoBiblioteca_AltaRiesgoBiblioteca(m_objRiesgoDeBiblioteca As RiesgoBiblioteca)

    On Error GoTo errores
   
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerComboFamilia Me.ComboFamilia, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
   
    RaiseEvent AltaRiesgoBiblioteca(m_objRiesgoDeBiblioteca)
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormRiesgoBiblioteca_AltaRiesgoBiblioteca se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub m_FormRiesgoBiblioteca_EdicionRiesgoBiblioteca(m_objRiesgoDeBiblioteca As RiesgoBiblioteca)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    ActualizarRiesgoBiblioteca m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerComboFamilia Me.ComboFamilia, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Filtrar m_Error
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
        m_Error = "Al m_FormRiesgoBiblioteca_EdicionRiesgoBiblioteca se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub





Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyReturn Then
        If Me.ComboFamilia.Enabled = False Then
            Exit Sub
        End If
        Me.ComboFamilia.SetFocus
        Me.PalabraClave.SetFocus
        If Nz(Me.PalabraClave, "") = m_PalabraClaveUltima Then
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Exit Sub
        End If
        
       Filtrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PalabraClaveUltima = Nz(Me.PalabraClave, "")
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    End If
End Sub
