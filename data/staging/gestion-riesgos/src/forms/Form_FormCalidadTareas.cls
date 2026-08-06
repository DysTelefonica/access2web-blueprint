Option Compare Database
Option Explicit
'````````````````````````````````````````````````````````````````````````````
'`                    Code from Tassos Filoxenidis                          `
'````````````````````````````````````````````````````````````````````````````

Private m_sngLeftResizePos As Integer
Private MeWidth As Integer
Private MeWidthHalf As Integer
Private ArbolLeft As Integer
Private SubFormLeft As Integer
Private Sepleft As Integer
Const SepWMin As Integer = 60
Const SepW As Integer = 60.1
Private m_Error As String
Private m_ObjArbolCalidad As ArbolTareasCalidad
Private m_Arbol As MSComctlLib.TreeView
Private m_Imagelist As ImageList
Public m_TipoNodoSeleccionado As EnumTipoNodoTareasCalidad

Private m_ObjSeleccionado As Object
Public m_NodoSeleccionado As MSComctlLib.Node
Public m_RiesgoSeleccionado As riesgo
Public m_EstadoRiesgoSeleccionado As EnumRiesgoEstado
Public m_RiesgoMaterializadoSeleccionado As RiesgoMaterializacion
Public m_ProyectoSeleccionado As Proyecto
Public m_EdicionSeleccionada As Edicion
Private m_ObjUsuario As Usuario


Private Function SeleccionarEdicion(Optional ByRef p_Error As String) As String

    On Error GoTo errores

    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If

    Set m_EdicionSeleccionada = m_ObjSeleccionado
    Me.FormDetalle.SourceObject = "FormCalidadTareasDetalleEdicion"

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarEdicion ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

' HR4-sister (HR3d follow-up #9, #4 - 2026-07-01):
'   Era Public Sub. Ahora Private. Es event handler de Access (TreeView NodeClick).
'   Hard rule 4 exception: event wiring. Sin external callers
'   (1 internal caller: SeleccionarNodoRiesgo via .Arbol_NodeClick m_Nodo).
'   Body intacto (UI orchestration + dispatch to SeleccionarRiesgoXxx).
Private Sub Arbol_NodeClick(ByVal Node As Object)

    Dim m_TipoNodoSeleccionado As EnumTipoNodoTareasCalidad
    Dim m_Riesgo As riesgo
    Dim m_RiesgoMaterializado As RiesgoMaterializacion
    Dim m_IdRiesgo As String
    Dim m_IDRiesgoMaterializado As String
    Dim dato
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    If m_ObjArbolCalidad Is Nothing Then
        Set m_ObjArbolCalidad = New ArbolTareasCalidad
        Set m_Arbol = Me.Arbol.Object
        Set m_Imagelist = Me.ListaImagenes.Object
    End If
    Set m_NodoSeleccionado = Node
    

    m_TipoNodoSeleccionado = m_ObjArbolCalidad.TipoNodoTareaSeleccionado(Node.key, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoNodoSeleccionado = 0 Then
        m_Error = "No se sabe el tipo de nodo que se ha seleccionado"
        Err.Raise 1000
    End If
    Set m_ObjSeleccionado = m_ObjArbolCalidad.ObjetoSeleccionado
    m_Error = m_ObjArbolCalidad.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.EDICIONESCADUCADAS Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.EDICIONESAPUNTODECADUCAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.EDICIONESPREPARADASPARAPUBLICAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSPENDIENTESRETIPIFICAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSACEPTADOSPORVISAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSRETIRADOSPORVISAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSMATERIALIZADOSPORDECIDIR Then
        AbrirExplicacion m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.EDICIONAPUNTODECADUCAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.EDICIONPREPARADAPARAPUBLICAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.EDICIONCADUCADA Then
        SeleccionarEdicion m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If

    If m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGOPENDIENTERETIPIFICAR Then
        SeleccionarRiesgoParaRetipificar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGOACEPTADOPORVISAR Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGORETIRADOPORVISAR Then
        SeleccionarRiesgoAceptadoRetirado m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasCalidad.RIESGOMATERIALIZADOPORDECIDIR Then
        If InStr(1, m_NodoSeleccionado.key, "|") Then
            dato = Split(m_NodoSeleccionado.key, "|")
            m_IDRiesgoMaterializado = dato(1)
            If IsNumeric(m_IDRiesgoMaterializado) Then
                Set m_RiesgoMaterializado = Constructor.getRiesgoMaterializado(m_IDRiesgoMaterializado, m_Error)
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not m_RiesgoMaterializado Is Nothing Then
                    Set m_ObjRiesgoMaterializadoActivo = m_RiesgoMaterializado
                    
                End If
                If Not m_ObjRiesgoMaterializadoActivo Is Nothing Then
                    Set m_ObjSeleccionado = m_ObjRiesgoMaterializadoActivo
                End If
                SeleccionarRiesgoMaterializado m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Arbol_NodeClick se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
'
Private Sub ComandoActualizarContador_Click()

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    SyncResponsableCalidadDesdeCombo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CargarArbol m_Error
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
        m_Error = "Al ComandoActualizarContador_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
'
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
'
'
'
Private Sub ComboResponsableCalidad_AfterUpdate()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    ComandoActualizarContador_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComboResponsableCalidad_AfterUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
'
Private Sub Form_Load()

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    If EsTecnico = EnumSiNo.Sí Then
        m_Error = "Es un formulario exclusivo para los miembros de Calidad. Perdone las molestias"
        Err.Raise 1000
    End If
    Me.ComboResponsableCalidad.Enabled = True
    EstablecerComboUsuariosCalidad Me.ComboResponsableCalidad, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If EsCalidad = EnumSiNo.Sí Then
        Me.ComboResponsableCalidad = m_ObjUsuarioConectado.Nombre
    End If
    SyncResponsableCalidadDesdeCombo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    Me.Caption = "TAREAS PENDIENTES CALIDAD " & m_ObjEntorno.TituloUsuarioConectado
    CargarArbol m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Sep.Left = Me.FormDetalle.Left - Me.Sep.Width
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    'Me.SetFocus

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

Private Function SyncResponsableCalidadDesdeCombo(Optional ByRef p_Error As String) As String

    Dim m_NombreResponsable As String

    On Error GoTo errores

    p_Error = ""
    m_NombreResponsable = Nz(Me.ComboResponsableCalidad, "")

    If m_NombreResponsable = "" Then
        Set m_ObjUsuarioParaTareas = Nothing
        Exit Function
    End If

    If Not m_ObjUsuarioParaTareas Is Nothing Then
        If m_ObjUsuarioParaTareas.Nombre = m_NombreResponsable Then
            Exit Function
        End If
    End If

    Set m_ObjUsuarioParaTareas = Constructor.getUsuario(, , m_NombreResponsable, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SyncResponsableCalidadDesdeCombo ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

' HR4-sister (HR3d follow-up #9, #4 - 2026-07-01):
'   MANTIENE Public porque tiene 2 callers EXTERNOS al form:
'     - Form_FormTecnicoTareas.cls:305 (2 callers)
'     - Form_FormRiesgosGestion.cls:1696 (4 callers)
'   Hard rule 4: Public business methods son anti-patron, PERO cuando hay
'   cross-form access requerido, el form debe mantener un Public Function
'   que actua como thin adapter delegando al helper.
'   Body: thin adapter que prepara m_ObjArbolCalidad y delega al helper.
Public Function CargarArbol(Optional ByRef p_Error As String) As String

    On Error GoTo errores

    p_Error = ""
    m_Error = ""

    Me.FormDetalle.SourceObject = ""
    Set m_ObjArbolCalidad = New ArbolTareasCalidad
    Set m_Arbol = Me.Arbol.Object
    Set m_Imagelist = Me.ListaImagenes.Object

    ' HR4-sister: el helper valida que m_ObjArbolCalidad este disponible.
    '   El form-side hace el setup + la llamada real al metodo del helper
    '   que tiene el TreeView (interfaz boundary).
    modCalidadTareasHelper.CalidadTareas_CargarArbol m_ObjArbolCalidad, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    ' La llamada real al cargar sigue aqui (UI-specific: requiere TreeView control)
    m_ObjArbolCalidad.CargarArbolTareasCalidad m_Arbol, m_Imagelist, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CargarArbol ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



'
Private Sub Sep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        MouseCursor IDC_SIZEWE
        m_sngLeftResizePos = X
        MeWidthHalf = Me.InsideWidth / 2
        MeWidth = Me.InsideWidth - 100
        ArbolLeft = Me.Arbol.Left
        SubFormLeft = Me.FormDetalle.Left
    End If

End Sub
'
Private Sub Sep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    MouseCursor IDC_SIZEWE
    If Button = 1 Then
        On Error Resume Next
        With Me.Sep
            If .Left > MeWidthHalf Then Exit Sub    '.Left = MeWidthHalf - 1:
            If .Left < SepWMin Then
                .Left = SepW
                Me.Arbol.Width = .Left - ArbolLeft
                Me.FormDetalle.Left = .Left + .Width
                Me.FormDetalle.Width = MeWidth - .Left

            Else
                .Move .Left + X - m_sngLeftResizePos
                Me.Arbol.Width = .Left - ArbolLeft
                Me.FormDetalle.Left = .Left + .Width
                Me.FormDetalle.Width = MeWidth - .Left

            End If
        End With
    End If
End Sub

Private Sub Sep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        With Me.Sep
            If .Left > MeWidthHalf Then .Left = MeWidthHalf - 0.1
            Me.Arbol.Width = IIf(Me.Arbol.Width < 60, 60, .Left - ArbolLeft)
            .Left = Me.Arbol.Width + Me.Arbol.Left
            Me.FormDetalle.Left = .Left + .Width
            Me.FormDetalle.Width = MeWidth - .Left

        End With
    End If
End Sub
'
' HR4-sister (HR3d follow-up #9, #4 - 2026-07-01):
'   Era Public Function. Ahora Private. Sin external callers
'   (codegraph blast radius: no muestra callers externos al form).
'   Hard rule 4: Public business methods son anti-patron.
'   Body: thin adapter que valida input via helper + hace el lookup en
'   el TreeView + Arbol_NodeClick (UI-specific).
Private Function SeleccionarNodoRiesgo( _
                                        p_TipoRiesgoTarea As EnumTipoRiesgoTarea, _
                                        p_ObjRiesgo As riesgo, _
                                        Optional p_Error As String) As String

    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    Dim m_Estado As EnumRiesgoEstado
    On Error GoTo errores

    ' HR4-sister: delega al helper para la validacion de inputs +
    '   construccion del key. El helper retorna el m_Key via p_OutKey.
    Dim m_OutKey As String
    modCalidadTareasHelper.CalidadTareas_SeleccionarNodoRiesgo _
        p_ObjRiesgo, p_TipoRiesgoTarea, m_OutKey, p_Error
    If p_Error <> "" Then
        ' El helper ya seteo p_Error (Nothing, tipo invalido, no mappable).
        Exit Function
    End If
    m_Key = m_OutKey

    ' Lookup en el TreeView (UI-specific).
    On Error Resume Next
    Set m_Nodo = m_Arbol.Nodes(m_Key)
    If Err.Number <> 0 Then
        p_Error = "No se encuentra el riesgo"
        Err.Raise 1000
    End If
    On Error GoTo errores
    m_Arbol.SelectedItem = m_Nodo
    Arbol_NodeClick m_Nodo

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo SeleccionarNodoRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
'
Private Function AbrirExplicacion(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    
    Me.FormDetalle.SourceObject = "FormCalidadTareaExplicacion"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirExplicacion ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

Private Function SeleccionarRiesgoParaRetipificar(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    
    Set m_RiesgoSeleccionado = m_ObjSeleccionado
    m_EstadoRiesgoSeleccionado = m_RiesgoSeleccionado.EstadoEnum
    Me.FormDetalle.SourceObject = "FormCalidadTareaRiesgosRetipificacion"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarRiesgoParaRetipificar ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
Private Function SeleccionarRiesgoAceptadoRetirado(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    
    Set m_RiesgoSeleccionado = m_ObjSeleccionado
    m_EstadoRiesgoSeleccionado = m_RiesgoSeleccionado.EstadoEnum
    Me.FormDetalle.SourceObject = "FormCalidadTareaRiesgosAceptadosRetirados"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarRiesgoAceptadoRetirado ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
Private Function SeleccionarRiesgoMaterializado(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    
    Set m_RiesgoMaterializadoSeleccionado = m_ObjSeleccionado
    
    Me.FormDetalle.SourceObject = "FormCalidadTareaRiesgosMaterializadosPorDecidir"
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarRiesgoMaterializado ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
