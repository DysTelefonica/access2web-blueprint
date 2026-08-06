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

Private m_ObjArbolTecnico As ArbolTareasTecnico
Private m_Arbol As MSComctlLib.TreeView
Private m_Imagelist As ImageList
Public m_TipoNodoSeleccionado As EnumTipoNodoTareasTecnico

Private m_ObjSeleccionado As Object
Public m_NodoSeleccionado As MSComctlLib.Node
Public m_RiesgoSeleccionado As riesgo
Public m_EdicionSeleccionada As Edicion
Private m_Error As String

Private Function EstablecerIconoCarpetas(Optional ByRef p_Error As String) As String
    Dim m_ObjNodo As MSComctlLib.Node
    On Error GoTo errores
    For Each m_ObjNodo In m_Arbol.Nodes
        If m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32) Or _
            m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32) Then
            If m_ObjNodo.Expanded = True Then
                m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
            Else
                m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            End If
        ElseIf m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32) Or _
                m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32) Then
            If m_ObjNodo.Expanded = True Then
                m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
            Else
                m_ObjNodo.Image = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            End If
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerIconoCarpetas ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
Private Function SeleccionarEdicion(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    
    Set m_EdicionSeleccionada = m_ObjSeleccionado
    Me.FormDetalle.SourceObject = "FormTecnicoTareasDetalleEdicion"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarEdicion ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function


Private Function SeleccionarRiesgoAceptadoRetirado(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    
    Set m_RiesgoSeleccionado = m_ObjSeleccionado
    Me.FormDetalle.SourceObject = "FormTecnicoTareaRiesgosAceptadosRetirados"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarRiesgoAceptadoRetirado ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

Private Function AbrirExplicacion(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    
    Me.FormDetalle.SourceObject = "FormTecnicoTareaExplicacion"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirExplicacion ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
Public Sub Arbol_NodeClick(ByVal Node As Object)
    
    
    
    Dim m_IdRiesgo As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_ObjArbolTecnico Is Nothing Then
        Set m_ObjArbolTecnico = New ArbolTareasTecnico
        Set m_Arbol = Me.Arbol.Object
        Set m_Imagelist = Me.ListaImagenes.Object
    End If
    Set m_NodoSeleccionado = Node
    EstablecerIconoCarpetas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_TipoNodoSeleccionado = m_ObjArbolTecnico.TipoNodoTareaSeleccionado(Node.key, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoNodoSeleccionado = 0 Then
        m_Error = "No se sabe el tipo de nodo que se ha seleccionado"
        Err.Raise 1000
    End If
    Set m_ObjSeleccionado = m_ObjArbolTecnico.ObjetoSeleccionado
    m_Error = m_ObjArbolTecnico.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.EDICIONESAPUNTODECADUCARSINPROPUESTA Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.EDICIONESPROPUESTASRECHAZADAS Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.RIESGOSACEPTADOSRECHAZADOS Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.RIESGOSRETIRADOSRECHAZADOS Then
        AbrirExplicacion m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.EDICIONAPUNTODECADUCARSINPROPUESTA Then
        SeleccionarEdicion m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.EDICIONPROPUESTARECHAZADA Then
        SeleccionarEdicion m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.RIESGOACEPTADORECHAZADO Or _
        m_TipoNodoSeleccionado = EnumTipoNodoTareasTecnico.RIESGORETIRADORECHAZADO Then
        SeleccionarRiesgoAceptadoRetirado m_Error
        If m_Error <> "" Then
            Err.Raise 1000
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

Private Sub ComandoActualizarContador_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    EstablecerTareasTecnico EnumSiNo.Sí, m_Error
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
    
Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    
    Me.Caption = "TAREAS PENDIENTES " & m_ObjEntorno.TituloUsuarioConectado
    CargarArbol m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Sep.Left = Me.FormDetalle.Left - Me.Sep.Width
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
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
Public Function CargarArbol(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.FormDetalle.SourceObject = ""
    Set m_ObjArbolTecnico = New ArbolTareasTecnico
    Set m_Arbol = Me.Arbol.Object
    Set m_Imagelist = Me.ListaImagenes.Object
    
    m_ObjArbolTecnico.CargarArbolTareasTecnico m_Arbol, m_Imagelist, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbol ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function




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
