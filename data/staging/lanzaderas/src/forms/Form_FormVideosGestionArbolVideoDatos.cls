Option Compare Database
Option Explicit
Private m_Error As String
Private m_DatosCargados As EnumSino


Private m_ObjVideoAlInicio As Video

Private m_Arbol As MSComctlLib.TreeView
Private ParaEncuesta As Boolean
Private ParaAyuda As Boolean
Private ParaTecnicos As String
Private ParaCalidad As String
Private ParaAdministrador As String


Public Function CargarDatos(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    If m_ObjAplicacionActiva Is Nothing Then
        p_Error = "No se ha seleccionado ningún nodo"
        Err.Raise 1000
    End If
    With m_ObjAplicacionActiva
        Me.lblNombreAplicacion.Caption = UCase(.NombreAplicacion)
        Me.ImagenAplicacion.Picture = .URLIconoAcceso
    End With
    
    If m_ObjVideoActivo Is Nothing Then
        Me.lblOperacion.Caption = "ALTA DE VÍDEO"
        Me.ComandoEliminar.Visible = False
        Me.ComandoRegistrar.Enabled = True
        Me.ComandoAnadirVideo.Enabled = True
        Me.URLLocalVideo = Null
        Me.Titulo = Null
        ParaAyuda = False
        ParaEncuesta = False
        Me.Descripcion = Null
        Me.Observaciones = Null
        Me.ComboDirigidoA = "Para Todos"
        Me.ComboUso = "Vídeo-Guía"
        Exit Function
    End If
    Set m_ObjVideoAlInicio = Constructor.getVideo(m_ObjVideoActivo.IDVideo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblOperacion.Caption = "EDICIÓN DE DATOS DE VÍDEO"
    Me.ComandoAnadirVideo.Enabled = False
    Me.ComandoEliminar.Visible = True
    Me.ComandoEliminar.Enabled = True
    Me.ComandoRegistrar.Enabled = True
    
    With m_ObjVideoActivo
        Me.Titulo = .Titulo
        If .EsDeEncuesta = EnumSino.Sí Then
            If .EsDeAyuda = EnumSino.Sí Then
                Me.ComboUso = "Ambos"
            Else
                Me.ComboUso = "Encuesta"
            End If
        Else
            If .EsDeAyuda = EnumSino.Sí Then
                Me.ComboUso = "Vídeo-Guía"
            End If
        End If
        If .Descripcion = "" Then
            Me.Descripcion = Null
        Else
            Me.Descripcion = .Descripcion
        End If
        If .Observaciones = "" Then
            Me.Observaciones = Null
        Else
            Me.Observaciones = .Observaciones
        End If
        If .ParaAdministrador = "Sí" And .ParaCalidad = "Sí" And .ParaTecnicos = "Sí" Then
            Me.ComboDirigidoA = "Para Todos"
        Else
            If .ParaTecnicos = "No" Then
                If .ParaCalidad = "Sí" Then
                    Me.ComboDirigidoA = "No Para Técnicos"
                Else
                    Me.ComboDirigidoA = "No Para Calidad"
                End If
            Else
                Me.ComboDirigidoA = "Para Todos"
            End If
        End If
        
        
    End With
    
        
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




Private Sub ComandoAnadirVideo_Click()
    Dim m_URL As String
    On Error GoTo errores
    
    m_URL = Seleccionar(True, "Seleccione el Vídeo a subir", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        Me.URLLocalVideo = Null
        Exit Sub
    End If
    Me.URLLocalVideo = m_URL
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ComandoAnadirVideo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
    Dim m_IDVideo As String
    
    On Error GoTo errores
    
    
    m_Error = ""
    
    If m_ObjVideoActivo Is Nothing Then
        m_Error = "No hay ningún vídeo activo"
        Err.Raise 1000
    End If
    Pregunta = MsgBox("Desea eliminar el vídeo seleccionado", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Vídeo")
    If Pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    
    m_IDVideo = m_ObjVideoActivo.IDVideo
    m_ObjVideoActivo.Eliminar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Form_FormVideosGestion.ComandoActualizarContador_Click
    
    'RaiseEvent VideoEliminado(m_IDVideo)
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
     Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ComandoEliminar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrar_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Not HaHabidoCambios(m_Error) Then
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "Ningún cambio que grabar"
        Err.Raise 1000
    End If
    Set m_ObjVideoActivo = New Video
    With m_ObjVideoActivo
        If m_ObjVideoAlInicio Is Nothing Then
            .URLLocal = Nz(Me.URLLocalVideo, "")
        End If
        
        .Titulo = Nz(Me.Titulo, "")
        .Descripcion = Nz(Me.Descripcion, "")
        .Observaciones = Nz(Me.Observaciones, "")
        If ParaAyuda = True Then
            .EsDeAyuda = EnumSino.Sí
        Else
            .EsDeAyuda = EnumSino.No
        End If
        If ParaEncuesta = True Then
            .EsDeEncuesta = EnumSino.Sí
        Else
            .EsDeEncuesta = EnumSino.No
        End If
        .ParaAdministrador = ParaAdministrador
        .ParaCalidad = ParaCalidad
        .ParaTecnicos = ParaTecnicos
        .IDAplicacion = m_ObjAplicacionActiva.IDAplicacion
        .Registrar m_ObjVideoAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If .IDVideo = "" Then
            If Not m_ObjVideoAlInicio Is Nothing Then
                .IDVideo = m_ObjVideoAlInicio.IDVideo
            End If
        End If
    End With
    Set m_ObjVideoActivo = Constructor.getVideo(m_ObjVideoActivo.IDVideo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjAplicacionActiva = Constructor.getAplicacion(m_ObjVideoActivo.IDAplicacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Arbol Is Nothing Then
        Set m_Arbol = Form.Parent.Parent.Controls("Arbol").Object
    End If
    If m_ObjVideoAlInicio Is Nothing Then
        
        CrearNodo m_Arbol, m_Arbol.SelectedItem.Parent, m_ObjVideoActivo, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjVideoAlInicio.Titulo <> m_ObjVideoActivo.Titulo Then
            m_Arbol.SelectedItem.Text = m_ObjVideoActivo.Titulo
        End If
    End If
    

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ComandoRegistrar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    Set m_Arbol = Me.Parent.Parent.Controls("Arbol").Object
    
End Sub

Private Sub TransponerDatoUso(Optional ByRef p_Error As String)
    
    Dim m_Uso As String
    On Error GoTo errores
    m_Uso = Nz(Me.ComboUso, "")
    If m_Uso = "" Then
        ParaEncuesta = False
        ParaAyuda = False
        Exit Sub
    End If
    If m_Uso = "Vídeo-Guía" Then
        ParaEncuesta = False
        ParaAyuda = True
        Exit Sub
    End If
    If m_Uso = "Encuesta" Then
        ParaEncuesta = True
        ParaAyuda = False
        Exit Sub
    End If
    If m_Uso = "Ambos" Then
        ParaEncuesta = True
        ParaAyuda = True
        Exit Sub
    End If
    
    Exit Sub
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método TransponerDatoUso ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Sub
Private Sub TransponerDirigidoA(Optional ByRef p_Error As String)
    
    Dim m_DirigidoA As String
    On Error GoTo errores
    m_DirigidoA = Nz(Me.ComboDirigidoA, "")
    If m_DirigidoA = "" Then
        ParaTecnicos = "Sí"
        ParaCalidad = "Sí"
        ParaAdministrador = "Sí"
        Exit Sub
    End If
    If m_DirigidoA = "Para Todos" Then
        ParaTecnicos = "Sí"
        ParaCalidad = "Sí"
        ParaAdministrador = "Sí"
        Exit Sub
    End If
    If m_DirigidoA = "No Para Técnicos" Then
        ParaTecnicos = "No"
        ParaCalidad = "Sí"
        ParaAdministrador = "Sí"
        Exit Sub
    End If
    If m_DirigidoA = "No Para Calidad" Then
        ParaTecnicos = "No"
        ParaCalidad = "No"
        ParaAdministrador = "Sí"
        Exit Sub
    End If
    Exit Sub
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método TransponerDirigidoA ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Sub
Private Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    On Error GoTo errores
    
    TransponerDatoUso p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    TransponerDirigidoA p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjVideoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    
    With m_ObjVideoAlInicio
        If .Titulo <> Nz(Me.Titulo, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Descripcion <> Nz(Me.Descripcion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .EsDeAyuda = EnumSino.Sí Then
            If ParaAyuda = False Then
                HaHabidoCambios = True
                Exit Function
            End If
        Else
            If ParaAyuda = True Then
                HaHabidoCambios = True
                Exit Function
            End If
        End If
        If .EsDeEncuesta = EnumSino.Sí Then
            If ParaEncuesta = False Then
                HaHabidoCambios = True
                Exit Function
            End If
        Else
            If ParaEncuesta = True Then
                HaHabidoCambios = True
                Exit Function
            End If
        End If
        If .Observaciones <> Nz(Me.Observaciones, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .ParaAdministrador <> ParaAdministrador Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .ParaCalidad <> ParaCalidad Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .ParaTecnicos <> ParaTecnicos Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    Exit Function
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function