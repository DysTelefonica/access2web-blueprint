Option Compare Database
Option Explicit
Private m_Error As String
Private fso As New FileSystemObject
Private m_ObjInforme As Informe
Private m_ParaRegularizar As Boolean
Dim m_ObjEventoSeleccionado As Evento
Private m_Col As Scripting.Dictionary
Private WithEvents m_FormEvento As Form_FormEventoEdicion
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub ComandoDetalleEvento_Click()
    Dim m_EstaFacturado As String
    Dim m_EstaFranqueado As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjEventoSeleccionado Is Nothing Then
        Set m_ObjEventoSeleccionado = Constructor.getEvento(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjEventoSeleccionado Is Nothing Then
            m_Error = "Seleccione un elemento de esta lista"
            Err.Raise 1000
        End If
    End If
    If IsDate(m_ObjEventoSeleccionado.FechaFacturado) Then
        m_EstaFacturado = "Sí"
    Else
        m_EstaFacturado = "No"
    End If
    If IsDate(m_ObjEventoSeleccionado.FechaFinal) Then
        m_EstaFranqueado = "Sí"
    Else
        m_EstaFranqueado = "No"
    End If
    If FormularioAbierto("FormEventoEdicion") Then
        DoCmd.Close acForm, "FormEventoEdicion", acSaveNo
    End If
    DoCmd.OpenForm "FormEventoEdicion", , , , , , m_ObjEventoSeleccionado.IDEVENTO & "|" & m_EstaFacturado & "|" & m_EstaFranqueado
    If FormularioAbierto("FormEventoEdicion") Then
        Set m_FormEvento = Forms("FormEventoEdicion")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoDetalleEvento_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoExportarExcel_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    '------------------------------------------
    '   -Devuelve:
    '       GenerarConsultas = Descriptivo
    '       GenerarConsultas = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = GenerarConsultas("Eventos para Informe", Me.hWnd, , Me.ListaFiltrados, "No")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método GenerarConsultas ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub




Private Sub ComandoGenerarInforme_Click()
    
    
    Dim m_URLInforme As String
    Dim m_Alcanzable As Boolean
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_Col Is Nothing Then
        m_Error = "No hay elementos en la lista"
        Err.Raise 1000
    End If
    Set m_ObjInforme = New Informe
    m_ObjInforme.EventosRAC m_Col, m_URLInforme, True, m_Error
    Set m_ObjInforme = Nothing
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLInforme) Then
        m_Error = "No se ha podido obtener la URL del informe"
        Err.Raise 1000
    End If
    
    Ejecutar Me.hWnd, "open", m_URLInforme, "", "", 1
    GrabarFechaInformeRACACole m_Col
    Me.cmdSalir.SetFocus
    m_Linea = "Estableciendo la lista de Averías pendientes de Informe"
    Avance m_Linea
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarInforme_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRecargar_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRecargar_Click ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AjustarTamaño Me


    If m_ObjEntorno Is Nothing Then
        EVE
    End If
    
    Me.Caption = Me.lblTitulo1.Caption
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaFiltrados_Click()

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Me.ComandoDetalleEvento.Enabled = False
    
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Set m_ObjEventoSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjEventoSeleccionado Is Nothing Then
        Set m_ObjEventoSeleccionado = Constructor.getEvento(Me.ListaFiltrados.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjEventoSeleccionado.IDEVENTO <> Me.ListaFiltrados.Column(0) Then
            Set m_ObjEventoSeleccionado = Constructor.getEvento(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    Me.ComandoDetalleEvento.Enabled = True
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Clickse ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If Me.ComandoDetalleEvento.Enabled Then
        ComandoDetalleEvento_Click
    End If
    Application.Echo True
End Sub

Private Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    
    Dim m_ID As Variant
    Dim m_Evento As Evento
    Dim lst As ListBox
    
    On Error GoTo errores
    Set lst = Me.ListaFiltrados
    lst.RowSource = "P.TRABAJO;DESCRIPCIÓN;FINICIAL;FFINAL"
    Set m_ObjEventoSeleccionado = Nothing
    Me.ComandoDetalleEvento.Enabled = False
    Me.ComandoExportarExcel.Enabled = False
    Me.ComandoGenerarInforme.Enabled = False
    Set m_Col = Constructor.getEventosSinFechaInformeRAC(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Me.ComandoExportarExcel.Enabled = False
        Me.ComandoGenerarInforme.Enabled = False
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Evento = m_Col(m_ID)
        lst.AddItem m_ID & ";" & m_Evento.Descripcion & ";" & m_Evento.FECHAALTAEVENTO & ";" & m_Evento.FechaFinal
        Set m_Evento = Nothing
    Next
    Me.lblFiltrados.Caption = "Sin  Informe Previo (" & m_Col.count & ")"
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaFiltrados_Click
    End If
    Me.ComandoExportarExcel.Enabled = True
    Me.ComandoGenerarInforme.Enabled = True
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Sub m_FormEvento_FechaInformeRACambiada()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormEvento_FechaInformeRACambiada ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub