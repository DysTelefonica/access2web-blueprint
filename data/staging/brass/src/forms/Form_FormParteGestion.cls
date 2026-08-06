Option Compare Database
Option Explicit

Private m_Error As String
Private m_ObjEventoSeleccionado As Evento
Private m_CentroSeleccionado As String
Private m_ObjParteSeleccionado As Parte
Private fso As New Scripting.FileSystemObject
Private m_ObjFacturaElegida As Factura
Public Event Adjuntado()
Private Sub ComandoAdjuntarParteFirmado_Click()
    
    Dim m_URLLocal As String, strNombreAnexo As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_ObjParteSeleccionado Is Nothing Then
        m_Error = "Se ha de seleccionar un elemento de la lista superior"
        Err.Raise 1000
    End If
    If m_ObjParteSeleccionado.URLParte <> "" Then
        If fso.FileExists(m_ObjParteSeleccionado.URLParte) Then
            DoCmd.Hourglass False
            pregunta = MsgBox("Actualmente ya hay un anexo registrado. " & vbCrLf & _
                    "¿Desea sobreescribirlo con otro?", vbExclamation + vbYesNo + vbDefaultButton2, "Ya existe parte anexo")
            If pregunta <> vbYes Then
                Exit Sub
            End If
            If FicheroAbierto(m_ObjParteSeleccionado.URLParte) Then
                m_Error = "Ha de cerrar el anexo primero"
                Err.Raise 1000
            End If
        End If
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    flag = Seleccionar(True, "Seleccione el documento")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método Seleccionar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_URLLocal = flag
    If Not fso.FileExists(m_URLLocal) Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_ObjParteSeleccionado.AnexarParte m_URLLocal, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ListaFiltrados_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAdjuntarParteFirmado_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoEliminarParte_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjParteSeleccionado Is Nothing Then
        m_Error = "Ha de seleccionar el parte"
        Err.Raise 1000
    End If
    pregunta = MsgBox("Se va a proceder a eliminar los partes de la Factura elegida" & vbNewLine & _
                "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar partes")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjParteSeleccionado.Factura Is Nothing Then
        m_Error = "No se sabe de qué factura es el parte seleccionado"
        Err.Raise 1000
    End If
    m_ObjParteSeleccionado.Factura.EliminarPartes m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ImagenParteAnexado.Visible = False
    Set m_ObjFacturaElegida = Nothing
    IDFactura_Exit False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarParte_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoExportarExcelEventos_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    '------------------------------------------
    '   -Devuelve:
    '       GenerarConsultas = Descriptivo
    '       GenerarConsultas = "#ERR" & "|" & m_Error
    '--------------------------------------------
    flag = GenerarConsultas("EventosEnInformeAverias", Me.hWnd, , Me.ListaEventosEnIDParte, "No")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método GenerarConsultas ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoExportarExcelEventos_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub


Private Sub ComandoGenerarPartesDeTrabajo_Click()
    
    Dim m_URLDirectorio As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_ObjFacturaElegida Is Nothing Then
        If Nz(Me.IDFactura.Column(0), "") <> "" Then
            Set m_ObjFacturaElegida = Constructor.getFactura(Me.IDFactura.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaElegida Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    m_URLDirectorio = Seleccionar(False, "Indique la carpeta de los partes")
    If Not fso.FolderExists(m_URLDirectorio) Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    m_ObjFacturaElegida.GenerarPartesDeFactura m_URLDirectorio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjFacturaElegida.ColPartes = Nothing
    ListaFiltrados_Click
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPartesParaCentros_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoGenerarWordParaFirma_Click()
    Dim m_URL As String
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjParteSeleccionado Is Nothing Then
        m_Error = "Seleccione un Parte de la lista "
        Err.Raise 1000
    End If
    m_URL = m_ObjParteSeleccionado.GeneraWordParaFirma(m_ObjFacturaElegida, m_ObjParteSeleccionado.Centro, m_ObjParteSeleccionado.colEventos, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If fso.FileExists(m_URL) Then
        Ejecutar Me.hWnd, "open", m_URL, "", "", 1
    End If
    Me.lblEstado.Visible = False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarWordParaFirma_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerDetalleEventoEnParte_Click()
    
    On Error GoTo errores
    m_Error = ""
    If m_ObjEventoSeleccionado Is Nothing Then
        m_Error = "Seleccione un evento de la lista superior"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '------------------------------------------
    '   -Devuelve:
    '       FicheroAbierto = true or false
    '       FicheroAbierto =False-->Error
    '------------------------------------------
    If Not FormularioAbierto("FormEventoGestion") Then
        DoCmd.OpenForm "FormEventoGestion"
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = m_ObjEventoSeleccionado.IDEVENTO
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    Else
        Forms("FormEventoGestion").SetFocus
        If Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value <> m_ObjEventoSeleccionado.IDEVENTO Then
            Form_FormEventoGestion.ComandoActualizarTodo_Click
            Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = m_ObjEventoSeleccionado.IDEVENTO
            Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerDetalleEventoEnParte_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub Form_Open(Canel As Integer)
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    Me.ListaFiltrados.RowSource = "IDParte;Centro;FInicial;FFinal"
    Me.ListaEventosEnIDParte.RowSource = "P.TRABAJO;DESCRIPCIÓN"
    Me.lblNumeroPartes.Caption = "Elementos en la lista (0)"
    Me.lblNumeroEventos.Caption = "Elementos en la lista (0)"
'    m_SQL = "SELECT TbCodActividad.CODIGO, TbCodActividad.ACTIVIDADNOPROGRAMADA FROM TbCodActividad ORDER BY TbCodActividad.CODIGO;"
'    Set cbo = Me.ComboActividad
'    RellenaCombo cbo, m_SQL

    
    getValoresDistintos Me.IDFactura, "TbFacturaPrincipal", "IDFactura", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.lblEstado.Visible = False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub IDFactura_Change()
    Dim m_IDFactura As String
    
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    m_IDFactura = Nz(Me.IDFactura, "")
    Me.ListaFiltrados.RowSource = "IDParte;Centro;FInicial;FFinal"
    Me.ListaEventosEnIDParte.RowSource = "P.TRABAJO;DESCRIPCIÓN"
    Me.lblNumeroEventos.Caption = "Elementos En La lista (0)"
    Me.lblNumeroPartes.Caption = "Elementos En La lista (0)"
    Me.ImagenParteAnexado.Visible = False
    
    If m_IDFactura = "" Then
        Set m_ObjFacturaElegida = Nothing
        Me.ComandoGenerarPartesDeTrabajo.Enabled = False
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_ObjFacturaElegida Is Nothing Then
        Set m_ObjFacturaElegida = Constructor.getFactura(m_IDFactura, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjFacturaElegida.IDFactura <> m_IDFactura Then
            Set m_ObjFacturaElegida = Constructor.getFactura(m_IDFactura, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjFacturaElegida.ColPartes Is Nothing Then
        Me.ComandoGenerarPartesDeTrabajo.Enabled = True
    Else
        EstablecerListaPartes m_ObjFacturaElegida, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al IDFactura_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub IDFactura_Exit(Cancel As Integer)
    
    Dim m_IDFactura As String
    
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    m_IDFactura = Nz(Me.IDFactura, "")
    Me.ListaFiltrados.RowSource = "IDParte;Centro;FInicial;FFinal"
    Me.ListaEventosEnIDParte.RowSource = "P.TRABAJO;DESCRIPCIÓN"
    Me.lblNumeroEventos.Caption = "Elementos En La lista (0)"
    Me.lblNumeroPartes.Caption = "Elementos En La lista (0)"
    Me.ImagenParteAnexado.Visible = False
    
    If m_IDFactura = "" Then
        Set m_ObjFacturaElegida = Nothing
        Me.ComandoGenerarPartesDeTrabajo.Enabled = False
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_ObjFacturaElegida Is Nothing Then
        Set m_ObjFacturaElegida = Constructor.getFactura(m_IDFactura, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjFacturaElegida.IDFactura <> m_IDFactura Then
            Set m_ObjFacturaElegida = Constructor.getFactura(m_IDFactura, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjFacturaElegida.ColPartes Is Nothing Then
        Me.ComandoGenerarPartesDeTrabajo.Enabled = True
    Else
        EstablecerListaPartes m_ObjFacturaElegida, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al IDFactura_Exit se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ImagenParteAnexado_Click()
    
    On Error GoTo errores
    DoCmd.Hourglass True
    
    flag = AbrirEnLocal(m_ObjParteSeleccionado.URLParte, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenParteAnexado_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ListaEventosEnIDParte_Click()
    
    Dim m_IDEvento As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_IDEvento = Nz(Me.ListaEventosEnIDParte.Column(0), "")
    If m_IDEvento = "" Then
        Set m_ObjEventoSeleccionado = Nothing
        Me.ComandoVerDetalleEventoEnParte.Enabled = False
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_ObjEventoSeleccionado Is Nothing Then
        Set m_ObjEventoSeleccionado = Constructor.getEvento(m_IDEvento, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    Else
        If m_ObjEventoSeleccionado.IDEVENTO <> m_IDEvento Then
            Set m_ObjEventoSeleccionado = Constructor.getEvento(m_IDEvento, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            
        End If
        
    End If
    Me.ComandoVerDetalleEventoEnParte.Enabled = True
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaEventosEnIDParte_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ListaFiltrados_Click()
    
    Dim m_URL As String
    Dim m_IDParte As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_IDParte = Nz(Me.ListaFiltrados.Column(0), "")
    If m_IDParte = "" Then
        Set m_ObjParteSeleccionado = Nothing
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_ObjParteSeleccionado Is Nothing Then
        Set m_ObjParteSeleccionado = Constructor.getParte(m_IDParte, , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    Else
        If m_ObjParteSeleccionado.IDParte <> m_IDParte Then
            Set m_ObjParteSeleccionado = Constructor.getParte(m_IDParte, , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            
        End If
        
    End If
    Me.ComandoAdjuntarParteFirmado.Enabled = True
    Me.ComandoEliminarParte.Enabled = True
    Me.ComandoGenerarWordParaFirma.Enabled = True
    m_URL = m_ObjParteSeleccionado.URLParte
    m_Error = m_ObjParteSeleccionado.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL <> "" Then
        Me.ImagenParteAnexado.Visible = True
    Else
        Me.ImagenParteAnexado.Visible = False
    End If
    EstablecerListaEventos m_ObjParteSeleccionado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub

Public Function EstablecerListaPartes( _
                                        ByVal p_ObjFactura As Factura, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim lstPartes As ListBox
    Dim lstEventos As ListBox
    Dim m_ColPartes As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_objParte As Parte
    
    
    On Error GoTo errores
    Set lstPartes = Me.ListaFiltrados
    Set lstEventos = Me.ListaEventosEnIDParte
    lstPartes.RowSource = "IDParte;Centro;FInicial;FFinal"
    lstEventos.RowSource = "P.TRABAJO;DESCRIPCIÓN"
    Me.ComandoGenerarPartesDeTrabajo.Enabled = False
    Me.ComandoAdjuntarParteFirmado.Enabled = False
    Me.ComandoEliminarParte.Enabled = False
    Me.ComandoGenerarWordParaFirma.Enabled = False
    If p_ObjFactura Is Nothing Then
        Exit Function
    End If
    Set m_ColPartes = p_ObjFactura.ColPartes
    p_Error = p_ObjFactura.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColPartes Is Nothing Then
        
        Exit Function
    End If
    Me.ComandoGenerarPartesDeTrabajo.Enabled = True
    
    For Each m_ID In m_ColPartes.Keys
        Set m_objParte = m_ColPartes(m_ID)
        lstPartes.AddItem m_ID & ";" & m_objParte.Centro & ";" & m_objParte.FechaInicial & ";" & m_objParte.FechaFinal
        Set m_objParte = Nothing
    Next
    Me.lblNumeroPartes.Caption = "Elementos en la lista (" & lstPartes.ListCount - 1 & ")"
    
    If lstPartes.ListCount = 2 Then
        lstPartes.Selected(1) = True
        ListaFiltrados_Click
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerListaPartes ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function EstablecerListaEventos( _
                                        ByVal p_ObjParte As Parte, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    
    Dim lstEventos As ListBox
    Dim m_ColEventos As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjEvento As Evento
    
    On Error GoTo errores
    
    Set lstEventos = Me.ListaEventosEnIDParte
    lstEventos.RowSource = "P.TRABAJO;DESCRIPCIÓN"
    Me.ComandoVerDetalleEventoEnParte.Enabled = False
    If p_ObjParte Is Nothing Then
        Exit Function
    End If
    Set m_ColEventos = p_ObjParte.colEventos
    p_Error = p_ObjParte.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColEventos Is Nothing Then
        Exit Function
    End If
    
    
    For Each m_ID In m_ColEventos
        Set m_ObjEvento = m_ColEventos(m_ID)
        lstEventos.AddItem m_ID & ";" & m_ObjEvento.Descripcion
        Set m_ObjEvento = Nothing
    Next
    Me.lblNumeroEventos.Caption = "Elementos en la lista (" & lstEventos.ListCount - 1 & ")"
    If lstEventos.ListCount = 2 Then
        lstEventos.Selected(1) = True
        ListaEventosEnIDParte_Click
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerListaEventos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function


