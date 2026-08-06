Option Compare Database
Option Explicit

Private fso As New FileSystemObject
Private m_ObjFacturaSeleccionada As Factura
Private WithEvents m_Form As Form_FormFacturaAlta
Private WithEvents m_FormPartes As Form_FormParteGestion
Private m_Error As String



Public Sub ActualizarLista()
    
    Dim lst As ListBox
    Dim m_ColFacturas As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjFactura As Factura
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set lst = Me.ListaFiltrados
    lst.RowSource = "ID;Título;Fecha"
    
    'Me.ComandoVerDatosFactura.Enabled = False
    Me.ComandoAnexarFirmados.Enabled = False
    Me.ComandoPartesParaCentros.Enabled = False
    Me.ComandoEliminarFactura.Enabled = False
    Me.ComandoObtenerLI.Enabled = False
    Me.ComandoAnexoAutorizacionMateriales.Enabled = False
    Me.ComandoQuitarAnexoAutorizacionMateriales.Enabled = False
    Me.ComandoVerDetalles.Enabled = False
    Me.ImagenFacturacionAnexada.Visible = False
    
    If Nz(Me.ComboIDFactura, "") <> "" Then
        Set m_ObjFactura = Constructor.getFactura(Me.ComboIDFactura, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_ObjFactura Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    
        With m_ObjFactura
            lst.AddItem .IDFactura & ";" & Replace(.Titulo1, ";", ":") & ";" & Format(.FechaFactura, "dd/mm/yyyy")
        End With
        
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaFiltrados_Click
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ActualizarLista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoActualizar_Click()
    On Error Resume Next
    RellenarComboIDFactura
End Sub

Private Sub ComandoAnexarFirmados_Click()
    
    On Error GoTo errores
    m_Error = ""
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    If m_ObjFacturaSeleccionada Is Nothing Then
        m_Error = "La factura seleccionada no tiene ningún parte de trabajo generado"
        Err.Raise 1000
    End If
    If Not FormularioAbierto("FormParteGestion") Then
        DoCmd.OpenForm "FormParteGestion"
        Forms("FormParteGestion").Controls("IDFactura") = m_ObjFacturaSeleccionada.IDFactura
        Form_FormParteGestion.IDFactura_Exit False
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexarFirmados_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAnexoAutorizacionMateriales_Click()
    
    Dim strURL As String
    Dim m_ColAnexos As Scripting.Dictionary
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    flag = Seleccionar(True, "Seleccione la autorización para materiales")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método Seleccionar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strURL = flag
    If strURL = "" Then
        Exit Sub
    End If
    If fso.FileExists(strURL) Then
        If FicheroAbierto(strURL) Then
            m_Error = "Tiene el fichero abierto"
            Err.Raise 1000
        End If
    Else
        Exit Sub
    End If
    pregunta = MsgBox("¿Desea anexar el documento a todos los materiales de la factura seleccionada?", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Anexo de autorización")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Linea = "Anexando Autorizaciones............"
    Avance m_Linea
    m_ObjFacturaSeleccionada.AnexoAutorizacionEnMaterialesPorFactura strURL, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_ColAnexos = m_ObjFacturaSeleccionada.ColAnexosAutorizacionMaterialesEnFactura
    m_Error = m_ObjFacturaSeleccionada.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColAnexos Is Nothing Then
        Me.ComandoQuitarAnexoAutorizacionMateriales.Enabled = False
    Else
        Me.ComandoQuitarAnexoAutorizacionMateriales.Enabled = True
    End If
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox(flag, vbInformation, "Anexo de autorización Resultado")
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexoAutorizacionMateriales_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Me.lblEstado.Visible = False
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    Exit Sub
End Sub

Private Sub ComandoEliminarFactura_Click()
    
    On Error GoTo errores
    m_Error = ""
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    pregunta = MsgBox("¿Desea eliminar la factura?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de Factura")
    If pregunta <> vbYes Then
       Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjFacturaSeleccionada.Eliminar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ComboIDFactura = Null
    RellenarComboIDFactura m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ComboIDFactura_Change
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarFactura_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoNuevaFactura_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If FormularioAbierto("FormFacturaAlta") Then
        DoCmd.Close acForm, "FormFacturaAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormFacturaAlta"
    Set m_Form = Forms("FormFacturaAlta")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNuevaFactura_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoObtenerLI_Click()
    
    Dim m_URLLI As String
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    m_Linea = "Obteniendo la lista............"
    Avance m_Linea
    m_URLLI = m_ObjFacturaSeleccionada.ObtenerLI(EnumSino.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLLI) Then
        m_Error = "El método ObtenerLI ha devuelto un resultado con formato desconocido"
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_URLLI, "", "", 1
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoObtenerLI_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoPartesParaCentros_Click()
    
    Dim m_URLDirectorio As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
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
    m_Linea = "Obteniendo Partes por centro............"
    Avance m_Linea
    m_ObjFacturaSeleccionada.GenerarPartesDeFactura m_URLDirectorio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_ObjFacturaSeleccionada.ColPartes = Nothing
    ListaFiltrados_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Ya se han generado los partes de trabajo por centro", vbInformation, "Partes Generados")
    Ejecutar Me.hWnd, "open", m_URLDirectorio, "", "", 1
    Me.lblEstado.Visible = False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPartesParaCentros_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoQuitarAnexoAutorizacionMateriales_Click()
    
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_Col = Constructor.getAnexosAutorizacionMaterialesEnFactura(m_ObjFacturaSeleccionada.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_Error = "No hay anexos de autorización que eliminar"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea eliminar los " & m_Col.count & " anexos de autorización de materiales para la factura seleccionada?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Anexos de Autorización de materiales")
    If pregunta <> vbYes Then
       Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjFacturaSeleccionada.EliminarAnexosDeAutorizacionEnMaterialesPorFactura m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Col = Constructor.getAnexosAutorizacionMaterialesEnFactura(m_ObjFacturaSeleccionada.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.cmdSalir.SetFocus
    If m_Col Is Nothing Then
        Me.ComandoQuitarAnexoAutorizacionMateriales.Enabled = True
    Else
        Me.ComandoQuitarAnexoAutorizacionMateriales.Enabled = False
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    pregunta = MsgBox("Anexos de autorización eliminados", vbInformation, "Eliminar Anexos de Autorización de materiales")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoQuitarAnexoAutorizacionMateriales_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerDetalles_Click()
    
    Dim m_ObjFactura As Factura
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjFacturaActiva = m_ObjFacturaSeleccionada
    If FormularioAbierto("FormFacturaResultado") Then
        DoCmd.Close acForm, "FormFacturaResultado", acSaveNo
    End If
    m_ObjFacturaActiva.ParaPorFacturar = EnumSino.No
    DoCmd.OpenForm "FormFacturaResultado"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerDetalles_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Public Sub ComboIDFactura_Change()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboIDFactura_Change se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboIDFactura_Enter()
    On Error Resume Next
    Me.ComboIDFactura.Requery
End Sub

Private Sub Form_Open(Canel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    RellenarComboIDFactura m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblEstado.Visible = False
    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ImagenFacturacionAnexada_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjFacturaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjFacturaSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar una factura de la lista"
            Err.Raise 1000
        End If
    End If
    
    flag = AbrirEnLocal(m_ObjFacturaSeleccionada.URLFactura, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenFacturacionAnexada_Clickse ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ListaFiltrados_Click()
    
    Dim m_ID As String
   
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComandoAnexarFirmados.Enabled = False
    Me.ComandoPartesParaCentros.Enabled = False
    Me.ComandoEliminarFactura.Enabled = False
    Me.ComandoObtenerLI.Enabled = False
    Me.ComandoAnexoAutorizacionMateriales.Enabled = False
    Me.ComandoQuitarAnexoAutorizacionMateriales.Enabled = False
    Me.ComandoVerDetalles.Enabled = False
    Me.ImagenFacturacionAnexada.Visible = False
    
    m_ID = Nz(Me.ListaFiltrados.Column(0), "")
    If m_ID = "" Then
        Set m_ObjFacturaSeleccionada = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjFacturaSeleccionada Is Nothing Then
        Set m_ObjFacturaSeleccionada = Constructor.getFactura(m_ID, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ID <> m_ObjFacturaSeleccionada.IDFactura Then
            Set m_ObjFacturaSeleccionada = Constructor.getFactura(m_ID, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjFacturaSeleccionada Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If Not m_ObjFacturaSeleccionada.ColPartes Is Nothing Then
        Me.ComandoAnexarFirmados.Enabled = True
    End If
    
    Me.ComandoPartesParaCentros.Enabled = True
    Me.ComandoEliminarFactura.Enabled = True
    Me.ComandoObtenerLI.Enabled = True
    If Not m_ObjFacturaSeleccionada.ColAnexosAutorizacionMaterialesEnFactura Is Nothing Then
        Me.ComandoQuitarAnexoAutorizacionMateriales.Enabled = True
    Else
        Me.ComandoAnexoAutorizacionMateriales.Enabled = True
    End If
    Me.ComandoVerDetalles.Enabled = True
    
    Me.ImagenFacturacionAnexada.Visible = False
   
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Clicka se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Form_Alta()
    Set m_ObjFacturaSeleccionada = Constructor.getFactura(m_ObjFacturaSeleccionada.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ComboIDFactura.Requery
    Me.ComboIDFactura = m_ObjFacturaSeleccionada.IDFactura
    ActualizarLista
End Sub
Public Function RellenarComboIDFactura(Optional ByRef p_Error As String) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjFactura As Factura
    On Error GoTo errores
    
    Me.ComboIDFactura.RowSource = ""
    Set m_Col = Constructor.getFacturas(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjFactura = m_Col(m_ID)
        With m_ObjFactura
            Me.ComboIDFactura.AddItem m_ID & ";" & Replace(.Titulo1, ";", ":")
        End With
        Set m_ObjFactura = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboIDFactura ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
                
End Function