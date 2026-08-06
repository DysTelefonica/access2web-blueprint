Option Compare Database
Option Explicit

Private fso As New FileSystemObject
Private m_Error As String

Private m_ObjColEventosEnFactura As Scripting.Dictionary
Private m_ObjActividadesDeEventosEnFactura As Scripting.Dictionary
Private m_ObjActividadesNoDeEventosEnFactura As Scripting.Dictionary
Private m_ObjMaterialesEnFactura As Scripting.Dictionary
Private m_ObjGastosEnFactura As Scripting.Dictionary
Private m_ObjSubContratacionesEnFactura As Scripting.Dictionary

Public Event Alta()

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Public Function RellenarListaPerfiles(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ObjColPrecios As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjTipoPrecio As TipoTecnicoPrecio
    
    On Error GoTo errores
    Set lst = Me.ListaPreciosPerfil
    lst.RowSource = "TIPO;€HLAB;€HEXTRA;FECHA INICIO;FECHA FIN"
    
    Set m_ObjColPrecios = m_ObjEntorno.ColPerfilesActivos
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColPrecios Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjColPrecios
        Set m_ObjTipoPrecio = m_ObjColPrecios(m_ID)
        With m_ObjTipoPrecio
            lst.AddItem .TIPOTECNICO & ";" & Format(.PrecioHoraLab, "#,##0.00 €") & ";" & Format(.PrecioHoraExtra, "#,##0.00 €") & ";" & .FechaInicial & ";" & .FechaFinal
        End With
        Set m_ObjTipoPrecio = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListaPerfiles ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function



Private Sub ComandoGenerarFactura_Click()
    
    Dim m_ObjFactura As Factura
    
    On Error GoTo errores
    m_Error = ""
    
    pregunta = MsgBox("El proceso puede tardar varios minutos. Tenga paciencia", vbExclamation, "Generación de Factura")
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_ObjFacturaActiva = Nothing
    Set m_ObjFactura = New Factura
    With m_ObjFactura
        .ParaPorFacturar = EnumSino.Sí
        .Titulo1 = Nz(Me.Titulo1, "")
        .Titulo2 = Nz(Me.Titulo2, "")
        .FechaFinalEventos = Nz(Me.FechaFinalEventos, "")
        .FechaFactura = .FechaFinalEventos
        .Recargo = Nz(Me.RecargoSobreSuministros, "")
        .TipoImpositivo = Nz(Me.TipoImpositivo, "")
        .ValidacionFactura EnumSino.No, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        .GrabarFactura m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Set m_ObjFacturaActiva = Constructor.getFactura(m_ObjFactura.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjFacturaActiva Is Nothing Then
        m_Error = "No se ha podido recuperar la factura recién generada"
        Err.Raise 1000
    End If
    m_ObjFacturaActiva.SalidaAExcel EnumSino.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblEstado.Visible = False
    Ejecutar Me.hWnd, "open", m_ObjFacturaActiva.URLFactura, "", "", 1
    m_ObjFacturaActiva.ParaPorFacturar = EnumSino.No
    If FormularioAbierto("FormFacturaResultado") Then
        DoCmd.Close acForm, "FormFacturaResultado", acSaveNo
    End If
    DoCmd.OpenForm "FormFacturaResultado"
    If FormularioAbierto("FormFacturacionGestion") Then
        Form_FormFacturacionGestion.RellenarComboIDFactura
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarFactura_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoPerfiles_Click()
    On Error GoTo errores
    
    m_Error = ""
    DoCmd.OpenForm "FormTipoTecnicoPreciosGestion"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPerfiles_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoSimulacro_Click()
    
    Dim m_ObjFactura As Factura
    
    On Error GoTo errores
    
    pregunta = MsgBox("El proceso puede tardar varios minutos. Tenga paciencia", vbExclamation, "Simulacro de Factura")
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Set m_ObjFactura = New Factura
    With m_ObjFactura
        .FechaFinalEventos = Nz(Me.FechaFinalEventos, "")
        .Titulo1 = Nz(Me.Titulo1, "")
        .Titulo2 = Nz(Me.Titulo2, "")
        .TipoImpositivo = Nz(Me.TipoImpositivo, "")
        .Recargo = Nz(Me.RecargoSobreSuministros, "")
        .FechaFactura = Nz(Me.FechaFinalEventos, "")
        .ValidacionFactura EnumSino.No, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_ObjFacturaActiva = m_ObjFactura
    With m_ObjFacturaActiva
        .SalidaAExcel EnumSino.No, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
'        m_Linea = "Obteniendo la lista............"
'        Avance m_Linea
'        .ObtenerLI EnumSino.Sí, m_Error
'        If m_Error <> "" Then
'            Err.Raise 1000
'        End If
'        m_Linea = "Obteniendo Partes por centro............"
'        Avance m_Linea
'        .GenerarPartesDeFactura .URLSubCarpetaLocal, m_Error
'        If m_Error <> "" Then
'            Err.Raise 1000
'        End If
        Me.lblEstado.Visible = False
        Ejecutar Me.hWnd, "open", .URLFactura, "", "", 1
        m_ObjFacturaActiva.ParaPorFacturar = EnumSino.Sí
        If FormularioAbierto("FormFacturaResultado") Then
            DoCmd.Close acForm, "FormFacturaResultado", acSaveNo
        End If
        DoCmd.OpenForm "FormFacturaResultado"
    End With
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSimulacro_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub Form_Load()
        
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    
    Me.Caption = Me.lblTitulo1.Caption
    Me.lblEstado.Visible = False
   
    RellenarListaPerfiles m_Error
    If m_Error <> "" Then
         Err.Raise 1000
    End If
    Me.TipoImpositivo = "21"
    Me.RecargoSobreSuministros = "15"
    Me.FechaFinalEventos = Date
    
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub









