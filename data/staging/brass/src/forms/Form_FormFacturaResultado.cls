Option Compare Database
Option Explicit


Private fso As New FileSystemObject
Private m_Error As String


Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_FechaFinalEventos As String
    Dim m_FechaFactura As String
    Dim m_Recargo As String
    Dim m_PrecioSuministro As String
    Dim m_PrecioSuministroConRecargo As String
    Dim m_PrecioSubcontrataciones As String
    Dim m_PrecioSubcontratacionesConRecargo As String
    Dim m_PrecioDietasYGastos As String
    
    
    Dim m_PrecioIncurrido As String
    Dim m_PrecioBaseImponible As String
    Dim m_TipoImpositivo As String
    Dim m_PrecioImpuestos As String
    Dim m_ImporteAFacturar As String
    
    On Error GoTo errores
    
    If m_ObjFacturaActiva Is Nothing Then
        m_Error = "No hay ninguna factura activa"
        Err.Raise 1000
    End If
    
    Me.ComandoObtenerLI.Visible = False
    With m_ObjFacturaActiva
        If .IDFactura = "" Then
            
            Me.ComandoObtenerLI.Visible = True
           
            Me.ComandoObtenerLI.Enabled = True
        Else
            
            Me.ComandoObtenerLI.Visible = False
        End If
        m_FechaFinalEventos = .FechaFinalEventos
        m_FechaFactura = .FechaFactura
        m_Recargo = .Recargo
        m_TipoImpositivo = .TipoImpositivo
        If .ParaPorFacturar = 0 Then
            .ParaPorFacturar = EnumSino.Sí
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioSuministro = .PrecioSuministroPorFacturar
        Else
            m_PrecioSuministro = .PrecioSuministroCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioSuministroConRecargo = .PrecioSuministroConRecargoPorFacturar
        Else
            m_PrecioSuministroConRecargo = .PrecioSuministroConRecargoCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioSubcontrataciones = .PrecioSubcontratacionesPorFacturar
        Else
            m_PrecioSubcontrataciones = .PrecioSubcontratacionesCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioSubcontratacionesConRecargo = .PrecioSubcontratacionesConRecargoForFacturar
        Else
            m_PrecioSubcontratacionesConRecargo = .PrecioSubcontratacionesConRecargoCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioDietasYGastos = .PrecioDietasYGastosPorFacturar
        Else
            m_PrecioDietasYGastos = .PrecioDietasYGastosCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioIncurrido = .PrecioIncurridoPorFacturar
        Else
            m_PrecioIncurrido = .PrecioIncurridoCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioBaseImponible = .PrecioBaseImponiblePorFacturar
        Else
            m_PrecioBaseImponible = .PrecioBaseImponibleCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_PrecioImpuestos = .PrecioImpuestosPorFacturar
        Else
            m_PrecioImpuestos = .PrecioImpuestosCalculado
        End If
        If .ParaPorFacturar = EnumSino.Sí Then
            m_ImporteAFacturar = .ImporteAFacturarPorFacturar
        Else
            m_ImporteAFacturar = .ImporteAFacturarCalculado
        End If
    End With
    If IsDate(m_FechaFinalEventos) Then
        Me.FechaFinalEventos = Format(m_FechaFinalEventos, "dd/mm/yyyy")
    End If
    If IsDate(m_FechaFactura) Then
        Me.FechaFactura = Format(m_FechaFactura, "dd/mm/yyyy")
    End If
    If IsNumeric(m_Recargo) Then
        Me.Recargo = Format(m_Recargo, "#,##0.00")
    End If
    If IsNumeric(m_PrecioSuministro) Then
        Me.PrecioSuministro = Format(m_PrecioSuministro, "#,##0.00 €")
    End If
    If IsNumeric(m_PrecioSuministroConRecargo) Then
        Me.PrecioSuministroConRecargo = Format(m_PrecioSuministroConRecargo, "#,##0.00 €")
    End If
    If IsNumeric(m_PrecioSubcontrataciones) Then
        Me.PrecioSubcontrataciones = Format(m_PrecioSubcontrataciones, "#,##0.00 €")
    End If
    If IsNumeric(m_PrecioSubcontratacionesConRecargo) Then
        Me.PrecioSubcontratacionesConRecargo = Format(m_PrecioSubcontratacionesConRecargo, "#,##0.00 €")
    End If
    If IsNumeric(m_PrecioDietasYGastos) Then
        Me.PrecioDietasYGastos = Format(m_PrecioDietasYGastos, "#,##0.00 €")
    End If
    If IsNumeric(m_PrecioIncurrido) Then
        Me.PrecioIncurrido = Format(m_PrecioIncurrido, "#,##0.00 €")
    End If
    If IsNumeric(m_PrecioBaseImponible) Then
        Me.PrecioBaseImponible = Format(m_PrecioBaseImponible, "#,##0.00 €")
    End If
    If IsNumeric(m_TipoImpositivo) Then
        Me.TipoImpositivo = Format(m_TipoImpositivo, "#,##0.00")
    End If
    If IsNumeric(m_PrecioImpuestos) Then
        Me.PrecioImpuestos = Format(m_PrecioImpuestos, "#,##0.00 €")
    End If
    If IsNumeric(m_ImporteAFacturar) Then
        Me.ImporteAFacturar = Format(m_ImporteAFacturar, "#,##0.00 €")
    End If
    RellenarPerfiles p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListaPerfiles ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Sub ComandoGenerarExcel_Click()
    
    Dim m_EsCalculado As EnumSino
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.lblEstado.Visible = False
    If m_ObjFacturaActiva Is Nothing Then
        m_Error = "No hay ninguna factura activa"
        Err.Raise 1000
    End If
    With m_ObjFacturaActiva
        If .ParaPorFacturar = EnumSino.Sí Then
            m_EsCalculado = EnumSino.No
        Else
            m_EsCalculado = EnumSino.Sí
        End If
         .SalidaAExcel m_EsCalculado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If fso.FileExists(.URLFactura) Then
            Ejecutar Me.hWnd, "open", .URLFactura, "", "", 1
        End If
    End With
    Me.lblEstado.Visible = False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarExcel_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If m_ObjFacturaActiva Is Nothing Then
        m_Error = "No hay una factura activa"
        Err.Raise 1000
    End If
    m_Linea = "Obteniendo la lista............"
    Avance m_Linea
    m_URLLI = m_ObjFacturaActiva.ObtenerLI(EnumSino.Sí, m_Error)
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


Private Sub Form_Load()
    
    Dim m_Titulo As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.lblEstado.Visible = False
    
    Me.Caption = Me.lblTitulo1.Caption
    If m_ObjFacturaActiva Is Nothing Then
        m_Error = "No hay ninguna factura activa"
        Err.Raise 1000
    End If
    If m_ObjFacturaActiva.ParaPorFacturar = EnumSino.Sí Then
        m_Titulo = "RESÚMEN DEL SIMULACRO DE LA FACTURA"
    Else
        m_Titulo = "RESÚMEN DE LA FACTURA " & m_ObjFacturaActiva.IDFactura
    End If
    Me.lblTitulo1.Caption = m_Titulo
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    EstablecerDatos m_Error
    If m_Error <> "" Then
         Err.Raise 1000
    End If
        
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

Private Function RellenarPerfiles(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_Col As Scripting.Dictionary
    Dim m_Perfil As FacturaPerfiles
    Dim m_ID As Variant
    Dim m_NHorasLab As String
    Dim m_PrecioHoraLab As String
    Dim m_ImporteHorasLab As String
    Dim m_NHorasExt As String
    Dim m_PrecioHoraExt As String
    Dim m_ImporteHorasExt As String
    Dim m_Importe As String
    
    On Error GoTo errores
    Set lst = Me.ListaPerfilFactura
    lst.RowSource = "Perfil;NHLab;€HLab;€Lab;NHExt;€HExt;€Ext;€Horas"
    If m_ObjFacturaActiva Is Nothing Then
        Exit Function
    End If
    If m_ObjFacturaActiva.ParaPorFacturar = EnumSino.Sí Then
        Set m_Col = m_ObjFacturaActiva.ColPerfilesPorFacturar
    Else
        Set m_Col = m_ObjFacturaActiva.ColPerfilesEnFactura
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Perfil = m_Col(m_ID)
        With m_Perfil
            m_NHorasLab = Format(.NHorasLab, "#,##0.00")
            m_PrecioHoraLab = Format(.PrecioHoraLab, "#,##0.00 €")
            m_ImporteHorasLab = Format(.ImporteHorasLab, "#,##0.00 €")
            m_NHorasExt = Format(.NHorasExt, "#,##0.00")
            m_PrecioHoraExt = Format(.PrecioHoraExt, "#,##0.00 €")
            m_ImporteHorasExt = Format(.ImporteHorasExt, "#,##0.00 €")
            m_Importe = Format(.Importe, "#,##0.00 €")
            lst.AddItem .Perfil & ";" & m_NHorasLab & ";" & m_PrecioHoraLab & ";" & m_ImporteHorasLab & ";" & m_NHorasExt & ";" & _
                        m_PrecioHoraExt & ";" & m_ImporteHorasExt & ";" & m_Importe
        End With
        Set m_Perfil = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarPerfiles ha devuelto el error: " & vbNewLine & Err.Number
    End If
End Function







