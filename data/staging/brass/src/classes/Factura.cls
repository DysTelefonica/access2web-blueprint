VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New Scripting.FileSystemObject

Public IDFactura As String
Public FechaFinalEventos As String
Public FechaFactura As String
Public Titulo1 As String
Public Titulo2 As String
Public Titulo3 As String
Public TipoImpositivo As String
Public Recargo As String
Public PrecioHoras As String
Public PrecioSuministro As String
Public PrecioSuministroConRecargo As String
Public PrecioDietasYGastos As String
Public PrecioIncurrido As String
Public PrecioSubcontrataciones As String
Public PrecioSubcontratacionesConRecargo As String
Public PrecioBaseImponible As String
Public PrecioImpuestos As String
Public ImporteAFacturar As String
Public NombreFacturaAnexa As String

Public ParaPorFacturar As EnumSino

Private m_sIDFacturaCalculada As String
Private m_sRecargoCalculado As String
Private m_sTipoImpositivoCalculado As String
Private m_sTitulo1Calculado As String
Private m_sTitulo2Calculado As String
Private m_sTitulo3Calculado As String

Private m_sPrecioHorasCalculado As String
Private m_sPrecioSuministroCalculado As String
Private m_sPrecioSuministroConRecargoCalculado As String
Private m_sPrecioSubcontratacionesCalculado As String
Private m_sPrecioSubcontratacionesConRecargoCalculado As String
Private m_sPrecioDietasYGastosCalculado As String
Private m_sPrecioIncurridoCalculado As String
Private m_sPrecioBaseImponibleCalculado As String
Private m_sPrecioImpuestosCalculado As String
Private m_sImporteAFacturarCalculado As String

Private m_sPrecioHorasPorFacturar As String
Private m_sPrecioSuministroPorFacturar As String
Private m_sPrecioSuministroConRecargoPorFacturar As String
Private m_sPrecioSubcontratacionesPorFacturar As String
Private m_sPrecioSubcontratacionesConRecargoForFacturar As String
Private m_sPrecioDietasYGastosPorFacturar As String
Private m_sPrecioIncurridoPorFacturar As String
Private m_sPrecioBaseImponiblePorFacturar As String
Private m_sPrecioImpuestosPorFacturar As String


Private m_sNombreFacturaAnexaCalculada As String
Private m_sURLFactura As String
Private m_sImporteConImpuestosPorFacturar As String
Private m_sImporteAFacturarPorFacturar As String




Private m_objcolListaEventosPorFacturar As Scripting.Dictionary
Private m_objcolListaActividadesPorFacturar As Scripting.Dictionary
Private m_objcolListaActividadesDeEventosPorFacturar As Scripting.Dictionary
Private m_objcolListaActividadesNODeEventosPorFacturar As Scripting.Dictionary
Private m_objcolListaMaterialesPorFacturar As Scripting.Dictionary
Private m_objcolListaGastosPorFacturar As Scripting.Dictionary
Private m_objcolListaSubContratacionesPorFacturar As Scripting.Dictionary
Private m_objColPerfilesPorFacturar As Scripting.Dictionary
Private m_objColPerfilesEnFactura As Scripting.Dictionary

Private m_objcolListaEventos As Scripting.Dictionary
Private m_objcolListaActividades As Scripting.Dictionary
Private m_objcolListaActividadesDeEventos As Scripting.Dictionary
Private m_objcolListaActividadesNODeEventos As Scripting.Dictionary
Private m_objcolListaMateriales As Scripting.Dictionary
Private m_objcolListaGastos As Scripting.Dictionary
Private m_objcolListaSubContrataciones As Scripting.Dictionary



Private m_ObjColPartes As Scripting.Dictionary
Private m_objColAnexosAutorizacionMaterialesEnFactura As Scripting.Dictionary


Private m_sSQLEventos As String
Private m_AppExcel As excel.Application
Private m_ObjPlantilla As excel.Workbook
Private m_ObjLibro As excel.Workbook
Private m_ObjHoja As excel.Worksheet
Private m_ObjHojaPlantilla As excel.Worksheet



Private m_FormActivo As Form
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get Titulo1Calculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sTitulo1Calculado = "MANTENIMIENTO DE LAS ESTACIONES RADIO Y SISTEMA BRASS"
    Titulo1Calculado = m_sTitulo1Calculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.Titulo1Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Titulo2Calculado() As String
    
      
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.FechaFactura = "" Then
        Exit Property
    End If
    m_sTitulo2Calculado = "FACTURACIÓN " & CDate(Format(Me.FechaFactura, "dd/mm/yyyy"))
    Titulo2Calculado = m_sTitulo2Calculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.Titulo2Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Titulo3Calculado() As String
    
    Dim intMes As Integer
    Dim m_NumeroTrimestre As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.FechaFactura = "" Then
        Exit Property
    End If
    intMes = CInt(Format(Me.FechaFactura, "mm"))
    If intMes >= 1 And intMes <= 4 Then
        m_NumeroTrimestre = "1"
    ElseIf intMes >= 5 And intMes <= 8 Then
        m_NumeroTrimestre = "2"
    ElseIf intMes >= 9 Then
        m_NumeroTrimestre = "3"
    End If
    m_sTitulo3Calculado = "TRABAJOS A FACTURAR TRIMESTRE-" & m_NumeroTrimestre & "-" & Format(Me.FechaFactura, "YYYY")
    Titulo3Calculado = m_sTitulo3Calculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.Titulo3Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TipoImpositivoCalculado() As String
    
    m_sTipoImpositivoCalculado = "21"
    TipoImpositivoCalculado = m_sTipoImpositivoCalculado

End Property

Public Property Get RecargoCalculado() As String
    
    m_sRecargoCalculado = "15"
    RecargoCalculado = m_sRecargoCalculado

End Property



Public Property Get PrecioHorasCalculado() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjPerfilFactura As FacturaPerfiles
    Dim m_ID As Variant
    
    Dim ImporteHorasLab As Double
    Dim ImporteHorasExtra As Double
    
    Dim Precio As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioHorasCalculado) Then
        PrecioHorasCalculado = m_sPrecioHorasCalculado
        Exit Property
    End If
    Set m_Col = Me.ColPerfilesEnFactura
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioHorasCalculado = "0"
        PrecioHorasCalculado = m_sPrecioHorasCalculado
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_ObjPerfilFactura = m_Col(m_ID)
        With m_ObjPerfilFactura
            If IsNumeric(.PrecioHoraLab) Then
                If IsNumeric(.NHorasLab) Then
                    ImporteHorasLab = CDbl(.NHorasLab * .PrecioHoraLab)
                End If
            End If
            If IsNumeric(.PrecioHoraExt) Then
                If IsNumeric(.NHorasExt) Then
                    ImporteHorasExtra = CDbl(.NHorasExt * .PrecioHoraExt)
                End If
            End If
            Precio = Precio + ImporteHorasLab + ImporteHorasExtra
        End With
        
        ImporteHorasLab = 0
        ImporteHorasExtra = 0
        Set m_ObjPerfilFactura = Nothing
    Next
    m_sPrecioHorasCalculado = CStr(Precio)
    
    PrecioHorasCalculado = m_sPrecioHorasCalculado
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioHorasCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioHorasPorFacturar() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjPerfilFactura As FacturaPerfiles
    Dim m_ID As Variant
    
    Dim ImporteHorasLab As Double
    Dim ImporteHorasExtra As Double
    
    Dim Precio As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioHorasPorFacturar) Then
        PrecioHorasPorFacturar = m_sPrecioHorasPorFacturar
        Exit Property
    End If
    Set m_Col = Me.ColPerfilesPorFacturar
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioHorasPorFacturar = "0"
        PrecioHorasPorFacturar = m_sPrecioHorasPorFacturar
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_ObjPerfilFactura = m_Col(m_ID)
        With m_ObjPerfilFactura
            If IsNumeric(.PrecioHoraLab) Then
                If IsNumeric(.NHorasLab) Then
                    ImporteHorasLab = CDbl(.NHorasLab * .PrecioHoraLab)
                End If
            End If
            If IsNumeric(.PrecioHoraExt) Then
                If IsNumeric(.NHorasExt) Then
                    ImporteHorasExtra = CDbl(.NHorasExt * .PrecioHoraExt)
                End If
            End If
            Precio = Precio + ImporteHorasLab + ImporteHorasExtra
        End With
        
        ImporteHorasLab = 0
        ImporteHorasExtra = 0
        Set m_ObjPerfilFactura = Nothing
    Next
    m_sPrecioHorasPorFacturar = CStr(Precio)
    
    PrecioHorasPorFacturar = m_sPrecioHorasPorFacturar
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioHorasPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioDietasYGastosCalculado() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Gasto As Gasto
    Dim m_ID As Variant
    Dim Precio As Double
    Dim ImporteGasto As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioDietasYGastosCalculado) Then
        PrecioDietasYGastosCalculado = m_sPrecioDietasYGastosCalculado
        Exit Property
    End If
    Set m_Col = Me.colListaGastos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioDietasYGastosCalculado = "0"
        PrecioDietasYGastosCalculado = m_sPrecioDietasYGastosCalculado
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_Gasto = m_Col(m_ID)
        With m_Gasto
            If IsNumeric(.ImporteUnitario) Then
                If IsNumeric(.NumeroUnidades) Then
                    ImporteGasto = CDbl(.ImporteUnitario * .NumeroUnidades)
                End If
            End If
            
            Precio = Precio + ImporteGasto
        End With
        
        ImporteGasto = 0
        
        Set m_Gasto = Nothing
    Next
    m_sPrecioDietasYGastosCalculado = CStr(Precio)
    
    PrecioDietasYGastosCalculado = m_sPrecioDietasYGastosCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioDietasYGastosCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get PrecioDietasYGastosPorFacturar() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Gasto As Gasto
    Dim m_ID As Variant
    Dim Precio As Double
    Dim ImporteGasto As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioDietasYGastosPorFacturar) Then
        PrecioDietasYGastosPorFacturar = m_sPrecioDietasYGastosPorFacturar
        Exit Property
    End If
    Set m_Col = Me.colListaGastosPorFacturar
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioDietasYGastosPorFacturar = "0"
        PrecioDietasYGastosPorFacturar = m_sPrecioDietasYGastosPorFacturar
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_Gasto = m_Col(m_ID)
        With m_Gasto
            If IsNumeric(.ImporteUnitario) Then
                If IsNumeric(.NumeroUnidades) Then
                    ImporteGasto = CDbl(.ImporteUnitario * .NumeroUnidades)
                End If
            End If
            
            Precio = Precio + ImporteGasto
        End With
        
        ImporteGasto = 0
        
        Set m_Gasto = Nothing
    Next
    m_sPrecioDietasYGastosPorFacturar = CStr(Precio)
    
    PrecioDietasYGastosPorFacturar = m_sPrecioDietasYGastosPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioDietasYGastosPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSuministroCalculado() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Material As Material
    Dim m_ID As Variant
    Dim Precio As Double
    Dim ImporteSuministro As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSuministroCalculado) Then
        PrecioSuministroCalculado = m_sPrecioSuministroCalculado
        Exit Property
    End If
    Set m_Col = Me.colListaMateriales
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioSuministroCalculado = "0"
        PrecioSuministroCalculado = m_sPrecioSuministroCalculado
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_Material = m_Col(m_ID)
        With m_Material
            If IsNumeric(.COSTE) Then
                ImporteSuministro = CDbl(.COSTE)
            End If
            
            Precio = Precio + ImporteSuministro
        End With
        
        ImporteSuministro = 0
        
        Set m_Material = Nothing
    Next
    m_sPrecioSuministroCalculado = CStr(Precio)
    
    PrecioSuministroCalculado = m_sPrecioSuministroCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSuministroCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSuministroPorFacturar() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Material As Material
    Dim m_ID As Variant
    Dim Precio As Double
    Dim ImporteSuministro As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSuministroPorFacturar) Then
        PrecioSuministroPorFacturar = m_sPrecioSuministroPorFacturar
        Exit Property
    End If
    Set m_Col = Me.colListaMaterialesPorFacturar
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioSuministroPorFacturar = "0"
        PrecioSuministroPorFacturar = m_sPrecioSuministroPorFacturar
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_Material = m_Col(m_ID)
        With m_Material
            If IsNumeric(.COSTE) Then
                ImporteSuministro = CDbl(.COSTE)
            End If
            
            Precio = Precio + ImporteSuministro
        End With
        
        ImporteSuministro = 0
        
        Set m_Material = Nothing
    Next
    m_sPrecioSuministroPorFacturar = CStr(Precio)
    
    PrecioSuministroPorFacturar = m_sPrecioSuministroPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSuministroPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSuministroConRecargoCalculado() As String
    
    Dim m_Recargo As Double
    Dim m_PrecioSuministro As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSuministroConRecargoCalculado) Then
        PrecioSuministroConRecargoCalculado = m_sPrecioSuministroConRecargoCalculado
        Exit Property
    End If
    If IsNumeric(Me.Recargo) Then
        m_Recargo = CDbl(Me.Recargo)
        If IsNumeric(Me.PrecioSuministroCalculado) Then
            m_Error = Me.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_PrecioSuministro = CDbl(Me.PrecioSuministroCalculado)
            m_sPrecioSuministroConRecargoCalculado = CStr(m_PrecioSuministro * (1 + (m_Recargo / 100)))
        End If
    End If
    PrecioSuministroConRecargoCalculado = m_sPrecioSuministroConRecargoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSuministroConRecargoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSuministroConRecargoPorFacturar() As String
    
    Dim m_Recargo As Double
    Dim m_PrecioSuministro As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSuministroConRecargoPorFacturar) Then
        PrecioSuministroConRecargoPorFacturar = m_sPrecioSuministroConRecargoPorFacturar
        Exit Property
    End If
    If IsNumeric(Me.Recargo) Then
        m_Recargo = CDbl(Me.Recargo)
        If IsNumeric(Me.PrecioSuministroPorFacturar) Then
            m_Error = Me.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_PrecioSuministro = CDbl(Me.PrecioSuministroPorFacturar)
            m_sPrecioSuministroConRecargoPorFacturar = CStr(m_PrecioSuministro * (1 + (m_Recargo / 100)))
        End If
    End If
    PrecioSuministroConRecargoPorFacturar = m_sPrecioSuministroConRecargoPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSuministroConRecargoPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSubcontratacionesCalculado() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_SubContratacion As Subcontratacion
    Dim m_ID As Variant
    Dim Precio As Double
    Dim ImporteSubContratacion As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSubcontratacionesCalculado) Then
        PrecioSubcontratacionesCalculado = m_sPrecioSubcontratacionesCalculado
        Exit Property
    End If
    Set m_Col = Me.colListaSubContrataciones
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioSubcontratacionesCalculado = "0"
        PrecioSubcontratacionesCalculado = m_sPrecioSubcontratacionesCalculado
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_SubContratacion = m_Col(m_ID)
        With m_SubContratacion
            If IsNumeric(.Importe) Then
                ImporteSubContratacion = CDbl(.Importe)
            End If
            
            Precio = Precio + ImporteSubContratacion
        End With
        
        ImporteSubContratacion = 0
        
        Set m_SubContratacion = Nothing
    Next
    m_sPrecioSubcontratacionesCalculado = CStr(Precio)
    
    PrecioSubcontratacionesCalculado = m_sPrecioSubcontratacionesCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSubcontratacionesCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSubcontratacionesPorFacturar() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_SubContratacion As Subcontratacion
    Dim m_ID As Variant
    Dim Precio As Double
    Dim m_ImporteSubContratacion As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSubcontratacionesPorFacturar) Then
        PrecioSubcontratacionesPorFacturar = m_sPrecioSubcontratacionesPorFacturar
        Exit Property
    End If
    Set m_Col = Me.colListaSubContratacionesPorFacturar
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sPrecioSubcontratacionesPorFacturar = "0"
        PrecioSubcontratacionesPorFacturar = m_sPrecioSubcontratacionesPorFacturar
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_SubContratacion = m_Col(m_ID)
        With m_SubContratacion
            If IsNumeric(.Importe) Then
                m_ImporteSubContratacion = CDbl(.Importe)
            End If
            
            Precio = Precio + m_ImporteSubContratacion
        End With
        
        m_ImporteSubContratacion = 0
        
        Set m_SubContratacion = Nothing
    Next
    m_sPrecioSubcontratacionesPorFacturar = CStr(Precio)
    
    PrecioSubcontratacionesPorFacturar = m_sPrecioSubcontratacionesPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSubcontratacionesPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSubcontratacionesConRecargoCalculado() As String
    
    Dim m_PrecioSubContratacionCalculado As Double
    Dim m_Recargo As Double
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSubcontratacionesConRecargoCalculado) Then
        PrecioSubcontratacionesConRecargoCalculado = m_sPrecioSubcontratacionesConRecargoCalculado
        Exit Property
    End If
    If IsNumeric(Me.PrecioSubcontratacionesCalculado) Then
        m_PrecioSubContratacionCalculado = CDbl(Me.PrecioSubcontratacionesCalculado)
        If Not IsNumeric(Me.Recargo) Then
            m_Recargo = CDbl(Me.RecargoCalculado)
            
        Else
             m_Recargo = CDbl(Me.Recargo)
        End If
        m_sPrecioSubcontratacionesConRecargoCalculado = CStr(m_PrecioSubContratacionCalculado * (1 + (m_Recargo / 100)))
    End If
    PrecioSubcontratacionesConRecargoCalculado = m_sPrecioSubcontratacionesConRecargoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSubcontratacionesConRecargoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioSubcontratacionesConRecargoForFacturar() As String
    
    Dim m_PrecioSubContratacionPorFacturar As Double
    Dim m_Recargo As Double
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioSubcontratacionesConRecargoForFacturar) Then
        PrecioSubcontratacionesConRecargoForFacturar = m_sPrecioSubcontratacionesConRecargoForFacturar
        Exit Property
    End If
    If IsNumeric(Me.PrecioSubcontratacionesPorFacturar) Then
        m_PrecioSubContratacionPorFacturar = CDbl(Me.PrecioSubcontratacionesPorFacturar)
        If Not IsNumeric(Me.Recargo) Then
            m_Recargo = CDbl(Me.RecargoCalculado)
            
        Else
             m_Recargo = CDbl(Me.Recargo)
        End If
        m_sPrecioSubcontratacionesConRecargoForFacturar = CStr(m_PrecioSubContratacionPorFacturar * (1 + (m_Recargo / 100)))
    End If
    PrecioSubcontratacionesConRecargoForFacturar = m_sPrecioSubcontratacionesConRecargoForFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioSubcontratacionesConRecargoForFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioIncurridoCalculado() As String
    
    Dim m_PrecioHoras As Double
    Dim m_PrecioSuministros As Double
    Dim m_PrecioGastos As Double
    Dim m_PrecioSubcontrataciones As Double
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioIncurridoCalculado) Then
        PrecioIncurridoCalculado = m_sPrecioIncurridoCalculado
        Exit Property
    End If
    If IsNumeric(Me.PrecioHorasCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioHoras = CDbl(Me.PrecioHorasCalculado)
    End If
    If IsNumeric(Me.PrecioSuministroConRecargoCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioSuministros = CDbl(Me.PrecioSuministroConRecargoCalculado)
    End If
    If IsNumeric(Me.PrecioDietasYGastosCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioGastos = CDbl(Me.PrecioDietasYGastosCalculado)
    End If
    
    If IsNumeric(Me.PrecioSubcontratacionesConRecargoCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioSubcontrataciones = CDbl(Me.PrecioSubcontratacionesConRecargoCalculado)
    End If
    
    m_sPrecioIncurridoCalculado = CStr(m_PrecioHoras + m_PrecioSuministros + m_PrecioGastos + m_PrecioSubcontrataciones)
    PrecioIncurridoCalculado = m_sPrecioIncurridoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioIncurridoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioIncurridoPorFacturar() As String
    
    Dim m_PrecioHoras As Double
    Dim m_PrecioSuministros As Double
    Dim m_PrecioGastos As Double
    Dim m_PrecioSubcontrataciones As Double
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sPrecioIncurridoPorFacturar) Then
        PrecioIncurridoPorFacturar = m_sPrecioIncurridoPorFacturar
        Exit Property
    End If
    If IsNumeric(Me.PrecioHorasPorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioHoras = CDbl(Me.PrecioHorasPorFacturar)
    End If
    If IsNumeric(Me.PrecioSuministroConRecargoPorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioSuministros = CDbl(Me.PrecioSuministroConRecargoPorFacturar)
    End If
    If IsNumeric(Me.PrecioDietasYGastosPorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioGastos = CDbl(Me.PrecioDietasYGastosPorFacturar)
    End If
    
    If IsNumeric(Me.PrecioSubcontratacionesConRecargoForFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioSubcontrataciones = CDbl(Me.PrecioSubcontratacionesConRecargoForFacturar)
    End If
    
    m_sPrecioIncurridoPorFacturar = CStr(m_PrecioHoras + m_PrecioSuministros + m_PrecioGastos + m_PrecioSubcontrataciones)
    PrecioIncurridoPorFacturar = m_sPrecioIncurridoPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioIncurridoPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ImporteAFacturarCalculado() As String
    
    Dim m_BaseImponible As Double
    Dim m_PrecioImpuestos As Double
    Dim m_PrecioGastos As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sImporteAFacturarCalculado) Then
        ImporteAFacturarCalculado = m_sImporteAFacturarCalculado
        Exit Property
    End If
    If IsNumeric(Me.PrecioBaseImponibleCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_BaseImponible = CDbl(Me.PrecioBaseImponibleCalculado)
    End If
    If IsNumeric(Me.PrecioImpuestosCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioImpuestos = CDbl(Me.PrecioImpuestosCalculado)
    End If
    If IsNumeric(Me.PrecioDietasYGastosCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioGastos = CDbl(Me.PrecioDietasYGastosCalculado)
    End If
    m_sImporteAFacturarCalculado = CStr(m_BaseImponible + m_PrecioImpuestos + m_PrecioGastos)
    ImporteAFacturarCalculado = m_sImporteAFacturarCalculado
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.ImporteAFacturarCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ImporteAFacturarPorFacturar() As String
    
    Dim m_BaseImponible As Double
    Dim m_PrecioImpuestos As Double
    Dim m_PrecioGastos As Double
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sImporteAFacturarPorFacturar) Then
        ImporteAFacturarPorFacturar = m_sImporteAFacturarPorFacturar
        Exit Property
    End If
    If IsNumeric(Me.PrecioBaseImponiblePorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_BaseImponible = CDbl(Me.PrecioBaseImponiblePorFacturar)
    End If
    If IsNumeric(Me.PrecioImpuestosCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioImpuestos = CDbl(Me.PrecioImpuestosPorFacturar)
    End If
    If IsNumeric(Me.PrecioDietasYGastosPorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_PrecioGastos = CDbl(Me.PrecioDietasYGastosPorFacturar)
    End If
    m_sImporteAFacturarPorFacturar = CStr(m_BaseImponible + m_PrecioImpuestos + m_PrecioGastos)
    ImporteAFacturarPorFacturar = m_sImporteAFacturarPorFacturar
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.ImporteAFacturarPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NombreFacturaAnexaCalculada() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreFacturaAnexaCalculada <> "" Then
        NombreFacturaAnexaCalculada = m_sNombreFacturaAnexaCalculada
        Exit Property
    End If
    
    If Me.FechaFactura = "" Then
        Exit Property
    End If
    If Me.IDFactura = "" Then
        m_sNombreFacturaAnexaCalculada = "000" & _
                            "FACTURA_" & Format(Me.FechaFactura, "YY") & Format(Me.FechaFactura, "MM") & _
                            Format(Me.FechaFactura, "DD") & ".xlsx"
    Else
        m_sNombreFacturaAnexaCalculada = Format(Me.IDFactura, "000") & _
                            "FACTURA_" & Format(Me.FechaFactura, "YY") & Format(Me.FechaFactura, "MM") & _
                            Format(Me.FechaFactura, "DD") & ".xlsx"
    End If
    
    NombreFacturaAnexaCalculada = m_sNombreFacturaAnexaCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.NombreFacturaAnexaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDFacturaCalculada() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sIDFacturaCalculada <> "" Then
        IDFacturaCalculada = m_sIDFacturaCalculada
        Exit Property
    End If
    m_sIDFacturaCalculada = DameID("TbFacturaPrincipal", "IDFactura")
    
    IDFacturaCalculada = m_sIDFacturaCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.IDFacturaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLFactura() As String
    
    
    Dim m_URLDirectorio As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_sURLFactura <> "" Then
        URLFactura = m_sURLFactura
        Exit Property
    End If
    If Me.NombreFacturaAnexaCalculada = "" Then
        Exit Property
    End If
    m_URLDirectorio = m_ObjEntorno.URLDirectorioFacturas
    
    If m_URLDirectorio = "" Then
        Exit Property
    End If
    
    m_sURLFactura = m_URLDirectorio & Me.NombreFacturaAnexaCalculada
    
    URLFactura = m_sURLFactura

    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.URLFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get colListaEventos() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaEventos Is Nothing Then
        Set colListaEventos = m_objcolListaEventos
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objcolListaEventos = Constructor.getEventosEnFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaEventos = m_objcolListaEventos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaEventos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colListaEventos(ByVal objNewValue As Variant)

    Set m_objcolListaEventos = objNewValue

End Property

Public Property Get colListaActividades() As Scripting.Dictionary
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjActividad As Actividad
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaActividades Is Nothing Then
        Set colListaActividades = m_objcolListaActividades
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_Col = Me.colListaActividadesDeEventos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            If m_objcolListaActividades Is Nothing Then
                Set m_objcolListaActividades = New Scripting.Dictionary
                m_objcolListaActividades.CompareMode = TextCompare
            End If
            If Not m_objcolListaActividades.Exists(m_ObjActividad.IDActividad) Then
                m_objcolListaActividades.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
        Next
    End If
    Set m_Col = Me.colListaActividadesNODeEventos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            If m_objcolListaActividades Is Nothing Then
                Set m_objcolListaActividades = New Scripting.Dictionary
                m_objcolListaActividades.CompareMode = TextCompare
            End If
            If Not m_objcolListaActividades.Exists(m_ObjActividad.IDActividad) Then
                m_objcolListaActividades.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
        Next
    End If
    Set colListaActividades = m_objcolListaActividades
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaActividades ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colListaActividades(ByVal objNewValue As Variant)

    Set m_objcolListaActividades = objNewValue

End Property

Public Property Get colListaActividadesPorFacturar() As Scripting.Dictionary
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjActividad As Actividad
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaActividadesPorFacturar Is Nothing Then
        Set colListaActividadesPorFacturar = m_objcolListaActividadesPorFacturar
        Exit Property
    End If
    Set m_Col = Me.colListaActividadesDeEventosPorFacturar
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            If m_objcolListaActividadesPorFacturar Is Nothing Then
                Set m_objcolListaActividadesPorFacturar = New Scripting.Dictionary
                m_objcolListaActividadesPorFacturar.CompareMode = TextCompare
            End If
            If Not m_objcolListaActividadesPorFacturar.Exists(m_ObjActividad.IDActividad) Then
                m_objcolListaActividadesPorFacturar.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
        Next
    End If
    Set m_Col = Me.colListaActividadesNODeEventosPorFacturar
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            If m_objcolListaActividadesPorFacturar Is Nothing Then
                Set m_objcolListaActividadesPorFacturar = New Scripting.Dictionary
                m_objcolListaActividadesPorFacturar.CompareMode = TextCompare
            End If
            If Not m_objcolListaActividadesPorFacturar.Exists(m_ObjActividad.IDActividad) Then
                m_objcolListaActividadesPorFacturar.Add m_ObjActividad.IDActividad, m_ObjActividad
            End If
            Set m_ObjActividad = Nothing
        Next
    End If
    Set colListaActividadesPorFacturar = m_objcolListaActividadesPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaActividadesPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colListaActividadesPorFacturar(ByVal objNewValue As Variant)

    Set m_objcolListaActividadesPorFacturar = objNewValue

End Property

Public Property Get ColAnexosAutorizacionMaterialesEnFactura() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColAnexosAutorizacionMaterialesEnFactura Is Nothing Then
        Set ColAnexosAutorizacionMaterialesEnFactura = m_objColAnexosAutorizacionMaterialesEnFactura
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objColAnexosAutorizacionMaterialesEnFactura = Constructor.getAnexosAutorizacionMaterialesEnFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexosAutorizacionMaterialesEnFactura = m_objColAnexosAutorizacionMaterialesEnFactura
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.ColAnexosAutorizacionMaterialesEnFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColAnexosAutorizacionMaterialesEnFactura(ByVal objNewValue As Variant)

    Set m_objColAnexosAutorizacionMaterialesEnFactura = objNewValue

End Property


Public Property Get colListaMateriales() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaMateriales Is Nothing Then
        Set colListaMateriales = m_objcolListaMateriales
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objcolListaMateriales = Constructor.getMaterialesEnFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaMateriales = m_objcolListaMateriales
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaMateriales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colListaMateriales(ByVal objNewValue As Variant)

    Set m_objcolListaMateriales = objNewValue

End Property

Public Property Get colListaGastos() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaGastos Is Nothing Then
        Set colListaGastos = m_objcolListaGastos
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objcolListaGastos = Constructor.getGastosEnFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaGastos = m_objcolListaGastos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaGastos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colListaGastos(ByVal objNewValue As Variant)

    Set m_objcolListaGastos = objNewValue

End Property

Public Property Get ColPartes() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColPartes Is Nothing Then
        Set ColPartes = m_ObjColPartes
        Exit Property
    End If
    If Me.IDFactura = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se conoce el IDFactura"
        Err.Raise 1000
    End If
    Set m_ObjColPartes = Constructor.getPartesDeFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPartes = m_ObjColPartes
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.ColPartes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColPartes(ByVal objNewValue As Variant)

    Set m_ObjColPartes = objNewValue

End Property


Public Property Get colListaSubContrataciones() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaSubContrataciones Is Nothing Then
        Set colListaSubContrataciones = m_objcolListaSubContrataciones
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objcolListaSubContrataciones = Constructor.getSubContratacionesEnFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaSubContrataciones = m_objcolListaSubContrataciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaSubContrataciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Set colListaSubContrataciones(ByVal objNewValue As Variant)

    Set m_objcolListaSubContrataciones = objNewValue

End Property

Public Property Get SQLEventos() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sSQLEventos <> "" Then
        SQLEventos = m_sSQLEventos
        Exit Property
    End If
    If Not IsDate(Me.FechaFinalEventos) Then
        Exit Property
    End If
    m_sSQLEventos = Constructor.getSQLEventosDeFactura(Me.FechaFinalEventos, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    SQLEventos = m_sSQLEventos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.SQLEventos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colListaEventosPorFacturar() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaEventosPorFacturar Is Nothing Then
        Set colListaEventosPorFacturar = m_objcolListaEventosPorFacturar
        Exit Property
    End If
    If Me.SQLEventos = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objcolListaEventosPorFacturar = Constructor.getEventosParaFactura(Me.SQLEventos, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaEventosPorFacturar = m_objcolListaEventosPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaEventosPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colListaActividadesDeEventos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaActividadesDeEventos Is Nothing Then
        Set colListaActividadesDeEventos = m_objcolListaActividadesDeEventos
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objcolListaActividadesDeEventos = Constructor.getActividadesDeEventos(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaActividadesDeEventos = m_objcolListaActividadesDeEventos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaActividadesDeEventos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get colListaActividadesDeEventosPorFacturar() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaActividadesDeEventosPorFacturar Is Nothing Then
        Set colListaActividadesDeEventosPorFacturar = m_objcolListaActividadesDeEventosPorFacturar
        Exit Property
    End If
    If Me.SQLEventos = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objcolListaActividadesDeEventosPorFacturar = Constructor.getActividadesDeEventosParaFactura(Me.SQLEventos, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaActividadesDeEventosPorFacturar = m_objcolListaActividadesDeEventosPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaActividadesDeEventosPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colListaActividadesNODeEventos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaActividadesNODeEventos Is Nothing Then
        Set colListaActividadesNODeEventos = m_objcolListaActividadesNODeEventos
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objcolListaActividadesNODeEventos = Constructor.getActividadesNODeEventos(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaActividadesNODeEventos = m_objcolListaActividadesNODeEventos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaActividadesNODeEventos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get colListaActividadesNODeEventosPorFacturar() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaActividadesNODeEventosPorFacturar Is Nothing Then
        Set colListaActividadesNODeEventosPorFacturar = m_objcolListaActividadesNODeEventosPorFacturar
        Exit Property
    End If
    If Me.FechaFinalEventos = "" Then
        
        Exit Property
    End If
    Set m_objcolListaActividadesNODeEventosPorFacturar = Constructor.getActividadesNoDeEventosParaFactura(Me.FechaFinalEventos, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaActividadesNODeEventosPorFacturar = m_objcolListaActividadesNODeEventosPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaActividadesNODeEventosPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colListaMaterialesPorFacturar() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaMaterialesPorFacturar Is Nothing Then
        Set colListaMaterialesPorFacturar = m_objcolListaMaterialesPorFacturar
        Exit Property
    End If
    If Me.FechaFinalEventos = "" Then
        Exit Property
    End If
    Set m_objcolListaMaterialesPorFacturar = Constructor.getMaterialesParaFactura(Me.SQLEventos, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaMaterialesPorFacturar = m_objcolListaMaterialesPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaMaterialesPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colListaGastosPorFacturar() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaGastosPorFacturar Is Nothing Then
        Set colListaGastosPorFacturar = m_objcolListaGastosPorFacturar
        Exit Property
    End If
    If Me.FechaFinalEventos = "" Then
        Exit Property
    End If
    Set m_objcolListaGastosPorFacturar = Constructor.getGastosParaFactura(Me.FechaFinalEventos, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaGastosPorFacturar = m_objcolListaGastosPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaGastosPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colListaSubContratacionesPorFacturar() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolListaSubContratacionesPorFacturar Is Nothing Then
        Set colListaSubContratacionesPorFacturar = m_objcolListaSubContratacionesPorFacturar
        Exit Property
    End If
    If Me.FechaFinalEventos = "" Then
        Exit Property
    End If
    Set m_objcolListaSubContratacionesPorFacturar = Constructor.getSubContratacionesParaFactura(Me.FechaFinalEventos, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colListaSubContratacionesPorFacturar = m_objcolListaSubContratacionesPorFacturar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.colListaSubContratacionesPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColPerfilesEnFactura() As Scripting.Dictionary
    
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_objColPerfilesEnFactura Is Nothing Then
        Set ColPerfilesEnFactura = m_objColPerfilesEnFactura
        Exit Property
    End If
    If Me.IDFactura = "" Then
        Exit Property
    End If
    Set m_objColPerfilesEnFactura = Constructor.getListaPerfilesEnFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPerfilesEnFactura = m_objColPerfilesEnFactura
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.ColPerfilesEnFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Private Function getColTipos(Optional p_EsCalculado As EnumSino = EnumSino.Sí, Optional p_Error As String) As Scripting.Dictionary
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjActividad As Actividad
    Dim m_ObjTecnico As Tecnico
    Dim m_ObjTipoTecnico As TIPOTECNICO
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaActividadesDeEventos
    Else
        Set m_Col = Me.colListaActividadesDeEventosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            Set m_ObjTecnico = m_ObjActividad.Tecnico
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTecnico Is Nothing Then
                p_Error = "De la Activida " & m_ObjActividad.IDActividad & " no se ha podido obtener el técnico"
                Err.Raise 1000
            End If
            Set m_ObjTipoTecnico = Constructor.getTipoTecnico(m_ObjTecnico.Tipo, , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTipoTecnico Is Nothing Then
                p_Error = "De la Activida " & m_ObjActividad.IDActividad & " no se ha podido obtener el tipo técnico"
                Err.Raise 1000
            End If
            If getColTipos Is Nothing Then
                Set getColTipos = New Scripting.Dictionary
                getColTipos.CompareMode = TextCompare
            End If
            If Not getColTipos.Exists(m_ObjTipoTecnico.Tipo) Then
                getColTipos.Add m_ObjTipoTecnico.Tipo, m_ObjTipoTecnico
            End If
            Set m_ObjTipoTecnico = Nothing
            Set m_ObjTecnico = Nothing
            Set m_ObjActividad = Nothing
        Next
    End If
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaActividadesNODeEventos
    Else
        Set m_Col = Me.colListaActividadesNODeEventosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            Set m_ObjTecnico = m_ObjActividad.Tecnico
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTecnico Is Nothing Then
                p_Error = "De la Activida " & m_ObjActividad.IDActividad & " no se ha podido obtener el técnico"
                Err.Raise 1000
            End If
            Set m_ObjTipoTecnico = Constructor.getTipoTecnico(m_ObjTecnico.Tipo, , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTipoTecnico Is Nothing Then
                p_Error = "De la Activida " & m_ObjActividad.IDActividad & " no se ha podido obtener el tipo técnico"
                Err.Raise 1000
            End If
            If getColTipos Is Nothing Then
                Set getColTipos = New Scripting.Dictionary
                getColTipos.CompareMode = TextCompare
            End If
            If Not getColTipos.Exists(m_ObjTipoTecnico.Tipo) Then
                getColTipos.Add m_ObjTipoTecnico.Tipo, m_ObjTipoTecnico
            End If
            Set m_ObjTipoTecnico = Nothing
            Set m_ObjTecnico = Nothing
            Set m_ObjActividad = Nothing
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.getColTipos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Property Get ColPerfilesPorFacturar() As Scripting.Dictionary
    
    Dim m_ObjTipoTecnico As TIPOTECNICO
    Dim m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio
    Dim m_ID As Variant
    Dim m_ColActividades As Scripting.Dictionary
    Dim m_ObjActividad As Actividad
    Dim m_ObjFacturaPerfil As FacturaPerfiles
    Dim m_Horas As String
    Dim m_HorasLab As Double
    Dim m_HOrasExtra As Double
    Dim m_ColTipos As Scripting.Dictionary
    Dim m_Tipo As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColPerfilesPorFacturar Is Nothing Then
        Set ColPerfilesPorFacturar = m_objColPerfilesPorFacturar
        Exit Property
    End If
    
    Set m_ColTipos = getColTipos(EnumSino.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColTipos Is Nothing Then
        For Each m_Tipo In m_ColTipos
            'Debug.Print m_Tipo
            Set m_ObjTipoTecnico = m_ColTipos(m_Tipo)
            m_Horas = getHorasPerfil(CStr(m_Tipo), EnumSino.No, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If InStr(1, m_Horas, "|") <> 0 Then
                dato = Split(m_Horas, "|")
                m_HorasLab = CDbl(dato(0))
                m_HOrasExtra = CDbl(dato(1))
            End If
            Set m_ObjFacturaPerfil = New FacturaPerfiles
            With m_ObjFacturaPerfil
                .Perfil = CStr(m_Tipo)
                .PrecioHoraExt = m_ObjTipoTecnico.PrecioHoraExtraActual
                m_Error = m_ObjTipoTecnico.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not IsNumeric(.PrecioHoraExt) Then
                    m_Error = "No se ha podido obtener el Precio hora extra del perfil " & .Perfil
                    Err.Raise 1000
                End If
                .PrecioHoraLab = m_ObjTipoTecnico.PrecioHoraLabActual
                m_Error = m_ObjTipoTecnico.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not IsNumeric(.PrecioHoraLab) Then
                    m_Error = "No se ha podido obtener el Precio hora Lab del perfil " & .Perfil
                    Err.Raise 1000
                End If
                .NHorasLab = m_HorasLab
                .NHorasExt = m_HOrasExtra
                
            End With
            If m_objColPerfilesPorFacturar Is Nothing Then
                Set m_objColPerfilesPorFacturar = New Scripting.Dictionary
                m_objColPerfilesPorFacturar.CompareMode = TextCompare
            End If
            If Not m_objColPerfilesPorFacturar.Exists(m_ObjFacturaPerfil.Perfil) Then
                m_objColPerfilesPorFacturar.Add m_ObjFacturaPerfil.Perfil, m_ObjFacturaPerfil
            End If
            m_HorasLab = 0
            m_HOrasExtra = 0
            m_Horas = ""
            
            Set m_ObjTipoTecnico = Nothing
            Set m_ObjFacturaPerfil = Nothing
        Next
        
    End If
    Set ColPerfilesPorFacturar = m_objColPerfilesPorFacturar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.ColPerfilesPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property







Public Property Get PrecioBaseImponibleCalculado() As String
    
    Dim dblPrecioIncurridoCalculado As Double
    Dim dblPrecioDietasYGastosCalculado As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not IsNumeric(Me.PrecioIncurridoCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    dblPrecioIncurridoCalculado = CDbl(Me.PrecioIncurridoCalculado)
    If IsNumeric(Me.PrecioDietasYGastosCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        dblPrecioDietasYGastosCalculado = CDbl(Me.PrecioDietasYGastosCalculado)
    End If
    m_sPrecioBaseImponibleCalculado = CStr(dblPrecioIncurridoCalculado - dblPrecioDietasYGastosCalculado)
    
    PrecioBaseImponibleCalculado = m_sPrecioBaseImponibleCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioBaseImponibleCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioBaseImponiblePorFacturar() As String
    
    Dim dblPrecioIncurridoPorFacturar As Double
    Dim dblPrecioDietasYGastosPorFacturar As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not IsNumeric(Me.PrecioIncurridoPorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    dblPrecioIncurridoPorFacturar = CDbl(Me.PrecioIncurridoPorFacturar)
    If IsNumeric(Me.PrecioDietasYGastosPorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        dblPrecioDietasYGastosPorFacturar = CDbl(Me.PrecioDietasYGastosPorFacturar)
    End If
    m_sPrecioBaseImponiblePorFacturar = CStr(dblPrecioIncurridoPorFacturar - dblPrecioDietasYGastosPorFacturar)
    
    PrecioBaseImponiblePorFacturar = m_sPrecioBaseImponiblePorFacturar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioBaseImponiblePorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioImpuestosCalculado() As String
    
    Dim dblTipoImpositivo As Double
    Dim dblPrecioBaseImponibleCalculado As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not IsNumeric(Me.PrecioBaseImponibleCalculado) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    dblPrecioBaseImponibleCalculado = CDbl(Me.PrecioBaseImponibleCalculado)
    If IsNumeric(Me.TipoImpositivo) Then
        dblTipoImpositivo = CDbl(Me.TipoImpositivo) / 100
    Else
        dblTipoImpositivo = 0
    End If
    m_sPrecioImpuestosCalculado = CStr(dblPrecioBaseImponibleCalculado * dblTipoImpositivo)
    
    PrecioImpuestosCalculado = m_sPrecioImpuestosCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioImpuestosCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PrecioImpuestosPorFacturar() As String
    
    Dim dblTipoImpositivo As Double
    Dim dblPrecioBaseImponiblePorFacturar As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not IsNumeric(Me.PrecioBaseImponiblePorFacturar) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    dblPrecioBaseImponiblePorFacturar = CDbl(Me.PrecioBaseImponiblePorFacturar)
    If IsNumeric(Me.TipoImpositivo) Then
        dblTipoImpositivo = CDbl(Me.TipoImpositivo) / 100
    Else
        dblTipoImpositivo = 0
    End If
    m_sPrecioImpuestosPorFacturar = CStr(dblPrecioBaseImponiblePorFacturar * dblTipoImpositivo)
    
    PrecioImpuestosPorFacturar = m_sPrecioImpuestosPorFacturar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.PrecioImpuestosPorFacturar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Function ValidacionFactura( _
                                    Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjEvento As Evento
    Dim m_ID As Variant
    Dim m_ObjActividad As Actividad
    Dim m_ObjMaterial As Material
    Dim m_ObjGasto As Gasto
    Dim m_ObjSubContratacion As Subcontratacion
    Dim m_ObjEquipo As Equipo
    Dim m_ColPerfilesFactura As Scripting.Dictionary
    Dim m_ObjTecnico As Tecnico
    Dim m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio
    Dim m_IDFacturaCalculada As String
    Dim m_RecargoCalculado As String
    Dim m_TipoImpositivoCalculado As String
    Dim m_PrecioHoras As String
    Dim m_PrecioSuministro As String
    Dim m_PrecioSuministroConRecargo As String
    Dim m_PrecioDietasYGastos As String
    Dim m_PrecioIncurrido As String
    Dim m_PrecioSubcontrataciones As String
    Dim m_PrecioSubcontratacionesConRecargo As String
    Dim m_PrecioBaseImponible As String
    Dim m_PrecioImpuestos As String
    Dim m_ImporteAFacturar As String
    Dim m_NombreFacturaAnexaCalculada As String
    
    
    On Error GoTo errores
    If Not IsDate(Me.FechaFinalEventos) Then
        p_Error = "No se ha indicado la fecha final de los eventos"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaFactura) Then
        Me.FechaFactura = Me.FechaFinalEventos
    End If
    Set m_FormActivo = Screen.ActiveForm
    If m_FormActivo.Name = "" Then
    
    ElseIf m_FormActivo.Name = "" Then
        
    End If
    
    
    If p_EsCalculado = EnumSino.Sí Then
        m_Linea = "Validando ..... IDFactura"
        Avance m_Linea
        If Me.IDFactura = "" Then
            m_IDFacturaCalculada = Me.IDFacturaCalculada
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not IsNumeric(m_IDFacturaCalculada) Then
                p_Error = "No se ha podido obtener el IDFactura nuevo"
                Err.Raise 1000
            End If
            Me.IDFactura = m_IDFacturaCalculada
            
        End If
    End If
    If Me.Recargo = "" Then
        m_Linea = "Validando ..... Recargo"
        Avance m_Linea
        m_RecargoCalculado = Me.RecargoCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not IsNumeric(m_RecargoCalculado) Then
            p_Error = "No se ha podido obtener el Recargo"
            Err.Raise 1000
        End If
        Me.Recargo = m_RecargoCalculado
    End If
    
    If Not IsNumeric(Me.TipoImpositivo) Then
        m_Linea = "Validando ..... TipoImpositivo"
        Avance m_Linea
        m_TipoImpositivoCalculado = Me.TipoImpositivoCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not IsNumeric(m_TipoImpositivoCalculado) Then
            p_Error = "No se ha podido obtener el TipoImpositivo"
            Err.Raise 1000
        End If
        Me.TipoImpositivo = m_TipoImpositivoCalculado
    End If
    'validación de eventos
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaEventos
    Else
        Set m_Col = Me.colListaEventosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        m_Linea = "Validando ..... Eventos sin Equipos"
        Avance m_Linea
        For Each m_ID In m_Col
            Set m_ObjEvento = m_Col(m_ID)
            Set m_ObjEquipo = m_ObjEvento.Equipo
            p_Error = m_ObjEvento.Error
            If p_Error <> "" Then
                p_Error = "El Evento " & m_ID & " no ha podido resolver el Equipo"
                Err.Raise 1000
            End If
            If m_ObjEquipo Is Nothing Then
                p_Error = "El Evento " & m_ID & " no ha podido resolver el Equipo"
                Err.Raise 1000
            End If
            Set m_ObjEvento = Nothing
        Next
    End If
    
    'validación de Atividades de Eventos
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaActividadesDeEventos
    Else
        Set m_Col = Me.colListaActividadesDeEventosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        m_Linea = "Validando ..... Actividades de Eventos"
        Avance m_Linea
        For Each m_ID In m_Col
            'If CStr(m_ID) = "jsf_180904155801" Then Stop
            'Debug.Print CStr(m_ID)
            Set m_ObjActividad = m_Col(m_ID)
            Set m_ObjEvento = m_ObjActividad.Evento
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                p_Error = "La actividad " & m_ObjActividad.IDActividad & " no ha podido obtener el evento"
                Err.Raise 1000
            End If
            Set m_ObjEquipo = m_ObjEvento.Equipo
            p_Error = m_ObjEvento.Error
            If p_Error <> "" Then
                p_Error = "La actividad " & m_ObjActividad.IDActividad & " que proviene del el evento " & m_ObjEvento.IDEVENTO & _
                         " no ha podido resolver el Equipo"
                Err.Raise 1000
            End If
            If m_ObjEquipo Is Nothing Then
                p_Error = "La actividad " & m_ObjActividad.IDActividad & " no ha podido obtener el Equipo desde el IDEvento " & m_ObjEvento.IDEVENTO
                Err.Raise 1000
            End If
           
            Set m_ObjTecnico = m_ObjActividad.Tecnico
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                p_Error = "La actividad " & m_ObjActividad.IDActividad & " no ha podido obtener el Técnico"
                Err.Raise 1000
            End If
            If m_ObjTecnico Is Nothing Then
                p_Error = "La actividad " & m_ObjActividad.IDActividad & " no ha podido obtener el Técnico"
                Err.Raise 1000
            End If
            Set m_ObjTipoTecnicoPrecio = m_ObjTecnico.TipoTecnicoPrecio
            p_Error = m_ObjTecnico.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTipoTecnicoPrecio Is Nothing Then
                p_Error = "La actividad " & m_ObjActividad.IDActividad & " Tiene el técnico " & _
                            m_ObjTecnico.Alias & " de tipo " & m_ObjTecnico.Tipo & _
                            " pero no ha podido resolver su precio activo"
                Err.Raise 1000
            End If
            Set m_ObjEquipo = Nothing
            Set m_ObjTipoTecnicoPrecio = Nothing
            Set m_ObjTecnico = Nothing
            Set m_ObjActividad = Nothing
             
        Next
    End If
    
    'validación de Atividades de NO Eventos
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaActividadesNODeEventos
    Else
        Set m_Col = Me.colListaActividadesNODeEventosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        m_Linea = "Validando ..... Actividades NO de Eventos"
        Avance m_Linea
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            Set m_ObjTecnico = m_ObjActividad.Tecnico
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                p_Error = "La actividad " & m_ObjActividad.IDActividad & " no ha podido obtener el Técnico"
                Err.Raise 1000
            End If
            If m_ObjTecnico Is Nothing Then
                p_Error = "La actividad No de Evento " & m_ObjActividad.IDActividad & " no ha podido obtener el Técnico"
                Err.Raise 1000
            End If
            Set m_ObjTipoTecnicoPrecio = m_ObjTecnico.TipoTecnicoPrecio
            p_Error = m_ObjTecnico.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTipoTecnicoPrecio Is Nothing Then
                p_Error = "La actividad No de Evento " & m_ObjActividad.IDActividad & " Tiene el técnico " & _
                            m_ObjTecnico.Alias & " de tipo " & m_ObjTecnico.Tipo & _
                            " pero no ha podido resolver su precio activo"
                Err.Raise 1000
            End If
            Set m_ObjTipoTecnicoPrecio = Nothing
            Set m_ObjTecnico = Nothing
            
            Set m_ObjActividad = Nothing
        Next
    End If
    'validación de Perfiles en factura
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.ColPerfilesEnFactura
    Else
        Set m_Col = Me.ColPerfilesPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        p_Error = "No se ha podido obtener el listado de perfiles distintos con su precio"
        Err.Raise 1000
    End If
    'validación de Materiales
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaMateriales
    Else
        Set m_Col = Me.colListaMaterialesPorFacturar
    End If
    If Not m_Col Is Nothing Then
        m_Linea = "Validando ..... Materiales sin Actividad"
        Avance m_Linea
        For Each m_ID In m_Col
            Set m_ObjMaterial = m_Col(m_ID)
            Set m_ObjActividad = m_ObjMaterial.Actividad
            p_Error = m_ObjMaterial.Error
            If p_Error <> "" Then
                p_Error = "EL material " & m_ObjMaterial.IDMaterial & " no ha podido resolver la actividad"
                Err.Raise 1000
            End If
            If m_ObjActividad Is Nothing Then
                p_Error = "EL material " & m_ObjMaterial.IDMaterial & " no ha podido resolver la actividad"
                Err.Raise 1000
            End If
            Set m_ObjEvento = m_ObjActividad.Evento
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                p_Error = "EL material " & m_ObjMaterial.IDMaterial & " no ha podido resolver El Evento"
                Err.Raise 1000
            End If
            If m_ObjEvento Is Nothing Then
                p_Error = "EL material " & m_ObjMaterial.IDMaterial & " no ha podido resolver El Evento"
                Err.Raise 1000
            End If
            Set m_ObjEquipo = m_ObjEvento.Equipo
            p_Error = m_ObjEvento.Error
            If p_Error <> "" Then
                p_Error = "EL material " & m_ObjMaterial.IDMaterial & " no ha podido resolver El Equipo"
                Err.Raise 1000
            End If
            If m_ObjEquipo Is Nothing Then
                p_Error = "EL material " & m_ObjMaterial.IDMaterial & " no ha podido resolver El Equipo"
                Err.Raise 1000
            End If
            Set m_ObjEquipo = Nothing
            Set m_ObjEvento = Nothing
            Set m_ObjActividad = Nothing
            
            Set m_ObjMaterial = Nothing
        Next
    End If
    'COMPROBACIONES DE RESULTADOS TOTALES
    
    If p_EsCalculado = EnumSino.Sí Then
        m_Linea = "Validando ..... Resultado de la factura"
        Avance m_Linea
        m_PrecioHoras = Me.PrecioHorasCalculado
    Else
        m_Linea = "Validando ..... Resultado del simulacro de la factura"
        Avance m_Linea
        m_PrecioHoras = Me.PrecioHorasPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioSuministro = Me.PrecioSuministroCalculado
    Else
        m_PrecioSuministro = Me.PrecioSuministroPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioSuministroConRecargo = Me.PrecioSuministroConRecargoCalculado
    Else
        m_PrecioSuministroConRecargo = Me.PrecioSuministroConRecargoPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioDietasYGastos = Me.PrecioDietasYGastosCalculado
    Else
        m_PrecioDietasYGastos = Me.PrecioDietasYGastosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioIncurrido = Me.PrecioIncurridoCalculado
    Else
        m_PrecioIncurrido = Me.PrecioIncurridoPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioSubcontrataciones = Me.PrecioSubcontratacionesCalculado
    Else
        m_PrecioSubcontrataciones = Me.PrecioSubcontratacionesPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioSubcontratacionesConRecargo = Me.PrecioSubcontratacionesConRecargoCalculado
    Else
        m_PrecioSubcontratacionesConRecargo = Me.PrecioSubcontratacionesConRecargoForFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioBaseImponible = Me.PrecioBaseImponibleCalculado
    Else
        m_PrecioBaseImponible = Me.PrecioBaseImponiblePorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioImpuestos = Me.PrecioImpuestosCalculado
    Else
        m_PrecioImpuestos = Me.PrecioImpuestosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_EsCalculado = EnumSino.Sí Then
        m_ImporteAFacturar = Me.ImporteAFacturarCalculado
    Else
        m_ImporteAFacturar = Me.ImporteAFacturarPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_NombreFacturaAnexaCalculada = Me.NombreFacturaAnexaCalculada
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ValidacionFactura = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.ValidacionFactura ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function EliminarFacturasAnteriores( _
                                            p_IDFactura As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
     
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_IDFactura = "" Then
        Exit Function
    End If
    m_SQL = "DELETE TbFacturaEventosInvolucrados.IDFactura " & _
                "FROM TbFacturaEventosInvolucrados " & _
                "WHERE (((TbFacturaEventosInvolucrados.IDFactura)=" & p_IDFactura & "));"
    getdb().Execute m_SQL
    
    m_SQL = "DELETE TbFacturaActividadesInvolucradas.IDFactura " & _
                "FROM TbFacturaActividadesInvolucradas " & _
                "WHERE (((TbFacturaActividadesInvolucradas.IDFactura)=" & p_IDFactura & "));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbFacturaMaterialesInvolucrados.IDFactura " & _
                "FROM TbFacturaMaterialesInvolucrados " & _
                "WHERE (((TbFacturaMaterialesInvolucrados.IDFactura)=" & p_IDFactura & "));"
    getdb().Execute m_SQL
    
    m_SQL = "DELETE TbFacturaSubcontratacionesInvolucradas.IDFactura " & _
                "FROM TbFacturaSubcontratacionesInvolucradas " & _
                "WHERE (((TbFacturaSubcontratacionesInvolucradas.IDFactura)=" & p_IDFactura & "));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbFacturaGastosInvolucrados.IDFactura " & _
                "FROM TbFacturaGastosInvolucrados " & _
                "WHERE (((TbFacturaGastosInvolucrados.IDFactura)=" & p_IDFactura & "));"
    getdb().Execute m_SQL
    
    m_SQL = "DELETE TbFacturaPrincipal.IDFactura " & _
                "FROM TbFacturaPrincipal " & _
                "WHERE (((TbFacturaPrincipal.IDFactura)=" & p_IDFactura & "));"
    CurrentDb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.EliminarFacturasAnteriores ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function GrabarFacturaPrincipal( _
                                        Optional ByRef p_Error As String _
                                        ) As String
     
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    
    Me.IDFactura = Me.IDFacturaCalculada
    m_SQL = "TbFacturaPrincipal"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDFactura") = Me.IDFactura
            .Fields("FechaFinalEventos") = Me.FechaFinalEventos
            .Fields("FechaFactura") = Me.FechaFactura
           
            If Me.Titulo1 = "" Then
                Me.Titulo1 = Me.Titulo1Calculado
            End If
            If Me.Titulo1 <> "" Then
                .Fields("Titulo1") = Me.Titulo1
            End If
            If Me.Titulo2 <> "" Then
                .Fields("Titulo2") = Me.Titulo2
            End If
            If Me.Titulo3 <> "" Then
                .Fields("Titulo3") = Me.Titulo3
            End If
            If Me.TipoImpositivo = "" Then
                Me.TipoImpositivo = Me.TipoImpositivoCalculado
            End If
            If Me.TipoImpositivo <> "" Then
                .Fields("TipoImpositivo") = Me.TipoImpositivo
            End If
            
            If Me.Recargo = "" Then
                Me.Recargo = Me.RecargoCalculado
            End If
            If Me.Recargo <> "" Then
                .Fields("Recargo") = Me.Recargo
            End If
            If IsNumeric(Me.PrecioHorasPorFacturar) Then
                .Fields("PrecioHoras") = Me.PrecioHorasPorFacturar
            End If
            If IsNumeric(Me.PrecioSuministroPorFacturar) Then
                .Fields("PrecioSuministro") = Me.PrecioSuministroPorFacturar
            End If
            If IsNumeric(Me.PrecioSuministroConRecargoPorFacturar) Then
                .Fields("PrecioSuministroConRecargo") = Me.PrecioSuministroConRecargoPorFacturar
            End If
            If IsNumeric(Me.PrecioDietasYGastosPorFacturar) Then
                .Fields("PrecioDietasYGastos") = Me.PrecioDietasYGastosPorFacturar
            End If
            If IsNumeric(Me.PrecioIncurridoPorFacturar) Then
                .Fields("PrecioIncurrido") = Me.PrecioIncurridoPorFacturar
            End If
            If IsNumeric(Me.PrecioSubcontratacionesPorFacturar) Then
                .Fields("PrecioSubcontrataciones") = Me.PrecioSubcontratacionesPorFacturar
            End If
            If IsNumeric(Me.PrecioSubcontratacionesConRecargoForFacturar) Then
                .Fields("PrecioSubcontratacionesConRecargo") = Me.PrecioSubcontratacionesConRecargoForFacturar
            End If
            If IsNumeric(Me.PrecioBaseImponiblePorFacturar) Then
                .Fields("PrecioBaseImponible") = Me.PrecioBaseImponiblePorFacturar
            End If
            If IsNumeric(Me.PrecioImpuestosPorFacturar) Then
                .Fields("PrecioImpuestos") = Me.PrecioImpuestosPorFacturar
            End If
            If IsNumeric(Me.ImporteAFacturarPorFacturar) Then
                .Fields("ImporteAFacturar") = Me.ImporteAFacturarPorFacturar
            End If
            If Me.NombreFacturaAnexa = "" Then
                Me.NombreFacturaAnexa = Me.NombreFacturaAnexaCalculada
            End If
            If Me.NombreFacturaAnexa <> "" Then
                .Fields("NombreFacturaAnexa") = Me.NombreFacturaAnexa
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.GrabarFacturaPrincipal ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Function GrabarEventos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_ObjColEventosParaFactura As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjEvento As Evento
    
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha introducido un número de factura existente"
        Err.Raise 1000
    End If
    Set m_ObjColEventosParaFactura = Me.colListaEventosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColEventosParaFactura Is Nothing Then
        Exit Function
    End If
    If Not m_ObjColEventosParaFactura Is Nothing Then
        
        
        m_SQL = "TbFacturaEventosInvolucrados"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        
        For Each m_ID In m_ObjColEventosParaFactura
            If Nz(m_ID, "") = "" Then
                GoTo siguienteEvento
            End If
            Set m_ObjEvento = m_ObjColEventosParaFactura(m_ID)
            With rcdDatos
                .AddNew
                    .Fields("IDFactura") = Me.IDFactura
                    .Fields("IDEvento") = m_ObjEvento.IDEVENTO
                    .Fields("NODO") = m_ObjEvento.NODO
                    .Fields("BUI") = m_ObjEvento.BUI
                    .Fields("SUBSISTEMA") = m_ObjEvento.SubSistema
                    .Fields("EQUIPO") = m_ObjEvento.Equipo.Equipo
                    If m_ObjEvento.PMPR <> "" Then
                        .Fields("PMPR") = m_ObjEvento.PMPR
                    End If
                    .Fields("TIPOEVENTO") = m_ObjEvento.TipoEvento
                    .Fields("CRITICIDAD") = m_ObjEvento.Criticidad
                    .Fields("TIPOTECNICO") = m_ObjEvento.Tecnico.Tipo
                    .Fields("ORIGINADOR") = m_ObjEvento.Originador
                    .Fields("FECHAALTA") = m_ObjEvento.FECHAALTAEVENTO
                    .Fields("HORAALTA") = m_ObjEvento.HORAINICIALEVENTO
                    .Fields("TIEMPORESPUESTA") = m_ObjEvento.TIEMPORESPUESTAEVENTO
                    .Fields("FECHAFIN") = m_ObjEvento.FechaFinal
                    .Fields("HORAFIN") = m_ObjEvento.HORAFINALEVENTO
                    .Fields("CAUSAFIN") = m_ObjEvento.CausaFin
                    If m_ObjEvento.Contacto <> "" Then
                        .Fields("CONTACTO") = m_ObjEvento.Contacto
                    End If
                    .Fields("DESCRIPCION") = m_ObjEvento.Descripcion
                    
                .Update
                
            End With
siguienteEvento:
            Set m_ObjEvento = Nothing
        Next
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    GrabarEventos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarEventos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GrabarActividades( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_ID As Variant
    Dim m_ObjActividad As Actividad
    Dim m_ObjEvento As Evento
    Dim m_Tecnico As Tecnico
    Dim m_TipoTecnicoPrecio As TipoTecnicoPrecio
    Dim m_ObjColActividadesDeEventosParaFactura As Scripting.Dictionary
    Dim m_ObjColActividadesNoDeEventosParaFactura As Scripting.Dictionary
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha introducido un número de factura existente"
        Err.Raise 1000
    End If
    Set m_ObjColActividadesDeEventosParaFactura = Me.colListaActividadesDeEventosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjColActividadesNoDeEventosParaFactura = Me.colListaActividadesNODeEventosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
        
    If Not m_ObjColActividadesDeEventosParaFactura Is Nothing Then
        m_SQL = "TbFacturaActividadesInvolucradas"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        
        For Each m_ID In m_ObjColActividadesDeEventosParaFactura
            If Nz(m_ID, "") = "" Then
                GoTo siguienteActividadDeEvento
            End If
            
            Set m_ObjActividad = m_ObjColActividadesDeEventosParaFactura(m_ID)
            Set m_ObjEvento = m_ObjActividad.Evento
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjEvento Is Nothing Then
                p_Error = "No se ha podido obtener el Evento de la actividad " & m_ID
                Err.Raise 1000
            End If
            With rcdDatos
                .AddNew
                    .Fields("IDFactura") = Me.IDFactura
                    .Fields("IDActividad") = m_ObjActividad.IDActividad
                    .Fields("NODO") = m_ObjEvento.NODO
                    .Fields("BUI") = m_ObjEvento.BUI
                    .Fields("SubSistema") = m_ObjEvento.SubSistema
                    .Fields("EQUIPO") = m_ObjEvento.Equipo.Equipo
                    .Fields("IDEVENTO") = m_ObjEvento.IDEVENTO
                    Set m_Tecnico = m_ObjActividad.Tecnico
                    p_Error = m_ObjActividad.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    
                    If m_Tecnico Is Nothing Then
                        p_Error = "NO se ha podido obtener el técnico de la Actividad: " & m_ObjActividad.IDActividad
                        Err.Raise 1000
                    End If
                    Set m_TipoTecnicoPrecio = m_Tecnico.TipoTecnicoPrecio
                    p_Error = m_Tecnico.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If m_TipoTecnicoPrecio Is Nothing Then
                        p_Error = "NO se ha podido obtener el m_TipoTecnicoPrecio de la Actividad: " & m_ObjActividad.IDActividad
                        Err.Raise 1000
                    End If
                    .Fields("AliasTecnico") = m_ObjActividad.ALIASTECNICO
                    .Fields("TIPOTECNICO") = m_Tecnico.Tipo
                    .Fields("Perfil") = m_Tecnico.Tipo
                    .Fields("PrecioHoraLaborablePerfil") = m_TipoTecnicoPrecio.PrecioHoraLab
                    .Fields("PrecioHoraExtraPerfil") = m_TipoTecnicoPrecio.PrecioHoraExtra
                    .Fields("TipoActividad") = m_ObjActividad.Actividad
                    .Fields("Descripcion") = m_ObjActividad.Descripcion
                    .Fields("FechaAlta") = m_ObjActividad.FechaAlta
                    .Fields("HorasLaborables") = m_ObjActividad.HorasLaborables
                    .Fields("HorasExtras") = m_ObjActividad.HorasExtras
                    
                    If m_ObjActividad.Ubicacion <> "" Then
                        .Fields("Ubicacion") = m_ObjActividad.Ubicacion
                    End If
                    
                .Update
                
            End With
siguienteActividadDeEvento:
            Set m_Tecnico = Nothing
            Set m_TipoTecnicoPrecio = Nothing
            Set m_ObjActividad = Nothing
        Next
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If m_ObjColActividadesNoDeEventosParaFactura Is Nothing Then
        GrabarActividades = "OK"
        Exit Function
    End If
    m_SQL = "TbFacturaActividadesInvolucradas"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    For Each m_ID In m_ObjColActividadesNoDeEventosParaFactura
        If Nz(m_ID, "") = "" Then
            GoTo siguienteActividadNODeEvento
        End If
        Set m_ObjActividad = m_ObjColActividadesNoDeEventosParaFactura(m_ID)
        With rcdDatos
            .AddNew
                .Fields("IDFactura") = Me.IDFactura
                .Fields("IDActividad") = m_ObjActividad.IDActividad
                
                Set m_Tecnico = m_ObjActividad.Tecnico
                p_Error = m_ObjActividad.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                
                If m_Tecnico Is Nothing Then
                    p_Error = "NO se ha podido obtener el técnico de la Actividad: " & m_ObjActividad.IDActividad
                    Err.Raise 1000
                End If
                Set m_TipoTecnicoPrecio = m_Tecnico.TipoTecnicoPrecio
                p_Error = m_Tecnico.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_TipoTecnicoPrecio Is Nothing Then
                    p_Error = "NO se ha podido obtener el m_TipoTecnicoPrecio de la Actividad: " & m_ObjActividad.IDActividad
                    Err.Raise 1000
                End If
                If m_ObjActividad.ALIASTECNICO <> "" Then
                    .Fields("AliasTecnico") = m_ObjActividad.ALIASTECNICO
                End If
                If m_TipoTecnicoPrecio.TIPOTECNICO <> "" Then
                    .Fields("TIPOTECNICO") = m_Tecnico.Tipo
                End If
                If m_Tecnico.Tipo <> "" Then
                    .Fields("Perfil") = m_Tecnico.Tipo
                End If
                .Fields("PrecioHoraLaborablePerfil") = m_TipoTecnicoPrecio.PrecioHoraLab
                .Fields("PrecioHoraExtraPerfil") = m_TipoTecnicoPrecio.PrecioHoraExtra
                
                If m_ObjActividad.Actividad <> "" Then
                    .Fields("TipoActividad") = m_ObjActividad.Actividad
                End If
                If m_ObjActividad.Descripcion <> "" Then
                    .Fields("Descripcion") = m_ObjActividad.Descripcion
                End If
                If m_ObjActividad.FechaAlta <> "" Then
                    .Fields("FechaAlta") = m_ObjActividad.FechaAlta
                End If
                .Fields("HorasLaborables") = m_ObjActividad.HorasLaborables
                .Fields("HorasExtras") = m_ObjActividad.HorasExtras
                If m_ObjActividad.Ubicacion <> "" Then
                    .Fields("Ubicacion") = m_ObjActividad.Ubicacion
                End If
                
            .Update
           
        End With
siguienteActividadNODeEvento:
        Set m_Tecnico = Nothing
        Set m_TipoTecnicoPrecio = Nothing
        Set m_ObjActividad = Nothing
    Next
    rcdDatos.Close
    Set rcdDatos = Nothing
        
    
    GrabarActividades = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarActividades ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GrabarMateriales( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_ObjColMaterialesParaFactura As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjEvento As Evento
    Dim m_ObjActividad As Actividad
    Dim m_ObjMaterial As Material
    Dim m_ObjEquipo As Equipo
    Dim m_Nodo As String
    Dim m_BUI As String
    Dim m_SUBSISTEMA As String
    Dim m_IDEvento As String
    Dim m_Equipo As String
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha introducido un número de factura existente"
        Err.Raise 1000
    End If
    Set m_ObjColMaterialesParaFactura = Me.colListaMaterialesPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjColMaterialesParaFactura Is Nothing Then
        m_SQL = "TbFacturaMaterialesInvolucrados"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        
        For Each m_ID In m_ObjColMaterialesParaFactura
            If Nz(m_ID, "") = "" Then
                GoTo siguienteMaterial
            End If
            
            Set m_ObjMaterial = m_ObjColMaterialesParaFactura(m_ID)
            Set m_ObjActividad = m_ObjMaterial.Actividad
            p_Error = m_ObjMaterial.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjActividad Is Nothing Then
                p_Error = "NO se ha podido obtener la actividad del material: " & CStr(m_ID)
                Err.Raise 1000
            End If
            Set m_ObjEvento = m_ObjActividad.Evento
            p_Error = m_ObjActividad.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjEvento Is Nothing Then
                p_Error = "NO se ha podido obtener El Evento de la actividad del material: " & CStr(m_ID)
                Err.Raise 1000
            End If
            m_IDEvento = m_ObjEvento.IDEVENTO
            Set m_ObjEquipo = m_ObjMaterial.Equipo
            p_Error = m_ObjMaterial.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_ObjEquipo Is Nothing Then
                m_Equipo = m_ObjEquipo.Equipo
                m_Nodo = m_ObjEquipo.NODO
                m_BUI = m_ObjEquipo.BUI
                m_SUBSISTEMA = m_ObjEquipo.SubSistema
            Else
                If Not m_ObjEvento Is Nothing Then
                    m_Nodo = m_ObjEvento.NODO
                    m_BUI = m_ObjEvento.BUI
                    m_SUBSISTEMA = m_ObjEvento.SubSistema
                    m_IDEvento = m_ObjEvento.IDEVENTO
                End If
            End If
            With rcdDatos
                .AddNew
                    .Fields("IDFactura") = Me.IDFactura
                    .Fields("IDMaterial") = m_ObjMaterial.IDMaterial
                    'Debug.Print m_ObjMaterial.IDMaterial
                    If m_Nodo <> "" Then
                        .Fields("NODO") = m_Nodo
                    End If
                    If m_BUI <> "" Then
                        .Fields("BUI") = m_BUI
                    End If
                    If m_SUBSISTEMA <> "" Then
                        .Fields("SubSistema") = m_SUBSISTEMA
                    End If
                    
                    If m_Equipo <> "" Then
                        .Fields("EQUIPO") = m_Equipo
                    End If
                    
                    If m_IDEvento <> "" Then
                        .Fields("IDEvento") = m_IDEvento
                    End If
                    If m_ObjMaterial.TipoAccion <> "" Then
                        .Fields("TIPOACCION") = m_ObjMaterial.TipoAccion
                    End If
                    
                    If m_ObjMaterial.PN <> "" Then
                        .Fields("PN") = m_ObjMaterial.PN
                    End If
                    If m_ObjMaterial.NS <> "" Then
                        .Fields("NS") = m_ObjMaterial.NS
                    End If
                    
                    If Not m_ObjEquipo Is Nothing Then
                        If m_ObjEquipo.Descripcion <> "" Then
                            .Fields("DESCRIPCIONEQUIPO") = m_ObjEquipo.Descripcion
                        End If
                        
                    End If
                    
                    If m_ObjMaterial.FechaEntrega <> "" Then
                        .Fields("FECHAENTREGA") = m_ObjMaterial.FechaEntrega
                    End If
                    If IsNumeric(m_ObjMaterial.COSTE) Then
                        .Fields("PRECIOSINIVA") = m_ObjMaterial.COSTE
                    End If
                    If m_ObjMaterial.ReparadoPor <> "" Then
                        .Fields("ReparadoPor") = m_ObjMaterial.ReparadoPor
                    End If
                .Update
                
            End With
            
siguienteMaterial:
            m_Nodo = ""
            m_BUI = ""
            m_SUBSISTEMA = ""
            m_Nodo = ""
            m_IDEvento = ""
            m_Equipo = ""
            Set m_ObjMaterial = Nothing
            Set m_ObjActividad = Nothing
            Set m_ObjEvento = Nothing
            Set m_ObjEquipo = Nothing
        Next
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
     
      
    GrabarMateriales = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarMateriales ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GrabarGastos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As Variant
    Dim m_ColGastosPorFacturar As Scripting.Dictionary
    Dim m_ObjGasto As Gasto
    
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha introducido un número de factura existente"
        Err.Raise 1000
    End If
    Set m_ColGastosPorFacturar = Me.colListaGastosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColGastosPorFacturar Is Nothing Then
        
       
        
        m_SQL = "TbFacturaGastosInvolucrados"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        For Each m_ID In m_ColGastosPorFacturar
            If Nz(m_ID, "") = "" Then
                GoTo siguienteGasto
            End If
            If Not IsNumeric(m_ID) Then
                p_Error = "Se ha colado un gasto no numérico en la colección m_ColGastosPorFacturar" & vbNewLine & m_ID
                Err.Raise 1000
            End If
            
            Set m_ObjGasto = m_ColGastosPorFacturar(m_ID)
            With rcdDatos
                .AddNew
                    .Fields("IDFactura") = Me.IDFactura
                    .Fields("IDGasto") = m_ObjGasto.IDGasto
                    If m_ObjGasto.TipoGasto <> "" Then
                        .Fields("TipoGasto") = m_ObjGasto.TipoGasto
                    End If
                    
                    If IsNumeric(m_ObjGasto.ImporteUnitario) Then
                        .Fields("ImporteUnitario") = m_ObjGasto.ImporteUnitario
                    End If
                    If IsNumeric(m_ObjGasto.NumeroUnidades) Then
                        .Fields("NumeroUnidades") = m_ObjGasto.NumeroUnidades
                    End If
                    If IsDate(m_ObjGasto.FechaImputacionGasto) Then
                        .Fields("FechaImputacionGasto") = m_ObjGasto.FechaImputacionGasto
                    End If
                    .Fields("ALIAS") = m_ObjGasto.Alias
                    If m_ObjGasto.Observaciones <> "" Then
                        .Fields("Observaciones") = m_ObjGasto.Observaciones
                    End If
                .Update
                
            End With
siguienteGasto:
            Set m_ObjGasto = Nothing
        Next
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
        
      
    GrabarGastos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarGastos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function GrabarSubcontrataciones( _
                                            Optional ByRef p_Error As String _
                                        ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_ObjSubContratacion As Subcontratacion
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha introducido un número de factura existente"
        Err.Raise 1000
    End If
    Set m_Col = Me.colListaSubContratacionesPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        GrabarSubcontrataciones = "OK"
        Exit Function
    End If
    m_SQL = "TbFacturaSubcontratacionesInvolucradas"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    For Each m_ID In m_Col
        If Nz(m_ID, "") = "" Then
            GoTo siguienteSubContratacion
        End If
        Set m_ObjSubContratacion = m_Col(m_ID)
        With rcdDatos
            .AddNew
                .Fields("IDFactura") = Me.IDFactura
                .Fields("IDSubcontratacion") = m_ObjSubContratacion.IDSubContratacion
                .Fields("FechaAlta") = m_ObjSubContratacion.FechaAlta
                If IsDate(m_ObjSubContratacion.FechaFin) Then
                    .Fields("FechaFin") = m_ObjSubContratacion.FechaFin
                End If
                If m_ObjSubContratacion.Empresa <> "" Then
                    .Fields("Empresa") = m_ObjSubContratacion.Empresa
                End If
                If m_ObjSubContratacion.Descripcion <> "" Then
                    .Fields("Descripcion") = m_ObjSubContratacion.Descripcion
                End If
                
                If IsNumeric(m_ObjSubContratacion.Importe) Then
                    .Fields("Importe") = m_ObjSubContratacion.Importe
                End If
            .Update
            
        End With
siguienteSubContratacion:
        Set m_ObjSubContratacion = Nothing
    Next
    rcdDatos.Close
    Set rcdDatos = Nothing
    GrabarSubcontrataciones = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarSubcontrataciones ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function getHorasPerfil( _
                                p_Perfil As String, _
                                Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_HLAb As Double
    Dim m_HExtra As Double
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjActividad As Actividad
    Dim m_ObjTecnico As Tecnico
    
    
    
    On Error GoTo errores
    
    If p_Perfil = "" Then
        Exit Function
    End If
   'set m_ObjColPerfiles=
    If p_EsCalculado = EnumSino.Sí Then
        If Not IsNumeric(Me.IDFactura) Then
            p_Error = "No se ha introducido un número de factura existente"
            Err.Raise 1000
        End If
        m_SQL = "SELECT Sum(TbFacturaActividadesInvolucradas.HORASLABORABLES) AS SumaDeHORASLABORABLES, " & _
                "Sum(TbFacturaActividadesInvolucradas.HORASEXTRAS) AS SumaDeHORASEXTRAS " & _
                "FROM TbFacturaActividadesInvolucradas " & _
                "WHERE (((TbFacturaActividadesInvolucradas.IDFactura)=" & Me.IDFactura & _
                ") AND ((TbFacturaActividadesInvolucradas.PERFIL)='" & p_Perfil & "'));"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                If IsNumeric(Nz(.Fields("SumaDeHORASLABORABLES"), "0")) Then
                    m_HLAb = CDbl(.Fields("SumaDeHORASLABORABLES"))
                End If
                If IsNumeric(Nz(.Fields("SumaDeHORASEXTRAS"), "0")) Then
                    m_HExtra = CDbl(.Fields("SumaDeHORASEXTRAS"))
                End If
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        getHorasPerfil = CStr(m_HLAb) & "|" & CStr(m_HExtra)
        Exit Function
    
    End If
    
    Set m_Col = Me.colListaActividadesDeEventosPorFacturar
    
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            Set m_ObjTecnico = m_ObjActividad.Tecnico
            If m_ObjTecnico.Tipo = p_Perfil Then
                If IsNumeric(m_ObjActividad.HorasLaborables) Then
                    If CDbl(m_ObjActividad.HorasLaborables) <> 0 Then
                        m_HLAb = m_HLAb + CDbl(m_ObjActividad.HorasLaborables)
                    End If
                End If
                If IsNumeric(m_ObjActividad.HorasExtras) Then
                    If CDbl(m_ObjActividad.HorasExtras) <> 0 Then
                        m_HExtra = m_HExtra + CDbl(m_ObjActividad.HorasExtras)
                    End If
                End If
            End If
            
            Set m_ObjActividad = Nothing
        Next
    End If
    Set m_Col = Me.colListaActividadesNODeEventosPorFacturar
    
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_ObjActividad = m_Col(m_ID)
            Set m_ObjTecnico = m_ObjActividad.Tecnico
            If m_ObjTecnico.Tipo = p_Perfil Then
                If IsNumeric(m_ObjActividad.HorasLaborables) Then
                    If CDbl(m_ObjActividad.HorasLaborables) <> 0 Then
                        m_HLAb = m_HLAb + CDbl(m_ObjActividad.HorasLaborables)
                    End If
                End If
                If IsNumeric(m_ObjActividad.HorasExtras) Then
                    If CDbl(m_ObjActividad.HorasExtras) <> 0 Then
                        m_HExtra = m_HExtra + CDbl(m_ObjActividad.HorasExtras)
                    End If
                End If
            End If
            
            Set m_ObjActividad = Nothing
        Next
    End If
    getHorasPerfil = CStr(m_HLAb) & "|" & CStr(m_HExtra)
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHorasPerfil ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GrabarTablasRelacionadas(Optional ByRef p_Error As String) As String
    
    Dim blnGrabadoFacturaPrincipal As Boolean
    
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        Me.IDFactura = Me.IDFacturaCalculada
    End If
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha introducido un número de factura existente"
        Err.Raise 1000
    End If
    GrabarFacturaPrincipal p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    blnGrabadoFacturaPrincipal = True
    GrabarEventos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    GrabarActividades p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    GrabarMateriales p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    GrabarGastos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    GrabarSubcontrataciones p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    GrabarPerfilesPrecios p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    GrabarTablasRelacionadas = "OK"
    Exit Function
errores:
    If blnGrabadoFacturaPrincipal = True And IsNumeric(Me.IDFactura) Then
        m_SQL = "DELETE * FROM TbFacturaPrincipal WHERE IDFactura=" & Me.IDFactura & ";"
        getdb().Execute m_SQL
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarTablasRelacionadas ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Function GrabarPerfilesPrecios(Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjPerfilFactura As FacturaPerfiles
    Dim m_IDPerfilFactura As String
    
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha introducido un número de factura existente"
        Err.Raise 1000
    End If
    Set m_Col = Me.ColPerfilesPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        GrabarPerfilesPrecios = "OK"
        Exit Function
    End If
    m_SQL = "TbFacturaPrincipalPerfiles"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
    For Each m_ID In m_Col
        Set m_ObjPerfilFactura = m_Col(m_ID)
        m_IDPerfilFactura = DameID("TbFacturaPrincipalPerfiles", "IDFacturaPerfil")
        With rcdDatos
            .AddNew
                .Fields("IDFacturaPerfil") = m_IDPerfilFactura
                .Fields("IDFactura") = Me.IDFactura
                .Fields("Perfil") = m_ObjPerfilFactura.Perfil
                .Fields("PrecioHoraLab") = m_ObjPerfilFactura.PrecioHoraLab
                .Fields("PrecioHoraExt") = m_ObjPerfilFactura.PrecioHoraExt
                .Fields("NHorasLab") = m_ObjPerfilFactura.NHorasLab
                .Fields("NHorasExt") = m_ObjPerfilFactura.NHorasExt
            .Update
            
        End With
        Set m_ObjPerfilFactura = Nothing
    Next
    
    GrabarPerfilesPrecios = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarPerfilesPrecios ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function GrabarFactura( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ColEventosPorFacturar As Scripting.Dictionary
    Dim m_ColActividadesDeEventosPorFacturar As Scripting.Dictionary
    Dim m_ColActividadesNODeEventosPorFactruar As Scripting.Dictionary
    Dim m_ColMaterialesPorFacturar As Scripting.Dictionary
    Dim m_ColGastosPorFacturar As Scripting.Dictionary
    Dim m_ColSubcontratacionesPorFacturar As Scripting.Dictionary
    Dim m_ColPerfilesPorFacturar As Scripting.Dictionary
    Dim m_ObjFacturaGrabada As Factura
    On Error GoTo errores
    
    
    'ValidacionFactura EnumSino.no, p_Error
    
    
    Set m_ColEventosPorFacturar = Me.colListaEventosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColActividadesDeEventosPorFacturar = Me.colListaActividadesDeEventosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColActividadesNODeEventosPorFactruar = Me.colListaActividadesNODeEventosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColMaterialesPorFacturar = Me.colListaMaterialesPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColGastosPorFacturar = Me.colListaGastosPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColSubcontratacionesPorFacturar = Me.colListaSubContratacionesPorFacturar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColEventosPorFacturar Is Nothing And _
        m_ColActividadesDeEventosPorFacturar Is Nothing And _
        m_ColActividadesNODeEventosPorFactruar Is Nothing And _
        m_ColMaterialesPorFacturar Is Nothing And _
        m_ColGastosPorFacturar Is Nothing And _
        m_ColSubcontratacionesPorFacturar Is Nothing Then
        p_Error = "No hay nada que facturar para esa fecha"
        Err.Raise 1000
    End If
    EliminarFacturasAnteriores Me.IDFactura, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    GrabarTablasRelacionadas p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    
    GrabarFactura = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarFactura ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    
End Function


Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
   
    
    Dim m_SQL As String
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "No se ha indicado un IDFactura numérico válido"
        Err.Raise 1000
    End If
    If fso.FileExists(Me.URLFactura) Then
        If FicheroAbierto(Me.URLFactura) Then
            p_Error = "Ha de cerrar la factura primero"
            Err.Raise 1000
        End If
    End If
    m_SQL = "DELETE TbFacturaEventosInvolucrados.IDFactura " & _
            "FROM TbFacturaEventosInvolucrados " & _
            "WHERE (((TbFacturaEventosInvolucrados.IDFactura)=" & Me.IDFactura & "));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbFacturaActividadesInvolucradas.IDFactura " & _
            "FROM TbFacturaActividadesInvolucradas " & _
            "WHERE (((TbFacturaActividadesInvolucradas.IDFactura)=" & Me.IDFactura & "));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbFacturaMaterialesInvolucrados.IDFactura " & _
            "FROM TbFacturaMaterialesInvolucrados " & _
            "WHERE (((TbFacturaMaterialesInvolucrados.IDFactura)=" & Me.IDFactura & "));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbFacturaGastosInvolucrados.IDFactura " & _
            "FROM TbFacturaGastosInvolucrados " & _
            "WHERE (((TbFacturaGastosInvolucrados.IDFactura)=" & Me.IDFactura & "));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbFacturaSubcontratacionesInvolucradas.IDFactura " & _
            "FROM TbFacturaSubcontratacionesInvolucradas " & _
            "WHERE (((TbFacturaSubcontratacionesInvolucradas.IDFactura)=" & Me.IDFactura & "));"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbFacturaPrincipalPerfiles.IDFactura " & _
            "FROM TbFacturaPrincipalPerfiles " & _
            "WHERE (((TbFacturaPrincipalPerfiles.IDFactura)=" & Me.IDFactura & "));"
    getdb().Execute m_SQL
    
    m_SQL = "DELETE TbFacturaPrincipal.* " & _
            "FROM TbFacturaPrincipal " & _
            "WHERE (((TbFacturaPrincipal.IDFactura)=" & Me.IDFactura & "));"
    getdb().Execute m_SQL
    
    
    If fso.FileExists(Me.URLFactura) Then
        fso.DeleteFile Me.URLFactura, True
    End If
    
    Eliminar = "Factura eliminada correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function


Public Function ObtenerLI( _
                            Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_ObjColEventosDeFactura As Scripting.Dictionary
    Dim m_ObjColMaterialesDeFactura As Scripting.Dictionary
    Dim m_ObjColSubContratacionesDeFactura As Scripting.Dictionary
    Dim p_NombreLI As String
    Dim p_URLCompletaLI As String
    Dim appexcel As excel.Application
    Dim WbLibro As excel.Workbook
    Dim wbHoja As excel.Worksheet
    Dim m_URLDirectorio As String
    
    On Error GoTo errores
    
    If p_EsCalculado = 0 Then
        p_EsCalculado = EnumSino.Sí
    End If
    
    If Me.IDFactura <> "" Then
        m_URLDirectorio = m_ObjEntorno.URLDirectorioCertificaciones
    Else
        m_URLDirectorio = m_ObjEntorno.URLDirectorioLocal
        
    End If
    
    
    m_Linea = "LI Obtención de colecciones ....."
    Avance m_Linea
    
    m_Linea = "LI Obtención de colecciones .....Eventos"
    Avance m_Linea
    If Me.IDFactura = "" Then
        Set m_ObjColEventosDeFactura = Me.colListaEventosPorFacturar
    Else
        Set m_ObjColEventosDeFactura = Me.colListaEventos
    End If
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Linea = "LI Obtención de colecciones .....Actividades"
    Avance m_Linea
    If Me.IDFactura = "" Then
        Set m_ObjColMaterialesDeFactura = Me.colListaMaterialesPorFacturar
    Else
        Set m_ObjColMaterialesDeFactura = Me.colListaMateriales
    End If
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Linea = "LI Obtención de colecciones .....SubContrataciones"
    Avance m_Linea
    If Me.IDFactura = "" Then
        Set m_ObjColSubContratacionesDeFactura = Me.colListaSubContratacionesPorFacturar
    Else
        Set m_ObjColSubContratacionesDeFactura = Me.colListaSubContrataciones
    End If
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColEventosDeFactura Is Nothing And m_ObjColMaterialesDeFactura Is Nothing And m_ObjColSubContratacionesDeFactura Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    
    p_NombreLI = "LI" & Format(Me.FechaFactura, "yyyy") & Format(Me.FechaFactura, "MM") & Format(Me.FechaFactura, "dd")
    p_NombreLI = p_NombreLI & ".xlsx"
    p_URLCompletaLI = m_URLDirectorio & p_NombreLI
    If fso.FileExists(p_URLCompletaLI) Then
        
        If FicheroAbierto(p_URLCompletaLI) Then
            p_Error = "Ha de cerrar el listado actual en excel"
            Err.Raise 1000
        End If
        fso.DeleteFile p_URLCompletaLI, True
    End If
    m_Linea = "LI Excel .....Abriendo"
    Avance m_Linea
    Set appexcel = New excel.Application
    appexcel.Visible = False
    Set WbLibro = appexcel.Workbooks.Add
    
    Set wbHoja = WbLibro.Worksheets(1)
    wbHoja.Name = "SUBCONT_ASISTENCIA_ING"
    GenerarExcelSubContratacionesLI wbHoja, m_ObjColSubContratacionesDeFactura, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set wbHoja = WbLibro.Worksheets.Add
    wbHoja.Name = "MATERIALES"
    GenerarExcelMaterialesLI wbHoja, m_ObjColMaterialesDeFactura, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set wbHoja = WbLibro.Worksheets.Add
    wbHoja.Name = "EVENTOS"
    GenerarExcelEventosLI wbHoja, m_ObjColEventosDeFactura, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Linea = "LI Excel .....Grabando documento generado"
    Avance m_Linea
    WbLibro.SaveAs p_URLCompletaLI
    Set wbHoja = Nothing
    WbLibro.Close
    Set WbLibro = Nothing
    appexcel.Quit
    Set appexcel = Nothing
    
    ObtenerLI = p_URLCompletaLI
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.ObtenerLI ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
     If Not wbHoja Is Nothing Then
        Set wbHoja = Nothing
    End If
    If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not appexcel Is Nothing Then
       appexcel.Quit
        Set appexcel = Nothing
    End If
End Function

Private Function GenerarExcelEventosLI( _
                                        ByRef wbHoja As excel.Worksheet, _
                                        ByRef p_ObjColEventosDeFactura As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String _
                                        ) As String
       
    Dim intFila As Integer
    Dim lngNumeroRegistro As Long
    Dim lngNumerRegistrosTotales As Long
    Dim m_IDEvento As Variant
    Dim m_ObjEvento As Evento
    On Error GoTo errores
    
    
    '-----------
    ' parte inicial de la cabecera
    '--------------
    intFila = 1
    With wbHoja
        .Cells(intFila, 1).Value = "NODO"
        .Cells(intFila, 2).Value = "BUI"
        .Cells(intFila, 3).Value = "SUBSISTEMA"
        .Cells(intFila, 4).Value = "EQUIPO"
        .Cells(intFila, 5).Value = "PARTE TRABAJO"
        .Cells(intFila, 6).Value = "PM/PR"
        .Cells(intFila, 7).Value = "NºHORAS_LAB"
        .Cells(intFila, 8).Value = "NºHORAS_EXT"
    End With
    With wbHoja
        .Range(.Cells(1, 1), .Cells(1, 8)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 8)).Interior.ColorIndex = 15
        .Range(.Cells(1, 1), .Cells(1, 8)).HorizontalAlignment = xlCenter
    End With
    If Not p_ObjColEventosDeFactura Is Nothing Then
        If p_ObjColEventosDeFactura.count > 0 Then
            lngNumerRegistrosTotales = p_ObjColEventosDeFactura.count
            lngNumeroRegistro = 1
            intFila = intFila + 1
            For Each m_IDEvento In p_ObjColEventosDeFactura.Keys
                Set m_ObjEvento = p_ObjColEventosDeFactura(m_IDEvento)
                m_Linea = "LI Excel .....Eventos " & m_ObjEvento.IDEVENTO & " ( " & lngNumeroRegistro & "\" & lngNumerRegistrosTotales & ")"""
                Avance m_Linea
                
                If m_ObjEvento.NODO <> "" Then
                    wbHoja.Cells(intFila, 1).Value = m_ObjEvento.NODO
                End If
                If m_ObjEvento.BUI <> "" Then
                    wbHoja.Cells(intFila, 2).Value = m_ObjEvento.BUI
                End If
                If m_ObjEvento.SubSistema <> "" Then
                    wbHoja.Cells(intFila, 3).Value = m_ObjEvento.SubSistema
                End If
                If m_ObjEvento.NombreEquipo <> "" Then
                    wbHoja.Cells(intFila, 4).Value = m_ObjEvento.NombreEquipo
                End If
                If m_ObjEvento.IDEVENTO <> "" Then
                    wbHoja.Cells(intFila, 5).Value = m_IDEvento
                End If
                If m_ObjEvento.PMPR <> "" Then
                    wbHoja.Cells(intFila, 6).Value = "'" & m_ObjEvento.PMPR
                End If
                If m_ObjEvento.HorasLabTotales <> "" Then
                    wbHoja.Cells(intFila, 7).Value = "'" & m_ObjEvento.HorasLabTotales
                End If
                If m_ObjEvento.HorasExtTotales <> "" Then
                    wbHoja.Cells(intFila, 8).Value = "'" & m_ObjEvento.HorasExtTotales
                End If
                Set m_ObjEvento = Nothing
                intFila = intFila + 1
                lngNumeroRegistro = lngNumeroRegistro + 1
            Next
            RecuadrarRango wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(intFila - 1, 8)), "Multiple"
           
            
        Else
            RecuadrarRango wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 8)), "Multiple"
        End If
    End If
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
    End With
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarExcelEventosLI ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GenerarExcelMaterialesLI( _
                                        ByRef wbHoja As excel.Worksheet, _
                                        ByRef p_ObjColMaterialesDeFactura As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String _
                                        ) As String
       
    Dim intFila As Integer
    Dim lngNumeroRegistro As Long
    Dim lngNumerRegistrosTotales As Long
    Dim m_IDMaterial As Variant
    Dim m_ObjMaterial As Material
    Dim m_ObjActividad As Actividad
    Dim m_ObjEvento As Evento
    Dim m_IDEvento As String
    Dim m_Nodo As String
    Dim m_BUI As String
    Dim m_SUBSISTEMA As String
    Dim m_NombreEquipo As String
    Dim m_ObjEquipo As Equipo
    On Error GoTo errores
    '-----------
    ' parte inicial de la cabecera
    '--------------
    intFila = 1
    With wbHoja
        .Cells(intFila, 1).Value = "NODO"
        .Cells(intFila, 2).Value = "BUI"
        .Cells(intFila, 3).Value = "SUBSISTEMA"
        .Cells(intFila, 4).Value = "ID_EQUIPO"
        .Cells(intFila, 5).Value = "PARTE TRABAJO"
        .Cells(intFila, 6).Value = "TIPO ACCIÓN"
        .Cells(intFila, 7).Value = "MATERIAL"
        .Cells(intFila, 8).Value = "P/N"
        .Cells(intFila, 9).Value = "S/N"
        .Cells(intFila, 10).Value = "DESCRIPCIÓN EQUIPO/MATERIAL"
    End With
    With wbHoja
        .Range(.Cells(1, 1), .Cells(1, 10)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 10)).Interior.ColorIndex = 15
        .Range(.Cells(1, 1), .Cells(1, 10)).HorizontalAlignment = xlCenter
    End With
    If Not p_ObjColMaterialesDeFactura Is Nothing Then
        If p_ObjColMaterialesDeFactura.count > 0 Then
            lngNumerRegistrosTotales = p_ObjColMaterialesDeFactura.count
            lngNumeroRegistro = 1
            intFila = intFila + 1
            For Each m_IDMaterial In p_ObjColMaterialesDeFactura.Keys
                Set m_ObjMaterial = p_ObjColMaterialesDeFactura(m_IDMaterial)
                m_Linea = "LI Excel .....Materiales ( " & lngNumeroRegistro & "\" & lngNumerRegistrosTotales & ")"""
                Avance m_Linea
                
                'm_ProvieneDeActividad = m_ObjMaterial.ProvieneDeActividad
                If m_ObjMaterial.IDActividad <> "" Then
                    Set m_ObjActividad = m_ObjMaterial.Actividad
                    p_Error = m_ObjMaterial.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_ObjActividad Is Nothing Then
                        Set m_ObjEvento = m_ObjActividad.Evento
                        p_Error = m_ObjActividad.Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        m_IDEvento = m_ObjEvento.IDEVENTO
                        m_Nodo = m_ObjEvento.NODO
                        m_BUI = m_ObjEvento.BUI
                        m_SUBSISTEMA = m_ObjEvento.SubSistema
                        m_NombreEquipo = m_ObjEvento.Equipo.Descripcion
                    End If
                Else
                    m_IDEvento = ""
                    Set m_ObjEquipo = m_ObjMaterial.Equipo
                    p_Error = m_ObjMaterial.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_ObjEquipo Is Nothing Then
                        m_Nodo = m_ObjEquipo.NODO
                        m_BUI = m_ObjEquipo.BUI
                        m_SUBSISTEMA = m_ObjEquipo.SubSistema
                        m_NombreEquipo = m_ObjEquipo.Descripcion
                    End If
                End If
                
                If m_Nodo <> "" Then
                    wbHoja.Cells(intFila, 1).Value = m_Nodo
                End If
                If m_BUI <> "" Then
                    wbHoja.Cells(intFila, 2).Value = m_BUI
                End If
                If m_SUBSISTEMA <> "" Then
                    wbHoja.Cells(intFila, 3).Value = m_SUBSISTEMA
                End If
                If m_NombreEquipo <> "" Then
                    wbHoja.Cells(intFila, 4).Value = m_NombreEquipo
                End If
                If m_ObjEvento.IDEVENTO <> "" Then
                    If CStr(m_IDEvento) <> "" Then
                        wbHoja.Cells(intFila, 5).Value = m_IDEvento
                    End If
                
                End If
                If m_ObjMaterial.TipoAccion <> "" Then
                    wbHoja.Cells(intFila, 6).Value = m_ObjMaterial.TipoAccion
                End If
                If m_ObjMaterial.Material <> "" Then
                    wbHoja.Cells(intFila, 7).Value = m_ObjMaterial.Material
                End If
                If m_ObjMaterial.PN <> "" Then
                    wbHoja.Cells(intFila, 8).Value = m_ObjMaterial.PN
                End If
                If m_ObjMaterial.NS <> "" Then
                    wbHoja.Cells(intFila, 9).Value = m_ObjMaterial.NS
                End If
                If m_ObjMaterial.Descripcion <> "" Then
                    wbHoja.Cells(intFila, 10).Value = m_ObjMaterial.Descripcion
                End If
                Set m_ObjMaterial = Nothing
                intFila = intFila + 1
                lngNumeroRegistro = lngNumeroRegistro + 1
            Next
            Set m_ObjActividad = Nothing
            Set m_ObjEvento = Nothing
            Set m_ObjEquipo = Nothing
            RecuadrarRango wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(intFila - 1, 10)), "Multiple"
        Else
            RecuadrarRango wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 10)), "Multiple"
        End If
    End If
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarExcelMaterialesLI ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GenerarExcelSubContratacionesLI( _
                                                ByRef wbHoja As excel.Worksheet, _
                                                ByRef p_ObjColSobContratacionesDeFactura As Scripting.Dictionary, _
                                                Optional ByRef p_Error As String _
                                                ) As String
       
    Dim intFila As Integer
    Dim lngNumeroRegistro As Long
    Dim lngNumerRegistrosTotales As Long
    Dim m_IDSubContratacion As Variant
    Dim m_ObjSubContratacion As Subcontratacion
    
    On Error GoTo errores
    
    intFila = 1
    With wbHoja
        .Cells(intFila, 1).Value = "FECHA INICIAL"
        .Cells(intFila, 2).Value = "FECHA FINAL"
        .Cells(intFila, 3).Value = "EMPRESA"
        .Cells(intFila, 4).Value = "DESCRIPCIÓN"
    End With
    With wbHoja
       .Range(.Cells(1, 1), .Cells(1, 4)).Font.Bold = True
       .Range(.Cells(1, 1), .Cells(1, 4)).Interior.ColorIndex = 15
       .Range(.Cells(1, 1), .Cells(1, 4)).HorizontalAlignment = xlCenter
    End With
    If Not p_ObjColSobContratacionesDeFactura Is Nothing Then
        If p_ObjColSobContratacionesDeFactura.count > 0 Then
            
            lngNumerRegistrosTotales = p_ObjColSobContratacionesDeFactura.count
            lngNumeroRegistro = 1
            intFila = intFila + 1
            For Each m_IDSubContratacion In p_ObjColSobContratacionesDeFactura.Keys
                Set m_ObjSubContratacion = p_ObjColSobContratacionesDeFactura(m_IDSubContratacion)
                m_Linea = "LI Excel .....Subcontrataciones ( " & lngNumeroRegistro & "\" & lngNumerRegistrosTotales & ")"""
                Avance m_Linea
                If m_ObjSubContratacion.FechaAlta <> "" Then
                    wbHoja.Cells(intFila, 1).Value = CDate(Format(m_ObjSubContratacion.FechaAlta, "dd/mm/yyyy"))
                     wbHoja.Range(wbHoja.Cells(intFila, 1), wbHoja.Cells(intFila, 1)).NumberFormat = "m/d/yyyy"
                End If
                If m_ObjSubContratacion.FechaFin <> "" Then
                    wbHoja.Cells(intFila, 2).Value = CDate(Format(m_ObjSubContratacion.FechaFin, "dd/mm/yyyy"))
                    wbHoja.Range(wbHoja.Cells(intFila, 2), wbHoja.Cells(intFila, 2)).NumberFormat = "m/d/yyyy"
                End If
                If m_ObjSubContratacion.Empresa <> "" Then
                    wbHoja.Cells(intFila, 3).Value = m_ObjSubContratacion.Empresa
                End If
                If m_ObjSubContratacion.Descripcion <> "" Then
                    wbHoja.Cells(intFila, 4).Value = m_ObjSubContratacion.Descripcion
                End If
                Set m_ObjSubContratacion = Nothing
                intFila = intFila + 1
                lngNumeroRegistro = lngNumeroRegistro + 1
            Next
            RecuadrarRango wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(intFila - 1, 4)), "Multiple"
        Else
            RecuadrarRango wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 4)), "Multiple"
        End If
    End If
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarExcelSubContratacionesLI ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function AnexoAutorizacionEnMaterialesPorFactura( _
                                                            p_URLDocumento As String, _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
   
    Dim rcdDatos As DAO.Recordset
    Dim m_ObjColMaterialesEnFactura As Scripting.Dictionary
    Dim m_ObjColAnexosMaterialesDeFactura As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjMaterial As Material
    Dim m_ObjAnexo As Anexo
    
    On Error GoTo errores
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "Se ha de indicar la factura"
        Err.Raise 1000
    End If
    If Not fso.FileExists(p_URLDocumento) Then
        p_Error = "Se ha de indicar el Anexo"
        Err.Raise 1000
    Else
        If FicheroAbierto(p_URLDocumento) Then
            p_Error = "Ha de cerrar el anexo"
            Err.Raise 1000
        End If
    End If
    m_Linea = "Obteniendo la lista de Materiales............"
    Avance m_Linea
    Set m_ObjColMaterialesEnFactura = Constructor.getMaterialesEnFactura(Me.IDFactura, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColMaterialesEnFactura Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjColMaterialesEnFactura
        Set m_ObjMaterial = m_ObjColMaterialesEnFactura(m_ID)
        Set m_ObjColAnexosMaterialesDeFactura = Constructor.getAnexosAutorizacionMateriales(m_ObjMaterial.IDMaterial, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjColAnexosMaterialesDeFactura Is Nothing Then
            Set m_ObjAnexo = New Anexo
            m_ObjAnexo.IDMaterial = m_ObjMaterial.IDMaterial
            m_ObjAnexo.RegistrarAutorizacionMaterial p_URLDocumento, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjAnexo = Nothing
        End If
        Set m_ObjColAnexosMaterialesDeFactura = Nothing
        Set m_ObjMaterial = Nothing
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.AnexoAutorizacionEnMaterialesPorFactura ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
   
End Function


Public Function EliminarAnexosDeAutorizacionEnMaterialesPorFactura( _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As String
    
    
    Dim m_colMaterialesConAnexo As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjAnexo As Anexo
    
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) Then
        p_Error = "Se ha de indicar la factura"
        Err.Raise 1000
    End If
    Set m_colMaterialesConAnexo = Constructor.getAnexosAutorizacionMaterialesEnFactura(Me.IDFactura, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_colMaterialesConAnexo Is Nothing Then
        Exit Function
    End If
    
    For Each m_ID In m_colMaterialesConAnexo
        Set m_ObjAnexo = m_colMaterialesConAnexo(m_ID)
        m_ObjAnexo.Eliminar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjAnexo = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.EliminarAnexosDeAutorizacionEnMaterialesPorFactura ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function GenerarPartesDeFactura( _
                                        p_URLDirectorio As String, _
                                        Optional ByRef p_Error As String) As String
    
    Dim m_objParte As Parte
    Dim m_ObjColEventosPorBUIEnFactura As Scripting.Dictionary
    Dim m_BUI As Variant
    
    Dim VarItem As Variant
    Dim m_ObjColEventosPorBUI As Scripting.Dictionary
    
    On Error GoTo errores
    
    
    If Not fso.FolderExists(p_URLDirectorio) Then
        p_Error = "No se alcanza el directorio"
        Err.Raise 1000
    End If
    If Right(p_URLDirectorio, 1) <> "\" Then
        p_URLDirectorio = p_URLDirectorio & "\"
    End If
    m_Linea = "Eventos por BUI ....."
    Avance m_Linea

    Set m_ObjColEventosPorBUIEnFactura = Constructor.getEventosPorBUIEnFactura(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColEventosPorBUIEnFactura Is Nothing Then
        p_Error = "No se ha encontrado ningún evento por centro de la factura"
        Err.Raise 1000
    End If
    For Each m_BUI In m_ObjColEventosPorBUIEnFactura
        'm_ObjColEventosPorBUIEnFactura.Add m_BUI , m_ObjColEventosBUI
        m_Linea = "Eventos por BUI ....." & m_BUI
        Avance m_Linea

        Set m_ObjColEventosPorBUI = m_ObjColEventosPorBUIEnFactura(m_BUI)
        Set m_objParte = New Parte
        With m_objParte
            .Centro = CStr(m_BUI)
            Set .colEventos = m_ObjColEventosPorBUI
            .IDFactura = Me.IDFactura
            Set .Factura = Me
            .Alta m_ObjColEventosPorBUI, Me.IDFactura, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End With
        Set m_objParte = Nothing
        Set m_ObjColEventosPorBUI = Nothing
    Next
    
    GenerarWordsInformeAveriasEnCarpeta p_URLDirectorio, m_ObjColEventosPorBUIEnFactura, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarPartesDeFactura ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function InformeRAC( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_ObjEvento As Evento
    Dim m_IDEvento As Variant
    Dim m_IntFilaFinalActividades As Integer
    Dim m_intFilaFinalMateriales As Integer
    Dim m_ColActividades As Scripting.Dictionary
    Dim m_ColMateriales As Scripting.Dictionary
    Dim m_ColMaterialesTotales As Scripting.Dictionary
    Dim m_IDMaterial As Variant
    Dim m_ObjMaterial As Material
    Dim m_IDActividad As Variant
    Dim m_ObjActividad As Actividad
    Dim intNumero As Integer
    Dim intNumeroTotal As Integer
    Dim m_URLInforme As String
    Dim m_ColEventos As Scripting.Dictionary
    On Error GoTo errores
    
    p_Error = ""
    Set m_ColEventos = Constructor.getEventosEnFactura(Me.IDFactura, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColEventos Is Nothing Then
        Exit Function
    End If
    m_Linea = "Regularización de eventos Franqueados sin fecha Fin"
    Avance m_Linea
    m_Linea = "Abriendo excel"
    Avance m_Linea
    
    m_URLInforme = AbrirNuevoExcel(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    intNumero = 1
    intNumeroTotal = m_ColEventos.count
    'copiamos para cada evento
    For Each m_IDEvento In m_ColEventos
        'If m_IDEvento = "SAN2008009" Then Stop
        Set m_ObjEvento = m_ColEventos(m_IDEvento)
        'Debug.Print m_IDEvento, intNumero, intNumeroTotal
        m_AppExcel.Workbooks(m_ObjPlantilla.Name).Activate
        m_Linea = "Copiando plantila de excel para " & CStr(m_IDEvento) & " (" & intNumero & "\" & intNumeroTotal & ")"
        Avance m_Linea
        
       ' m_ObjPlantilla.Worksheets(1).Select
        m_AppExcel.Workbooks(m_ObjPlantilla.Name).Worksheets(1).Copy before:=m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1)
        
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Name = m_ObjEvento.IDEVENTO
        RellenarDatosEvento m_ObjEvento, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ColActividades = m_ObjEvento.ColActividades
        p_Error = m_ObjEvento.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        m_Linea = "Rellenando Actividades para " & CStr(m_IDEvento)
        Avance m_Linea
        RellenarActividades m_ColActividades, m_IntFilaFinalActividades, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColActividades Is Nothing Then
            If m_ColActividades.count > 0 Then
                For Each m_IDActividad In m_ColActividades.Keys
                    Set m_ObjActividad = m_ColActividades(m_IDActividad)
                    Set m_ColMateriales = m_ObjActividad.ColMateriales
                    p_Error = m_ObjActividad.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_ColMateriales Is Nothing Then
                        If m_ColMateriales.count > 0 Then
                            For Each m_IDMaterial In m_ColMateriales.Keys
                                If Nz(m_IDMaterial, "") <> "" Then
                                    Set m_ObjMaterial = m_ColMateriales(m_IDMaterial)
                                    If m_ColMaterialesTotales Is Nothing Then
                                        Set m_ColMaterialesTotales = New Scripting.Dictionary
                                        m_ColMaterialesTotales.Add m_IDMaterial & "|" & m_ObjActividad.IDActividad, m_ObjMaterial
                                    Else
                                        If Not m_ColMaterialesTotales.Exists(m_IDMaterial) Then
                                            m_ColMaterialesTotales.Add m_IDMaterial & "|" & m_ObjActividad.IDActividad, m_ObjMaterial
                                        End If
                                    End If
                                    Set m_ObjMaterial = Nothing
                                End If
                            Next
                        End If
                    End If
                    Set m_ObjActividad = Nothing
                Next
            End If
            
        End If
        m_Linea = "Rellenando Materiales para " & CStr(m_IDEvento)
        Avance m_Linea
       
        RellenarMateriales m_ColMaterialesTotales, m_IntFilaFinalActividades, m_intFilaFinalMateriales, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        'Debug.Print m_IDEvento
        VBA.DoEvents
        intNumero = intNumero + 1
    Next
    m_Linea = "Formateando Excel"
    Avance m_Linea
    PonerBordeFinal m_intFilaFinalMateriales, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    GeneraHojaResumenInformeRAC m_AppExcel.Workbooks(m_ObjLibro.Name), m_ColEventos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    PonerEnlaceAResumen m_AppExcel.Workbooks(m_ObjLibro.Name), m_ColEventos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_AppExcel.Workbooks(m_ObjPlantilla.Name).Close False
    Set m_ObjPlantilla = Nothing
    m_AppExcel.Workbooks(m_ObjLibro.Name).SaveAs m_URLInforme
    m_AppExcel.Workbooks(m_ObjLibro.Name).Close False
    Set m_ObjLibro = Nothing
    m_AppExcel.Quit
    Set m_AppExcel = Nothing
    InformeRAC = m_URLInforme
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método InformeRAC ha devuelto: " & vbNewLine & Err.Description
    End If
    If Not m_AppExcel Is Nothing Then
        If Not m_ObjHoja Is Nothing Then
            Set m_ObjHoja = Nothing
        End If
        If Not m_ObjLibro Is Nothing Then
            m_ObjLibro.Close False
            Set m_ObjLibro = Nothing
        End If
        If Not m_ObjHojaPlantilla Is Nothing Then
            Set m_ObjHojaPlantilla = Nothing
        End If
        If Not m_ObjPlantilla Is Nothing Then
            m_ObjPlantilla.Close False
            Set m_ObjPlantilla = Nothing
        End If
        m_AppExcel.Quit
        Set m_AppExcel = Nothing
    End If
    
End Function

Private Function RellenarDatosEvento( _
                                    p_Evento As Evento, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim m_int As Integer
    Dim m_NombreCampoParaInforme As String
    Dim m_NombreCampoParaDato As String
    Dim m_intFilaTitular As Integer
    Dim m_IntFila As Integer
    Dim m_intFilaInicial As Integer
    Dim m_intFilaFinal As Integer
    
    Dim m_ObjRango As excel.Range
    Dim m_ValorEnRango As String
    Dim m_ValorDato As Variant
    Dim m_IDNombreCampo As Long
    On Error GoTo errores
    
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    
    m_intFilaTitular = 2
    Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("G2")
    m_ValorDato = p_Evento.IDEVENTO
    p_Error = p_Evento.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjRango.Value = m_ValorDato
    
    Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("C9")
    m_ValorDato = p_Evento.HorasLabTotales
    p_Error = p_Evento.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjRango.Value = m_ValorDato
    
    Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("H9")
    m_ValorDato = p_Evento.HorasExtTotales
    p_Error = p_Evento.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjRango.Value = m_ValorDato
    
    
    m_intFilaInicial = m_intFilaTitular + 2
    m_intFilaFinal = 8
    
    For m_IntFila = m_intFilaInicial To m_intFilaFinal
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFila))
        m_ValorEnRango = m_ObjRango.Value
        m_ValorEnRango = Replace(m_ValorEnRango, "{{", "")
        m_ValorEnRango = Replace(m_ValorEnRango, "}}", "")
        If IsNumeric(m_ValorEnRango) Then
            
            m_IDNombreCampo = CLng(m_ValorEnRango)
            m_NombreCampoParaInforme = m_ColParaEnumNombreCamposEnInforme(CStr(m_IDNombreCampo))
            dato = Split(m_NombreCampoParaInforme, ".")
            m_ObjRango.Value = dato(1)
            m_NombreCampoParaDato = m_ColParaNombreCamposInformeTabla(m_NombreCampoParaInforme)
            'If m_NombreCampoParaDato = "NombreEquipo" Then Stop
            dato = Split(m_NombreCampoParaDato, ".")
            m_NombreCampoParaDato = dato(1)
            'If m_NombreCampoParaDato = "NombreEquipo" Then Stop
            m_ValorDato = p_Evento.SetPropiedad(m_NombreCampoParaDato, m_ValorDato, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("C" & CStr(m_IntFila))
            If m_NombreCampoParaDato = "FECHAALTAEVENTO" Or m_NombreCampoParaDato = "FechaFinalCalculada" Or m_NombreCampoParaDato = "FechaFinal" Then
                If IsDate(m_ValorDato) Then
                    m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                    m_ObjRango.Value = m_ValorDato
                    m_ObjRango.NumberFormat = "dd/mm/yyyy"
                End If
            Else
                If m_NombreCampoParaDato = "PMPR" Then
                    m_ValorDato = "'" & m_ValorDato
                
                End If
                m_ObjRango.Value = m_ValorDato
            End If
            
            
        End If
        
    Next
    For m_IntFila = m_intFilaInicial To m_intFilaFinal
        
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("F" & CStr(m_IntFila))
        m_ValorEnRango = m_ObjRango.Value
        m_ValorEnRango = Replace(m_ValorEnRango, "{{", "")
        m_ValorEnRango = Replace(m_ValorEnRango, "}}", "")
        If IsNumeric(m_ValorEnRango) Then
            
            m_IDNombreCampo = CLng(m_ValorEnRango)
            m_NombreCampoParaInforme = m_ColParaEnumNombreCamposEnInforme(CStr(m_IDNombreCampo))
            dato = Split(m_NombreCampoParaInforme, ".")
            m_ObjRango.Value = dato(1)
            m_NombreCampoParaDato = m_ColParaNombreCamposInformeTabla(m_NombreCampoParaInforme)
            'If m_NombreCampoParaDato = "TbEventos.FranqueadoParaInforme" Then Stop
            dato = Split(m_NombreCampoParaDato, ".")
            m_NombreCampoParaDato = dato(1)
            'If p_Evento = "SAN2103019" And m_NombreCampoParaDato = "FechaFinal" Then Stop
            m_ValorDato = p_Evento.getPropiedad(m_NombreCampoParaDato, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("H" & CStr(m_IntFila))
            If m_NombreCampoParaDato = "FECHAALTAEVENTO" Or m_NombreCampoParaDato = "FechaFinalCalculada" Or m_NombreCampoParaDato = "FechaFinal" Then
                If IsDate(m_ValorDato) Then
                    m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                    m_ObjRango.Value = m_ValorDato
                    m_ObjRango.NumberFormat = "dd/mm/yyyy"
                End If
            Else
                m_ObjRango.Value = m_ValorDato
            End If
            
            
        End If
        
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarDatosEvento ha devuelto: " & vbNewLine & Err.Description
    End If
    
    
End Function
Private Function PonerEnlaceAResumen( _
                                    p_ObjLibro As excel.Workbook, _
                                    p_ObjColEventos As Scripting.Dictionary, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    
    Dim m_IDEvento As Variant
    Dim m_NombreHoja As String
    
    On Error GoTo errores
    
    If p_ObjColEventos Is Nothing Then
        Exit Function
    End If
    For Each m_IDEvento In p_ObjColEventos.Keys
        m_NombreHoja = CStr(m_IDEvento)
        m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(m_NombreHoja).Hyperlinks.Add m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(m_NombreHoja).Range("B1:B1"), "", m_AppExcel.Workbooks(p_ObjLibro.Name).Sheets(1).Name & "!A1", , "Ir a Resumen"
    Next
        
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PonerEnlaceAResumen ha devuelto: " & vbNewLine & Err.Description
    End If
    
    
End Function

Private Function RellenarActividades( _
                                        p_ColActividades As Scripting.Dictionary, _
                                        p_IntFilaUltimaActividad As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_ColColumnasNombreCampos As Scripting.Dictionary
    Dim m_ColNombreCamposLetraColumna As Scripting.Dictionary
    Dim m_IntFila As Integer
    Dim m_IDActividad As Variant
    Dim m_ObjActividad As Actividad
    Dim m_NombreCampoParaInformeCompleto As String
    Dim m_NombreCampoParaInforme As String
    Dim m_NombreCampoParaDatoCompleto As String
    Dim m_NombreCampoParaDato As String
    Dim m_ObjRango As excel.Range
    Dim m_ValorEnRango As String
    Dim m_ValorDato As Variant
    Dim m_NumeroActividad As Integer
    Dim m_Letra As Variant
    Dim m_Resto As String
    Dim m_intFilaInicial As Integer
    Dim m_intFilaTitulosActividades As Integer
    Dim m_intNumeroActividades As Integer
    Dim m_inti As Integer
    Dim m_IntFilaPrimeraActividad As Integer
    Dim m_IntFilaFinalActividad As Integer
    On Error GoTo errores
    
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    Set m_ColNombreCamposLetraColumna = New Scripting.Dictionary
    With m_ColNombreCamposLetraColumna
        .Add "B", m_ColParaEnumNombreCamposEnInforme("29") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("29"))
        .Add "D", m_ColParaEnumNombreCamposEnInforme("13") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("13"))
        .Add "E", m_ColParaEnumNombreCamposEnInforme("16") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("16"))
        .Add "F", m_ColParaEnumNombreCamposEnInforme("11") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("11"))
        .Add "G", m_ColParaEnumNombreCamposEnInforme("14") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("14"))
        .Add "H", m_ColParaEnumNombreCamposEnInforme("15") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("15"))
        .Add "I", m_ColParaEnumNombreCamposEnInforme("12") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("12"))
        .Add "J", m_ColParaEnumNombreCamposEnInforme("17") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("17"))
    End With
    m_intFilaTitulosActividades = 12
    'titulares
    For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_intFilaTitulosActividades))
        m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
        dato = Split(m_Resto, "|")
        m_NombreCampoParaInformeCompleto = dato(0)
        If InStr(1, m_NombreCampoParaInformeCompleto, ".") <> 0 Then
            dato = Split(m_NombreCampoParaInformeCompleto, ".")
            m_NombreCampoParaInforme = dato(1)
        Else
            m_NombreCampoParaInforme = m_NombreCampoParaInformeCompleto
        End If
        m_ObjRango.Value = m_NombreCampoParaInforme
    Next
    If p_ColActividades Is Nothing Then
        p_IntFilaUltimaActividad = m_intFilaTitulosActividades + 1
        Exit Function
    End If
    If p_ColActividades.count = 0 Then
        p_IntFilaUltimaActividad = m_intFilaTitulosActividades + 1
        Exit Function
    End If
    If p_ColActividades.count > 1 Then
        GenerarFilasDeActividades p_ColActividades, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_IntFila = m_intFilaTitulosActividades + 1
    m_NumeroActividad = 1
    m_IntFilaPrimeraActividad = m_IntFila
    For Each m_IDActividad In p_ColActividades.Keys
        m_Linea = "Rellenando Actividad: " & CStr(m_IDActividad)
        Avance m_Linea
        Set m_ObjActividad = p_ColActividades(m_IDActividad)
        For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
            Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_IntFila))
            m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
            dato = Split(m_Resto, "|")
            m_NombreCampoParaDatoCompleto = dato(1)
            If InStr(1, m_NombreCampoParaDatoCompleto, ".") <> 0 Then
                dato = Split(m_NombreCampoParaDatoCompleto, ".")
                m_NombreCampoParaDato = dato(1)
            Else
                m_NombreCampoParaDato = m_NombreCampoParaDatoCompleto
            End If
            
            m_ValorDato = m_ObjActividad.getPropiedad(m_NombreCampoParaDato, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_NombreCampoParaDato = "FechaAlta" Then
                If IsDate(m_ValorDato) Then
                    m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                    m_ObjRango.Value = m_ValorDato
                    m_ObjRango.NumberFormat = "dd/mm/yyyy"
                End If
            Else
                m_ObjRango.Value = m_ValorDato
            End If
        Next
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).EntireRow.AutoFit
        
        If m_NumeroActividad <> p_ColActividades.count Then
            m_IntFila = m_IntFila + 1
        End If
        Set m_ObjActividad = Nothing
        m_NumeroActividad = m_NumeroActividad + 1
    Next
    m_IntFilaFinalActividad = m_IntFila
    If p_ColActividades.count > 1 Then
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaFinalActividad) & ":M" & CStr(m_IntFilaFinalActividad)).Copy
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaPrimeraActividad) & _
                            ":M" & CStr(m_IntFilaFinalActividad - 1)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                            SkipBlanks:=False, Transpose:=False
        m_AppExcel.CutCopyMode = False
    End If
    
    p_IntFilaUltimaActividad = m_IntFilaFinalActividad
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarActividades ha devuelto: " & vbNewLine & Err.Description
    End If
    
    
End Function
Private Function GenerarFilasDeActividades( _
                                            ByRef p_ColActividades As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_IntFila As Integer
    Dim m_intFilaTitulosActividades As Integer
    Dim m_intNumeroActividades As Integer
    Dim m_inti As Integer
    On Error GoTo errores
    
    m_intFilaTitulosActividades = 11
    m_IntFila = m_intFilaTitulosActividades + 1
    'AHORA SE HA DE COPIAR Y PEGAR TANTAS VECES LA FILA 12 PARA CADA ACTIVIDAD QUE HAYA MENOS UNA
    m_intNumeroActividades = p_ColActividades.count
    If m_intNumeroActividades > 1 Then
        For m_inti = 1 To m_intNumeroActividades - 1
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).Copy
            
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila + 1) & ":" & CStr(m_IntFila + 1)).Insert Shift:=xlDown
            'Selection.Insert Shift:=xlDown
            m_IntFila = m_IntFila + 1
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarFilasDeActividades ha devuelto: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function GenerarFilasDeMateriales( _
                                            ByRef p_ColMateriales As Scripting.Dictionary, _
                                            P_IntFilaMaterialesTitulares As Integer, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_IntFila As Integer
    Dim m_intFilaTitulosMateriales As Integer
    Dim m_intNumeroMateriales As Integer
    Dim m_inti As Integer
    
    On Error GoTo errores
    
    m_IntFila = P_IntFilaMaterialesTitulares + 1
    'AHORA SE HA DE COPIAR Y PEGAR TANTAS VECES LA FILA 12 PARA CADA Material QUE HAYA MENOS UNA
    m_intNumeroMateriales = p_ColMateriales.count
    If m_intNumeroMateriales > 1 Then
        For m_inti = 1 To m_intNumeroMateriales - 1
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).Copy
            m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila + 1) & ":" & CStr(m_IntFila + 1)).Insert Shift:=xlDown
            
            m_IntFila = m_IntFila + 1
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarFilasDeMateriales ha devuelto: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function RellenarMateriales( _
                                        p_ColMateriales As Scripting.Dictionary, _
                                        p_IntFilaUltimaMaterial As Integer, _
                                        ByRef p_intFilaUltimaMateriales As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    Dim m_ColColumnasNombreCampos As Scripting.Dictionary
    Dim m_ColNombreCamposLetraColumna As Scripting.Dictionary
    Dim m_IntFila As Integer
    Dim m_IDMaterial As String
    Dim m_VarITem As Variant
    
    Dim m_ObjMaterial As Material
    Dim m_NombreCampoParaInformeCompleto As String
    Dim m_NombreCampoParaInforme As String
    Dim m_NombreCampoParaDatoCompleto As String
    Dim m_NombreCampoParaDato As String
    Dim m_ObjRango As excel.Range
    Dim m_ValorEnRango As String
    Dim m_ValorDato As Variant
    Dim m_NumeroMaterial As Integer
    Dim m_Letra As Variant
    Dim m_Resto As String
    Dim m_intFilaInicial As Integer
    Dim dato1
    Dim m_IntFilaTitularesMateriales As Integer
    Dim m_IntFilaPrimeraMaterial As Integer
    Dim m_IntFilaFinalMaterial As Integer
    On Error GoTo errores
    
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    Set m_ColNombreCamposLetraColumna = New Scripting.Dictionary
    With m_ColNombreCamposLetraColumna
        .Add "B", m_ColParaEnumNombreCamposEnInforme("29") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("29"))
        .Add "C", m_ColParaEnumNombreCamposEnInforme("30") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("30"))
        .Add "D", m_ColParaEnumNombreCamposEnInforme("18") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("18"))
        .Add "E", m_ColParaEnumNombreCamposEnInforme("21") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("21"))
        .Add "F", m_ColParaEnumNombreCamposEnInforme("22") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("22"))
        .Add "G", m_ColParaEnumNombreCamposEnInforme("23") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("23"))
        .Add "H", m_ColParaEnumNombreCamposEnInforme("19") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("19"))
        .Add "I", m_ColParaEnumNombreCamposEnInforme("26") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("26"))
        .Add "J", m_ColParaEnumNombreCamposEnInforme("20") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("20"))
        .Add "L", m_ColParaEnumNombreCamposEnInforme("24") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("24"))
        .Add "M", m_ColParaEnumNombreCamposEnInforme("25") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("25"))
    End With
    
    
    m_IntFilaTitularesMateriales = p_IntFilaUltimaMaterial + 3
    'titulares
    For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
        Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_IntFilaTitularesMateriales))
        m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
        dato = Split(m_Resto, "|")
        m_NombreCampoParaInformeCompleto = dato(0)
        If InStr(1, m_NombreCampoParaInformeCompleto, ".") <> 0 Then
            dato = Split(m_NombreCampoParaInformeCompleto, ".")
            m_NombreCampoParaInforme = dato(1)
        Else
            m_NombreCampoParaInforme = m_NombreCampoParaInformeCompleto
        End If
        m_ObjRango.Value = m_NombreCampoParaInforme
    Next
    If p_ColMateriales Is Nothing Then
        p_intFilaUltimaMateriales = m_IntFilaTitularesMateriales + 1
        Exit Function
    End If
    If p_ColMateriales.count = 0 Then
        p_intFilaUltimaMateriales = m_IntFilaTitularesMateriales + 1
        Exit Function
    End If
    If p_ColMateriales.count > 1 Then
        GenerarFilasDeMateriales p_ColMateriales, m_IntFilaTitularesMateriales, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_NumeroMaterial = 1
    m_IntFila = m_IntFilaTitularesMateriales + 1
    m_IntFilaPrimeraMaterial = m_IntFila
    For Each m_VarITem In p_ColMateriales.Keys
        dato = Split(m_VarITem, "|")
        m_IDMaterial = dato(0)
        
        If p_ColMateriales.Exists(m_VarITem) Then
            Set m_ObjMaterial = p_ColMateriales(m_VarITem)
            For Each m_Letra In m_ColNombreCamposLetraColumna.Keys
                Set m_ObjRango = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range(CStr(m_Letra) & CStr(m_IntFila))
                If m_Letra = "B" Then
                    m_ObjRango.Value = m_IDMaterial
                Else
                    m_Resto = m_ColNombreCamposLetraColumna(m_Letra)
                    dato1 = Split(m_Resto, "|")
                    m_NombreCampoParaDatoCompleto = dato1(1)
                    If InStr(1, m_NombreCampoParaDatoCompleto, ".") <> 0 Then
                        
                        dato1 = Split(m_NombreCampoParaDatoCompleto, ".")
                        m_NombreCampoParaDato = dato1(1)
                    Else
                        m_NombreCampoParaDato = m_NombreCampoParaDatoCompleto
                    End If
                    m_ValorDato = m_ObjMaterial.getPropiedad(m_NombreCampoParaDato, p_Error)
                   
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If m_NombreCampoParaDato = "FechaIncio" Or m_NombreCampoParaDato = "FechaEntrega" Then
                        If IsDate(m_ValorDato) Then
                            m_ValorDato = CDate(Format(m_ValorDato, "dd/mm/yyyy"))
                            m_ObjRango.Value = m_ValorDato
                            m_ObjRango.NumberFormat = "dd/mm/yyyy"
                        End If
                    Else
                        m_ObjRango.Value = m_ValorDato
                    End If
                End If
            Next
        End If
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Rows(CStr(m_IntFila) & ":" & CStr(m_IntFila)).EntireRow.AutoFit
        If m_NumeroMaterial <> p_ColMateriales.count Then
            m_IntFila = m_IntFila + 1
        End If
        Set m_ObjMaterial = Nothing
        m_NumeroMaterial = m_NumeroMaterial + 1
    Next
    m_IntFilaFinalMaterial = m_IntFila
    If p_ColMateriales.count > 1 Then
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaFinalMaterial) & ":M" & CStr(m_IntFilaFinalMaterial)).Copy
        m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1).Range("B" & CStr(m_IntFilaPrimeraMaterial) & _
                            ":M" & CStr(m_IntFilaFinalMaterial - 1)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                            SkipBlanks:=False, Transpose:=False
        m_AppExcel.CutCopyMode = False
    End If
    p_intFilaUltimaMateriales = m_IntFilaFinalMaterial
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarMateriales ha devuelto: " & m_ObjRango.Address & vbNewLine & Err.Description
    End If
    
    
    
End Function

Private Function PonerBordeFinal( _
                                    ByRef m_intFilaFinalMateriales, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_intFilaInicial As Integer
    Dim m_intFilaFinal As Integer
    Dim m_Hoja As excel.Worksheet
    Dim m_Rango As excel.Range
    
    On Error GoTo errores
    
    Set m_Hoja = m_AppExcel.Workbooks(m_ObjLibro.Name).Worksheets(1)
    m_intFilaInicial = 2
    
    m_intFilaFinal = m_intFilaFinalMateriales
    If m_intFilaFinalMateriales < m_intFilaInicial Then
        p_Error = "La fila final de los materiales es menor que la inicial"
        Err.Raise 1000
    End If
    
    Set m_Rango = m_Hoja.Range("B2:M" & CStr(m_intFilaFinal))
    With m_Rango
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    Set m_Rango = m_Hoja.Range("B2:M3")
    With m_Rango
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PonerBordeFinal ha devuelto: " & vbNewLine & Err.Description
    End If

End Function




Private Function GeneraHojaResumenInformeRAC( _
                                            p_ObjLibro As excel.Workbook, _
                                            ByRef p_ObjColEventos As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                            
    Dim m_IDEvento As Variant
    Dim m_ObjEvento As Evento
    
    Dim intFila As Integer
    Dim lngNumeroRegistro As Long
    Dim lngNumerRegistrosTotales As Long
    On Error GoTo errores
    
    If p_ObjColEventos Is Nothing Then
        Exit Function
    End If
    m_Linea = "Generando Hoja Resúmen"
    Avance m_Linea
    m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets.Add m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(1)
    m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets(1).Name = "Resumen"
    '-----------
    ' parte inicial de la cabecera
    '--------------
    intFila = 1
    With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
        .Cells(intFila, 1).Value = "PARTE TRABAJO"
        .Cells(intFila, 2).Value = "NODO"
        .Cells(intFila, 3).Value = "BUI"
        .Cells(intFila, 4).Value = "DESCRIPCIÓN"
        .Cells(intFila, 5).Value = "SUBSISTEMA"
        .Cells(intFila, 6).Value = "EQUIPO"
        .Cells(intFila, 7).Value = "PM/PR"
        .Cells(intFila, 8).Value = "FECHA INICIO"
        .Cells(intFila, 9).Value = "FECHA FIN"
        .Cells(intFila, 10).Value = "NºHORAS_LAB"
        .Cells(intFila, 11).Value = "NºHORAS_EXT"
    End With
    With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
        .Range(.Cells(1, 1), .Cells(1, 9)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 9)).Interior.ColorIndex = 15
        .Range(.Cells(1, 1), .Cells(1, 9)).HorizontalAlignment = xlCenter
    End With
    lngNumerRegistrosTotales = p_ObjColEventos.count
    lngNumeroRegistro = 1
    intFila = intFila + 1
    For Each m_IDEvento In p_ObjColEventos.Keys
        Set m_ObjEvento = p_ObjColEventos(m_IDEvento)
        m_Linea = "Evento " & m_ObjEvento.IDEVENTO & " ( " & lngNumeroRegistro & "\" & lngNumerRegistrosTotales & ")"
        Avance m_Linea
        
        
        
        With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
            If m_ObjEvento.IDEVENTO <> "" Then
                .Hyperlinks.Add Anchor:=.Range("A" & CStr(intFila)), _
                                     Address:="", SubAddress:=CStr(m_IDEvento) & "!A1", _
                                     TextToDisplay:=CStr(m_IDEvento)
             End If
             If m_ObjEvento.NODO <> "" Then
                 .Cells(intFila, 2).Value = m_ObjEvento.NODO
             End If
             If m_ObjEvento.BUI <> "" Then
                 .Cells(intFila, 3).Value = m_ObjEvento.BUI
             End If
             If m_ObjEvento.BUI <> "" Then
                 .Cells(intFila, 4).Value = m_ObjEvento.Descripcion
             End If
             If m_ObjEvento.SubSistema <> "" Then
                 .Cells(intFila, 5).Value = m_ObjEvento.SubSistema
             End If
             If m_ObjEvento.NombreEquipo <> "" Then
                 .Cells(intFila, 6).Value = m_ObjEvento.NombreEquipo
             End If
            
             If m_ObjEvento.PMPR <> "" Then
                 .Cells(intFila, 7).Value = "'" & m_ObjEvento.PMPR
             End If
             If m_ObjEvento.FECHAALTAEVENTO <> "" Then
                 .Cells(intFila, 8).Value = Format(m_ObjEvento.FECHAALTAEVENTO, "mm/dd/yyyy")
             End If
             If m_ObjEvento.FechaFinal <> "" Then
                 .Cells(intFila, 9).Value = Format(m_ObjEvento.FechaFinal, "mm/dd/yyyy")
             End If
             If m_ObjEvento.HorasLabTotales <> "" Then
                 .Cells(intFila, 10).Value = "'" & m_ObjEvento.HorasLabTotales
             End If
             If m_ObjEvento.HorasExtTotales <> "" Then
                 .Cells(intFila, 11).Value = "'" & m_ObjEvento.HorasExtTotales
             End If
         End With
        Set m_ObjEvento = Nothing
        intFila = intFila + 1
        lngNumeroRegistro = lngNumeroRegistro + 1
    Next
    With m_AppExcel.Workbooks(p_ObjLibro.Name).Worksheets("Resumen")
        RecuadrarRango .Range(.Cells(1, 1), .Cells(intFila - 1, 11)), "Multiple"
    
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
    End With
    
    
    GeneraHojaResumenInformeRAC = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Certificacion.GeneraHojaResumenInformeRAC ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function AbrirNuevoExcel( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim m_URLInforme As String
    Dim m_NombreInforme As String
   
    On Error GoTo errores
    
    If m_ObjEntorno.URLDirectorioLocal = "" Then
        p_Error = "No se ha definido el directorioLocal"
        Err.Raise 1000
    End If
    
    If Not fso.FileExists(m_ObjEntorno.URLPlantillaInforme) Then
        p_Error = "No se alcanza m_ObjEntorno.URLPlantillaInforme" & vbNewLine & m_ObjEntorno.URLPlantillaInforme
        Err.Raise 1000
    End If
    
    m_NombreInforme = "Informe de Partes de Trabajo.xlsx"
    
    m_URLInforme = m_ObjEntorno.URLDirectorioLocal & m_NombreInforme
    If fso.FileExists(m_URLInforme) Then
        If FicheroAbierto(m_URLInforme) Then
            p_Error = "Ha de cerrar el informe abierto"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLInforme, True
    End If
    
    If m_AppExcel Is Nothing Then
        Set m_AppExcel = New excel.Application
    End If
    m_AppExcel.Visible = False
    Set m_ObjPlantilla = m_AppExcel.Workbooks.Open(m_ObjEntorno.URLPlantillaInforme)
    Set m_ObjHojaPlantilla = m_ObjPlantilla.Worksheets(1)
    Set m_ObjLibro = m_AppExcel.Workbooks.Add
    AbrirNuevoExcel = m_URLInforme
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirNuevoExcel ha devuelto: " & vbNewLine & Err.Description
    End If
    On Error Resume Next
    If Not m_AppExcel Is Nothing Then
        If Not m_ObjHoja Is Nothing Then
            Set m_ObjHoja = Nothing
        End If
        If Not m_ObjLibro Is Nothing Then
            m_ObjLibro.Close False
            Set m_ObjLibro = Nothing
        End If
        m_AppExcel.Quit
        Set m_AppExcel = Nothing
    End If
    
End Function

Public Function GenerarWordsInformeAveriasEnCarpeta( _
                                                     m_URLCarpeta As String, _
                                                     Optional ByRef p_ObjColEventosPorBUI As Scripting.Dictionary, _
                                                    Optional ByRef p_Error As String _
                                                     ) As String
                
    Dim m_ID As Variant
    Dim objEvento As Evento
    Dim URLInformeGenerado As String
    Dim VarItem As Variant
    Dim m_BUI As Variant
    Dim m_FechaMinima As String
    Dim FechaMaxima As String
    Dim m_ObjColEventos As Scripting.Dictionary
    Dim flag As String
    Dim appWord As Word.Application
    Dim Plantilla As Word.Document
    On Error GoTo errores
    
    
    If Not fso.FolderExists(m_URLCarpeta) Then
        p_Error = "No es alcanzable la carpeta que ha definido el usuario para establecer los informes"
        Err.Raise 1000
    End If
    If p_ObjColEventosPorBUI Is Nothing Then
        Set p_ObjColEventosPorBUI = Constructor.getEventosPorBUIEnFactura(Me, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_Linea = "Generación de Words ....."
    Avance m_Linea
    If p_ObjColEventosPorBUI Is Nothing Then
        p_Error = "No hay eventos para el informe"
        Err.Raise 1000
    End If
    Set appWord = New Word.Application
    'abrimos la plantilla
    Set Plantilla = appWord.Documents.Open(m_ObjEntorno.URLPlantillaInformeAverias)
    For Each m_BUI In p_ObjColEventosPorBUI
        'm_ObjColEventosPorBUIEnFactura.Add m_BUI , m_ObjColEventosBUI
        m_Linea = "Generación de Words .....BUI " & m_BUI
        Avance m_Linea
        Set m_ObjColEventos = p_ObjColEventosPorBUI(m_BUI)
        flag = getFechaInicialYFinalDeColeEventos(m_ObjColEventos, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        dato = Split(flag, "|")
        m_FechaMinima = dato(0)
        FechaMaxima = dato(1)
        m_Linea = "Generando Word para el centro " & CStr(m_BUI)
        Avance m_Linea
        RellenarInformeEnWord appWord, Plantilla, m_ObjColEventos, CStr(m_BUI), m_URLCarpeta, m_FechaMinima, FechaMaxima, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjColEventos = Nothing
    Next
    appWord.Documents(m_ObjEntorno.URLPlantillaInformeAverias).Close False
    Set Plantilla = Nothing
    appWord.Quit
    Set appWord = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarInformeAverias ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    If Not appWord.Documents(m_ObjEntorno.URLPlantillaInformeAverias) Is Nothing Then
        appWord.Documents(m_ObjEntorno.URLPlantillaInformeAverias).Close False
        Set Plantilla = Nothing
    End If
    If Not appWord Is Nothing Then
        appWord.Quit
        Set appWord = Nothing
    End If
End Function



Public Function EliminarPartes(Optional ByRef p_Error As String) As String
    
    Dim m_objParte As Parte
    Dim m_ID As Variant
    Dim m_ObjColPartes As Scripting.Dictionary
    On Error GoTo errores
    Set m_ObjColPartes = Constructor.getPartesDeFactura(Me.IDFactura, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColPartes Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjColPartes
        Set m_objParte = m_ObjColPartes(m_ID)
        If fso.FileExists(m_objParte.URLParte) Then
            If FicheroAbierto(m_objParte.URLParte) Then
                p_Error = "Al menos un archivo de los partes está abierto"
                Err.Raise 1000
            End If
            
        End If
        m_objParte.Eliminar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_objParte = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarPartes ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function SalidaAExcel( _
                                Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim WbLibro As excel.Workbook
    Dim wbHoja As excel.Worksheet
    Dim m_URLFactura As String
    
    On Error GoTo errores
    
    
    If p_EsCalculado = 0 Then
        p_EsCalculado = EnumSino.Sí
    End If
    If Not IsNumeric(Me.IDFactura) And p_EsCalculado = EnumSino.Sí Then
        
        p_Error = "No se ha introducido un valor de factura válido"
        Err.Raise 1000
    End If
    m_URLFactura = Me.URLFactura
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If fso.FileExists(m_URLFactura) Then
        If FicheroAbierto(m_URLFactura) Then
            p_Error = "Ya existe una factura y parece estar abierta"
            Err.Raise 1000
        End If
    End If
    
    
    
    m_Linea = "Excel ..... Abriendo documento"
    Avance m_Linea
    '---------------
    ' crear el libro excel
    '-------------
    Set appexcel = New excel.Application
    appexcel.Visible = False
    Set WbLibro = appexcel.Workbooks.Add
    
    Set wbHoja = WbLibro.Worksheets.Add
    wbHoja.Name = "DETALLE GETRONICS"
    CreaHojaDetalleGetronics wbHoja, p_EsCalculado, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set wbHoja = WbLibro.Worksheets.Add
    wbHoja.Name = "MATERIALES"
    CreaHojaMateriales wbHoja, p_EsCalculado, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set wbHoja = WbLibro.Worksheets.Add
    wbHoja.Name = "ACTIVIDADES"
    CreaHojaActividades wbHoja, p_EsCalculado, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set wbHoja = WbLibro.Worksheets.Add
    wbHoja.Name = "EVENTOS"
    CreaHojaEventos wbHoja, p_EsCalculado, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set wbHoja = WbLibro.Worksheets.Add
    wbHoja.Name = "FACTURA"
    CreaHojaFactura wbHoja, p_EsCalculado, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    If fso.FileExists(m_URLFactura) Then
        fso.DeleteFile m_URLFactura, True
    End If
    '-------------
    ' GUARDAR el libro excel
    '-----------
    WbLibro.SaveAs m_URLFactura
    Set wbHoja = Nothing
    WbLibro.Close
    Set WbLibro = Nothing
    appexcel.Quit
    Set appexcel = Nothing
    
    
    SalidaAExcel = m_URLFactura
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Facturacion.SalidaAExcel ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not wbHoja Is Nothing Then
        Set wbHoja = Nothing
    End If
    If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not appexcel Is Nothing Then
       appexcel.Quit
        Set appexcel = Nothing
    End If
   
    
End Function
Public Function CreaHojaEventos( _
                                ByRef wbHoja As excel.Worksheet, _
                                Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ColEventos As Scripting.Dictionary
    Dim m_ObjEvento As Evento
    Dim m_ID As Variant
    Dim intFila As Integer
    Dim intFilaCabecera As Integer
    Dim intFilaInicio As Integer
    Dim intFilaFin As Integer
    
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) And p_EsCalculado = EnumSino.Sí Then
        p_Error = "No se ha introducido un IDFactura numérico existente válido"
        Err.Raise 1000
    End If
    m_Linea = "Excel Hoja Eventos..... "
    Avance m_Linea
    '------------------
    '   CABECERA DE LA HOJA EXCEL
    '---------------
    intFila = 1
    intFilaCabecera = intFila
    With wbHoja
        .Cells(1, 1).Value = "NODO"
        .Cells(1, 2).Value = "BUI"
        .Cells(1, 3).Value = "SUBSISTEMA"
        .Cells(1, 4).Value = "EQUIPO"
        .Cells(1, 5).Value = "ID_EVENTO (PT)"
        .Cells(1, 6).Value = "PM/PR"
        .Cells(1, 7).Value = "TIPO_EVENTO"
        .Cells(1, 8).Value = "CRITICIDAD"
        .Cells(1, 9).Value = "TÉCNICO"
        .Cells(1, 10).Value = "ORIGINADOR"
        .Cells(1, 11).Value = "FECHA_ALTA"
        .Cells(1, 12).Value = "HORA_ALTA"
        .Cells(1, 13).Value = "TIEMPO_RESPUESTA"
        .Cells(1, 14).Value = "FECHA_FIN"
        .Cells(1, 15).Value = "HORA_FIN"
        .Cells(1, 16).Value = "CAUSA_FIN"
        .Cells(1, 17).Value = "CONTACTO"
        .Cells(1, 18).Value = "EVENTO_GENERADO"
        .Cells(1, 19).Value = "DESCRIPCIÓN"
    End With
    With wbHoja
        .Range(.Cells(1, 1), .Cells(1, 19)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 19)).HorizontalAlignment = xlCenter
        .Range(.Cells(1, 1), .Cells(1, 19)).Interior.ColorIndex = 15
    End With
    '--------------
    ' recorrer la tabla TbFacturaEventosInvolucrados de los datos
    '--------------
    intFila = intFila + 1
    If p_EsCalculado = EnumSino.Sí Then
        Set m_ColEventos = Me.colListaEventos
    Else
        Set m_ColEventos = Me.colListaEventosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColEventos Is Nothing Then
        Exit Function
    End If
    intFila = intFila - 1
    intFilaInicio = intFila
    For Each m_ID In m_ColEventos
        intFila = intFila + 1
        Set m_ObjEvento = m_ColEventos(m_ID)
        m_Linea = "Excel Hoja Eventos..... " & m_ObjEvento.IDEVENTO
        Avance m_Linea
        With m_ObjEvento
            If .NODO <> "" Then
                wbHoja.Cells(intFila, 1) = .NODO
            End If
            If .BUI <> "" Then
                wbHoja.Cells(intFila, 2) = .BUI
            End If
            If .SubSistema <> "" Then
                wbHoja.Cells(intFila, 3) = .SubSistema
            End If
            If Not .Equipo Is Nothing Then
                If .Equipo.IDEquipo <> "" Then
                    wbHoja.Cells(intFila, 4) = .Equipo.Equipo
                End If
            End If
            If .IDEVENTO <> "" Then
                wbHoja.Cells(intFila, 5) = .IDEVENTO
            End If
            If .PMPR <> "" Then
                wbHoja.Cells(intFila, 6) = .PMPR
            End If
            If .TipoEvento <> "" Then
                wbHoja.Cells(intFila, 7) = .TipoEvento
            End If
            If .Criticidad <> "" Then
                wbHoja.Cells(intFila, 8) = .Criticidad
            End If
            If Not .Tecnico Is Nothing Then
                If .Tecnico.Tipo <> "" Then
                    wbHoja.Cells(intFila, 9) = .Tecnico.Tipo
                End If
            End If
            If .Originador <> "" Then
                wbHoja.Cells(intFila, 10) = .Originador
            End If
            If IsDate(.FECHAALTAEVENTO) Then
                wbHoja.Cells(intFila, 11) = CDate(Format(.FECHAALTAEVENTO, "dd/mm/yyyy"))
            End If
            If IsDate(.HORAINICIALEVENTO) Then
                wbHoja.Cells(intFila, 12) = CDate(Format(.HORAINICIALEVENTO, "hh:mm"))
            End If
            If IsDate(.TIEMPORESPUESTAEVENTO) Then
                wbHoja.Cells(intFila, 13) = CDate(Format(.TIEMPORESPUESTAEVENTO, "hh:mm"))
            End If
            If IsDate(.FechaFinal) Then
                wbHoja.Cells(intFila, 14) = CDate(Format(.FechaFinal, "dd/mm/yyyy"))
            End If
            If IsDate(.HORAFINALEVENTO) Then
                wbHoja.Cells(intFila, 15) = CDate(Format(.HORAFINALEVENTO, "hh:mm"))
            End If
            If .CausaFin <> "" Then
                wbHoja.Cells(intFila, 16) = .CausaFin
            End If
            If .Contacto <> "" Then
                wbHoja.Cells(intFila, 17) = .Contacto
            End If

            If .Descripcion <> "" Then
                wbHoja.Cells(intFila, 19) = .Descripcion
            End If
            
        End With
        Set m_ObjEvento = Nothing
    Next
    intFilaFin = intFila
   RecuadrarRango wbHoja.Range(wbHoja.Cells(intFilaCabecera, 1), wbHoja.Cells(intFilaFin, 19)), "Multiple"
        
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").EntireColumn.AutoFit
        .Columns("N:N").EntireColumn.AutoFit
        .Columns("O:O").EntireColumn.AutoFit
        .Columns("P:P").EntireColumn.AutoFit
        .Columns("Q:Q").EntireColumn.AutoFit
        .Columns("R:R").EntireColumn.AutoFit
        .Columns("S:S").ColumnWidth = 95
        .Columns("S:S").MergeCells = False
         .Columns("S:S").WrapText = True
    End With
    CreaHojaEventos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CreaHojaEventos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function CreaHojaActividades( _
                                    ByRef wbHoja As excel.Worksheet, _
                                    Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    
    Dim m_ColTipoTecnicos As Scripting.Dictionary
    Dim m_IDFacturaPerfil As Variant
    Dim m_ObjFacturaPerfiles As FacturaPerfiles
    
    Dim m_ObjColActividades As Scripting.Dictionary
    Dim m_IDActividad As Variant
    Dim m_ObjActividad As Actividad
    Dim m_ObjEvento As Evento
    Dim m_ObjEquipo As Equipo
    Dim m_Nodo As String
    Dim m_BUI As String
    Dim m_SUBSISTEMA As String
    Dim m_Equipo As String
    
    
    Dim intFila As Integer
    Dim intFilaCabecera As Integer
    Dim intFilaInicio As Integer
    Dim intFilaFin As Integer
    
    On Error GoTo errores
   
    If Not IsNumeric(Me.IDFactura) And p_EsCalculado = EnumSino.Sí Then
        p_Error = "No se ha introducido un IDFactura numérico existente válido"
        Err.Raise 1000
    End If
    
    m_Linea = "Excel Hoja Actividades....."
    Avance m_Linea
    If p_EsCalculado = EnumSino.Sí Then
        Set m_ColActividades = Me.colListaActividades
    Else
        Set m_ColActividades = Me.colListaActividadesPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColActividades Is Nothing Then
        Exit Function
    End If
    If p_EsCalculado = EnumSino.Sí Then
        Set m_ColTipoTecnicos = Me.ColPerfilesEnFactura
    Else
        Set m_ColTipoTecnicos = Me.ColPerfilesPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColTipoTecnicos Is Nothing Then
        Exit Function
    End If
    intFila = 1
    
    'SE HA DE RECORRE CADA UNO DE LOS PERFILES QUE TIENEN ACTIVIDADES
    For Each m_IDFacturaPerfil In m_ColTipoTecnicos
        With wbHoja
            .Cells(intFila, 1).Value = "NODO"
            .Cells(intFila, 2).Value = "BUI"
            .Cells(intFila, 3).Value = "SUBSISTEMA"
            .Cells(intFila, 4).Value = "ID_EQUIPO"
            .Cells(intFila, 5).Value = "ID_EVENTO (PT)"
            .Cells(intFila, 6).Value = "TÉCNICO"
            .Cells(intFila, 7).Value = "INICIALES_TÉCNICO"
            .Cells(intFila, 8).Value = "ACTIVIDAD"
            .Cells(intFila, 9).Value = "FECHA_ACTIVIDAD"
            .Cells(intFila, 10).Value = "HORAS_LABORABLES"
            .Cells(intFila, 11).Value = "HORAS_EXTRAS"
            .Cells(intFila, 12).Value = "UBICACIÓN"
            .Cells(intFila, 13).Value = "DESCRIPCIÓN"
        End With
        With wbHoja
            .Range(.Cells(intFila, 1), .Cells(intFila, 13)).HorizontalAlignment = xlCenter
            .Range(.Cells(intFila, 1), .Cells(intFila, 13)).Interior.ColorIndex = 15
            .Range(.Cells(intFila, 1), .Cells(intFila, 13)).Font.Bold = True
        End With
        Set m_ObjFacturaPerfiles = m_ColTipoTecnicos(m_IDFacturaPerfil)
        m_Linea = "Excel Hoja Actividades..... " & m_ObjFacturaPerfiles.Perfil
        Avance m_Linea
        intFilaCabecera = intFila
        For Each m_IDActividad In m_ColActividades
            Set m_ObjActividad = m_ColActividades(m_IDActividad)
            With m_ObjActividad
                If .Tecnico.Tipo <> m_ObjFacturaPerfiles.Perfil Then
                    GoTo siguienteActividad
                End If
                
                Set m_ObjEvento = m_ObjActividad.Evento
                If Not m_ObjEvento Is Nothing Then
                    m_Nodo = m_ObjEvento.NODO
                    m_BUI = m_ObjEvento.BUI
                    m_SUBSISTEMA = m_ObjEvento.SubSistema
                    Set m_ObjEquipo = m_ObjEvento.Equipo
                    If Not m_ObjEquipo Is Nothing Then
                        m_Equipo = m_ObjEquipo.Equipo
                    End If
                End If
                intFila = intFila + 1
                If intFilaInicio = 0 Then
                    intFilaInicio = intFila
                End If
                If m_Nodo <> "" Then
                    wbHoja.Cells(intFila, 1).Value = m_Nodo
                End If
                If m_BUI <> "" Then
                    wbHoja.Cells(intFila, 2).Value = m_BUI
                End If
                If m_SUBSISTEMA <> "" Then
                    wbHoja.Cells(intFila, 3).Value = m_SUBSISTEMA
                End If
                If m_Equipo <> "" Then
                    wbHoja.Cells(intFila, 4).Value = m_Equipo
                End If
                If .IDEVENTO <> "" Then
                    wbHoja.Cells(intFila, 5).Value = .IDEVENTO
                End If
                If .ALIASTECNICO <> "" Then
                    wbHoja.Cells(intFila, 6).Value = .ALIASTECNICO
                End If
                If m_ObjFacturaPerfiles.Perfil <> "" Then
                    wbHoja.Cells(intFila, 7).Value = m_ObjFacturaPerfiles.Perfil
                End If
                If .Actividad <> "" Then
                    wbHoja.Cells(intFila, 8).Value = .Actividad
                End If
                If IsDate(.FechaAlta) Then
                    wbHoja.Cells(intFila, 9) = CDate(Format(.FechaAlta, "dd/mm/yyyy"))
                End If
                If IsNumeric(.HorasLaborables) Then
                    wbHoja.Cells(intFila, 10) = CDbl(.HorasLaborables)
                End If
                If IsNumeric(.HorasExtras) Then
                    wbHoja.Cells(intFila, 11) = CDbl(.HorasExtras)
                End If
                If .Ubicacion <> "" Then
                    wbHoja.Cells(intFila, 12).Value = .Ubicacion
                End If
                If .Descripcion <> "" Then
                    wbHoja.Cells(intFila, 13).Value = .Descripcion
                End If
                
                
            End With
siguienteActividad:
            Set m_ObjActividad = Nothing
        Next
        intFilaFin = intFila
        intFila = intFila + 1
        With wbHoja
            RecuadrarRango .Range(.Cells(intFilaCabecera, 1), .Cells(intFilaFin, 13)), "Multiple"
        End With
        '-------------
        ' parte de HORAS PERFIL
        '-------------
        With wbHoja
            .Cells(intFila, 8).Value = "HORAS PERFIL " & UCase(m_ObjFacturaPerfiles.Perfil)
            .Cells(intFila, 8).Font.Bold = True
            .Cells(intFila, 8).HorizontalAlignment = xlCenter
            .Range(.Cells(intFila, 10), .Cells(intFila, 10)).Formula = "=SUM(J" & CStr(intFilaInicio) & ":J" & CStr(intFilaFin) & ")"
            .Cells(intFila, 10).Font.Bold = True
            .Range(.Cells(intFila, 11), .Cells(intFila, 11)).Formula = "=SUM(K" & CStr(intFilaInicio) & ":K" & CStr(intFilaFin) & ")"
            .Cells(intFila, 11).Font.Bold = True
        End With
        Set m_ObjFacturaPerfiles = Nothing
        intFilaInicio = 0
        intFila = intFila + 2
    Next
    
    
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").ColumnWidth = 95
        .Columns("M:M").MergeCells = False
        .Columns("M:M").WrapText = True
    End With

    CreaHojaActividades = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CreaHojaActividades ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function CreaHojaMateriales( _
                                    ByRef wbHoja As excel.Worksheet, _
                                    Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim intFila As Integer
    Dim intFilaCabecera As Integer
    Dim intFilaInicio As Integer
    Dim intFilaFin As Integer
    
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjMaterial As Material
    Dim m_ObjActividad As Actividad
    Dim m_ObjEvento As Evento
    Dim m_ObjEquipo As Equipo
    
    Dim m_ObjSubContratacion As Subcontratacion
    Dim m_ObjGasto As Gasto
    
    
    On Error GoTo errores
    
    If Not IsNumeric(Me.IDFactura) And p_EsCalculado = EnumSino.Sí Then
        p_Error = "No se ha introducido un IDFactura numérico existente válido"
        Err.Raise 1000
    End If
    
    m_Linea = "Excel Hoja Materiales....."
    Avance m_Linea
    intFila = 1
    With wbHoja
        .Cells(intFila, 1).Value = "MATERIALES"
        .Range(.Cells(intFila, 1), .Cells(intFila, 1)).Font.Bold = True
    End With
   
    intFila = intFila + 2
    intFilaCabecera = intFila
    '-------------------
    '  CABECERA DE LOS MATERIALES
    '--------------------
    With wbHoja
        .Cells(intFila, 1).Value = "NODO"
        .Cells(intFila, 2).Value = "BUI"
        .Cells(intFila, 3).Value = "SUBSISTEMA"
        .Cells(intFila, 4).Value = "ID_EQUIPO"
        .Cells(intFila, 5).Value = "ID_EVENTO (PT)"
        .Cells(intFila, 6).Value = "TIPO ACCIÓN"
        .Cells(intFila, 7).Value = "MATERIAL"
        .Cells(intFila, 8).Value = "P/N"
        .Cells(intFila, 9).Value = "S/N"
        .Cells(intFila, 10).Value = "DESCRIPCIÓN EQUIPO/MATERIAL"
        .Cells(intFila, 11).Value = "FECHA INSTALACIÓN / DESINSTALACION / ENTREGA"
        .Cells(intFila, 12).Value = "PRECIO (SIN IVA)"
        .Cells(intFila, 13).Value = "REPARADO POR"
    End With
    With wbHoja
        .Range(.Cells(intFila, 1), .Cells(intFila, 13)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFila, 1), .Cells(intFila, 13)).Interior.ColorIndex = 15
        .Range(.Cells(intFila, 1), .Cells(intFila, 13)).Font.Bold = True
    End With
    
    intFila = intFila + 1
    intFilaInicio = intFila
    '-----------------------
    ' MATERIALES
    '--------------------------
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaMateriales
    Else
       Set m_Col = Me.colListaMaterialesPorFacturar
    End If
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        GoTo Subcontrataciones
    End If
    intFila = intFila - 1
    For Each m_ID In m_Col
        intFila = intFila + 1
        Set m_ObjMaterial = m_Col(m_ID)
        Set m_ObjActividad = m_ObjMaterial.Actividad
        p_Error = m_ObjMaterial.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjActividad Is Nothing Then
            p_Error = "No se conoce la actividad del material " & m_ID
            Err.Raise 1000
        End If
        Set m_ObjEvento = m_ObjActividad.Evento
        p_Error = m_ObjActividad.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjEvento Is Nothing Then
            p_Error = "No se conoce El evento de la actividad " & m_ObjActividad.IDActividad
            Err.Raise 1000
        End If
        Set m_ObjEquipo = m_ObjEvento.Equipo
        p_Error = m_ObjEvento.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjEquipo Is Nothing Then
            p_Error = "No se conoce El Equipo del evento " & m_ObjEvento.IDEVENTO
            Err.Raise 1000
        End If
        m_Linea = "Excel Hoja Materiales..... " & m_ObjMaterial.Descripcion
        Avance m_Linea
        With wbHoja
            If m_ObjEvento.NODO <> "" Then
                .Cells(intFila, 1).Value = m_ObjEvento.NODO
            End If
            If m_ObjEvento.BUI <> "" Then
                .Cells(intFila, 2).Value = m_ObjEvento.BUI
            End If
            If m_ObjEvento.SubSistema <> "" Then
                .Cells(intFila, 3).Value = m_ObjEvento.SubSistema
            End If
            If m_ObjEquipo.Equipo <> "" Then
                .Cells(intFila, 4).Value = m_ObjEquipo.Equipo
            End If
            If m_ObjEvento.IDEVENTO <> "" Then
                .Cells(intFila, 5).Value = m_ObjEvento.IDEVENTO
            End If
            If m_ObjMaterial.TipoAccion <> "" Then
                .Cells(intFila, 6).Value = m_ObjMaterial.TipoAccion
            End If
            If m_ObjMaterial.Material <> "" Then
                .Cells(intFila, 7).Value = m_ObjMaterial.Material
            End If
            If m_ObjMaterial.PN <> "" Then
                .Cells(intFila, 8).Value = m_ObjMaterial.PN
            End If
            If m_ObjMaterial.NS <> "" Then
                .Cells(intFila, 9).Value = m_ObjMaterial.NS
            End If
            If m_ObjEquipo.Descripcion <> "" Then
                .Cells(intFila, 10).Value = m_ObjEquipo.Descripcion
            End If
            If IsDate(m_ObjMaterial.FechaEntrega) Then
                .Cells(intFila, 11).Value = CDate(Format(m_ObjMaterial.FechaEntrega, "dd/mm/yyyy"))
            End If
            If IsNumeric(m_ObjMaterial.COSTE) Then
                .Cells(intFila, 12).Value = CDbl(m_ObjMaterial.COSTE)
            End If
            
            If m_ObjMaterial.ReparadoPor <> "" Then
                .Cells(intFila, 13) = m_ObjMaterial.ReparadoPor
            End If
            .Range(.Cells(intFila, 12), .Cells(intFila, 12)).NumberFormat = "#,##0.00 $"
        End With
        Set m_ObjMaterial = Nothing
    Next
    intFilaFin = intFila
    With wbHoja
        RecuadrarRango .Range(.Cells(intFilaCabecera, 1), .Cells(intFilaFin, 13)), "Multiple"
        intFila = intFila + 1
        .Range(.Cells(intFila, 12), .Cells(intFila, 12)).Formula = "=SUM(L" & CStr(intFilaInicio) & ":L" & CStr(intFilaFin) & ")"
        .Cells(intFila, 12).Font.Bold = True
        .Range(.Cells(intFila, 12), .Cells(intFila, 12)).NumberFormat = "#,##0.00 $"
    End With
    
Subcontrataciones:
    '-------------------
    '  SUBCONTRATACIONES
    '--------------------
    intFila = intFila + 2
    wbHoja.Cells(intFila, 1).Value = "SUBCONTRATACIONES Y ASISTENCIA DE INGENIERÍA"
    wbHoja.Range(wbHoja.Cells(intFila, 1), wbHoja.Cells(intFila, 1)).Font.Bold = True
    intFila = intFila + 2
    intFilaCabecera = intFila
    With wbHoja
        .Cells(intFila, 1).Value = "FECHA INICIAL"
        .Cells(intFila, 2).Value = "FECHA FINAL"
        .Cells(intFila, 3).Value = "EMPRESA"
        .Cells(intFila, 4).Value = "DESCRIPCION"
        .Cells(intFila, 5).Value = "IMPORTE"
    End With
    With wbHoja
        .Range(.Cells(intFila, 1), .Cells(intFila, 5)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFila, 1), .Cells(intFila, 5)).Interior.ColorIndex = 15
        .Range(.Cells(intFila, 1), .Cells(intFila, 5)).Font.Bold = True
    End With
    
    intFila = intFila + 1
    intFilaInicio = intFila
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaSubContrataciones
    Else
        Set m_Col = Me.colListaSubContratacionesPorFacturar
    End If
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        GoTo Gastos
    End If
    intFila = intFila - 1
    For Each m_ID In m_Col
        
        Set m_ObjSubContratacion = m_Col(m_ID)
        m_Linea = "Excel Hoja Materiales..... " & m_ObjSubContratacion.Descripcion
        Avance m_Linea
        intFila = intFila + 1
        With m_ObjSubContratacion
            If IsDate(.FechaAlta) Then
                wbHoja.Cells(intFila, 1).Value = CDate(Format(.FechaAlta, "dd/mm/yyyy"))
            End If
            If IsDate(.FechaFin) Then
                wbHoja.Cells(intFila, 2).Value = CDate(Format(.FechaFin, "dd/mm/yyyy"))
            End If
            If .Empresa <> "" Then
                wbHoja.Cells(intFila, 3).Value = .Empresa
            End If
            If .Descripcion <> "" Then
                wbHoja.Cells(intFila, 4).Value = .Descripcion
            End If
            If IsNumeric(.Importe) Then
                wbHoja.Cells(intFila, 5).Value = CDbl(.Importe)
            End If
            wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).NumberFormat = "#,##0.00 $"
        End With
        Set m_ObjSubContratacion = Nothing
    Next
    intFilaFin = intFila
    intFila = intFila + 1
    With wbHoja
        RecuadrarRango .Range(.Cells(intFilaCabecera, 1), .Cells(intFilaFin, 5)), "Multiple"
        .Range(.Cells(intFila, 5), .Cells(intFila, 5)).Formula = "=SUM(E" & CStr(intFilaInicio) & ":E" & CStr(intFilaFin) & ")"
        .Cells(intFila, 5).Font.Bold = True
        .Range(.Cells(intFila, 5), .Cells(intFila, 5)).NumberFormat = "#,##0.00 $"
    End With
    
Gastos:
    '-------------------
    ' GASTOS Y DIETAS
    '--------------------
    intFila = intFila + 2
    wbHoja.Cells(intFila, 1).Value = "GASTOS Y DIETAS"
    wbHoja.Range(wbHoja.Cells(intFila, 1), wbHoja.Cells(intFila, 1)).Font.Bold = True
    intFila = intFila + 2
    intFilaCabecera = intFila
    With wbHoja
        .Cells(intFila, 1).Value = "TIPO GASTO"
        .Cells(intFila, 2).Value = "IMPORTE UNITARIO"
        .Cells(intFila, 3).Value = "Nº UNIDADES "
        .Cells(intFila, 4).Value = "IMPORTE "
        .Cells(intFila, 5).Value = "FECHA GASTO"
        .Cells(intFila, 6).Value = "TÉCNICO"
        .Cells(intFila, 7).Value = "OBSERVACIONES"
    End With
    With wbHoja
        .Range(.Cells(intFila, 1), .Cells(intFila, 7)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFila, 1), .Cells(intFila, 7)).Interior.ColorIndex = 15
        .Range(.Cells(intFila, 1), .Cells(intFila, 7)).Font.Bold = True
    End With
    intFila = intFila + 1
    intFilaInicio = intFila
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.colListaGastos
    Else
        Set m_Col = Me.colListaGastosPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    intFila = intFila - 1
    For Each m_ID In m_Col
        intFila = intFila + 1
        Set m_ObjGasto = m_Col(m_ID)
        m_Linea = "Excel Hoja Materiales..... " & m_ObjGasto.TipoGasto
        Avance m_Linea
        With m_ObjGasto
            If .TipoGasto <> "" Then
                wbHoja.Cells(intFila, 1).Value = .TipoGasto
            End If
            If IsNumeric(.ImporteUnitario) Then
                 wbHoja.Cells(intFila, 2).Value = CDbl(.ImporteUnitario)
                 wbHoja.Range(wbHoja.Cells(intFila, 2), wbHoja.Cells(intFila, 2)).NumberFormat = "#,##0.00 $"
            End If
            If IsNumeric(.NumeroUnidades) Then
                 wbHoja.Cells(intFila, 3).Value = CDbl(.NumeroUnidades)
            End If
            If IsNumeric(.ImporteUnitario) And IsNumeric(.NumeroUnidades) Then
                wbHoja.Cells(intFila, 4).Formula = "=B" & CStr(intFila) & "*" & "C" & CStr(intFila)
                wbHoja.Range(wbHoja.Cells(intFila, 4), wbHoja.Cells(intFila, 4)).NumberFormat = "#,##0.00 $"
            End If
            If IsDate(.FechaImputacionGasto) Then
                wbHoja.Cells(intFila, 5).Value = .FechaImputacionGasto
                 wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).NumberFormat = "mm/dd/yyyy"
            End If
            If .Alias <> "" Then
                wbHoja.Cells(intFila, 6).Value = .Alias
            End If
            If .Observaciones <> "" Then
                wbHoja.Cells(intFila, 7).Value = .Observaciones
            End If
        End With
        
        Set m_ObjGasto = Nothing
    Next
    intFilaFin = intFila
    intFila = intFila + 1
    With wbHoja
        RecuadrarRango .Range(.Cells(intFilaCabecera, 1), .Cells(intFilaFin, 7)), "Multiple"
        .Range(wbHoja.Cells(intFila, 4), wbHoja.Cells(intFila, 4)).Formula = "=SUM(D" & CStr(intFilaInicio) & ":D" & CStr(intFilaFin) & ")"
        .Cells(intFila, 4).Font.Bold = True
        .Range(.Cells(intFila, 4), .Cells(intFila, 4)).NumberFormat = "#,##0.00 $"
    End With
    
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").EntireColumn.AutoFit
    End With
    CreaHojaMateriales = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CreaHojaMateriales ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function CreaHojaFactura( _
                                    ByRef wbHoja As excel.Worksheet, _
                                    Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    
    Dim m_ObjColPerfiles As Scripting.Dictionary
    Dim m_ObjFacturaPerfil As FacturaPerfiles
    Dim m_ID As Variant
    Dim intFila As Integer
    Dim m_NHorasLab As Double
    Dim m_NHorasExt As Double
    Dim m_PrecioHoraLab As Double
    Dim m_PrecioHoraExt As Double
    Dim m_PrecioLab As Double
    Dim m_PrecioExt As Double
    Dim m_Precio As Double

    On Error GoTo errores
    If Not IsNumeric(Me.IDFactura) And p_EsCalculado = EnumSino.Sí Then
        p_Error = "No se ha introducido un IDFactura numérico existente válido"
        Err.Raise 1000
    End If
    
    m_Linea = "Excel Hoja Factura....."
    Avance m_Linea
    If p_EsCalculado = EnumSino.Sí Then
        Set m_ObjColPerfiles = Me.ColPerfilesEnFactura
    Else
        Set m_ObjColPerfiles = Me.ColPerfilesPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.Titulo1 = "" Then
        Me.Titulo1 = Me.Titulo1Calculado
    End If
    If Me.Titulo2 = "" Then
        Me.Titulo2 = Me.Titulo2Calculado
    End If
    If Me.Titulo3 = "" Then
        Me.Titulo3 = Me.Titulo3Calculado
    End If
    If Me.TipoImpositivo = "" Then
        Me.TipoImpositivo = Me.TipoImpositivoCalculado
    End If
    
    '----------------
    ' TÍTULO1
    '   MANTENIMIENTO DE LAS ESTACIONES RADIO Y SISTEMA BRASS PERIODO 2011-2013
    '----------------------
    intFila = 3
    With wbHoja
        .Range(.Cells(intFila, 2), .Cells(intFila, 8)).MergeCells = True
        .Cells(intFila, 2).Value = UCase(Me.Titulo1)
        .Range(.Cells(intFila, 2), .Cells(intFila, 2)).Font.Bold = True
        .Range(.Cells(intFila, 2), .Cells(intFila, 2)).HorizontalAlignment = xlCenter
    End With
    '----------------
    ' TÍTULO1
    '   MANTENIMIENTO DE LAS ESTACIONES RADIO Y SISTEMA BRASS PERIODO 2011-2013
    '----------------------
    intFila = intFila + 2
    With wbHoja
        .Range(.Cells(intFila, 2), .Cells(intFila, 8)).MergeCells = True
        .Cells(intFila, 2).Value = UCase(Me.Titulo2)
        .Range(.Cells(intFila, 2), .Cells(intFila, 2)).Font.Bold = True
        .Range(.Cells(intFila, 2), .Cells(intFila, 2)).HorizontalAlignment = xlCenter
    End With
    '----------------
    ' TRABAJOS EFECTIVAMENTE REALIZADOS
    '----------------------
    intFila = intFila + 4
    With wbHoja
        .Cells(intFila, 1).Value = "TRABAJOS EFECTIVAMENTE REALIZADOS"
        .Range(.Cells(intFila, 1), .Cells(intFila, 1)).Font.Bold = True
    End With
    '----------------
    ' LAB EXT LAB + EXT
    '----------------------
    intFila = intFila + 2
    With wbHoja
        .Cells(intFila, 7).Value = "LAB"
        .Cells(intFila, 8).Value = "EXT"
        .Cells(intFila, 9).Value = "LAB + EXT"
        .Range(.Cells(intFila, 7), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 7), .Cells(intFila, 9)).HorizontalAlignment = xlCenter
    End With
    For Each m_ID In m_ObjColPerfiles
        Set m_ObjFacturaPerfil = m_ObjColPerfiles(m_ID)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        With m_ObjFacturaPerfil
            If IsNumeric(.NHorasLab) Then
                m_NHorasLab = CDbl(.NHorasLab)
            Else
                m_NHorasLab = 0
            End If
            If IsNumeric(.NHorasExt) Then
                m_NHorasExt = CDbl(.NHorasExt)
            Else
                m_NHorasExt = 0
            End If
            If IsNumeric(.PrecioHoraLab) Then
                m_PrecioHoraLab = CDbl(.PrecioHoraLab)
            Else
                m_PrecioHoraLab = 0
            End If
            If IsNumeric(.PrecioHoraExt) Then
                m_PrecioHoraExt = CDbl(.PrecioHoraExt)
            Else
                m_PrecioHoraExt = 0
            End If
            m_PrecioLab = m_NHorasLab * m_PrecioHoraLab
            m_PrecioExt = m_NHorasExt * m_PrecioHoraExt
            m_Precio = m_PrecioLab + m_PrecioExt
        End With
        
        intFila = intFila + 1
        With wbHoja
            .Cells(intFila, 2).Value = "HORAS PERFIL " & m_ObjFacturaPerfil.Perfil
            .Range(.Cells(intFila, 2), .Cells(intFila, 8)).Font.Bold = True
            .Cells(intFila, 7).Value = m_NHorasLab
            .Cells(intFila, 8).Value = m_NHorasExt
            .Range(.Cells(intFila, 7), .Cells(intFila, 8)).NumberFormat = "0.00"
            .Range(.Cells(intFila, 8), .Cells(intFila, 8)).NumberFormat = "0.00"
            .Range(.Cells(intFila, 7), .Cells(intFila, 9)).HorizontalAlignment = xlRight
        End With
        intFila = intFila + 1
        With wbHoja
            .Cells(intFila, 2).Value = "TARIFA"
            .Range(.Cells(intFila, 2), .Cells(intFila, 8)).Font.Bold = True
            .Cells(intFila, 7).Value = m_PrecioHoraLab
            .Cells(intFila, 8).Value = m_PrecioHoraExt
            .Range(.Cells(intFila, 7), .Cells(intFila, 7)).NumberFormat = "#,##0.00 $"
            .Range(.Cells(intFila, 8), .Cells(intFila, 8)).NumberFormat = "#,##0.00 $"
            .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
            .Range(.Cells(intFila, 7), .Cells(intFila, 9)).HorizontalAlignment = xlRight
        End With
        
        
        intFila = intFila + 1
        With wbHoja
            .Cells(intFila, 2).Value = "PMO " & m_ObjFacturaPerfil.Perfil
            .Range(.Cells(intFila, 2), .Cells(intFila, 9)).Font.Bold = True
            .Cells(intFila, 7).Value = m_PrecioLab
            .Range(.Cells(intFila, 7), .Cells(intFila, 7)).NumberFormat = "#,##0.00 $"
            .Cells(intFila, 8).Value = m_PrecioExt
            .Range(.Cells(intFila, 8), .Cells(intFila, 8)).NumberFormat = "#,##0.00 $"
            .Cells(intFila, 9).Value = m_Precio
            .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
        End With
siguienteTipoPrecio:
        Set m_ObjFacturaPerfil = Nothing
        
    Next
    '----------------------------------------------------------------
    ' SUMINISTROS, REPARACIONES Y SUBCONTRATACION DE SERVICIOS
    '----------------------------------------------------------------
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 1).Value = "SUMINISTROS, REPARACIONES Y SUBCONTRATACION DE SERVICIOS"
        .Range(.Cells(intFila, 1), .Cells(intFila, 1)).Font.Bold = True
    End With
    '----------------------------------------------------------------
    ' RECARGO                                               strRecargo
    '----------------------------------------------------------------
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 2).Value = "RECARGO"
        .Range(.Cells(intFila, 2), .Cells(intFila, 7)).Font.Bold = True
        .Range(.Cells(intFila, 7), .Cells(intFila, 7)).Formula = 1 + (CDbl(Me.Recargo) / 100)
        .Range(.Cells(intFila, 7), .Cells(intFila, 7)).NumberFormat = "0.00"
    End With
    
    '----------------------------------------------------------------
    ' PRECIO SUMINISTRO REALIZADO
    '----------------------------------------------------------------
    Dim m_PrecioSuministro As String
    Dim m_PrecioSuministroConRecargo As String
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioSuministro = Me.PrecioSuministroCalculado
        m_PrecioSuministroConRecargo = Me.PrecioSuministroConRecargoCalculado
    Else
        m_PrecioSuministro = Me.PrecioSuministroPorFacturar
        m_PrecioSuministroConRecargo = Me.PrecioSuministroConRecargoPorFacturar
    End If
    If Not IsNumeric(m_PrecioSuministro) Then
        m_PrecioSuministro = "0"
    End If
    If Not IsNumeric(m_PrecioSuministroConRecargo) Then
        m_PrecioSuministroConRecargo = "0"
    End If
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 2).Value = "PRECIO SUMINISTRO REALIZADO"
        .Range(.Cells(intFila, 2), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 7), .Cells(intFila, 7)).Formula = CDbl(m_PrecioSuministro)
        .Range(.Cells(intFila, 7), .Cells(intFila, 7)).NumberFormat = "#,##0.00 $"
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(m_PrecioSuministroConRecargo)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
    End With
    '----------------------------------------------------------------
    ' PRECIO SUBCONTRATACIONES REALIZADAS
    '----------------------------------------------------------------
    Dim m_PrecioSubcontrataciones As String
    Dim m_PrecioSubcontratacionesConRecargo As String
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioSubcontrataciones = Me.PrecioSubcontratacionesCalculado
        m_PrecioSubcontratacionesConRecargo = Me.PrecioSubcontratacionesConRecargoCalculado
    Else
        m_PrecioSubcontrataciones = Me.PrecioSubcontratacionesPorFacturar
        m_PrecioSubcontratacionesConRecargo = Me.PrecioSubcontratacionesConRecargoForFacturar
    End If
    If Not IsNumeric(m_PrecioSubcontrataciones) Then
        m_PrecioSubcontrataciones = "0"
    End If
    If Not IsNumeric(m_PrecioSubcontratacionesConRecargo) Then
        m_PrecioSubcontratacionesConRecargo = "0"
    End If
    
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 2).Value = "PRECIO SUBCONTRATACIONES REALIZADAS"
        .Range(.Cells(intFila, 2), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 7), .Cells(intFila, 7)).Formula = CDbl(m_PrecioSubcontrataciones)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(m_PrecioSubcontratacionesConRecargo)
        .Range(.Cells(intFila, 7), .Cells(intFila, 7)).NumberFormat = "#,##0.00 $"
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
        
    End With
    '----------------------------------------------------------------
    ' DIETAS Y GASTOS
    '----------------------------------------------------------------
    Dim m_PrecioDietasYGastos As String
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioDietasYGastos = Me.PrecioDietasYGastosCalculado
    Else
        m_PrecioDietasYGastos = Me.PrecioDietasYGastosPorFacturar
    End If
    If Not IsNumeric(m_PrecioDietasYGastos) Then
        m_PrecioDietasYGastos = "0"
    End If
    intFila = intFila + 2
    With wbHoja
        .Cells(intFila, 1).Value = "DIETAS Y GASTOS"
        .Range(.Cells(intFila, 1), .Cells(intFila, 1)).Font.Bold = True
        .Range(.Cells(intFila, 2), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(m_PrecioDietasYGastos)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
        
    End With
    '----------------------------------------------------------------
    ' PRECIO INCURRIDO
    '----------------------------------------------------------------
    Dim m_PrecioIncurrido As String
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioIncurrido = Me.PrecioIncurridoCalculado
    Else
        m_PrecioIncurrido = Me.PrecioIncurridoPorFacturar
    End If
    If Not IsNumeric(m_PrecioIncurrido) Then
        m_PrecioIncurrido = "0"
    End If
    intFila = intFila + 2
    With wbHoja
        .Cells(intFila, 1).Value = "PRECIO INCURRIDO"
        .Range(.Cells(intFila, 1), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(m_PrecioIncurrido)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
        
    End With
    '----------------------------------------------------------------
    ' BASE IMPONIBLE
    '----------------------------------------------------------------
    Dim m_PrecioBaseImponible As String
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioBaseImponible = Me.PrecioBaseImponibleCalculado
    Else
        m_PrecioBaseImponible = Me.PrecioBaseImponiblePorFacturar
    End If
    If Not IsNumeric(m_PrecioBaseImponible) Then
        m_PrecioBaseImponible = "0"
    End If
    intFila = intFila + 2
    With wbHoja
        .Cells(intFila, 1).Value = "BASE IMPONIBLE"
        .Range(.Cells(intFila, 1), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(m_PrecioBaseImponible)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
        
    End With
    '----------------------------------------------------------------
    ' TIPO IMPONIBLE
    '----------------------------------------------------------------
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 1).Value = "TIPO IMPONIBLE"
        .Range(.Cells(intFila, 1), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(Me.TipoImpositivo)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "0.00"
    End With
    
    '----------------------------------------------------------------
    ' IMPUESTOS
    '----------------------------------------------------------------
    Dim m_PrecioImpuestos As String
    If p_EsCalculado = EnumSino.Sí Then
        m_PrecioImpuestos = Me.PrecioImpuestosCalculado
    Else
        m_PrecioImpuestos = Me.PrecioImpuestosPorFacturar
    End If
    If Not IsNumeric(m_PrecioImpuestos) Then
        m_PrecioImpuestos = "0"
    End If
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 1).Value = "IMPUESTOS"
        .Range(.Cells(intFila, 1), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(m_PrecioImpuestos)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
    End With
    '------------------------------------------------------------------------
    ' TÍTULO 3
    '   TRABAJOS A FACTURAR TRIMESTRE-1-2 013                        X,XX
    '-------------------------------------------------------------------------
    Dim m_ImporteAFacturar As String
    If p_EsCalculado = EnumSino.Sí Then
        m_ImporteAFacturar = Me.ImporteAFacturarCalculado
    Else
        m_ImporteAFacturar = Me.ImporteAFacturarPorFacturar
    End If
    If Not IsNumeric(m_ImporteAFacturar) Then
        m_ImporteAFacturar = "0"
    End If
    intFila = intFila + 2
    With wbHoja
        .Range(.Cells(intFila, 1), .Cells(intFila, 7)).MergeCells = True
        .Cells(intFila, 1).Value = UCase(Me.Titulo3)
        .Range(.Cells(intFila, 1), .Cells(intFila, 9)).Font.Bold = True
        .Range(.Cells(intFila, 1), .Cells(intFila, 1)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Formula = CDbl(m_ImporteAFacturar)
        .Range(.Cells(intFila, 9), .Cells(intFila, 9)).NumberFormat = "#,##0.00 $"
    End With
    With wbHoja
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
    End With
    CreaHojaFactura = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CreaHojaFactura ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function CreaHojaDetalleGetronics( _
                                        ByRef wbHoja As excel.Worksheet, _
                                        Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                        Optional ByRef p_Error As String _
                                        ) As String


    Dim m_EventosDeGetronicsEnFactura As Scripting.Dictionary
    Dim m_IDEvento As Variant
    Dim m_Evento As Evento
    Dim m_ActividadesGetronics As String
    
    'ANALISTA;150;7410|TECCOM;0;0|PROGRAMADOR;120;4284|11694

    Dim m_TrozoAnalista As String
    Dim m_TrozoTeccom As String
    Dim m_TrozoProgramador As String
    Dim m_TrozoCosteTotal As String
    
    Dim m_HorasTotales As String
    Dim m_Coste As String
    Dim m_CosteTotal As String


    Dim intFila As Integer
    Dim intFilaCabecera As Integer
    Dim intFilaInicio As Integer
    Dim intFilaFin As Integer
    Dim dato
    On Error GoTo errores

    If Not IsNumeric(Me.IDFactura) And p_EsCalculado = EnumSino.Sí Then
        p_Error = "No se ha introducido un IDFactura numérico existente válido"
        Err.Raise 1000
    End If

    m_Linea = "Excel Hoja Detalles Getronics....."
    Avance m_Linea
    If p_EsCalculado = EnumSino.Sí Then
        Set m_EventosDeGetronicsEnFactura = getEventosGetronics(Me.colListaEventos, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set m_EventosDeGetronicsEnFactura = getEventosGetronics(Me.colListaEventosPorFacturar, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
       
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EventosDeGetronicsEnFactura Is Nothing Then
        Exit Function
    End If
    
    intFila = 1
    With wbHoja
        .Cells(intFila, 1).Value = "TIPO TÉCNICO"
        .Cells(intFila, 2).Value = "EVENTO"
        .Cells(intFila, 3).Value = "DESCRIPCIÓN"
        .Cells(intFila, 4).Value = "HORAS TOTALES"
        .Cells(intFila, 5).Value = "COSTE"
    End With
    With wbHoja
        .Range(.Cells(intFila, 1), .Cells(intFila, 5)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFila, 1), .Cells(intFila, 5)).Interior.ColorIndex = 15
        .Range(.Cells(intFila, 1), .Cells(intFila, 5)).Font.Bold = True
    End With
    intFila = intFila + 1
    'SE HA DE RECORRE CADA UNO DE LOS PERFILES QUE TIENEN ACTIVIDADES
    For Each m_IDEvento In m_EventosDeGetronicsEnFactura
        Set m_Evento = m_EventosDeGetronicsEnFactura(m_IDEvento)
        m_ActividadesGetronics = Constructor.getActividadesGetronic(CStr(m_IDEvento), p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Linea = "Excel Hoja Detalles Getronics Evento.... " & m_IDEvento
        'ANALISTA;150;7410|TECCOM;0;0|PROGRAMADOR;120;4284|11694
        dato = Split(m_ActividadesGetronics, "|")
        m_TrozoAnalista = dato(0)
        m_TrozoTeccom = dato(1)
        m_TrozoProgramador = dato(2)
        m_TrozoCosteTotal = dato(3)
        
        m_CosteTotal = m_TrozoCosteTotal
        
        wbHoja.Cells(intFila, 2).Value = m_Evento.IDEVENTO
        wbHoja.Cells(intFila, 3).Value = m_Evento.Descripcion
        wbHoja.Cells(intFila, 5).Value = CDbl(m_CosteTotal)
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).NumberFormat = "#,##0.00 $"
        wbHoja.Range(wbHoja.Cells(intFila, 2), wbHoja.Cells(intFila, 2)).Font.Bold = True
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).Font.Bold = True
        wbHoja.Range(wbHoja.Cells(intFila, 3), wbHoja.Cells(intFila, 3)).WrapText = True
        intFila = intFila + 1
        
        
        dato = Split(m_TrozoAnalista, ";")
        m_HorasTotales = dato(1)
        m_Coste = dato(2)
        wbHoja.Cells(intFila, 1).Value = "ING.SISTEMAS"
        wbHoja.Cells(intFila, 2).Value = "ANALISTA"
        wbHoja.Cells(intFila, 4).Value = CDbl(m_HorasTotales)
        wbHoja.Range(wbHoja.Cells(intFila, 4), wbHoja.Cells(intFila, 4)).NumberFormat = "#,##0"
        wbHoja.Cells(intFila, 5).Value = CDbl(m_Coste)
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).NumberFormat = "#,##0.00 $"
        wbHoja.Range(wbHoja.Cells(intFila, 2), wbHoja.Cells(intFila, 2)).Font.Bold = True
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).Font.Bold = True
        
        wbHoja.Range(wbHoja.Cells(intFila, 3), wbHoja.Cells(intFila, 3)).WrapText = True
        intFila = intFila + 1
        
        dato = Split(m_TrozoTeccom, ";")
        m_HorasTotales = dato(1)
        m_Coste = dato(2)
        wbHoja.Cells(intFila, 1).Value = "TEC. COM."
        wbHoja.Cells(intFila, 2).Value = "TEC. COM."
        wbHoja.Cells(intFila, 4).Value = CDbl(m_HorasTotales)
        wbHoja.Range(wbHoja.Cells(intFila, 4), wbHoja.Cells(intFila, 4)).NumberFormat = "#,##0"
        wbHoja.Cells(intFila, 5).Value = CDbl(m_Coste)
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).NumberFormat = "#,##0.00 $"
        wbHoja.Range(wbHoja.Cells(intFila, 2), wbHoja.Cells(intFila, 2)).Font.Bold = True
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).Font.Bold = True
        wbHoja.Range(wbHoja.Cells(intFila, 3), wbHoja.Cells(intFila, 3)).WrapText = True
        intFila = intFila + 1
        
        dato = Split(m_TrozoProgramador, ";")
        m_HorasTotales = dato(1)
        m_Coste = dato(2)
        wbHoja.Cells(intFila, 1).Value = "TÉCNICO DE SISTEMAS"
        wbHoja.Cells(intFila, 2).Value = "PROGRAMADOR"
        wbHoja.Cells(intFila, 4).Value = CDbl(m_HorasTotales)
        wbHoja.Range(wbHoja.Cells(intFila, 4), wbHoja.Cells(intFila, 4)).NumberFormat = "#,##0"
        wbHoja.Cells(intFila, 5).Value = CDbl(m_Coste)
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).NumberFormat = "#,##0.00 $"
        wbHoja.Range(wbHoja.Cells(intFila, 2), wbHoja.Cells(intFila, 2)).Font.Bold = True
        wbHoja.Range(wbHoja.Cells(intFila, 5), wbHoja.Cells(intFila, 5)).Font.Bold = True
        wbHoja.Range(wbHoja.Cells(intFila, 3), wbHoja.Cells(intFila, 3)).WrapText = True
        intFila = intFila + 2
    Next
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        

    End With

    With wbHoja
        .Columns("C:C").ColumnWidth = 93
        
    End With
    CreaHojaDetalleGetronics = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CreaHojaDetalleGetronics ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If

End Function

Public Function RellenarListaResultado( _
                                        ByRef lst As ListBox, _
                                        Optional p_EsCalculado As EnumSino = EnumSino.Sí, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjPerfilFactura As FacturaPerfiles
    Dim m_NHorasLab As Double
    Dim m_NHorasExt As Double
    Dim m_PrecioHoraLab As Double
    Dim m_PrecioHoraExt As Double
    Dim m_PrecioLab As Double
    Dim m_PrecioExt As Double
    Dim m_Precio As Double
    
    On Error GoTo errores
    lst.RowSource = "PERFIL;NLAB;NEXT;€HLab;€HExt;PLab;PExt;Precio"
    If p_EsCalculado = EnumSino.Sí Then
        Set m_Col = Me.ColPerfilesEnFactura
    Else
        Set m_Col = Me.ColPerfilesPorFacturar
    End If
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjPerfilFactura = m_Col(m_ID)
        With m_ObjPerfilFactura
            If IsNumeric(.NHorasLab) Then
                m_NHorasLab = CDbl(.NHorasLab)
            Else
                m_NHorasLab = 0
            End If
            If IsNumeric(.NHorasExt) Then
                m_NHorasExt = CDbl(.NHorasExt)
            Else
                m_NHorasExt = 0
            End If
            If IsNumeric(.PrecioHoraLab) Then
                m_PrecioHoraLab = CDbl(.PrecioHoraLab)
            Else
                m_PrecioHoraLab = 0
            End If
            If IsNumeric(.PrecioHoraExt) Then
                m_PrecioHoraExt = CDbl(.PrecioHoraExt)
            Else
                m_PrecioHoraExt = 0
            End If
            If IsNumeric(.ImporteHorasLab) Then
                m_PrecioLab = CDbl(.ImporteHorasLab)
            Else
                m_PrecioLab = 0
            End If
            If IsNumeric(.ImporteHorasExt) Then
                m_PrecioExt = CDbl(.ImporteHorasExt)
            Else
                m_PrecioExt = 0
            End If
            If IsNumeric(.Importe) Then
                m_Precio = CDbl(.Importe)
            Else
                m_Precio = 0
            End If
            lst.AddItem .Perfil & ";" & Format(m_NHorasLab, "#,##0.00") & ";" & _
                        Format(m_NHorasExt, "#,##0.00") & ";" & Format(m_PrecioHoraLab, "#,##0.00 €") & ";" & _
                        Format(m_PrecioHoraExt, "#,##0.00 €") & ";" & Format(m_PrecioLab, "#,##0.00 €") & ";" & _
                        Format(m_PrecioExt, "#,##0.00 €") & ";" & Format(m_Precio, "#,##0.00 €")
        End With
        
        Set m_ObjPerfilFactura = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListaResultado ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "IDFactura"
        .Add "FechaFinalEventos"
        .Add "FechaFactura"
        .Add "Titulo1"
        .Add "Titulo2"
        .Add "Titulo3"
        .Add "TipoImpositivo"
        .Add "Recargo"
        .Add "PrecioHoras"
        .Add "PrecioSuministro"
        .Add "PrecioSuministroConRecargo"
        .Add "PrecioDietasYGastos"
        .Add "PrecioIncurrido"
        .Add "PrecioSubcontrataciones"
        .Add "PrecioSubcontratacionesConRecargo"
        .Add "PrecioBaseImponible"
        .Add "PrecioImpuestos"
        .Add "ImporteAFacturar"
        .Add "NombreFacturaAnexa"
       
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Factura.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDFactura" Then
        getPropiedad = Me.IDFactura
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinalEventos" Then
        getPropiedad = Me.FechaFinalEventos
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFactura" Then
        getPropiedad = Me.FechaFactura
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo1" Then
        getPropiedad = Me.Titulo1
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo2" Then
        getPropiedad = Me.Titulo2
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo3" Then
        getPropiedad = Me.Titulo3
        Exit Function
    End If
    If m_NombrePropiedad = "TipoImpositivo" Then
        getPropiedad = Me.TipoImpositivo
        Exit Function
    End If
    If m_NombrePropiedad = "Recargo" Then
        getPropiedad = Me.Recargo
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoras" Then
        getPropiedad = Me.PrecioHoras
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioSuministro" Then
        getPropiedad = Me.PrecioSuministro
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioSuministroConRecargo" Then
        getPropiedad = Me.PrecioSuministroConRecargo
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioDietasYGastos" Then
        getPropiedad = Me.PrecioDietasYGastos
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioIncurrido" Then
        getPropiedad = Me.PrecioIncurrido
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioSubcontrataciones" Then
        getPropiedad = Me.PrecioSubcontrataciones
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioSubcontratacionesConRecargo" Then
        getPropiedad = Me.PrecioSubcontratacionesConRecargo
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioBaseImponible" Then
        getPropiedad = Me.PrecioBaseImponible
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioImpuestos" Then
        getPropiedad = Me.PrecioImpuestos
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteAFacturar" Then
        getPropiedad = Me.ImporteAFacturar
        Exit Function
    End If
    If m_NombrePropiedad = "NombreFacturaAnexa" Then
        getPropiedad = Me.NombreFacturaAnexa
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Factura.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function
Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDFactura" Then
        Me.IDFactura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinalEventos" Then
        Me.FechaFinalEventos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFactura" Then
        Me.FechaFactura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo1" Then
        Me.Titulo1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo2" Then
        Me.Titulo2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo3" Then
        Me.Titulo3 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoImpositivo" Then
        Me.TipoImpositivo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Recargo" Then
        Me.Recargo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoras" Then
        Me.PrecioHoras = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "PrecioSuministro" Then
        Me.PrecioSuministro = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioSuministroConRecargo" Then
        Me.PrecioSuministroConRecargo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioDietasYGastos" Then
        Me.PrecioDietasYGastos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioIncurrido" Then
        Me.PrecioIncurrido = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioSubcontrataciones" Then
        Me.PrecioSubcontrataciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioSubcontratacionesConRecargo" Then
        Me.PrecioSubcontratacionesConRecargo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioBaseImponible" Then
        Me.PrecioBaseImponible = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioImpuestos" Then
        Me.PrecioImpuestos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteAFacturar" Then
        Me.ImporteAFacturar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreFacturaAnexa" Then
        Me.NombreFacturaAnexa = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Factura.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Factura.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function






