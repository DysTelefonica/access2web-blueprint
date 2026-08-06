Option Compare Database
Option Explicit
Dim strURLCompletaLog As String

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function Generar( _
                            strURLDirectorio As String, _
                            strAño As String _
                            ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 10/04/13
    '   -Modificaciones:
    '   -Funcionamiento:
   
    '   -Llamada desde
    
    '   -Devuelve:
    '       Generar = strURLExcel
    '       Generar = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------------
    Dim fso As New FileSystemObject, WbLibro As excel.Workbook, wbHoja As excel.Worksheet, strFechaAhora As String, _
        strNombreExcel As String, strURLExcel As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strAño) Then
        strTextoError = "Se ha de introducir el año"
        Err.Raise 1000
    End If
    If Not fso.FolderExists(strURLDirectorio) Then
        strTextoError = "Se ha de elegir una ubicación alcanzable"
        Err.Raise 1000
    End If
    If Right(strURLDirectorio, 1) <> "\" Then
        strURLDirectorio = strURLDirectorio & "\"
    End If
    strNombreExcel = "Informe_Planificados_" & strAño & ".xls"
    strURLExcel = strURLDirectorio & strNombreExcel
    If fso.FileExists(strURLExcel) Then
        '--------------------------------------------
        '   -Devuelve:
        '       FicheroAbierto = true or false
        '       FicheroAbierto =False-->Error
        '--------------------------------------------
        If FicheroAbierto(strURLExcel) Then
            strTextoError = "El documento existe y está abierto"
            Err.Raise 1000
        End If
        fso.DeleteFile strURLExcel, True
    End If
    Set appexcel = New excel.Application
    appexcel.Visible = False
    Set WbLibro = appexcel.Workbooks.Add
    Set wbHoja = WbLibro.Worksheets(1)
    wbHoja.Name = "EVENTOS"
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       ObtenerHojaEventosDeExportacion =Descriptivo
    '       ObtenerHojaEventosDeExportacion ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = ObtenerHojaEventosDeExportacion(wbHoja)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método ObtenerHojaEventosDeExportacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Set wbHoja = WbLibro.Worksheets(2)
    wbHoja.Name = "ACTIVIDAD REALIZADA"
     '-------------------------------------------------------------------
    '   -Devuelve:
    '       ObtenerHojaActividadesDeExportacion =Descriptivo
    '       ObtenerHojaActividadesDeExportacion ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = ObtenerHojaActividadesDeExportacion(wbHoja)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método ObtenerHojaActividadesDeExportacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Set wbHoja = WbLibro.Worksheets(3)
     wbHoja.Name = "MATERIAL UTILIZADO Y REPARADO"
     '-------------------------------------------------------------------
    '   -Devuelve:
    '       ObtenerHojaMaterialesDeExportacion =Descriptivo
    '       ObtenerHojaMaterialesDeExportacion ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = ObtenerHojaMaterialesDeExportacion(wbHoja)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método ObtenerHojaMaterialesDeExportacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    WbLibro.SaveAs strURLExcel
    Set wbHoja = Nothing
    WbLibro.Close
    Set WbLibro = Nothing
    appexcel.Quit
    Set appexcel = Nothing
    Generar = "Edición correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
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
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    Generar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoGenerar_Click()
    Dim fso As New FileSystemObject, strAño As String, strURLDirectorio As String, strURLExcel As String, _
        strTextoError As String
    On Error GoTo errores
    strAño = Nz(Me.AÑO, "")
    If Not IsNumeric(strAño) Then
        strTextoError = "Se ha de introducir primero el año"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Seleccionar = CStr(varFile)
    '       Seleccionar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = Seleccionar(False, "Seleccione la carpeta")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Seleccionar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strURLDirectorio = flag
    If Not fso.FolderExists(strURLDirectorio) Then
        If Not fso Is Nothing Then
            Set fso = Nothing
        End If
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '------------------------------------------------------
    '   -Devuelve:
    '       Generar = strURLExcel
    '       Generar = "#ERR" & "|" & strTextoError
    '------------------------------------------------------
    flag = Generar(strURLDirectorio, strAño)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Generar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strURLExcel = flag
    If Not fso.FileExists(strURLExcel) Then
        strTextoError = "El método Generar ha devuelto un resultado con formato desconocido"
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", strURLExcel, "", "", 1
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el calendario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Load()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub