Option Compare Database
Option Explicit
Dim strIDEquipo As String
Public Sub EstablecerListaSemana()
    Dim strAño As String, strMes As String, strCadenaSemana As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strAño = Nz(Me.AÑO, "")
    strMes = Nz(Me.ComboMes, "")
    If Not IsNumeric(strAño) Or Not IsNumeric(strMes) Then
        strCadenaSemana = "#ERR"
    Else
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       FechaViernes =dteFechaViernesSemanax para intSemana=x <>0
        '       FechaViernes =strCadenaFechas para intSemana=0
        '       strCadenaFechas=01 & dteFechaViernesSemana1 & vbcrlf & _
                                02 & dteFechaViernesSemana2 & vbcrlf & _
        '                       ...
        '                       intNumeroSemanas & dteFechaViernesSemanax
        '       FechaViernes="#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = FechaViernes(CInt(strMes), CLng(strAño))
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método FechaViernes ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strCadenaSemana = flag
        strCadenaSemana = Replace(strCadenaSemana, vbCrLf, ";")
    End If
    Me.ComboSemana.RowSource = strCadenaSemana
    Me.ComboSemana.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar cargar el combo de las semanas se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Año_Exit(Cancel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    EstablecerListaSemana
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoAlta_Click()
    Dim strAño As String, strMes As String, strSemana As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "No se puede determinar el equipo al que se le pretende aplicar la planificación"
        Err.Raise 1000
    End If
    strAño = Nz(Me.AÑO, "")
    strMes = Nz(Me.ComboMes, "")
    strSemana = Nz(Me.ComboSemana, "")
     If Not IsNumeric(strAño) Then
        strTextoError = "No se ha introducido un año numérico válido"
        Err.Raise 1000
    End If
    If Not IsNumeric(strMes) Then
        strTextoError = "No se ha introducido una mes no válido"
        Err.Raise 1000
    End If
    If Len(strSemana) < 3 Then
         strTextoError = "No se ha introducido una válida"
        Err.Raise 1000
    End If
    If Not IsNumeric(Left(strSemana, 2)) Then
        strTextoError = "No se ha introducido una válida"
        Err.Raise 1000
    End If
    strSemana = Left(strSemana, 2)
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AltaPlanificacion = strIDPlanificacion & ";" & strFechaPrevistaCierre
    '       AltaPlanificacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = AltaPlanificacion(strIDEquipo, CInt(strSemana), CInt(strMes), CLng(strAño))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AltaPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormPlanificacionEquipos") Then
        Form_FormPlanificacionEquipos.ActualizarLista
    End If
    '----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormPlanificacionesBusqueda") Then
        Forms("FormPlanificacionesBusqueda").Controls("ListaFiltrados").Requery
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de AltaPlanificacion una nueva planifiación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComboMes_Exit(Cancel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    EstablecerListaSemana
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComboSemana_Exit(Cancel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub Form_Load()
    Dim strEquipo As String, strBUI As String, strSUBSISTEMA As String, strArgumento As String, _
        strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipoSeleccionado & "|" & strIDEquipo
    strArgumento = Nz(Me.OpenArgs(), "")
    If strArgumento = "" Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 3 Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    strBUI = Nz(dato(0), "")
    strSUBSISTEMA = Nz(dato(1), "")
    strIDEquipo = Nz(dato(2), "")
    strEquipo = Nz(dato(3), "")
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
    If strBUI <> "" Then
        Me.BUI = strBUI
    End If
    If strSUBSISTEMA <> "" Then
        Me.SubSistema = strSUBSISTEMA
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub







