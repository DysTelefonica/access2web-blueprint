Option Compare Database
Option Explicit
Dim strIDEquipo As String, strNombreEquipo As String
Private Sub ComandoAlta_Click()
    Dim strAlias As String, strNumeroMeses As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strAlias = Nz(Me.Alias, "")
    strNumeroMeses = Nz(Me.PERIODICIDADMESES, "")
    If Not IsNumeric(strNumeroMeses) Then
        strNumeroMeses = "0"
    End If
    
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AltaEquipoEnPlanificacion = Descriptivo
    '       AltaEquipoEnPlanificacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = AltaEquipoEnPlanificacion(strIDEquipo, strAlias, , CInt(strNumeroMeses))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AltaEquipoEnPlanificacion ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta el equipo planificable se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Load()
    Dim strTextoError As String, strArgumento As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    VBA.DoEvents
    
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strIDEquipo & "|" & strNombreEquipo
    strArgumento = Nz(Me.OpenArgs(), "")
    If InStr(1, strArgumento, "|") = 0 Then
        strTextoError = "El formulario se ha abierto con datos insuficientes"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    strIDEquipo = Nz(dato(0), "")
    strNombreEquipo = Nz(dato(1), "")
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "El formulario se ha abierto con datos insuficientes"
        Err.Raise 1000
    End If
    If strNombreEquipo = "" Then
        strTextoError = "El formulario se ha abierto con datos insuficientes"
        Err.Raise 1000
    End If
    Me.ComboEquipoPlanificable = strNombreEquipo
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