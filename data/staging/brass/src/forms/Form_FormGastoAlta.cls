Option Compare Database
Option Explicit
Private Sub CalculoTotal()
    Dim dblImporteUnitario As Double, dblNUnidades As Double, dblImporteTotal As Double
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If IsNumeric(Nz(Me.ImporteUnitario, "")) And IsNumeric(Nz(Me.NUnidades, "")) Then
        dblImporteUnitario = CDbl(Me.ImporteUnitario)
        dblNUnidades = CDbl(Me.NUnidades)
        dblImporteTotal = dblImporteUnitario * dblNUnidades
    End If
    If dblImporteTotal = 0 Then
        Me.ImporteTotal = Null
    Else
        Me.ImporteTotal = Format(dblImporteTotal, "#,##0.00 €")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoAlta_Click()
    Dim strTipoGasto As String, strImporteUnitario As String, strNumeroUnidades As String, strNombreFormulario As String, _
        strFechaImputacionGasto As String, strAlias As String, strObservaciones As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    strTipoGasto = Nz(Me.ComboTipoGasto, "")
    strImporteUnitario = Nz(Me.ImporteUnitario, "")
    strNumeroUnidades = Nz(Me.NUnidades, "")
    strFechaImputacionGasto = Nz(Me.FechaImputacion, "")
    strAlias = Nz(Me.ComboAlias, "")
    strObservaciones = Nz(Me.Observaciones, "")
    '-----------------------------------------------
    '   -Devuelve:
    '       Alta = strID
    '       Alta = "#ERR" & "|" &strTextoError
    '-----------------------------------------------
    flag = Alta(strTipoGasto, strImporteUnitario, strNumeroUnidades, strFechaImputacionGasto, strAlias, strObservaciones)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '--------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------------
    If FormularioAbierto("FormGastoGestion") Then
        Form_FormGastoGestion.ActualizarLista
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta el gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComboTipoGasto_Change()
    Dim rcdDatos As DAO.Recordset, strImporteUnitarioDeTabla As String, strTipoGasto As String, strTextoError As String
    On Error GoTo errores
    strTipoGasto = Nz(Me.ComboTipoGasto, "")
    If strTipoGasto = "" Then
        Me.ImporteUnitario = Null
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_SQL = "SELECT TbGastosImportePorTipo.ImporteUnitario " & _
            "FROM TbGastosImportePorTipo " & _
            "WHERE (((TbGastosImportePorTipo.TipoGasto)='" & strTipoGasto & _
                "') AND ((TbGastosImportePorTipo.FechaFinalImporte) Is Null));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strImporteUnitarioDeTabla = Nz(.Fields("ImporteUnitario"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If IsNumeric(strImporteUnitarioDeTabla) Then
        Me.ImporteUnitario = Format(strImporteUnitarioDeTabla, "#,##0.00 €")
    Else
        Me.ImporteUnitario = Null
    End If
    CalculoTotal
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del campo Tipo Gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Load()

    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    RellenarComboTecnicos Me.ComboAlias, EnumSino.No, EnumSino.Sí, strTextoError
    If strTextoError <> "" Then
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
        strTextoError = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ImporteUnitario_Exit(Cancel As Integer)
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    CalculoTotal
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub NUnidades_Exit(Cancel As Integer)
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    CalculoTotal
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Public Function Alta( _
                        strTipoGasto As String, _
                        strImporteUnitario As String, _
                        strNumeroUnidades As String, _
                        Optional strFechaImputacionGasto As String, _
                        Optional strAlias As String, _
                        Optional strObservaciones As String _
                        ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 05/09/14
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla  TbGastos
    '   -Llamada desde
    '       -FormGastoAlta.ComandoRegistrar_Click
    '   -Devuelve:
    '       Alta = strID
    '       Alta = "#ERR" & "|" &strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strID As String, strTextoError As String
    On Error GoTo errores
    
    If strTipoGasto = "" Then
        strTextoError = "Se ha de introducir el tipo de gasto"
        Err.Raise 1000
    End If
    If Not IsNumeric(strImporteUnitario) Then
        strTextoError = "Se ha de introducir el importe unitario"
        Err.Raise 1000
    End If
    If Not IsNumeric(strNumeroUnidades) Then
        strTextoError = "Se ha de introducir el número de unidades"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaImputacionGasto) Then
        strFechaImputacionGasto = Now()
    End If
    '------------------------------------------
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameID("TbGastos", "IDGasto")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameID ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strID = flag
    m_SQL = "SELECT TbGastos.* " & _
            "FROM TbGastos " & _
            "WHERE (((TbGastos.IDGasto)=" & strID & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "la función dameid ha devuelto un valor ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDGasto") = strID
            .Fields("TipoGasto") = strTipoGasto
            .Fields("ImporteUnitario") = CDbl(strImporteUnitario)
            .Fields("NumeroUnidades") = CDbl(strNumeroUnidades)
            .Fields("FechaImputacionGasto") = Format(strFechaImputacionGasto, "dd/mm/yyyy")
            If strAlias <> "" Then
                .Fields("ALIAS") = strAlias
            End If
            If strObservaciones <> "" Then
                .Fields("Observaciones") = strObservaciones
            End If
        .Update
    End With
    Alta = strID
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Alta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Alta = "#ERR" & "|" & strTextoError
End Function

