Option Compare Database
Option Explicit
Dim strBUI As String
Private Sub ComandoAlta_Click()
    Dim strSUBSISTEMA As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strSUBSISTEMA = Nz(Me.SubSistema, "")
    '-----------------------------------------------
    '   -Devuelve:
    '       Alta = strDescriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Alta(strSUBSISTEMA)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '------------------------------------------------
    If FormularioAbierto("FormSubSistemaGestion") Then
        Form_FormSubSistemaGestion.ActualizarLista
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta el SubSistema se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Load()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    'strArgumento=strBUI
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    strBUI = Nz(Me.OpenArgs, "")
    If strBUI = "" Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuadosW"
        Err.Raise 1000
    End If
    Me.BUI = strBUI
    Me.SubSistema.SetFocus
    m_SQL = "SELECT distinct TbSubsistemaBui.SUBSISTEMA " & _
                "FROM TbSubsistemaBui " & _
                "WHERE ((Not (TbSubsistemaBui.BUI)='" & strBUI & "')) " & _
                "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    Me.SubSistema.RowSource = m_SQL
    Me.SubSistema.Requery
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
Public Function Alta(strSUBSISTEMA As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 29/12/13
    '   -Modificaciones:
    '       -9/7/13: Sólo se puede elegir entre los subsistemas que ya haya y no uno que ya esté en ese bui
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbSubsistemaBui
    '       -strSubSistema<>""
    '       -strSubSistema no ha de existir previamente
    '       -strBUI<>""
    '       -strBUI existente
    '   -llamada por:
    '       -Form_FormSubSistemaAlta.ComandoAlta_Click
    '   -Devuelve:
    '       Alta = Descriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strBUI = "" Then
        strTextoError = "No se ha indicado el nombre del BUI"
        Err.Raise 1000
    End If
    If strSUBSISTEMA = "" Then
        strTextoError = "No se ha indicado el nombre del SubSistema"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbSubsistemaBui.SUBSISTEMA, TbSubsistemaBui.BUI " & _
            "FROM TbSubsistemaBui " & _
            "WHERE (((TbSubsistemaBui.SUBSISTEMA)='" & strSUBSISTEMA & _
            "') AND ((TbSubsistemaBui.BUI)='" & strBUI & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya existe un subsistema llamado: " & strSUBSISTEMA & " para el BUI : " & strBUI
            Err.Raise 1000
        End If
        .AddNew
            .Fields("SUBSISTEMA") = strSUBSISTEMA
            .Fields("BUI") = strBUI
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Alta = "Se ha dado de alta correctamente el Subsistema llamado: " & strSUBSISTEMA & " para el BUI: " & strBUI
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







