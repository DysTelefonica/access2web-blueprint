Option Compare Database
Option Explicit
Dim strSubSistemaSeleccionado As String, strBUISeleccionado As String
Dim strBUIAnterior As String
Private Sub ComandoAlta_Click()
    Dim strTextoError As String, strBUI As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strBUI = Nz(Me.ComboBui, "")
    If strBUI = "" Then
        strTextoError = "No se ha indicado para que BUI es el Subsistema"
        Err.Raise 1000
    End If
    '----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormSubSistemaAlta") Then
        DoCmd.Close acForm, "FormSubSistemaAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormSubSistemaAlta", , , , , , strBUI
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un subsistema se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoEditar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strSubSistemaSeleccionado = Nz(Me.ListaSubSistema.Column(0), "")
    strBUISeleccionado = Nz(Me.ComboBui, "")
    If strSubSistemaSeleccionado = "" Then
        strTextoError = "Ha de seleccionar el SubSistema que pretende editar"
        Err.Raise 1000
    End If
    If strBUISeleccionado = "" Then
        strTextoError = "Ha de seleccionar el BUI del subsistema que se pretende editar"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormSubSistemaEdicion") Then
        DoCmd.Close acForm, "FormSubSistemaEdicion", acSaveNo
    End If
    DoCmd.OpenForm "FormSubSistemaEdicion", , , , , , strBUISeleccionado & "|" & strSubSistemaSeleccionado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar un SubSistema se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminar_Click()
    Dim strTextoError As String
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strSubSistemaSeleccionado = Nz(Me.ListaSubSistema.Column(0), "")
    strBUISeleccionado = Nz(Me.ComboBui, "")
    If strSubSistemaSeleccionado = "" Then
        strTextoError = "Ha de seleccionar el SubSistema que pretende editar"
        Err.Raise 1000
    End If
    If strBUISeleccionado = "" Then
        strTextoError = "Ha de seleccionar el BUI del subsistema que se pretende editar"
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente eliminar el SubSistema: " & strSubSistemaSeleccionado & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de SubSistema")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '--------------------------------------------------------
    flag = Eliminar(strBUISeleccionado, strSubSistemaSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ComandoAlta.SetFocus
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar el SubSistema se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoExportarExcel_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_SQL = Me.ListaSubSistema.RowSource
    GenerarConsultasSQL m_SQL, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoLimpiarBUI_Click()
    On Error Resume Next
    strBUIAnterior = ""
    If Nz(Me.ComboBui, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboBui = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComboBui_BeforeUpdate(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBui, "")
    If strValor = "" Then
        Me.ComandoAlta.Enabled = False
    Else
        Me.ComandoAlta.Enabled = True
    End If
    ActualizarLista
    strBUIAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


Private Sub Form_Open(Canel As Integer)
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
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ListaSubSistema_Click()
    On Error Resume Next
    strSubSistemaSeleccionado = Nz(Me.ListaSubSistema.Column(0), "")
    If strSubSistemaSeleccionado = "" Then
        Me.ComandoEditar.Enabled = False
        Me.ComandoEliminar.Enabled = False
    Else
         Me.ComandoEditar.Enabled = True
        Me.ComandoEliminar.Enabled = True
    End If
End Sub
Public Function Eliminar( _
                            strBUI As String, _
                            strSUBSISTEMA As String _
                            ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 29/12/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a eliminar de la tabla TbBUI el nodo con BUI=strNodo
    '       -strBUI<>""
    '       -strBUI Existente
    '       -strSubSistema<>""
    '       -strBUI, strSubSistema Existente
    
    '
    '       -strBUI, strSubSistema no ha de estar previamente en ningún evento-->me.EsteSubSistemaEstaYaEnEvento="No"
    '   -llamada por:
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strBUI = "" Then
        strTextoError = "No se ha indicado el nombre del BUI del Subsistema"
        Err.Raise 1000
    End If
    If strSUBSISTEMA = "" Then
        strTextoError = "No se ha indicado el nombre del SubSistema "
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbEventos", "IDEVENTO", "SUBSISTEMA", strSUBSISTEMA)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "El Subsistema está incluido en al menos un evento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbSubsistemaBui.SUBSISTEMA, TbSubsistemaBui.BUI " & _
            "FROM TbSubsistemaBui " & _
            "WHERE (((TbSubsistemaBui.SUBSISTEMA)='" & strSUBSISTEMA & _
            "') AND ((TbSubsistemaBui.BUI)='" & strBUI & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No está registrado el Subsistema : " & strSUBSISTEMA & " para el BUI llamado: " & strBUI
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Eliminar = "Se ha Eliminado correctamente el registro"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Eliminar = "#ERR" & "|" & strTextoError
End Function
Public Sub ActualizarLista()
    Dim strTextoError As String, intElementos As Integer, lst As ListBox, VarItem As Variant
    Dim strSubSistemaSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strSubSistemaSeleccionado = ""
    Set lst = Me.ListaSubSistema
    strSubSistemaSeleccionadoActualmente = Nz(lst.Column(0), "")
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoExportarExcel.Enabled = False
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    flag = Filtrar(Nz(Me.ComboBui, ""))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Filtrar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_SQL = flag
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount
    If lst.ListCount = 0 Then
        Me.ComandoExportarExcel.Enabled = False
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoExportarExcel.Enabled = True
    lst.Selected(0) = True
    ListaSubSistema_Click
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ActualizarLista la lista de Actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Filtrar(Optional strBUI As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    '       -me.ActualizarLista
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    Dim strParteWhere As String, strSQLParteInicial As String, strSQLFinal As String
    Dim strParteWhereBUI As String, strTextoError As String
    On Error GoTo errores
    strSQLParteInicial = "SELECT TbSubsistemaBui.SUBSISTEMA " & _
                        "FROM TbSubsistemaBui "
    If strBUI = "" Then
        strParteWhere = "WHERE (((TbSubsistemaBui.BUI)='-1'));"
        strSQLFinal = strSQLParteInicial & strParteWhere
        Filtrar = strSQLFinal
        Exit Function
    End If
    If strBUI = "" Then
        strParteWhereBUI = "((TbSubsistemaBui.BUI) Like '*' Or (TbSubsistemaBui.BUI) Is Null)"
    Else
        strParteWhereBUI = "((TbSubsistemaBui.BUI) = '" & strBUI & "')"
    End If
    strParteWhere = "WHERE (" & _
                            strParteWhereBUI & ");"
    strSQLFinal = strSQLParteInicial & strParteWhere
    Filtrar = strSQLFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Filtrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Filtrar = "#ERR" & "|" & strTextoError
End Function

