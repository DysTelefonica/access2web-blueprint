Option Compare Database
Option Explicit
Dim strNodoSeleccionado As String
Private Sub ComandoAñadirBUI_Click()
    Dim strBUI As String, strNodo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNodo = Nz(Me.ComboNodos, "")
    strBUI = Nz(Me.ComboBuiParaNodo, "")
    If strNodo = "" Then
        strTextoError = "No se puede determinar para qué NODO se quiere añadir el BUI"
        Err.Raise 1000
    End If
    If strBUI = "" Then
        strTextoError = "Se ha de seleccionar del desplegable un BUI a añadir"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AsignarBUIANodo = strDescriptivo
    '       AsignarBUIANodo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = AsignarBUIANodo(strBUI, strNodo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AsignarBUIANodo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaBuisEnNodo.Requery
    Me.ComboBuiParaNodo = Null
    Me.ComboBuiParaNodo.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar asignar un bui a un nodo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoIrABUI_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormBUIGestion"
End Sub
Private Sub ComandoIrANodos_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormNodoGestion"
End Sub
Private Sub ComandoQuitarBUI_Click()
    Dim strBUI As String, strNodo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strBUI = Nz(Me.ListaBuisEnNodo.Column(0), "")
    strNodo = Nz(Me.ListaBuisEnNodo.Column(1), "")
    If strNodo = "" Then
        strTextoError = "No se puede determinar para qué NODO se quiere desasignar el BUI"
        Err.Raise 1000
    End If
    If strBUI = "" Then
        strTextoError = "Se ha de seleccionar de la lista un BUI a desasignar"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DesAsignarBUIANodo = strDescriptivo
    '       DesAsignarBUIANodo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = DesAsignarBUIANodo(strBUI, strNodo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DesAsignarBUIANodo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaBuisEnNodo.Requery
    Me.ComboBuiParaNodo.Requery
    Me.ComandoAñadirBUI.SetFocus
    Me.ComandoQuitarBUI.Enabled = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar Desasignar un bui a un nodo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboBuiParaNodo_Enter()
    On Error Resume Next
    Me.ComboBuiParaNodo.Requery
End Sub
Private Sub ComboNodos_Change()
    Dim strNodo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNodo = Nz(Me.ComboNodos, "")
    If strNodo = "" Then
        strNodo = "-1"
    End If
    m_SQL = "SELECT TbNodoBUI.BUI, TbNodoBUI.NODO " & _
            "FROM TbNodoBUI " & _
            "WHERE (((TbNodoBUI.NODO)='" & strNodo & "')) " & _
            "ORDER BY TbNodoBUI.BUI;"
    Me.ListaBuisEnNodo.RowSource = m_SQL
    Me.ListaBuisEnNodo.Requery
    Me.ComandoQuitarBUI.Enabled = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del combo del NODO se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboNodos_Enter()
    On Error Resume Next
    Me.ComboNodos.Requery
End Sub

Private Sub ComboNodos_Exit(Cancel As Integer)
    Dim strNodo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNodo = Nz(Me.ComboNodos, "")
    If strNodo = "" Then
        strNodo = "-1"
    End If
    m_SQL = "SELECT TbNodoBUI.BUI, TbNodoBUI.NODO " & _
            "FROM TbNodoBUI " & _
            "WHERE (((TbNodoBUI.NODO)='" & strNodo & "')) " & _
            "ORDER BY TbNodoBUI.BUI;"
    Me.ListaBuisEnNodo.RowSource = m_SQL
    Me.ListaBuisEnNodo.Requery
    Me.ComandoQuitarBUI.Enabled = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del combo del NODO se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
Public Function AsignarBUIANodo( _
                                strBUI As String, _
                                strNodo As String _
                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 29/12/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de AsignarBUIANodo un registro en la tabla TbNodoBUI
    '       -strBUI<>"" ha de existir previamente
    '       -strNodo <>"" ha de existir
    '   -llamada por:
    '       -Form_FormNodoBUIGestion.ComandoAñadirBuiANodo_Click
    '   -Devuelve:
    '       AsignarBUIANodo = Descriptivo
    '       AsignarBUIANodo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strBUI = "" Then
        strTextoError = "No se ha indicado el nombre del BUI"
        Err.Raise 1000
    End If
    If strNodo = "" Then
        strTextoError = "No se ha indicado el nombre del NODO"
        Err.Raise 1000
    End If
    '-----------
    ' ahora vinculación de nodo y bui
    '---------------------
     m_SQL = "SELECT TbNodoBUI.NODO, TbNodoBUI.BUI " & _
                "FROM TbNodoBUI " & _
                "WHERE (((TbNodoBUI.NODO)='" & strNodo & _
                "') AND ((TbNodoBUI.BUI)='" & strBUI & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya está registrado para el Nodo: " & strNodo & " el BUI: " & strBUI
            Err.Raise 1000
        End If
        .AddNew
            .Fields("BUI") = strBUI
            .Fields("NODO") = strNodo
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    AsignarBUIANodo = "Se ha dado de AsignarBUIANodo correctamente el BUI llamado: " & strBUI & " para el nodo: " & strNodo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método AsignarBUIANodo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    AsignarBUIANodo = "#ERR" & "|" & strTextoError
End Function
Public Function DesAsignarBUIANodo( _
                                    strBUI As String, _
                                    strNodo As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 29/12/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de DesAsignarBUIANodo un registro en la tabla TbNodoBUI
    '       -strBUI<>"" ha de existir previamente
    '       -strNodo <>"" ha de existir
    '   -llamada por:
    
    '   -Devuelve:
    '       DesAsignarBUIANodo = Descriptivo
    '       DesAsignarBUIANodo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If strBUI = "" Then
        strTextoError = "No se ha indicado el nombre del BUI"
        Err.Raise 1000
    End If
    If strNodo = "" Then
        strTextoError = "No se ha indicado el nombre del NODO"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbEventos", "BUI", "BUI", strBUI)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "La pareja Nodo: " & strNodo & " BUI: " & strBUI & " se ha registrado en al menos un evento"
        Err.Raise 1000
    End If
    '-----------
    ' ahora eliminación de la relación de nodo y bui
    '---------------------
     m_SQL = "SELECT TbNodoBUI.NODO, TbNodoBUI.BUI " & _
                "FROM TbNodoBUI " & _
                "WHERE (((TbNodoBUI.NODO)='" & strNodo & _
                "') AND ((TbNodoBUI.BUI)='" & strBUI & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "Ya está registrado para el Nodo: " & strNodo & " el BUI: " & strBUI
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DesAsignarBUIANodo = "Se Desasignado correctamente el BUI llamado: " & strBUI & " para el nodo: " & strNodo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DesAsignarBUIANodo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DesAsignarBUIANodo = "#ERR" & "|" & strTextoError
End Function
Private Sub ListaBuisEnNodo_Click()
    On Error Resume Next
    strNodoSeleccionado = Nz(Me.ListaBuisEnNodo.Column(0), "")
    If strNodoSeleccionado = "" Then
        Me.ComandoQuitarBUI.Enabled = False
    Else
        Me.ComandoQuitarBUI.Enabled = True
    End If
End Sub