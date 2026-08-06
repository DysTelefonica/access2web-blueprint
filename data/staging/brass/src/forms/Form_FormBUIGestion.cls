Option Compare Database
Option Explicit
Dim strBUISeleccionado As String
Private Sub ComandoAlta_Click()
    Dim strBUI As String, strTextoError As String
    On Error GoTo errores
    strBUI = Nz(InputBox("Nombre del BUI"), "")
    If strBUI = "" Then
       Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '   -Devuelve:
    '       Alta = strDescriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Alta(strBUI)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Alta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaBUI.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un nodo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    If strBUISeleccionado = "" Then
        strTextoError = "Se ha de seleccionar el Nodo que desea eliminar"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea realmente eliminar el BUI: " & strBUISeleccionado & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de BUI")
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
    '----------------------------------------------------------------
    flag = Eliminar(strBUISeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strBUISeleccionado = ""
    Me.ListaBUI.Requery
    Me.ComandoAlta.SetFocus
    Me.ComandoEliminar.Enabled = False
    '-----------------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------------
    If FormularioAbierto("FormNodoBUIGestion") Then
        Forms("FormNodoBUIGestion").Controls("ListaBuisEnNodo").Requery
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar el BUI se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Canel As Integer)
    On Error Resume Next
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    Me.ComandoEliminar.Enabled = False
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ListaBUI_Click()
    On Error Resume Next
    strBUISeleccionado = Nz(Me.ListaBUI.Column(0), "")
    If strBUISeleccionado = "" Then
        Me.ComandoEliminar.Enabled = False
    Else
        Me.ComandoEliminar.Enabled = True
    End If
End Sub
Public Function Eliminar(strBUI As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 29/12/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a eliminar de la tabla TbBUI el nodo con BUI=strNodo
    '       -strBUI<>""
    '       -strBUI Existente
    '       -strBUI no ha de estar previamente en ningún evento-->me.EsteBUIEstaYaEnEvento="No"
    '       -luego he de eliminar todos los registros para ese bui de la tabla de unión entre nodos y buis
    '   -llamada por:
    
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If strBUI = "" Then
        strTextoError = "No se ha indicado el nombre del BUI"
        Err.Raise 1000
    End If
    '----------------------------------------------------------------
    '   -Devuelve:
    '       Dame = strValor
    '       Dame = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------
    flag = Dame("TbEventos", "IDEvento", "BUI", strBUI)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "El BUI " & strBUI & " forma parte de algún evento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbBUI.bui " & _
                "FROM TbBUI " & _
                "WHERE (((TbBUI.BUI)='" & strBUI & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El BUI " & strBUI & " no está registrado en la tabla de los BUIS"
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "DELETE TbNodoBUI.BUI " & _
            "FROM TbNodoBUI " & _
            "WHERE (((TbNodoBUI.BUI)='" & strBUI & "'));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    Eliminar = "BUI " & strBUI & " borrado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método BUI.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.SetWarnings True
    Eliminar = "#ERR" & "|" & strTextoError
End Function
Public Function Alta(strBUI As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 07/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de alta en la tabla TbBUI el bui strBUI
    '       -strBUI<>""
    '       -strBUI Existente
    '       -strBUI no ha de estar previamente en ningún evento-->me.EsteBUIEstaYaEnEvento="No"
    '   -llamada por:
    
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
    m_SQL = "SELECT TbBUI.BUI " & _
            "FROM TbBUI " & _
            "WHERE (((TbBUI.BUI)='" & strBUI & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "El BUI ya está dado de alta previamente "
            Err.Raise 1000
        End If
        .AddNew
            .Fields("BUI") = strBUI
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Alta = "Se ha dado correctamente de alta el BUI"
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

