Option Compare Database
Option Explicit
Dim strNodoSeleccionado As String
Private Sub ComandoAlta_Click()
    Dim strNodo As String, strTextoError As String
    On Error GoTo errores
    strNodo = Nz(InputBox("Nombre del Nodo"), "")
    If strNodo = "" Then
       Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '   -Devuelve:
    '       Alta = strDescriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Alta(strNodo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Alta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaNodos.Requery
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
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNodoSeleccionado = Nz(Me.ListaNodos.Column(0), "")
    If strNodoSeleccionado = "" Then
        strTextoError = "Se ha de seleccionar el Nodo que desea eliminar"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente eliminar el Nodo: " & strNodoSeleccionado & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de Nodo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '------------------------------------------------------
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '------------------------------------------------------
    flag = Eliminar(strNodoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ComandoEliminar.Enabled = False
    Me.ListaNodos.Requery
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------
    If FormularioAbierto("FormNodoBUIGestion") Then
        Forms("FormNodoBUIGestion").Controls("ListaNodos").Requery
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar el nodo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
Private Sub ListaNodos_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNodoSeleccionado = Nz(Me.ListaNodos.Column(0), "")
    If strNodoSeleccionado = "" Then
        Me.ComandoEliminar.Enabled = False
    Else
        Me.ComandoEliminar.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Public Function Eliminar(strNodo As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 28/12/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a eliminar de la tabla TbNodos el nodo con Nodo=strNodo
    '       -strNodo<>""
    '       -strNodo Existente
    '       -strNodo no ha de estar previamente en ningún evento-->me.EsteNodoEstaYaEnEvento="No"
    '       -luego he de eliminar todos los registros de la tabla de unión entre NOdos y buis que tengan ese nodo
    '   -llamada por:
    '       -Form_FormNodoGestion.ComandoEliminar_Click
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strTextoError As String
    On Error GoTo errores
    If strNodo = "" Then
        strTextoError = "No se ha indicado el nombre del nodo"
        Err.Raise 1000
    End If
    '------------------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '------------------------------------------------------
    flag = Dame("TbEventos", "IDEvento", "NODO", strNodo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "El nodo se ha empleado al menos en un evento: " & vbNewLine & flag
        Err.Raise 1000
    End If
    m_SQL = "DELETE TbNodos.NODO " & _
                "FROM TbNodos " & _
                "WHERE (((TbNodos.NODO)='" & strNodo & "'));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    m_SQL = "DELETE TbNodoBUI.NODO " & _
            "FROM TbNodoBUI " & _
            "WHERE (((TbNodoBUI.NODO)='" & strNodo & "'));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    Eliminar = "Nodo " & strNodo & " borrado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Nodo.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.SetWarnings True
    Eliminar = "#ERR" & "|" & strTextoError
End Function
Public Function Alta(strNodo As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 28/12/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbNodos
    '       -strNodo<>""
    '       -strNodo no ha de existir previamente
    
    '   -llamada por:
    '       -Form_FormNodoGestion.ComandoAlta_Click
    '   -Devuelve:
    '       Alta = Descriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strNodo = "" Then
        strTextoError = "No se ha indicado el nombre del nodo"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbNodos.NODO " & _
            "FROM TbNodos " & _
            "WHERE (((TbNodos.NODO)='" & strNodo & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya existe un nodo registrado con el nombre: " & strNodo
            Err.Raise 1000
        End If
        .AddNew
            .Fields("NODO") = strNodo
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Alta = "Se ha dado de alta correctamente el nodo llamado: " & strNodo
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



