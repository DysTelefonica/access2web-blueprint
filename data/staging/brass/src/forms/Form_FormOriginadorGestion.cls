Option Compare Database
Option Explicit
Dim strOriginadorSeleccionado As String
Private Sub ComandoAlta_Click()
    Dim strORIGINADOR As String, strTextoError As String
    On Error GoTo errores
    strORIGINADOR = Nz(InputBox("Nombre del Originador"), "")
    If strORIGINADOR = "" Then
       Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       Alta = strDescriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Alta(strORIGINADOR)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Alta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaOriginadores.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un originador se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If strOriginadorSeleccionado = "" Then
        strTextoError = "Se ha de seleccionar el Originador que desea eliminar"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea realmente eliminar el Originador: " & strOriginadorSeleccionado & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de Originador")
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
    flag = Eliminar(strOriginadorSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaOriginadores.Requery
    Me.ComandoAlta.SetFocus
    Me.ComandoEliminar.Enabled = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar la ubicación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
Private Sub ListaOriginadores_Click()
    On Error Resume Next
    strOriginadorSeleccionado = Nz(Me.ListaOriginadores.Column(0), "")
    If strOriginadorSeleccionado = "" Then
        Me.ComandoEliminar.Enabled = False
    Else
        Me.ComandoEliminar.Enabled = True
    End If
End Sub
Public Function Alta(strORIGINADOR As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 06/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbOriginador
    '       -strOriginador<>""
    '       -strOriginador no ha de existir previamente
    
    '   -llamada por:
    
    '   -Devuelve:
    '       Alta = Descriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strORIGINADOR = "" Then
        strTextoError = "No se ha indicado el nombre del nodo"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbOriginador.ORIGINADOR " & _
            "FROM TbOriginador " & _
            "WHERE (((TbOriginador.ORIGINADOR)='" & strORIGINADOR & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya existe un Originador registrado con el nombre: " & strORIGINADOR
            Err.Raise 1000
        End If
        .AddNew
            .Fields("ORIGINADOR") = strORIGINADOR
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Alta = "Se ha dado de alta correctamente el originador llamado: " & strORIGINADOR
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Originador.Alta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Alta = "#ERR" & "|" & strTextoError
End Function
Public Function Eliminar(strORIGINADOR As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 06/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a eliminar de la tabla TbOriginador el nodo con Originador=strOriginador
    '       -strOriginador<>""
    '       -strOriginador Existente
    '       -strOriginador no ha de estar previamente en ningún evento-->me.EsteOriginadorEstaYaEnEvento="No"
    '   -llamada por:
    
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strORIGINADOR = "" Then
        strTextoError = "No se ha indicado el nombre del originador"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbEventos", "Originador", "Originador", strORIGINADOR)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "El originador " & strORIGINADOR & " forma parte de algún evento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbOriginador.ORIGINADOR " & _
                "FROM TbOriginador " & _
                "WHERE (((TbOriginador.ORIGINADOR)='" & strORIGINADOR & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El Originador " & strORIGINADOR & " no está registrado en la tabla de los Originadores"
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Eliminar = "Originador " & strORIGINADOR & " borrado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Originador.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Eliminar = "#ERR" & "|" & strTextoError
End Function

