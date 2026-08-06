Option Compare Database
Option Explicit
Dim strTipoEventoSeleccionado As String
Private Sub ComandoAlta_Click()
    Dim strTipoEvento As String, strTextoError As String
    On Error GoTo errores
    strTipoEvento = Nz(InputBox("Nombre el Tipo de Evento"), "")
    If strTipoEvento = "" Then
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
    flag = Alta(strTipoEvento)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Alta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaTipoEventos.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    strTipoEventoSeleccionado = Nz(Me.ListaTipoEventos.Column(0), "")
    If strTipoEventoSeleccionado = "" Then
        strTextoError = "Se ha de seleccionar el Tipo de Evento que desea eliminar"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea realmente eliminar el Tipo de Evento: " & strTipoEventoSeleccionado & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de Tipo Evento")
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
    flag = Eliminar(strTipoEventoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaTipoEventos.Requery
    Me.ComandoAlta.SetFocus
    Me.ComandoEliminar.Enabled = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
Private Sub ListaTipoEventos_Click()
    On Error Resume Next
    strTipoEventoSeleccionado = Nz(Me.ListaTipoEventos.Column(0), "")
    If strTipoEventoSeleccionado = "" Then
        Me.ComandoEliminar.Enabled = False
    Else
        Me.ComandoEliminar.Enabled = True
    End If
End Sub
Public Function Alta(strTipoEvento As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 06/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbTipoEvento
    '       -strTipoEvento<>""
    '       -strTipoEvento no ha de existir previamente
    
    '   -llamada por:
    
    '   -Devuelve:
    '       Alta = Descriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
   
    If strTipoEvento = "" Then
        strTextoError = "No se ha indicado el nombre del Tipo de Evento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbTipoEvento.TIPOEVENTO " & _
            "FROM TbTipoEvento " & _
            "WHERE (((TbTipoEvento.TIPOEVENTO)='" & strTipoEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya existe un Tipo Evento registrado con el nombre: " & strTipoEvento
            Err.Raise 1000
        End If
        .AddNew
            .Fields("TipoEvento") = strTipoEvento
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Alta = "Se ha dado de alta correctamente el Tipo Evento llamado: " & strTipoEvento
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Tipo Alta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Alta = "#ERR" & "|" & strTextoError
End Function
Public Function Eliminar(strTipoEvento As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 06/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a eliminar de la tabla TbTipoEvento el Tipo de Evento con Originador=strOriginador
    '       -strTipoEvento<>""
    '       -strTipoEvento Existente
    '       -strTipoEvento no ha de estar previamente en ningún evento-->me.EsteTipoEventoEstaYaEnEvento="No"
    '   -llamada por:
    
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If strTipoEvento = "" Then
        strTextoError = "No se ha indicado el nombre del Tipo Evento"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbEventos", "TIPOEVENTO", "TIPOEVENTO", strTipoEvento)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "El Tipo Evento " & strTipoEvento & " forma parte de algún evento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbTipoEvento.TipoEvento " & _
                "FROM TbTipoEvento " & _
                "WHERE (((TbTipoEvento.TipoEvento)='" & strTipoEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El Tipo Evento " & strTipoEvento & " no está registrado en la tabla de los Tipos de eventos"
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Eliminar = "Tipo Evento " & strTipoEvento & " borrado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Tipo Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Eliminar = "#ERR" & "|" & strTextoError
End Function

