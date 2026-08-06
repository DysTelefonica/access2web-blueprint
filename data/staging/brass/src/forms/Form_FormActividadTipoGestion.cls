Option Compare Database
Option Explicit
Dim strCodTipoActividadSeleccionada As String
Private Sub ComandoAlta_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormActividadTipoAlta"
End Sub

Private Sub ComandoEditar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strCodTipoActividadSeleccionada = Nz(Me.ListaTiposDeActividad.Column(0), "")
    If strCodTipoActividadSeleccionada = "" Then
        strTextoError = "Se ha de seleccionar el código del tipo de actividad que desea editar"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------
    If FormularioAbierto("FormActividadTipoEdicion") Then
        DoCmd.Close acForm, "FormActividadTipoEdicion", acSaveNo
    End If
    DoCmd.OpenForm "FormActividadTipoEdicion", , , , , , strCodTipoActividadSeleccionada
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar el tipo de actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If strCodTipoActividadSeleccionada = "" Then
        strTextoError = "Se ha de seleccionar el tipo de Actividad que desea eliminar"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente eliminar el tipo de Actividad: " & strCodTipoActividadSeleccionada & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de tipo de Actividad")
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
    flag = Eliminar(strCodTipoActividadSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaTiposDeActividad.Requery
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
Private Sub ListaTiposDeActividad_click()
    On Error Resume Next
    strCodTipoActividadSeleccionada = Nz(Me.ListaTiposDeActividad.Column(0), "")
    If strCodTipoActividadSeleccionada = "" Then
        Me.ComandoEliminar.Enabled = False
        Me.ComandoEditar.Enabled = False
    Else
        Me.ComandoEliminar.Enabled = True
        Me.ComandoEditar.Enabled = True
    End If
End Sub
Public Function Eliminar(strCodTipoActividad As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 08/02/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a Eliminar de la tabla TbCodActividad el nodo con CODIGO=strNodo
    '       -strCodTipoActividad <>"" y existente
    
    '       -strCodTipoActividad no ha de estar previamente en ninguna actividad-->me.EsteTipoActividadYaEnActividad="No"
    
    '   -llamada por:
    
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If strCodTipoActividad = "" Then
        strTextoError = "No se ha indicado el nombre del tipo de actividad"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbActividades", "Actividad", "Actividad", strCodTipoActividad)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "El Tipo de actividad " & strCodTipoActividad & " forma parte de alguna actividad"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbCodActividad.CODIGO " & _
                "FROM TbCodActividad " & _
                "WHERE (((TbCodActividad.CODIGO)='" & strCodTipoActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El Tipo de actividad " & strCodTipoActividad & " no está registrado en la tabla TbCodActividad"
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Eliminar = "El Tipo de actividad " & strCodTipoActividad & " borrada correctamente"
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


