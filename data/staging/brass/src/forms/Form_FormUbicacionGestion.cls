Option Compare Database
Option Explicit
Dim strUbicacionSeleccionada As String
Private Sub ComandoAlta_Click()
    Dim strUbicacion As String, strTextoError As String
    On Error GoTo errores
    strUbicacion = Nz(InputBox("Nombre de la Ubicación"), "")
    If strUbicacion = "" Then
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
    flag = Alta(strUbicacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Alta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaUbicacion.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta una ubicación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If strUbicacionSeleccionada = "" Then
        strTextoError = "Se ha de seleccionar la ubicación que desea eliminar"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea realmente eliminar la ubicación: " & strUbicacionSeleccionada & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de ubicación")
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
    flag = Eliminar(strUbicacionSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strUbicacionSeleccionada = ""
    Me.ListaUbicacion.Requery
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
Public Function Alta(strUbicacion As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 07/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbUbicacion
    '       -strUbicacion <>""
    '       -strUbicacion no ha de existir previamente
    
    '   -llamada por:
    
    '   -Devuelve:
    '       Alta = Descriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strUbicacion = "" Then
        strTextoError = "No se ha indicado el nombre de la ubicación"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbUbicacion.UBICACION " & _
            "FROM TbUbicacion " & _
            "WHERE (((TbUbicacion.UBICACION)='" & strUbicacion & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya existe una ubicación registrada con el nombre: " & strUbicacion
            Err.Raise 1000
        End If
        .AddNew
            .Fields("UBICACION") = strUbicacion
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Alta = "Se ha dado de alta correctamente el originador llamado: " & strUbicacion
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Ubicacion.Alta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Alta = "#ERR" & "|" & strTextoError
End Function
Public Function Eliminar(strUbicacion As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 07/2/13
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a eliminar de la tabla TbUbicacion el nodo con Ubicacion=strUbicacion
    '       -strUbicacion <>""
    '       -strUbicacion Existente
    '       -strUbicacion no ha de estar previamente en ningún evento-->me.EstaUbicacionEstaYaEnActividad="No"
    '   -llamada por:
    
    '   -Devuelve:
    '       Eliminar = Descriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strUbicacion = "" Then
        strTextoError = "No se ha indicado el nombre de la ubicación"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbActividades", "Ubicacion", "Ubicacion", strUbicacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "La ubicación " & strUbicacion & " forma parte de alguna actividad"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbUbicacion.UBICACION " & _
                "FROM TbUbicacion " & _
                "WHERE (((TbUbicacion.UBICACION)='" & strUbicacion & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "La ubicación " & strUbicacion & " no está registrada en la tabla de las Ubicaciones"
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Eliminar = "Ubicación " & strUbicacion & " borrada correctamente"
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

Private Sub ListaUbicacion_Click()
    On Error Resume Next
    strUbicacionSeleccionada = Nz(Me.ListaUbicacion.Column(0), "")
    If strUbicacionSeleccionada = "" Then
        Me.ComandoEliminar.Enabled = False
    Else
        Me.ComandoEliminar.Enabled = True
    End If
End Sub