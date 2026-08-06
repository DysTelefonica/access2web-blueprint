Option Compare Database
Option Explicit
Private Sub ComandoAlta_Click()
    Dim strTipoActividad As String, strCodTipoActividad As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strCodTipoActividad = Nz(Me.CodTipo, "")
    strTipoActividad = Nz(Me.TIPOACTIVIDAD, "")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Alta = strDescriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Alta(strCodTipoActividad, strTipoActividad)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Alta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormActividadTipoGestion") Then
        Forms("FormActividadTipoGestion").Controls("ListaTiposDeActividad").Requery
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
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

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Load()
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
Public Function Alta( _
                        strCodTipoActividad As String, _
                        strTipoActividad As String _
                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 07/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a dar de Alta strCodTipoActividad,strTipoActividad en la tabla TbCodActividad
    '       -strCodTipoActividad <>"" y no existente
    '       -strTipoActividad <>"" y no existente
    
    '   -llamada por:
    
    '   -Devuelve:
    '       Alta = Descriptivo
    '       Alta = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If strCodTipoActividad = "" Then
        strTextoError = "No se ha indicado el código de la actividad"
        Err.Raise 1000
    End If
    If strTipoActividad = "" Then
        strTextoError = "No se ha indicado el tipo de actividad"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbCodActividad", "CODIGO", "CODIGO", strCodTipoActividad)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "Ya existe un código de actividad como ese"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbCodActividad", "ACTIVIDADNOPROGRAMADA", "ACTIVIDADNOPROGRAMADA", strTipoActividad)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "Ya existe un tipo de actividad como esa"
        Err.Raise 1000
    End If
    m_SQL = "TbCodActividad"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("CODIGO") = strCodTipoActividad
            .Fields("ACTIVIDADNOPROGRAMADA") = strTipoActividad
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Alta = "Se ha dado de Alta el Tipo de Actividad correctamente "
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






