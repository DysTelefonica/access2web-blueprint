Option Compare Database
Option Explicit
Private Sub ComandoAlta_Click()
    Dim strEmpresa As String, strFechaAlta As String, strDESCRIPCION As String, strFechaFin As String, strImporte As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strEmpresa = Nz(Me.ComboEmpresa, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strDESCRIPCION = Nz(Me.Descripcion, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strFechaFin = Nz(Me.FechaFin, "")
    strImporte = Nz(Me.Importe, "")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Alta = strID
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Alta(strEmpresa, strFechaAlta, strDESCRIPCION, strImporte, strFechaFin)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormSubContratacionGestion") Then
        Form_FormSubContratacionGestion.ActualizarLista
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta la subcontratación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
                        strEmpresa As String, _
                        strFechaInicial As String, _
                        strDESCRIPCION As String, _
                        strImporte As String, _
                        Optional strFechaFin As String _
                        ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 13/06/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbSubcontrataciones
    '   -Llamada desde
    '       -Form_FormSubContratacionAlta.ComandoAlta_Click
    '   -Devuelve:
    '       Alta = strID
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strID As String, strTextoError As String
    On Error GoTo errores
    If strEmpresa = "" Then
        strTextoError = "Se ha de introducir el nombre de la empresa subcontratada"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaInicial) Then
        strTextoError = "Se ha de introducir una fecha inicial válida"
        Err.Raise 1000
    End If
    If Not IsNumeric(strImporte) Then
        strTextoError = "Se ha de introducir un importe válido"
        Err.Raise 1000
    End If
    If strFechaFin <> "" Then
        If CDate(strFechaFin) < CDate(strFechaInicial) Then
            strTextoError = "Si se introduce una fecha de fin ha de ser mayor o igual que la inicial"
            Err.Raise 1000
        End If
    End If
    '------------------------------------------
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameID("TbSubcontrataciones", "IDSubcontratacion")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameID ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strID = flag
    If Not IsNumeric(strID) Then
        strTextoError = "la función dameid ha devuelto un valor con formato desconocido"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbSubcontrataciones.* " & _
            "FROM TbSubcontrataciones " & _
            "WHERE (((TbSubcontrataciones.IDSubcontratacion)=" & strID & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "la función dameid ha devuelto un valor ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDSubcontratacion") = strID
            .Fields("FechaAlta") = strFechaInicial
            .Fields("Empresa") = strEmpresa
            If strDESCRIPCION <> "" Then
                .Fields("Descripcion") = strDESCRIPCION
            End If
            If IsDate(strFechaFin) Then
                 .Fields("FechaFin") = strFechaFin
            End If
            If IsNumeric(strImporte) Then
                .Fields("Importe") = strImporte
            End If
        .Update
    End With
    Alta = strID
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





