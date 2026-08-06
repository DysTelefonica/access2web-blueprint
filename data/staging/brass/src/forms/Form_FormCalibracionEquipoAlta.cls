Option Compare Database
Option Explicit
Private Sub ComandoAlta_Click()
    Dim strEquipo As String, strModelo As String, strNS As String, strBUI As String, strObservaciones As String, _
        strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strEquipo = Nz(Me.Equipo, "")
    strModelo = Nz(Me.Modelo, "")
    strNS = Nz(Me.NS, "")
    strBUI = Nz(Me.ComboBui, "")
    strObservaciones = Nz(Me.Observaciones, "")
    '---------------------------------------------------
    '   -Devuelve:
    '       AltaEquipoACalibrar = strID
    '       AltaEquipoACalibrar = "#ERR" & "|" & strTextoError
    '---------------------------------------------------
    flag = AltaEquipoACalibrar(strBUI, strEquipo, strModelo, strNS, strObservaciones)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AltaEquipoACalibrar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto =True or false
    '           strEstado="1"---->True or false
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    If FormularioAbierto("FormCalibracionEquipoGestion") Then
        Form_FormCalibracionEquipoGestion.ActualizarLista
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un nuevo equipo calibrable se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
Public Function AltaEquipoACalibrar( _
                                        strBUI As String, _
                                        strEquipo As String, _
                                        Optional strModelo As String, _
                                        Optional strNS As String, _
                                        Optional strObservaciones As String _
                                        ) As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
    
    '   -Devuelve:
        
    '       AltaEquipoACalibrar = strID
    '       AltaEquipoACalibrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strID As String, strTextoError As String
    On Error GoTo errores
    
    If strBUI = "" Then
        strTextoError = "Es obligatorio indicar el BUI"
        Err.Raise 1000
    End If
    If strEquipo = "" Then
        strTextoError = "Es obligatorio indicar el Equipo"
        Err.Raise 1000
    End If
    If strModelo <> "" And strNS <> "" Then
        m_SQL = "SELECT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS, TbEquiposCalibrables.BUI " & _
                "FROM TbEquiposCalibrables " & _
                "WHERE (((TbEquiposCalibrables.Equipo)='" & strEquipo & _
                "') AND ((TbEquiposCalibrables.Modelo)='" & strModelo & _
                "') AND ((TbEquiposCalibrables.NS)='" & strNS & _
                "') AND ((TbEquiposCalibrables.BUI)='" & strBUI & "'));"
    ElseIf strModelo = "" And strNS = "" Then
        m_SQL = "SELECT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.BUI " & _
                "FROM TbEquiposCalibrables " & _
                "WHERE (((TbEquiposCalibrables.Equipo)='" & strEquipo & _
                "') AND ((TbEquiposCalibrables.BUI)='" & strBUI & "'));"
    ElseIf strModelo <> "" And strNS = "" Then
        m_SQL = "SELECT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS, TbEquiposCalibrables.BUI " & _
                "FROM TbEquiposCalibrables " & _
                "WHERE (((TbEquiposCalibrables.Equipo)='" & strEquipo & _
                "') AND ((TbEquiposCalibrables.Modelo)='" & strModelo & _
                "') AND ((TbEquiposCalibrables.BUI)='" & strBUI & "'));"
    ElseIf strModelo = "" And strNS <> "" Then
        m_SQL = "SELECT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS, TbEquiposCalibrables.BUI " & _
                "FROM TbEquiposCalibrables " & _
                "WHERE (((TbEquiposCalibrables.Equipo)='" & strEquipo & _
                "') AND ((TbEquiposCalibrables.NS)='" & strNS & _
                "') AND ((TbEquiposCalibrables.BUI)='" & strBUI & "'));"
    End If
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Para " & strBUI & " el equipo " & strEquipo & " ya ha sido registrado previamente"
            Err.Raise 1000
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    '------------------------------------------
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameID("TbEquiposCalibrables", "IDEquipoCalibrable")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameID ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strID = flag
    If Not IsNumeric(strID) Then
        strTextoError = "La función DameID ha devuelto un valor con formato desconocido"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrables.* " & _
            "FROM TbEquiposCalibrables " & _
            "WHERE (((TbEquiposCalibrables.IDEquipoCalibrable)=" & strID & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función DameID ha devuelto un valor con formato desconocido"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDEquipoCalibrable") = strID
            .Fields("Equipo") = strEquipo
            .Fields("BUI") = strBUI
            If strModelo <> "" Then
                .Fields("Modelo") = strModelo
            End If
            If strNS <> "" Then
                .Fields("NS") = strNS
            End If
            If strObservaciones <> "" Then
                .Fields("Observaciones") = strObservaciones
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    AltaEquipoACalibrar = strID
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Calibracion.AltaEquipoACalibrar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    AltaEquipoACalibrar = "#ERR" & "|" & strTextoError
End Function





