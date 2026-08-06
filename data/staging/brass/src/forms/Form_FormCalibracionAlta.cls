Option Compare Database
Option Explicit
Dim strIDEquipoCalibracion As String
Private Sub ComandoAlta_Click()
    Dim strFechaInicial As String, strFechaFinal As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDEquipoCalibracion) Then
        strTextoError = "No se puede determinar el equipo"
        Err.Raise 1000
    End If
    strFechaInicial = Nz(Me.FechaCalibracion, "")
    strFechaFinal = Nz(Me.FechaFinCalibracion, "")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AltaCalibracion = strID
    '       AltaCalibracion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = AltaCalibracion(strIDEquipoCalibracion, strFechaInicial, strFechaFinal)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método URLCompletaArchivoAyuda ha devuelto el error: " & vbNewLine & dato(1)
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
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto =True or false
    '           strEstado="1"---->True or false
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    If FormularioAbierto("FormCalibracionGestion") Then
        Forms("FormCalibracionGestion").Controls("ListaFiltrados").Requery
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta una nueva calibración se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Load()
    Dim rcdDatos As DAO.Recordset, strEquipo As String, strModelo As String, strNS As String, strBUI As String, _
        strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    strIDEquipoCalibracion = Nz(Me.OpenArgs(), "")
    If Not IsNumeric(strIDEquipoCalibracion) Then
        strTextoError = "No se ha abierto con los parámetros adecuados"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS,TbEquiposCalibrables.BUI " & _
            "FROM TbEquiposCalibrables " & _
            "WHERE (((TbEquiposCalibrables.IDEquipoCalibrable)=" & strIDEquipoCalibracion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha abierto con los parámetros adecuados"
            Err.Raise 1000
        End If
        strEquipo = Nz(.Fields("Equipo"), "")
        strModelo = Nz(.Fields("Modelo"), "")
        strNS = Nz(.Fields("NS"), "")
        strBUI = Nz(.Fields("BUI"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
     If strModelo <> "" Then
        Me.Modelo = strModelo
    End If
     If strNS <> "" Then
        Me.NS = strNS
    End If
    If strBUI <> "" Then
        Me.BUI = strBUI
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function AltaCalibracion( _
                                  strIDEquipoACalibrar As String, _
                                  strFechaCalibrado As String, _
                                  strFechaFinCalibrado As String _
                                    ) As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
    '       -Form_FormCalibracionAlta.ComandoAlta_Click
    '   -Devuelve:
    '       AltaCalibracion = strID
    '       AltaCalibracion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strID As String, strFechaBajaParaCalibracion As String, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strIDEquipoACalibrar) Then
        strTextoError = "No se ha indicado el equipo a calibrar"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaCalibrado) Then
        strTextoError = "Es obligatoria la fecha de calibrado"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaFinCalibrado) Then
        strTextoError = "Es obligatoria la fecha de fin de calibrado"
        Err.Raise 1000
    End If
    If CDate(strFechaFinCalibrado) < CDate(strFechaCalibrado) Then
        strTextoError = "La fecha de fin de calibrado ha de ser igual o mayor que la de calibrado"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbEquiposCalibrables", "FechaBajaParaCalibracion", "IDEquipoCalibrable", strIDEquipoACalibrar)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaBajaParaCalibracion = flag
    If IsDate(strFechaBajaParaCalibracion) Then
        strTextoError = "Este equipo está fuera de la calibración desde: " & strFechaBajaParaCalibracion
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameID("TbEquiposCalibrablesFechas", "IDCalibracion")
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
    m_SQL = "SELECT TbEquiposCalibrablesFechas.* " & _
            "FROM TbEquiposCalibrablesFechas " & _
            "WHERE (((TbEquiposCalibrablesFechas.IDCalibracion)=" & strID & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función DameID ha devuelto un valor con formato desconocido"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDCalibracion") = strID
            .Fields("IDEquipoCalibrable") = strIDEquipoACalibrar
            .Fields("FechaCalibrado") = strFechaCalibrado
            .Fields("FechaFinCalibrado") = strFechaFinCalibrado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    AltaCalibracion = strID
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Calibracion.AltaCalibracion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    AltaCalibracion = "#ERR" & "|" & strTextoError
End Function









