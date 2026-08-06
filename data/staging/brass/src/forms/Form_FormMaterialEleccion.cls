Option Compare Database
Option Explicit
Dim strIDEvento As String, strIDActividad As String, strArgumento As String
Dim strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, strEquipo As String
Private Sub cmdSalir_Click()
     On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAlta_Click()
    Dim rcdDatos As DAO.Recordset, strEsReparacion As String, strTextoError As String, strNombreFormulario As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    strEsReparacion = Nz(Me.ComboEsReparacion, "")
    If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        strTextoError = "No se ha elegido sí o no"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If strEsReparacion = "Sí" Then
        
        '-----------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '-----------------------------------------------------
        If FormularioAbierto("FormMaterialAltaReparacion") Then
            DoCmd.Close acForm, "FormMaterialAltaReparacion", acSaveNo
        End If
        If UBound(dato) = 1 Then
            m_SQL = "SELECT  TbEventos.BUI, TbEventos.SUBSISTEMA, TbEventos.IDEquipo, TbEquipos.Equipo " & _
                        "FROM TbEventos LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                    "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
        
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If .EOF Then
                    strTextoError = "Ne está registrado el evento: " & strIDEvento
                    Err.Raise 1000
                End If
                strBUI = Nz(.Fields("BUI"), "")
                strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
                strIDEquipo = Nz(.Fields("IDEquipo"), "")
                strEquipo = Nz(.Fields("Equipo"), "")
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
            'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipo & "|" & strIDEquipo & "|" & strIDEvento & "|" & strIDActividad
            strArgumento = strBUI & "|" & strSUBSISTEMA & "|" & strIDEquipo & "|" & strIDEquipo & "|" & strIDEvento & "|" & strIDActividad
        ElseIf UBound(dato) = 3 Then
            'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipo & "|" & strIDEquipo
            strArgumento = strBUI & "|" & strSUBSISTEMA & "|" & strIDEquipo & "|" & strIDEquipo
        End If
        DoCmd.OpenForm "FormMaterialAltaReparacion", , , , , , strArgumento
    Else
         m_SQL = "SELECT  TbEventos.BUI, TbEventos.SUBSISTEMA, TbEventos.IDEquipo, TbEquipos.Equipo " & _
                    "FROM TbEventos LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "Ne está registrado el evento: " & strIDEvento
                Err.Raise 1000
            End If
            strBUI = Nz(.Fields("BUI"), "")
            strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
            strIDEquipo = Nz(.Fields("IDEquipo"), "")
            strEquipo = Nz(.Fields("Equipo"), "")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        '-----------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '-----------------------------------------------------
        If FormularioAbierto("FormMaterialRepuestoAlta") Then
            DoCmd.Close acForm, "FormMaterialRepuestoAlta", acSaveNo
        End If
        If UBound(dato) <> 1 Then
            strTextoError = "No venimos de una actividad"
            Err.Raise 1000
        End If
        'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipo & "|" & strIDEquipo & "|" & strIDEvento & "|" & strIDActividad
        strArgumento = strBUI & "|" & strSUBSISTEMA & "|" & strIDEquipo & "|" & strEquipo & "|" & strIDEvento & "|" & strIDActividad
        DoCmd.OpenForm "FormMaterialRepuestoAlta", , , , , , strArgumento
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    '   0) Desde cero-->Sin argumentos
    '   1) Desde Actividad-->strArgumento=strIDEvento;strIDActividad
    '   2) Desde Equipo elegido-->strArgumento=strBUI;strSubsistema;strIDEquipo;strIdEquipo
    
    strArgumento = Nz(Me.OpenArgs, "")
    If strArgumento = "" Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 1 And UBound(dato) <> 3 Then
        strTextoError = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    If UBound(dato) = 1 Then
        strIDEvento = CStr(dato(0))
        strIDActividad = CStr(dato(1))
        If strIDEvento = "" Or strIDActividad = "" Then
            strTextoError = "No se ha abierto el formulario con los datos necesarios"
            Err.Raise 1000
        End If
    Else
        strBUI = CStr(dato(0))
        strSUBSISTEMA = CStr(dato(1))
        strIDEquipo = CStr(dato(2))
        strIDEquipo = CStr(dato(3))
        If strBUI = "" Or strSUBSISTEMA = "" Or strIDEquipo = "" Or strIDEquipo = "" Then
            strTextoError = "No se ha abierto el formulario con los datos necesarios"
            Err.Raise 1000
        End If
    End If
    Me.Caption = Me.lblTitulo1.Caption
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
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub






