Option Compare Database
Option Explicit
Dim strValorAnteriorBUI As String, strValorAnteriorSubSistema As String, strValorAnteriorIDEquipo As String
Private WithEvents m_FormEquipoMedida As Form_FormEventoEqMedida

Private Function Registrar() As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = strDescriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '------------------------------------------------------------------------------------
    Dim strIDEvento As String, strValorComboEquipo As String
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strPM As String, strTipoEvento As String, strCriticidad As String, strAliasTecnico As String, _
        strORIGINADOR As String, strDESCRIPCION As String, strFechaAlta As String, strContacto As String, strIDEquipo As String, strNotas As String, _
        strFechaRecepcionNotificacion As String, strFechaInicioContactoCliente As String, strIncidenciaAveriaOReparacion As String, _
        strUrgente As String, strEventoConServicioAfectado As String, strTextoError As String
    On Error GoTo errores
    strNodo = Nz(Me.ComboBui.Column(1), "")
    strBUI = Nz(Me.ComboBui.Column(0), "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    strPM = Nz(Me.PMPR, "")
    strTipoEvento = Nz(Me.ComboTipoEvento, "")
    strCriticidad = Nz(Me.ComboCriticidad, "")
    strAliasTecnico = Nz(Me.ComboTecnico.Column(0), "")
    strORIGINADOR = Nz(Me.ComboOriginador, "")
    strDESCRIPCION = Nz(Me.DescripcionEvento, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strContacto = Nz(Me.Contacto, "")
    strNotas = ""
    strFechaRecepcionNotificacion = Nz(Me.FechaRecepcionNotificacion, "")
    strFechaInicioContactoCliente = Nz(Me.FechaInicioContactoCliente, "")
    strIncidenciaAveriaOReparacion = Nz(Me.cmbIncidenciaAveriaOReparacion, "")
    If Me.chkUrgente = True Then
        strUrgente = "SI"
    ElseIf Me.chkUrgente = False Then
        strUrgente = "NO"
    Else
        strUrgente = ""
    End If
    strEventoConServicioAfectado = Nz(Me.cmbEventoConServicioAfectado, "")
    strValorComboEquipo = Nz(Me.ComboIDEquipo, "")
    If IsNumeric(strValorComboEquipo) Then
        strIDEquipo = strValorComboEquipo
    Else
        '------------------------------------------------------------------------------------
        '   -Devuelve:
        '       DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre = strIDEquipo
        '       DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre = "#ERR" & "|" & strTextoError
        '------------------------------------------------------------------------------------
        flag = DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre(strNodo, strBUI, strSUBSISTEMA, strValorComboEquipo)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEquipo = flag
    End If
    
    '----------------------------------------------
    '   -Devuelve:
    '       Alta = strIDEvento
    '       Alta = "#ERR" & "|" & strTextoError
    '----------------------------------------------
    flag = Alta(strNodo, strBUI, strSUBSISTEMA, strIDEquipo, strTipoEvento, strCriticidad, strAliasTecnico, _
                        strORIGINADOR, strDESCRIPCION, strFechaAlta, strContacto, strPM, , strNotas, _
                        strFechaRecepcionNotificacion, strFechaInicioContactoCliente, strIncidenciaAveriaOReparacion, _
                        strUrgente, strEventoConServicioAfectado)
    
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEvento = flag
    Registrar = strIDEvento
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoAlta_Click()
    Dim strIDEvento As String, strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    '--------------------------------------------------------------
    '   -Devuelve:
    '       Registrar = strDescriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '--------------------------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEvento = flag
    '--------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        Form_FormEventoGestion.ComandoActualizarTodo_Click
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda") = strIDEvento
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    '--------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        If Forms("FormEventoGestion").Controls("ComandoAltaActividad").Enabled = True Then
            pregunta = MsgBox("Se ha dado de alta el evento con código: " & strIDEvento & vbNewLine & "¿Desea dar de alta una actividad?", _
                        vbExclamation + vbYesNo + vbDefaultButton1, "Alta de Evento")
            If pregunta = vbYes Then
                Form_FormEventoGestion.ComandoAltaActividad_Click
            End If
        Else
            pregunta = MsgBox("Se ha dado de alta el evento con código: " & strIDEvento, vbInformation, "Alta de Evento")
        End If
    Else
        pregunta = MsgBox("Se ha dado de alta el evento con código: " & strIDEvento, vbInformation, "Alta de Evento")
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


Private Sub ComandoEquiposDeMedida_Click()
    Dim strTextoError As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_ObjEventoActivo = Nothing
    If Not FormularioAbierto("FormEventoEqMedida") Then
        DoCmd.OpenForm "FormEventoEqMedida"
    End If
    If FormularioAbierto("FormEventoEqMedida") Then
        Set m_FormEquipoMedida = Forms("FormEventoEqMedida")
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al ComandoEquiposDeMedida_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoNuevoIDEquipo_Click()
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNodo = Nz(Me.ComboBui.Column(1), "")
    strBUI = Nz(Me.ComboBui.Column(0), "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    If strNodo = "" Or strBUI = "" Or strSUBSISTEMA = "" Then
        strTextoError = "Para dar de alta un Equipo se ha de conocer primero para que NODO,BUI,SUBSISTEMA va."
        Err.Raise 1000
    End If
    '--------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
    If FormularioAbierto("FormEquipoAlta") Then
        DoCmd.Close acForm, "FormEquipoAlta", acSaveNo
    End If
    strArgumento = strNodo & "|" & strBUI & "|" & strSUBSISTEMA
    DoCmd.OpenForm "FormEquipoAlta", , , , , , strArgumento
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario de IDEquipo se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboBUI_AfterUpdate()
    Dim strValorBUI As String, strValorActualSubSistema As String, _
        strSQLSubSistema As String, intElementos As Integer, strTextoError As String, strSQLContacto As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValorBUI = Nz(Me.ComboBui.Column(0), "")
    strValorActualSubSistema = Nz(Me.ComboSubsistema, "")
    If strValorBUI <> "" And strValorActualSubSistema <> "" Then
        Me.ComandoNuevoIDEquipo.Enabled = True
        Me.ComandoEquiposDeMedida.Enabled = True
    Else
        Me.ComandoNuevoIDEquipo.Enabled = False
        Me.ComandoEquiposDeMedida.Enabled = False
    End If
    If strValorBUI = strValorAnteriorBUI Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboSubsistema = Null
    Me.ComboIDEquipo = Null
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerComboSubSistema = m_SQL
    '       EstablecerComboSubSistema = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = EstablecerComboSubSistema(strValorBUI)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerComboSubSistema ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strSQLSubSistema = flag
    Me.ComboSubsistema.RowSource = strSQLSubSistema
    Me.ComboSubsistema.Requery
    intElementos = Me.ComboSubsistema.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 1 Then
        Me.ComboSubsistema = Me.ComboSubsistema.Selected(1)
        ComboSubsistema_Change
    End If
    strValorAnteriorBUI = strValorBUI
   
    RellenarComboContacto strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del campo BUI se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


Private Sub ComboBUI_Enter()
    On Error Resume Next
    Me.ComboBui.Requery
End Sub


Private Sub ComboIDEquipo_Enter()
    On Error Resume Next
    Me.ComboIDEquipo.Requery
End Sub

Private Sub ComboOriginador_Enter()
    On Error Resume Next
    Me.ComboOriginador.Requery
End Sub
Private Sub ComboSubsistema_Change()
    Dim strValorSubsistema As String, strValorActualIDEquipo As String, strSQLIDEquipo As String, strValorBUI As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValorBUI = Nz(Me.ComboBui.Column(0), "")
    strValorSubsistema = Nz(Me.ComboSubsistema, "")
    If strValorBUI <> "" And strValorSubsistema <> "" Then
        Me.ComandoNuevoIDEquipo.Enabled = True
        Me.ComandoEquiposDeMedida.Enabled = True
    Else
        Me.ComandoNuevoIDEquipo.Enabled = False
        Me.ComandoEquiposDeMedida.Enabled = False
    End If
    If strValorSubsistema = strValorAnteriorSubSistema Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComboIDEquipo = Null
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerComboEquipo = m_SQL
    '       EstablecerComboEquipo = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = EstablecerComboEquipo(strValorBUI, strValorSubsistema)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerComboEquipo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strSQLIDEquipo = flag
    Me.ComboIDEquipo.RowSource = strSQLIDEquipo
    Me.ComboIDEquipo.Requery
    strValorAnteriorSubSistema = strValorSubsistema
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del campo subsistema se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboSubsistema_Enter()
    On Error Resume Next
    Me.ComboSubsistema.Requery
End Sub
Private Sub ComboTipoEvento_Enter()
    On Error Resume Next
    Me.ComboTipoEvento.Requery
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Contacto_Enter()
    Me.Requery
End Sub

Private Sub Form_Load()
    Dim strArgumento As String, strIDEvento As String, strTextoError  As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    strArgumento = Nz(Me.OpenArgs(), "")
    If strArgumento <> "" Then
        strIDEvento = strArgumento
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstablecerDatos = Descriptivo
        '       EstablecerDatos = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = EstablecerDatos(strIDEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstablecerDatos ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        Me.lblTitulo.Caption = "Alta de Evento de " & strIDEvento
        Me.FechaAlta.SetFocus
    Else
        Me.lblTitulo.Caption = "Alta de Evento"
        Me.ComboCriticidad = 3
        Me.FechaAlta = Format(Now(), "dd/mm/yyyy")
        
        Me.ComboBui.SetFocus
    End If
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.No, EnumSino.Sí, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    
    Me.Caption = Me.lblTitulo.Caption
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Function EstablecerDatos(strIDEvento As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim strBUI As String, strSUBSISTEMA As String, strPM As String, strTipoEvento As String, strCriticidad As String, strAliasTecnico As String, _
        strORIGINADOR As String, strDESCRIPCION As String, strFechaAlta As String, strHoraAlta As String, strTiempoRespuesta As String, _
        strContacto As String, strIDEquipo As String, strCAUSAFIN As String, strEquipo As String, strNotas As String, strTextoError As String
    On Error GoTo errores
    If strIDEvento = "" Then
        strTextoError = "No se ha indicado un IDEvento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha indicado un IDEvento existente"
            Err.Raise 1000
        End If
        strBUI = Nz(.Fields("BUI"), "")
        strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
        strPM = Nz(.Fields("PMPR"), "")
        strTipoEvento = Nz(.Fields("TIPOEVENTO"), "")
        strCriticidad = Nz(.Fields("CRITICIDAD"), "")
        strAliasTecnico = Nz(.Fields("ALIASTECNICO"), "")
        strORIGINADOR = Nz(.Fields("ORIGINADOR"), "")
        strDESCRIPCION = Nz(.Fields("DESCRIPCION"), "")
        strFechaAlta = Nz(.Fields("FECHAALTAEVENTO"), "")
        strHoraAlta = Nz(.Fields("HORAINICIALEVENTO"), "")
        strTiempoRespuesta = Nz(.Fields("TIEMPORESPUESTAEVENTO"), "")
        strContacto = Nz(.Fields("CONTACTO"), "")
        strIDEquipo = Nz(.Fields("IDEquipo"), "")
        strNotas = Nz(.Fields("Notas"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strBUI <> "" Then
        Me.ComboBui = strBUI
    End If
    If strSUBSISTEMA <> "" Then
        Me.ComboSubsistema = strSUBSISTEMA
    End If
    If strPM <> "" Then
        Me.PMPR = strPM
    End If
    If strTipoEvento <> "" Then
        Me.ComboTipoEvento = strTipoEvento
    End If
    If IsNumeric(strCriticidad) Then
        Me.ComboCriticidad = strCriticidad
    End If
    If strAliasTecnico <> "" Then
        Me.ComboTecnico = strAliasTecnico
        
    End If
    If strORIGINADOR <> "" Then
        Me.ComboOriginador = strORIGINADOR
    End If
    If strDESCRIPCION <> "" Then
        Me.DescripcionEvento = strDESCRIPCION
    End If
    If IsDate(strFechaAlta) Then
        Me.FechaAlta = strFechaAlta
    End If
    
    ' El control visual Contacto fue retirado del alta
    ' El control visual Notas fue retirado del alta
    If strIDEquipo <> "" Then
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbEquipos", "Equipo", "IDEquipo", strIDEquipo)
        If InStr(1, flag, "|") <> 0 Then
            strIDEquipo = strIDEquipo
        Else
            strIDEquipo = flag
        End If
        Me.ComboIDEquipo = strIDEquipo
    End If
    'establecimiento de los combos
    If strBUI <> "" And strSUBSISTEMA <> "" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstablecerComboSubSistema = m_SQL
        '       EstablecerComboSubSistema = "ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = EstablecerComboSubSistema(strBUI)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstablecerComboSubSistema ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        m_SQL = flag
        Me.ComboSubsistema.RowSource = m_SQL
        Me.ComboSubsistema.Requery
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstablecerComboEquipo = m_SQL
        '       EstablecerComboEquipo = "ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = EstablecerComboEquipo(strBUI, strSUBSISTEMA)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstablecerComboEquipo ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        m_SQL = flag
        Me.ComboIDEquipo.RowSource = m_SQL
        Me.ComboIDEquipo.Requery
    End If
    RellenarComboContacto strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatos = "Datos establecidos correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstablecerDatos = "#ERR" & "|" & strTextoError
End Function
Public Function Alta( _
                        strNodo As String, _
                        strBUI As String, SubSistema As String, strIDEquipo As String, _
                         strTipoEvento As String, _
                        strCriticidad As String, strAliasTecnico As String, _
                        strORIGINADOR As String, strDescripcionEvento As String, _
                        strFechaAlta As String, strContacto As String, Optional strPMPR As String, _
                        Optional strIDEventoManual As String, Optional strNotas As String, _
                        Optional strFechaRecepcionNotificacion As String = "", Optional strFechaInicioContactoCliente As String = "", _
                        Optional strIncidenciaAveriaOReparacion As String = "", Optional strUrgente As String = "", _
                        Optional strEventoConServicioAfectado As String = "" _
                        ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/01/13
    '   -Modificaciones:
    '       -30/05/2013: se ha de poder poner 0 horas laborables y 0 horas no laborables para poder meter un material que sólo dependa de evento
    '       -24/03/14: Se considera la fecha de obsolescencia en el IDEquipo. Se coteja con la fecha de strFechaAlta
    '   -Funcionamiento:
    '       -va a tomar todos los datos que entran y los va a meter en la tabla TbEventos
    '       -1) Verificación de datos
    '           -strIDEventoManual no existente en caso de que esté relleno
    '           -strNodo no vacío
    '           -strBUI no vacío
    '           -SubSistema no vacío
    
    '           -strTipoEvento no vacío
    '           -strCriticidad no vacío
    '           -strAliasTecnico no vacío
    '           -strOriginador no vacío
    '           -strDescripcionEvento no vacío
    '           -strFechaAlta no vacío
    '           -strHoraAlta no vacío
    '           -strIDEquipo no vacío
    '           -si strFranqueado="Sí" ha de rellenarse strCausaFin
    '           -Si strHoraFin<>"" ha de ser una fecha válida
    '   -Llamada desde
    '       -Form_FormEventoAlta.ComandoAlta_Click
    '   -Devuelve:
    '       Alta = strIDEvento
    '       Alta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset
    Dim m_Evento As Evento
    Dim m_ID As Variant
    Dim strIDEventoEnBD As String, strIDEventoResultante As String, strFechaObsoleto As String, strTextoError As String, strValorSiNo As String
    On Error GoTo errores
    If strNodo = "" Then
        strTextoError = "Se ha de introducir el nodo"
        Err.Raise 1000
    End If
    If strBUI = "" Then
        strTextoError = "Se ha de introducir el BUI"
        Err.Raise 1000
    End If
    If SubSistema = "" Then
        strTextoError = "Se ha de introducir el Subsistema"
        Err.Raise 1000
    End If
    If strIDEquipo = "" Then
        strTextoError = "Se ha de introducir el IDEquipo"
        Err.Raise 1000
    End If
    If strTipoEvento = "" Then
        strTextoError = "Se ha de introducir el tipo de evento"
        Err.Raise 1000
    End If
    If Not IsNumeric(strCriticidad) Then
        strTextoError = "El código de criticidad ha de ser un valor numérico válido"
        Err.Raise 1000
    End If
    
    If strAliasTecnico = "" Then
        strTextoError = "Se ha de introducir el técnico que da de alta el evento"
        Err.Raise 1000
    End If
    If strORIGINADOR = "" Then
        strTextoError = "Se ha de introducir el originador del evento"
        Err.Raise 1000
    End If
    If strDescripcionEvento = "" Then
        strTextoError = "Se ha de introducir la descripción del evento"
        Err.Raise 1000
    End If
    
    If Not IsDate(strFechaAlta) Then
        strTextoError = "Se ha de introducir la fecha de alta del evento"
        Err.Raise 1000
    Else
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbEquipos", "FechaObsoleto", "IDEquipo", strIDEquipo)
        If InStr(1, flag, "|") <> 0 Then
            flag = ""
        End If
        strFechaObsoleto = flag
        If IsDate(strFechaObsoleto) Then
            If CDate(strFechaObsoleto) < CDate(strFechaAlta) Then
                strTextoError = "El equipo seleccionado para el evento es obsoleto desde el " & strFechaObsoleto
                Err.Raise 1000
            End If
        End If
    End If
    
    If strIDEventoManual <> "" Then
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbEventos", "IDEvento", "IDEvento", strIDEventoManual)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEventoEnBD = flag
        If strIDEventoEnBD <> "" Then
            strTextoError = "Ha intentado dar de alta un Evento con IDEvento manual. El IDEvento: " & strIDEventoEnBD & " ya existe en la base de datos"
            Err.Raise 1000
        End If
        strIDEventoResultante = strIDEventoManual
    Else
        '-----------------------------------------------
        '   -Devuelve:
        '       DameIDEvento = strIDEvento
        '       DameIDEvento = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = DameIDEvento(strBUI, strFechaAlta)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método DameIDEvento ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEventoResultante = flag
        If strIDEventoResultante = "" Then
            strTextoError = "La función Evento.DameIDEvento no ha devuelto un valor con formato válido"
            Err.Raise 1000
        End If
    End If
    '-----------------
    ' alta del evento
    '-----------------
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE (((TbEventos.IDEvento)='" & strIDEventoResultante & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función Evento.DameIDEvento ha devuelto un IDEvento ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDEvento") = strIDEventoResultante
            .Fields("NODO") = strNodo
            .Fields("BUI") = strBUI
            .Fields("SUBSISTEMA") = SubSistema
            If strPMPR <> "" Then
                .Fields("PMPR") = strPMPR
            End If
            .Fields("TIPOEVENTO") = strTipoEvento
            .Fields("CRITICIDAD") = strCriticidad
            .Fields("ALIASTECNICO") = strAliasTecnico
            .Fields("ORIGINADOR") = strORIGINADOR
            .Fields("FECHAALTAEVENTO") = strFechaAlta
            .Fields("HORAINICIALEVENTO") = "7:30"
            .Fields("TIEMPORESPUESTAEVENTO") = "00:00"
            .Fields("IDEquipo") = strIDEquipo
            .Fields("DESCRIPCION") = strDescripcionEvento
            If strContacto <> "" Then
                .Fields("CONTACTO") = strContacto
            End If
            If strNotas <> "" Then
                .Fields("Notas") = strNotas
            End If

            ' SLA inicial desde argumentos de alta
            If IsDate(strFechaRecepcionNotificacion) Then
                .Fields("FechaRecepcionNotificacion") = CDate(strFechaRecepcionNotificacion)
            Else
                .Fields("FechaRecepcionNotificacion") = Null
            End If

            If IsDate(strFechaInicioContactoCliente) Then
                .Fields("FechaInicioContactoCliente") = CDate(strFechaInicioContactoCliente)
            Else
                .Fields("FechaInicioContactoCliente") = Null
            End If

            strValorSiNo = UCase(Replace(strIncidenciaAveriaOReparacion, "Í", "I"))
            If strValorSiNo = "SI" Then
                .Fields("IncidenciaAveriaOReparacion") = True
            ElseIf strValorSiNo = "NO" Then
                .Fields("IncidenciaAveriaOReparacion") = False
            Else
                .Fields("IncidenciaAveriaOReparacion") = Null
            End If

            strValorSiNo = UCase(Replace(strUrgente, "Í", "I"))
            If strValorSiNo = "SI" Then
                .Fields("Urgente") = True
            ElseIf strValorSiNo = "NO" Then
                .Fields("Urgente") = False
            Else
                .Fields("Urgente") = Null
            End If

            strValorSiNo = UCase(Replace(strEventoConServicioAfectado, "Í", "I"))
            If strValorSiNo = "SI" Then
                .Fields("EventoConServicioAfectado") = True
            ElseIf strValorSiNo = "NO" Then
                .Fields("EventoConServicioAfectado") = False
            Else
                .Fields("EventoConServicioAfectado") = Null
            End If

            .Fields("FechaRegistroAlta") = Now()
            .Fields("FechaRegistroModificacion") = Now()
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    If Not m_ColEquiposMedidaPorEventoActivos Is Nothing Then
        Set m_Evento = Constructor.getEvento(strIDEventoResultante)
        If Not m_Evento Is Nothing Then
            For Each m_ID In m_ColEquiposMedidaPorEventoActivos
                
                m_Evento.RegistrarEquipoMedida CStr(m_ID)
               
                
            Next
        End If
        
    End If
    
    Alta = strIDEventoResultante
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

Private Function RellenarComboContacto(Optional ByRef p_Error As String) As String
    
    Dim m_BUI As String
    On Error GoTo errores

    m_BUI = Nz(Me.ComboBui, "")

    If m_BUI = "" Then m_BUI = "-1"
    If m_BUI = "-1" Then
        Me.Contacto.LimitToList = True
    Else
        Me.Contacto.LimitToList = False
    End If
    m_SQL = "SELECT DISTINCT TbEventos.CONTACTO " & _
            "FROM TbEventos " & _
            "WHERE ((Not (TbEventos.CONTACTO) Is Null) " & _
            "AND ((TbEventos.BUI)='" & m_BUI & "')) " & _
            "ORDER BY TbEventos.CONTACTO;"
    Me.Contacto.RowSource = m_SQL
    Me.Contacto.Requery
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Alta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
                
End Function

Private Sub m_FormEquipoMedida_Alta()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not m_ColEquiposMedidaPorEventoActivos Is Nothing Then
        Me.ComandoEquiposDeMedida.Caption = "Alta equipos de Medida (" & m_ColEquiposMedidaPorEventoActivos.count & ")"
    Else
        Me.ComandoEquiposDeMedida.Caption = "Alta equipos de Medida (0)"
    End If
    'Forms("FormEventoEqMedida").SetFocus
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al m_FormEquipoMedida_Alta se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub m_FormEquipoMedida_Eliminado()
    
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not m_ColEquiposMedidaPorEventoActivos Is Nothing Then
        Me.ComandoEquiposDeMedida.Caption = "Alta equipos de Medida (" & m_ColEquiposMedidaPorEventoActivos.count & ")"
    Else
        Me.ComandoEquiposDeMedida.Caption = "Alta equipos de Medida (0)"
    End If
    'Forms("FormEventoEqMedida").SetFocus
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al m_FormEquipoMedida_Eliminado se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
