Option Compare Database
Option Explicit
Dim strValorAnteriorNodo As String, strValorAnteriorBUI As String, strValorAnteriorSubSistema As String
Dim strIDEvento As String, strBUIAlInicio As String, strSubSistemaAlInicio As String, strEquipoAlInicio As String, strPMPRAlInicio As String, _
    strTipoEventoAlInicio As String, strCriticidadAlInicio As String, strAliasTecnicoAlInicio As String, strOriginadorAlInicio As String, strDescripcionAlInicio As String, _
    strFECHAALTAAlInicio As String, strContactoAlInicio As String, strNotasAlInicio As String, strFechaEnInformeRACAlInicio As String
Dim strEscritura As String, strURLAnexo As String, strIDMantenimiento As String, strFranqueado As String, _
    strEstaFacturado As String, strFormularioQueAbre As String

Private WithEvents m_FormEquipoMedida As Form_FormEventoEqMedidaEdicion
Public Event FechaInformeRACambiada()
Private Function ActivarDesactivarImagenAnexo(p_Evento As Evento, Optional ByRef p_Error As String) As String
    
    Dim m_URL As String
    On Error GoTo errores
    
    m_URL = p_Evento.URLAnexos
    p_Error = p_Evento.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        Me.ImagenAnexo.Visible = False
    Else
        Me.ImagenAnexo.Visible = True
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActivarDesactivarImagenAnexo ha devuelto el error " & vbNewLine & Err.Description
    End If
    Me.ImagenAnexo.Visible = False
    
End Function
Private Function HaHabidoCambios(Optional p_ExceptoCasafin As EnumSino = EnumSino.No) As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  18/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True
    '-------------------------------------------------------------------
    Dim strBUI As String, strSUBSISTEMA As String, strEquipo As String, strPMPR As String, _
        strTipoEvento As String, strCriticidad As String, strAliasTecnico As String, strORIGINADOR As String, strDESCRIPCION As String, _
        strFechaAlta As String, strContacto As String, strCAUSAFIN As String, strNotas As String, strFechaEnInformeRAC As String
    On Error GoTo errores
    If p_ExceptoCasafin = 0 Then
        p_ExceptoCasafin = EnumSino.No
    End If
    strBUI = Nz(Me.ComboBui.Column(0), "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    strEquipo = Nz(Me.ComboIDEquipo.Column(1), "")
    strPMPR = Nz(Me.PMPR, "")
    strTipoEvento = Nz(Me.ComboTipoEvento, "")
    strCriticidad = Nz(Me.ComboCriticidad, "")
    strAliasTecnico = Nz(Me.ComboTecnico, "")
    strORIGINADOR = Nz(Me.ComboOriginador, "")
    strDESCRIPCION = Nz(Me.DescripcionEvento, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strContacto = Nz(Me.Contacto, "")
    
    strNotas = Nz(Me.Notas, "")
    strFechaEnInformeRAC = Nz(Me.FechaEnInformeRAC, "")
    If strBUI <> strBUIAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strSUBSISTEMA <> strSubSistemaAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strEquipo <> strEquipoAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strPMPR <> strPMPRAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strTipoEvento <> strTipoEventoAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strCriticidad <> strCriticidadAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strAliasTecnico <> strAliasTecnicoAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strORIGINADOR <> strOriginadorAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strDESCRIPCION <> strDescripcionAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strFechaAlta <> strFECHAALTAAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strContacto <> strContactoAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    
    If strNotas <> strNotasAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strFechaEnInformeRAC <> strFechaEnInformeRACAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Private Sub cmdSalir_Click()
    Dim strNombreFormulario As String, lst As ListBox, strTextoError As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    If HaHabidoCambios() And strEscritura = "Sí" Then
        pregunta = MsgBox("¿Desea guardar los cambios producidos?", vbExclamation + vbYesNo + vbDefaultButton1, "Guardar Cambios")
        If pregunta = vbYes Then
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            '-------------------------------------------------------------------
            '   -Devuelve:
            '       Registrar = strDescriptivo
            '       Registrar = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = Registrar()
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Registrar ha devuelto un error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            '-------------------------------------------------------------------
            '   -Devuelve:
            '       FormularioAbierto = True Or False
            '
            '           strEstado="1"----> True Or False
            '           strEstado="-1"--->False
            '-------------------------------------------------------------------
            If FormularioAbierto("FormEventoGestion") Then
                Set lst = Forms("FormEventoGestion").Controls("ListaEventos")
                lst.Requery
                lst.Selected(1) = True
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
        End If
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAdjuntar_Click()

    Dim m_Evento As Evento
     Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strIDEvento = "" Then
        strTextoError = "No se ha podido determinar el IDEvento"
        Err.Raise 1000
    End If
    Set m_Evento = Constructor.getEvento(strIDEvento, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    strFormularioQueAbre = Me.Name
    Set m_ObjEntidadParaAnexoActiva = m_Evento
    'strPrefijo="EV,AC,MA,SUB"
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoEditar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
   
    
    
    '-----------------------------------------------
    '   -Devuelve:
    '       Registrar = strDescriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
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
    
    If FormularioAbierto("FormEventoEqMedidaEdicion") Then
       DoCmd.Close acForm, "FormEventoEqMedidaEdicion", acSaveNo
    End If
     DoCmd.OpenForm "FormEventoEqMedidaEdicion"
    If FormularioAbierto("FormEventoEqMedidaEdicion") Then
        Set m_FormEquipoMedida = Forms("FormEventoEqMedidaEdicion")
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

Private Sub ComandoSLU_Click()
    Dim strTextoError As String
    On Error GoTo errores

    If Nz(strIDEvento, "") = "" Then
        strTextoError = "No se puede abrir SLA: IDEvento no disponible"
        Err.Raise 1000
    End If

    If FormularioAbierto("FormEventoSLA") Then
        DoCmd.Close acForm, "FormEventoSLA", acSaveNo
    End If

    DoCmd.OpenForm "FormEventoSLA", , , , , , strIDEvento
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al abrir el formulario SLA se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoNuevoIDEquipo_Click()
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------
    If FormularioAbierto("FormEquipoAlta") Then
        DoCmd.Close acForm, "FormEquipoAlta", acSaveNo
    End If
    strNodo = Nz(Me.ComboBui.Column(1), "")
    strBUI = Nz(Me.ComboBui.Column(0), "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    If strNodo <> "" And strBUI <> "" And strSUBSISTEMA <> "" Then
        strArgumento = strNodo & "|" & strBUI & "|" & strSUBSISTEMA
    End If
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

Private Sub ComandoQuitarFechaInforme_Click()
    On Error Resume Next
    Me.FechaEnInformeRAC = Null
End Sub

Private Sub ComboBUI_AfterUpdate()
    Dim strValorBUI As String, strValorActualSubSistema As String, strSQLSubSistema As String, _
        intElementos As Integer, strTextoError As String, strSQLContacto As String
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
    flag = EstablecerComboContacto(strValorBUI)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerComboContacto ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strSQLContacto = flag
    Me.Contacto.RowSource = strSQLContacto
    Me.Contacto.Requery
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
Private Function Registrar() As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = strDescriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '---------------------------------------------------------------------------------
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, strPMPR As String, _
        strTipoEvento As String, strCriticidad As String, strAliasTecnico As String, strORIGINADOR As String, strDESCRIPCION As String, _
        strFechaAlta As String, strContacto As String, strNotas As String, strFechaEnInformeRAC As String, strTextoError As String
    On Error GoTo errores
    If strIDEvento = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    If strEscritura = "No" Then
        strFechaEnInformeRAC = Nz(Me.FechaEnInformeRAC, "")
        If strFechaEnInformeRAC = strFechaEnInformeRACAlInicio Then
            strTextoError = "No ha habido cambios"
            Err.Raise 1000
        End If
        If Nz(Me.FechaEnInformeRAC, "") = "" Then
            m_ObjEventoActivo.InformeRACQuitarFecha strTextoError
            If strTextoError <> "" Then
                Err.Raise 1000
            End If
        Else
             m_ObjEventoActivo.InformeRACRegistrarFecha Me.FechaEnInformeRAC, strTextoError
            If strTextoError <> "" Then
                Err.Raise 1000
            End If
        End If
        RaiseEvent FechaInformeRACambiada
        Registrar = "Edición correcta"
        Exit Function
    End If
    strNodo = Nz(Me.ComboBui.Column(1), "")
    strBUI = Nz(Me.ComboBui.Column(0), "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    strIDEquipo = Nz(Me.ComboIDEquipo)
    strPMPR = Nz(Me.PMPR, "#")
    strTipoEvento = Nz(Me.ComboTipoEvento, "")
    strCriticidad = Nz(Me.ComboCriticidad, "")
    strAliasTecnico = Nz(Me.ComboTecnico, "")
    strORIGINADOR = Nz(Me.ComboOriginador, "")
    strDESCRIPCION = Nz(Me.DescripcionEvento, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strContacto = Nz(Me.Contacto, "#")
    strNotas = Nz(Me.Notas, "#")
    strFechaEnInformeRAC = Nz(Me.FechaEnInformeRAC, "#")
    If IsDate(strFechaEnInformeRAC) Then
        m_ObjEventoActivo.FechaEnInformeRAC = strFechaEnInformeRAC
    Else
        m_ObjEventoActivo.FechaEnInformeRAC = ""
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Edicion = strDescriptivo
    '       Edicion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Edicion(strIDEvento, strNodo, strBUI, strSUBSISTEMA, strIDEquipo, strPMPR, strTipoEvento, strCriticidad, strAliasTecnico, strORIGINADOR, strDESCRIPCION, _
                            strFechaAlta, strContacto, strNotas, strFechaEnInformeRAC)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Edicion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    
    Registrar = "Edición correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function

Private Sub ComboTipoEvento_Enter()
    On Error Resume Next
    Me.ComboTipoEvento.Requery
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    'strArgumento=strIDEvento & "|" & strEstaFacturado & "|" & strFranqueado
    
    strArgumento = Nz(Me.OpenArgs, "")
    If strArgumento = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 2 Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    strIDEvento = Nz(dato(0), "")
    
    strEstaFacturado = Nz(dato(1), "")
    strFranqueado = Nz(dato(2), "")
    If strIDEvento = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    If strFranqueado <> "Sí" And strFranqueado <> "No" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    
    If strEstaFacturado = "Sí" Or strFranqueado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    Me.PMPR.SetFocus
    
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerDatos ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, , strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
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
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function EstablecerDatos() As String
    
    Dim m_ColEquiposMedida As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset, strFechaUltimaActividad As String, ctl As Control, strHorasEvento As String, strFranqueada As String
    Dim strEventoSinNada As String, strTextoError As String
    
    Dim m_HorasLaborables As String
    Dim m_HorasExtras As String
    
    On Error GoTo errores
    If strIDEvento = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    
    
    Set m_ObjEventoActivo = Constructor.getEvento(strIDEvento, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    Set m_ColEquiposMedida = m_ObjEventoActivo.ColEquiposMedida
    strTextoError = m_ObjEventoActivo.Error
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    If m_ColEquiposMedida Is Nothing Then
        Me.ComandoEquiposDeMedida.Caption = "Equipos de Medida (0)"
        If strEscritura = "Sí" Then
            Me.ComandoEquiposDeMedida.Enabled = True
        Else
            Me.ComandoEquiposDeMedida.Enabled = False
        End If
        
    Else
         Me.ComandoEquiposDeMedida.Caption = "Ver equipos de Medida (" & m_ColEquiposMedida.count & ")"
    End If
    
     m_SQL = "SELECT TbEventos.*, TbEquipos.Equipo " & _
                "FROM TbEquipos RIGHT JOIN TbEventos ON TbEquipos.IDEquipo = TbEventos.IDEquipo " & _
                "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
    
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strBUIAlInicio = Nz(.Fields("BUI"), "")
        strSubSistemaAlInicio = Nz(.Fields("SUBSISTEMA"), "")
        strEquipoAlInicio = Nz(.Fields("Equipo"), "")
        strPMPRAlInicio = Nz(.Fields("PMPR"), "")
        strTipoEventoAlInicio = Nz(.Fields("TIPOEVENTO"), "")
        strCriticidadAlInicio = Nz(.Fields("CRITICIDAD"), "")
        strAliasTecnicoAlInicio = Nz(.Fields("ALIASTECNICO"), "")
        strOriginadorAlInicio = Nz(.Fields("ORIGINADOR"), "")
        strDescripcionAlInicio = Nz(.Fields("DESCRIPCION"), "")
        strFECHAALTAAlInicio = Nz(.Fields("FECHAALTAEVENTO"), "")
        strContactoAlInicio = Nz(.Fields("CONTACTO"), "")
        strFranqueada = Nz(.Fields("Franqueado"), "")
        strNotasAlInicio = Nz(.Fields("Notas"), "")
        strFechaEnInformeRACAlInicio = Nz(.Fields("FechaEnInformeRAC"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strBUIAlInicio <> "" Then
        Me.ComboBui = strBUIAlInicio
    End If
    If strSubSistemaAlInicio <> "" Then
       Me.ComboSubsistema.RowSource = EstablecerComboSubSistema(Me.ComboBui)
        Me.ComboSubsistema = strSubSistemaAlInicio
    End If
    If strEquipoAlInicio <> "" Then
        Me.ComboIDEquipo.RowSource = EstablecerComboEquipo(Me.ComboBui, Me.ComboSubsistema)
        Me.ComboIDEquipo = m_ObjEventoActivo.IDEquipo
    End If
    If strPMPRAlInicio <> "" Then
        Me.PMPR = strPMPRAlInicio
    End If
    If strTipoEventoAlInicio <> "" Then
        Me.ComboTipoEvento = strTipoEventoAlInicio
    End If
    If strCriticidadAlInicio <> "" Then
        Me.ComboCriticidad = strCriticidadAlInicio
    End If
    If strAliasTecnicoAlInicio <> "" Then
        Me.ComboTecnico = strAliasTecnicoAlInicio
    End If
    If strOriginadorAlInicio <> "" Then
        Me.ComboOriginador = strOriginadorAlInicio
    End If
    If strDescripcionAlInicio <> "" Then
        Me.DescripcionEvento = strDescripcionAlInicio
    End If
    If IsDate(strFECHAALTAAlInicio) Then
        Me.FechaAlta = strFECHAALTAAlInicio
    End If
    
    If strContactoAlInicio <> "" Then
        Me.Contacto = strContactoAlInicio
    End If
    If strFranqueada = "Verdadero" Then
        strFranqueada = "Sí"
    Else
        strFranqueada = "No"
    End If
        
    If strNotasAlInicio <> "" Then
        Me.Notas = strNotasAlInicio
    End If
    If IsDate(strFechaEnInformeRACAlInicio) Then
        Me.FechaEnInformeRAC = strFechaEnInformeRACAlInicio
    End If
    
    Me.Franqueada = strFranqueada
    
    '----------------------------------------------------------------
    '   -Devuelve:
    '       DameFechaFinEvento = strFechaUltimaActividad
    '       DameFechaFinEvento = #ERR
    '----------------------------------------------------------------
    strFechaUltimaActividad = DameFechaFinEvento(strIDEvento)
    If IsDate(strFechaUltimaActividad) Then
        Me.FechaUltimaActividad = strFechaUltimaActividad
    Else
        Me.FechaUltimaActividad = "Sin Acts"
    End If
    Me.FechaUltimaActividad.Enabled = False
    Me.FechaUltimaActividad.Locked = True
    For Each ctl In Me.Controls
        If ctl.Tag = "NOEDITABLE" Then
            ctl.BackColor = lngColorCampoNoEditable
            ctl.Enabled = True
            ctl.Locked = True
        End If
    Next
    If strEscritura = "Sí" Then
        Me.lblTitulo1.Caption = "Evento " & strIDEvento
        For Each ctl In Me.Controls
            If (ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox) And ctl.Tag <> "NOEDITABLE" Then
                ctl.BackColor = lngColorCampoEditable
                ctl.Enabled = True
                ctl.Locked = False
            End If
        Next
        Me.ComandoEditar.Enabled = True
        Me.ComandoNuevoIDEquipo.Enabled = True
        Me.ComandoEquiposDeMedida.Enabled = True
        '-----------------------------------------------
        '   -Devuelve:
        '       Dame = strValor
        '       Dame = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = Dame("TbActividades", "IDEvento", "IDEvento", strIDEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag = "" Then
            strEventoSinNada = "Sí"
        Else
            strEventoSinNada = "No"
        End If
    Else
        Me.lblTitulo1.Caption = "Evento " & strIDEvento & " (Sólo lectura)"
        Me.cmdSalir.SetFocus
        For Each ctl In Me.Controls
            If ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox Then
                ctl.BackColor = lngColorCampoNoEditable
                ctl.Enabled = True
                ctl.Locked = True
            End If
        Next
        Me.ComandoEditar.Enabled = False
        Me.ComandoNuevoIDEquipo.Enabled = False
        Me.ComandoEquiposDeMedida.Enabled = False
        
    End If
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    ActivarDesactivarImagenAnexo m_ObjEventoActivo, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    If strEstaFacturado = "Sí" Then
        Me.lblFacturado.Visible = True
    Else
        Me.lblFacturado.Visible = False
    End If
    If strFranqueado = "Sí" Then
        Me.lblFranqueado.Visible = True
    Else
        Me.lblFranqueado.Visible = False
    End If
    If IsDate(strFechaEnInformeRACAlInicio) Then
        Me.lblInformeRAC.Visible = True
    Else
        Me.lblInformeRAC.Visible = False
    End If
    
    
    With m_ObjEventoActivo
        m_HorasLaborables = .HorasLabTotales
        m_HorasExtras = .HorasExtTotales
    End With
    If Not IsNumeric(m_HorasLaborables) Then
        m_HorasLaborables = "0"
    End If
    If Not IsNumeric(m_HorasExtras) Then
        m_HorasExtras = "0"
    End If
    Me.HorasEvento = CDbl(m_HorasLaborables) + CDbl(m_HorasExtras)
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacion", "IDPlanificacion", "IDEvento", strIDEvento)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDMantenimiento = flag
    If IsNumeric(strIDMantenimiento) Then
        Me.ComandoVerProgramacion.Visible = True
        Me.ComandoVerProgramacion.Caption = "Cierra la Programacion nº " & strIDMantenimiento
    Else
        Me.ComandoVerProgramacion.Visible = False
    End If
     'establecimiento de los combos
    If Nz(Me.ComboBui.Column(0), "") <> "" And Nz(Me.ComboSubsistema, "") <> "" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstablecerComboSubSistema = m_SQL
        '       EstablecerComboSubSistema = "ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = EstablecerComboSubSistema(Nz(Me.ComboBui.Column(0), ""))
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
        flag = EstablecerComboEquipo(Nz(Me.ComboBui.Column(0), ""), Nz(Me.ComboSubsistema, ""))
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstablecerComboEquipo ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        m_SQL = flag
        Me.ComboIDEquipo.RowSource = m_SQL
        Me.ComboIDEquipo.Requery
    End If
'    If strEstaFacturado = "Sí" Or strFranqueado = "Sí" Then
'        Me.AllowEdits = False
'    Else
'        Me.AllowEdits = True
'    End If
    Me.FechaEnInformeRAC.Enabled = True
    Me.FechaEnInformeRAC.Locked = False
    Me.FechaEnInformeRAC.BackColor = lngColorCampoEditable
    Me.ComandoEditar.Enabled = True
    
    If IsDate(strFechaEnInformeRACAlInicio) Then
        Me.ComandoQuitarFechaInforme.Enabled = True
    Else
        Me.ComandoQuitarFechaInforme.Enabled = False
    End If
    
    Me.CausaFin.Enabled = True
    Me.CausaFin.Locked = True
    Me.ComandoEquiposDeMedida.Enabled = True
    EstablecerDatos = "Datos del evento establecidos correctamente"
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
Private Sub ComandoVerProgramacion_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If IsNumeric(strIDMantenimiento) Then
        '---------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '---------------------------------------------
        If FormularioAbierto("FormPlanificacionCierre") Then
            DoCmd.Close acForm, "FormPlanificacionCierre", acSaveNo
        End If
        DoCmd.OpenForm "FormPlanificacionCierre", , , , , , strIDMantenimiento
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ImagenAnexo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strURLAnexo = "" Then
        strTextoError = "En estos momentos el anexo seleccionado no es alcanzable por la herramienta. Revise que ubicación de la documentación remota esté bien establecida"
        Err.Raise 1000
    End If
    If InStr(1, strURLAnexo, ";") = 0 Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AbrirEnLocal = strURLAnexoLocal
        '       AbrirEnLocal = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = AbrirEnLocal(strURLAnexo, Me.hWnd)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        ComandoAdjuntar_Click
    End If
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
End Sub


Public Function Edicion( _
                            strIDEvento As String, _
                            Optional strNodo As String, _
                            Optional strBUI As String, Optional SubSistema As String, Optional strIDEquipo As String, _
                            Optional strPMPR As String, Optional strTipoEvento As String, _
                            Optional strCriticidad As String, Optional strAliasTecnico As String, _
                            Optional strORIGINADOR As String, Optional strDescripcionEvento As String, _
                            Optional strFechaAlta As String, _
                            Optional strContacto As String, Optional strNotas As String, Optional strFechaEnInformeRAC As String _
                            ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a tomar todos los datos que entran y los va a editar en la tabla TbEventos
    '       -1) Verificación de datos
    '           -strIDEvento  existente
    '           -strNodo no vacío y no #
    '           -strBUI no vacío y no #
    '           -SubSistema no vacío y no #
    
    '           -strTipoEvento no vacío y no #
    '           -strCriticidad no vacío y no # y numérico
    '           -strAliasTecnico no vacío y no #
    '           -strOriginador no vacío y no #
    '           -strDescripcionEvento no vacío y no #
    '           -strIDEquipo no vacío y no #
    '           -strFechaAlta fecha
    '           -strHoraAlta fecha
    '           -strCausaFin<>"" y strCausaFin<>"#" cuando strFranqueado="Sí"
    '           -si strHoraFin<>"" y strHoraFin<>"#" ha de ser una fecha
    '   -Llamada desde
    '       -Form_FormEventoEdicion.Editar
    '   -Devuelve:
    '       Edicion = Descriptivo
    '       Edicion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strBUIAnterior As String, strIDEventoNuevo As String, strFechaObsoleto As String, strTextoError As String
    On Error GoTo errores
    If strIDEvento = "" Then
        strTextoError = "No se puede dejar en blanco el IDEvento"
        Err.Raise 1000
    End If
    If strNodo = "" Or strNodo = "#" Then
        strTextoError = "Se ha de introducir el nodo"
        Err.Raise 1000
    End If
    If strBUI = "" Or strBUI = "#" Then
        strTextoError = "Se ha de introducir el BUI"
        Err.Raise 1000
    End If
    If SubSistema = "" Or SubSistema = "#" Then
        strTextoError = "Se ha de introducir el Subsistema"
        Err.Raise 1000
    End If
    If strIDEquipo = "" Or strIDEquipo = "#" Then
        strTextoError = "Se ha de introducir el IDEquipo"
        Err.Raise 1000
    End If
    If strTipoEvento = "" Or strTipoEvento = "#" Then
        strTextoError = "Se ha de introducir el tipo de evento"
        Err.Raise 1000
    End If
    If Not IsNumeric(strCriticidad) Then
        strTextoError = "El código de criticidad ha de ser un valor numérico válido"
        Err.Raise 1000
    End If
    If strAliasTecnico = "" Or strAliasTecnico = "#" Then
        strTextoError = "Se ha de introducir el técnico que da de Edicion el evento"
        Err.Raise 1000
    End If
    If strORIGINADOR = "" Or strORIGINADOR = "#" Then
        strTextoError = "Se ha de introducir el originador del evento"
        Err.Raise 1000
    End If
    If strDescripcionEvento = "" Or strDescripcionEvento = "#" Then
        strTextoError = "Se ha de introducir la descripción del evento"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaAlta) Then
        strTextoError = "Se ha de introducir una fecha de alta del evento correcta"
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
    
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbEventos", "BUI", "IDEvento", strIDEvento)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strBUIAnterior = flag
    If strBUIAnterior <> strBUI Then
        '------------------------------------------
        '   -Devuelve:
        '       DameIDEvento = strIDEvento
        '       DameIDEvento = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = DameIDEvento(strBUI, strFechaAlta)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEventoNuevo = flag
        If strIDEventoNuevo = "" Then
            strTextoError = "El método Evento.DameIDEvento ha devuelto un valor con formato no válido"
            Err.Raise 1000
        End If
    End If
     '-----------------
    ' Edicion del evento
    '-----------------
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No existe un evento con ese código"
            Err.Raise 1000
        End If
        .Edit
            If strIDEventoNuevo <> "" Then
                .Fields("IDEvento") = strIDEventoNuevo
            End If
            .Fields("NODO") = strNodo
            .Fields("BUI") = strBUI
            .Fields("SUBSISTEMA") = SubSistema
            .Fields("IDEquipo") = strIDEquipo
            If strPMPR <> "" And strPMPR <> "#" Then
                .Fields("PMPR") = strPMPR
            Else
                If strPMPR = "#" Then
                    .Fields("PMPR") = Null
                End If
            End If
            .Fields("TIPOEVENTO") = strTipoEvento
            .Fields("CRITICIDAD") = strCriticidad
            .Fields("ALIASTECNICO") = strAliasTecnico
            .Fields("ORIGINADOR") = strORIGINADOR
            .Fields("FECHAALTAEVENTO") = strFechaAlta
            .Fields("HORAINICIALEVENTO") = "7:30:00"
            .Fields("TIEMPORESPUESTAEVENTO") = "00:00"
            .Fields("DESCRIPCION") = strDescripcionEvento
            If strContacto <> "" And strContacto <> "#" Then
                .Fields("CONTACTO") = strContacto
            Else
                If strContacto = "#" Then
                    .Fields("CONTACTO") = Null
                End If
            End If
            
            If strNotas <> "" And strNotas <> "#" Then
                .Fields("Notas") = strNotas
            Else
                If strNotas = "#" Then
                    .Fields("Notas") = Null
                End If
            End If
            If IsDate(strFechaEnInformeRAC) Then
                .Fields("FechaEnInformeRAC") = strFechaEnInformeRAC
            Else
                .Fields("FechaEnInformeRAC") = Null
            End If
            .Fields("FechaRegistroModificacion") = Now()
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strIDEventoNuevo <> "" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       ReNombrarIDEvento = strDescriptivo
        '       ReNombrarIDEvento = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = ReNombrarIDEvento(strIDEvento, strIDEventoNuevo)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método ReNombrarIDEvento ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    
    Edicion = "Evento editado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Evento.Edicion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Edicion = "#ERR" & "|" & strTextoError
End Function
Public Function ReNombrarIDEvento( _
                                    strIDEventoActual As String, _
                                    strIDEventoNuevo As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 21/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a editar las tablas:
    '           -TbActividades,TbEventoAnexos
    '       -1) Verificación de datos
    '           -strIDEventoActual  no existente, se supone que ya se ha cambiado antes
    '           -strIDEventoNuevo  existente
   
    '   -Llamada desde
    '       -Edicion
    '   -Devuelve:
    '       ReNombrarIDEvento = Descriptivo
    '       ReNombrarIDEvento = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim strSQLRenombrar As String, strTextoError As String
    On Error GoTo errores
    
    If strIDEventoActual = "" Then
        strTextoError = "Se ha de introducir el IDEvento actual"
        Err.Raise 1000
    End If
    If strIDEventoNuevo = "" Then
        strTextoError = "Se ha de introducir el IDEvento Nuevo"
        Err.Raise 1000
    End If
    strSQLRenombrar = "UPDATE TbActividades SET TbActividades.IDEvento = '" & strIDEventoNuevo & "' " & _
                    "WHERE (((TbActividades.IDEvento)='" & strIDEventoActual & "'));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQLRenombrar
    DoCmd.SetWarnings True
   
    strSQLRenombrar = "UPDATE TbEventoAnexos SET TbEventoAnexos.IDEvento = '" & strIDEventoNuevo & "' " & _
                    "WHERE (((TbEventoAnexos.IDEvento)='" & strIDEventoActual & "'));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQLRenombrar
    DoCmd.SetWarnings True
    ReNombrarIDEvento = "IDEvento Renombrado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método ReNombrarIDEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.SetWarnings True
    ReNombrarIDEvento = "#ERR" & "|" & strTextoError
End Function


Private Sub m_FormEquipoMedida_Alta()
    Dim strTextoError As String
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Set m_Col = m_ObjEventoActivo.ColEquiposMedida
    If Not m_Col Is Nothing Then
        Me.ComandoEquiposDeMedida.Caption = "Ver equipos de Medida (" & m_Col.count & ")"
    Else
        Me.ComandoEquiposDeMedida.Caption = "Equipos de Medida (0)"
    End If
    
    
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
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_Col = m_ObjEventoActivo.ColEquiposMedida
    If Not m_Col Is Nothing Then
        Me.ComandoEquiposDeMedida.Caption = "Ver equipos de Medida (" & m_Col.count & ")"
    Else
        Me.ComandoEquiposDeMedida.Caption = "Equipos de Medida (0)"
    End If
    
    
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
