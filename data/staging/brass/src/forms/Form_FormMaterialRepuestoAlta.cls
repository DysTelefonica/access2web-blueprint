Option Compare Database
Option Explicit
Dim strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, strIDActividad As String, strIDEvento As String
Dim strIDMaterialUltimo As String, strUltimoMaterialNoEnCombo As String, strEquipo As String
Private Sub ComandoRegistrar_Click()
    Dim strMaterial As String, strGarantia As String, strFecha As String, strSuministrador As String, strCOSTE As String, _
        strTipoAccion As String, strPN As String, strNS As String, strObservaciones As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strIDActividad = "" Then
        strTextoError = "No se sabe para qué actividad es el repuesto"
        Err.Raise 1000
    End If
    strMaterial = Nz(Me.ComboMaterial, "")
    If strMaterial = "" Then
        strMaterial = Nz(Me.Material, "")
    End If
    strGarantia = Nz(Me.ComboGarantia, "")
    strFecha = Nz(Me.FechaIncio, "")
    strSuministrador = Nz(Me.ComboReparadoPor, "")
    strCOSTE = Nz(Me.COSTE, "")
    strTipoAccion = Nz(Me.ComboTipoAccion, "")
    strPN = Nz(Me.PN, "")
    strNS = Nz(Me.NS, "")
    strObservaciones = Nz(Me.Observaciones, "")
    '-----------------------------------------------
    '   -Devuelve:
    '       MaterialDeActividadAlta = strIDMaterial
    '       MaterialDeActividadAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = MaterialDeActividadAlta(strIDActividad, strMaterial, strCOSTE, _
                                             strTipoAccion, strPN, strNS, strSuministrador, _
                                            strObservaciones, strFecha, strGarantia, "No", strFecha)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método MaterialDeActividadAlta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
       Form_FormEventoGestion.ListaActividades_Click
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

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComandoVerActividad_Click()
    Dim strTextoError As String
    On Error GoTo errores
    If strIDActividad = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If Not FormularioAbierto("FormActividadGestion") Then
        DoCmd.OpenForm "FormActividadGestion"
        Forms("FormActividadGestion").Controls("ComboIDActividadBusqueda").Value = strIDActividad
        Form_FormActividadGestion.ComboIDActividadBusqueda_Change
    Else
        Forms("FormActividadGestion").SetFocus
        If Forms("FormActividadGestion").Controls("ComboIDActividadBusqueda").Value <> strIDActividad Then
            'Form_FormActividadGestion.ComandoActualizarTodo_Click
            Forms("FormActividadGestion").Controls("ComboIDActividadBusqueda").Value = strIDActividad
            Form_FormActividadGestion.ComboIDActividadBusqueda_Change
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver la actividad se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoVerEvento_Click()
    Dim strTextoError As String, strIDEvento As String
    On Error GoTo errores
    strIDEvento = Nz(Me.IDEVENTO, "")
    If strIDEvento = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If Not FormularioAbierto("FormEventoGestion") Then
        DoCmd.OpenForm "FormEventoGestion"
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    Else
        Forms("FormEventoGestion").SetFocus
        If Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value <> strIDEvento Then
            Form_FormEventoGestion.ComandoActualizarTodo_Click
            Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
            Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el Evento origen  se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboMaterial_Change()
    Dim strValor As String, strNS As String, strPN As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboMaterial.Column(0), "")
    If strValor <> "" Then
        Me.Material = Null
        strNS = Nz(Me.ComboMaterial.Column(2), "")
        strPN = Nz(Me.ComboMaterial.Column(1), "")
        Me.NS = strNS
        Me.PN = strPN
    Else
        If Nz(Me.Material, "") = "" Then
            Me.PN = Null
            Me.NS = Null
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al elegir el material se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    strArgumento = Nz(Me.OpenArgs(), "")
    If InStr(1, strArgumento, "|") = 0 Then
        strTextoError = "El formulario ha sido abierto con parámetros erróneos"
        Err.Raise 1000
    End If
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strEquipo & "|" & strIDEquipo & "|" & strIDEvento & "|" & strIDActividad
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 5 Then
        strTextoError = "El formulario ha sido abierto con parámetros erróneos"
        Err.Raise 1000
    End If
    strBUI = dato(0)
    strSUBSISTEMA = dato(1)
    strEquipo = dato(2)
    strIDEquipo = dato(3)
    strIDEvento = dato(4)
    strIDActividad = dato(5)
    If strBUI = "" Or strSUBSISTEMA = "" Or strIDEquipo = "" Or strIDEquipo = "" Then
        strTextoError = "El formulario ha sido abierto con parámetros erróneos"
        Err.Raise 1000
    End If
    If strIDEvento = "" Or strIDActividad = "" Then
        strTextoError = "El formulario ha sido abierto con parámetros erróneos"
        Err.Raise 1000
    End If
    Me.IDEVENTO = strIDEvento
    Me.IDActividad = strIDActividad
    Me.BUI = strBUI
    Me.SubSistema = strSUBSISTEMA
    Me.IDEquipo = strIDEquipo
    Me.ComandoVerEvento.Enabled = True
    Me.ComandoVerActividad.Enabled = True
    Me.ComboGarantia = "No"
    m_SQL = "SELECT DISTINCT TbMaterial.Material, TbMaterial.PN, TbMaterial.NS " & _
            "FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
            "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento " & _
            "WHERE ((Not (TbMaterial.Material) Is Null) And ((TbEventos.IDEquipo) =" & strIDEquipo & ")) " & _
            "ORDER BY TbMaterial.Material;"
    Me.ComboMaterial.RowSource = m_SQL
    Me.ComboMaterial.Requery
    Me.FechaIncio = Format(Now(), "dd/mm/yyyy")
    Me.ComboMaterial.SetFocus
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Material_Exit(Cancel As Integer)
    Dim strValor As String
    On Error Resume Next
    strValor = Nz(Me.Material, "")
    If strValor <> strUltimoMaterialNoEnCombo And strValor <> "" Then
        Me.ComboMaterial = Null
    End If
    strUltimoMaterialNoEnCombo = strValor
End Sub