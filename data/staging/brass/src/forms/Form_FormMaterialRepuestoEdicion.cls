Option Compare Database
Option Explicit
Dim strIDMaterial As String, strMaterialConocidoAlInicio As String, strMaterialNoConocidoAlInicio As String, strPNAlInicio As String, strNSAlinicio As String, strCosteAlInicio As String, strFechaAlInicio As String, strReparadoPorAlInicio As String, _
    strTipoAccionAlInicio As String, strObservacionesAlInicio As String, strGarantiaAlInicio As String, strFechaInicialAlInicio As String
Dim strURLAnexo As String, strEstaFacturado As String, strFranqueado As String, strEscritura As String, strFormularioQueAbre As String
Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  25/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '-------------------------------------------------------------------
    Dim strMaterialConocido As String, strMaterialNoConocido As String, strPN As String, strNS As String, strCOSTE As String, strFechaEntrega As String, strReparadoPor As String, _
        strTipoAccion As String, strDESCRIPCION As String, strFechaInicial As String, strGarantia As String
    On Error GoTo errores
    strMaterialConocido = Nz(Me.ComboMaterial, "")
    strMaterialNoConocido = Nz(Me.Material, "")
    strPN = Nz(Me.PN, "")
    strNS = Nz(Me.NS, "")
    strCOSTE = Nz(Me.COSTE, "")
    If IsNumeric(strCOSTE) Then
        strCOSTE = CStr(CDbl(strCOSTE))
    End If
    strReparadoPor = Nz(Me.ComboReparadoPor, "")
    strDESCRIPCION = Nz(Me.Observaciones, "")
    strTipoAccion = Nz(Me.ComboTipoAccion, "")
    strFechaInicial = Nz(Me.FechaIncio, "")
    strGarantia = Nz(Me.ComboGarantia, "")
    If IsDate(strFechaInicial) Then
        strFechaInicial = Format(strFechaInicial, "dd/mm/yyyy")
    End If
    If strMaterialConocido <> strMaterialConocidoAlInicio Or strMaterialNoConocido <> strMaterialNoConocidoAlInicio Or strPN <> strPNAlInicio Or strNS <> strNSAlinicio Or strCOSTE <> strCosteAlInicio Or _
        strReparadoPor <> strReparadoPorAlInicio Or strTipoAccion <> strTipoAccionAlInicio Or _
        strDESCRIPCION <> strObservacionesAlInicio Or strFechaInicial <> strFechaInicialAlInicio Or strGarantia <> strGarantiaAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Private Sub cmdSalir_Click()
    Dim strNombreFormulario As String, strTextoError As String
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
            '   -Devuelve
            '       FormularioAbierto = true or false
            '       FormularioAbierto =False-->Error
            '---------------------------------------------
            If FormularioAbierto("FormEventoGestion") Then
                Forms("FormEventoGestion").Controls("ListaMateriales").Requery
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
    Dim m_Material As Material
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strIDMaterial = "" Then
        strTextoError = "Se ha de Seleccionar un material primero"
        Err.Raise 1000
    End If
    
    '-----------------------------------------------
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '-----------------------------------------------
    If HaHabidoCambios Then
        '-------------------------------------------------------------------
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
        '-------------------------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '---------------------------------------------
        If FormularioAbierto("FormEventoGestion") Then
            Forms("FormEventoGestion").Controls("ListaMateriales").Requery
        End If
        
    End If
    Set m_Material = Constructor.getMaterial(strIDMaterial, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    strFormularioQueAbre = Me.Name
    Set m_ObjEntidadParaAnexoActiva = m_Material
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
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    '-----------------------------------------------
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '-----------------------------------------------
    If Not HaHabidoCambios() Then
        strTextoError = "No ha habido cambios"
        Err.Raise 1000
    End If
    '-----------------------------------------------
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
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        Forms("FormEventoGestion").Controls("ListaMateriales").Requery
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Function Registrar() As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 10/04/13
    '   -Modificaciones:
    '   -Funcionamiento:
   
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------------
    Dim strMaterial As String, strGarantia As String, strReparadoPor As String, strTipoAccion As String, _
        strFechaInicial As String, strCOSTE As String, strPN As String, strNS As String, strObservaciones As String, _
        strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    strMaterial = Nz(Me.ComboMaterial, "#")
    strCOSTE = Nz(Me.COSTE, "0")
    strTipoAccion = Nz(Me.ComboTipoAccion, "#")
    strPN = Nz(Me.PN, "#")
    strNS = Nz(Me.NS, "#")
    strReparadoPor = Nz(Me.ComboReparadoPor, "#")
    strObservaciones = Nz(Me.Observaciones, "#")
    strFechaInicial = Nz(Me.FechaIncio, "#")
    strGarantia = Nz(Me.ComboGarantia, "#")
    '-----------------------------------------------
    '   -Devuelve:
    '       EdicionMaterialEnActividad = Descriptivo
    '       EdicionMaterialEnActividad = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EdicionMaterialEnActividad(strIDMaterial, strMaterial, strCOSTE, strTipoAccion, _
                                                strPN, strNS, strReparadoPor, strObservaciones, strFechaInicial, strGarantia, "No", strFechaInicial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EdicionMaterialEnActividad ha devuelto un error: " & vbNewLine & dato(1)
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
    Dim strTextoError As String, strValor As String, strNS As String, strPN As String
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
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strIDMaterial  & "|" & strFranqueado & "|" & strEstaFacturado  & "|" & strURLAnexo
   
    strArgumento = Nz(Me.OpenArgs, "")
    If strArgumento = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 3 Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    strIDMaterial = dato(0)
    strFranqueado = dato(1)
    strEstaFacturado = dato(2)
    strURLAnexo = dato(3)
    
    If strIDMaterial = "" Then
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
    
    Dim rcdDatos As DAO.Recordset, fso As New FileSystemObject, strNombreArchivo As String, strNodo As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, _
        strIDEvento As String, strIDActividad As String, strIDEquipo As String, ctl As Control, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterial.IDEquipo, TbMaterial.IDActividad,TbMaterial.EsReparacion " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El formulario no se ha abierto con los datos suficientes"
            Err.Raise 1000
        End If
        strIDActividad = Nz(.Fields("IDActividad"), "")
        strIDEquipo = Nz(.Fields("IDEquipo"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
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
    '-----------
    ' carga de los datos del evento
    '-------------
    m_SQL = "SELECT TbActividades.IDEvento, TbEventos.NODO, TbEventos.BUI, TbEventos.SUBSISTEMA, TbEventos.IDEquipo, TbEquipos.Equipo " & _
            "FROM (TbActividades INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) INNER JOIN TbEquipos ON " & _
            "TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
            "WHERE (((TbActividades.IDActividad)='" & strIDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strBUI = Nz(.Fields("BUI"), "")
            strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
            strEquipo = Nz(.Fields("Equipo"), "")
            strIDEvento = Nz(.Fields("IDEvento"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Me.IDMaterial = strIDMaterial
    If strBUI <> "" Then
        Me.BUI = strBUI
    End If
    If strSUBSISTEMA <> "" Then
        Me.SubSistema = strSUBSISTEMA
    End If
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
    If strIDEvento <> "" Then
        Me.IDEVENTO = strIDEvento
        Me.ComandoVerEvento.Enabled = True
    Else
        Me.ComandoVerEvento.Enabled = False
    End If
    If strIDActividad <> "" Then
        Me.IDActividad = strIDActividad
    End If
    m_SQL = "SELECT DISTINCT TbMaterial.Material, TbMaterial.PN, TbMaterial.NS " & _
            "FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
            "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento " & _
            "WHERE ((Not (TbMaterial.Material) Is Null) And ((TbEventos.IDEquipo) =" & strIDEquipo & ")) " & _
            "ORDER BY TbMaterial.Material;"
    Me.ComboMaterial.RowSource = m_SQL
    Me.ComboMaterial.Requery
    '-----------
    ' carga de los datos del Material
    '-------------
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strMaterialConocidoAlInicio = Nz(.Fields("Material"), "")
        strPNAlInicio = Nz(.Fields("PN"), "")
        strNSAlinicio = Nz(.Fields("NS"), "")
        strCosteAlInicio = Nz(.Fields("COSTE"), "")
        strFechaAlInicio = Nz(.Fields("FechaEntrega"), "")
        strTipoAccionAlInicio = Nz(.Fields("TipoAccion"), "")
        strReparadoPorAlInicio = Nz(.Fields("ReparadoPor"), "")
        strObservacionesAlInicio = Nz(.Fields("Descripcion"), "")
        strFechaInicialAlInicio = Nz(.Fields("FechaIncio"), "")
        strGarantiaAlInicio = Nz(.Fields("Garantia"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strMaterialConocidoAlInicio <> "" Then
        Me.ComboMaterial = strMaterialConocidoAlInicio
    End If
    If strPNAlInicio <> "" Then
        Me.PN = strPNAlInicio
    End If
    If strNSAlinicio <> "" Then
        Me.NS = strNSAlinicio
    End If
    If IsNumeric(strCosteAlInicio) Then
        Me.COSTE = Format(strCosteAlInicio, "#,##0.00 €")
    End If
    If strTipoAccionAlInicio <> "" Then
        Me.ComboTipoAccion = strTipoAccionAlInicio
    End If
    If IsDate(strFechaAlInicio) Then
        strFechaAlInicio = Format(strFechaAlInicio, "dd/mm/yyyy")
        Me.FechaIncio = strFechaAlInicio
    End If
    If strReparadoPorAlInicio <> "" Then
        Me.ComboReparadoPor = strReparadoPorAlInicio
    End If
    If strObservacionesAlInicio <> "" Then
        Me.Observaciones = strObservacionesAlInicio
    End If
    If IsDate(strFechaInicialAlInicio) Then
        strFechaInicialAlInicio = Format(strFechaInicialAlInicio, "dd/mm/yyyy")
        Me.FechaIncio = strFechaInicialAlInicio
    End If
    If strGarantiaAlInicio <> "" Then
        Me.ComboGarantia = strGarantiaAlInicio
    End If
    If strURLAnexo <> "" Then
        
        Me.ImagenAnexoMaterial.Visible = True
    Else
         Me.ImagenAnexoMaterial.Visible = False
    End If
    Me.cmdSalir.SetFocus
    For Each ctl In Me.Controls
        If ctl.Tag = "FIJOS" Then
            ctl.BackColor = lngColorCampoNoEditable
            ctl.Enabled = True
            ctl.Locked = True
        End If
    Next
    If strEscritura = "No" Then
        Me.lblTitulo1.Caption = "DETALLE DE MATERIAL COMO REPUESTO"
        For Each ctl In Me.Controls
            If ctl.Tag <> "FIJOS" Then
                If (ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox) Then
                    ctl.BackColor = lngColorCampoNoEditable
                    ctl.Enabled = True
                    ctl.Locked = True
                End If
            End If
        Next
        Me.ComandoRegistrar.Enabled = False
        Me.lblFranqueado.Visible = True
    Else
        Me.lblTitulo1.Caption = "EDICIÓN DE MATERIAL COMO REPUESTO"
        For Each ctl In Me.Controls
            If ctl.Tag <> "FIJOS" Then
                If (ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox) Then
                    ctl.BackColor = lngColorCampoEditable
                    ctl.Enabled = True
                    ctl.Locked = False
                End If
            End If
        Next
        Me.ComandoRegistrar.Enabled = True
        Me.lblFranqueado.Visible = False
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
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
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    EstablecerDatos = "#ERR" & "|" & strTextoError
End Function
Private Sub ImagenAnexoMaterial_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If InStr(1, strURLAnexo, ";") <> 0 Then
        ComandoAdjuntar_Click
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
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





