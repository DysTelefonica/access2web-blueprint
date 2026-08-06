Option Compare Database
Option Explicit
Dim strIDMaterial As String, strEsDeActividad As String, strEscritura As String, intNumeroBarras As Integer
Dim strMaterialAlInicio As String, strPNAlInicio As String, strNSAlinicio As String, strCosteAlInicio As String, strFechaEntregaAlInicio As String, strReparadoPorAlInicio As String, _
    strTipoAccionAlInicio As String, strDescripcionAlInicio As String, strFechaInicialAlInicio As String, strGarantiaAlInicio As String
Public strURLAnexo As String, strEstaFacturado As String, strFranqueado As String, strFormularioQueAbre As String
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
    Dim strMaterial As String, strPN As String, strNS As String, strCOSTE As String, strFechaEntrega As String, strReparadoPor As String, _
        strTipoAccion As String, strDESCRIPCION As String, strFechaInicial As String, strGarantia As String
    On Error GoTo errores
    strMaterial = Nz(Me.ComboMaterial, "")
    strPN = Nz(Me.PN, "")
    strNS = Nz(Me.NS, "")
    strCOSTE = Nz(Me.COSTE, "")
    strReparadoPor = Nz(Me.ComboReparadoPor, "")
    strDESCRIPCION = Nz(Me.Descripcion, "")
    strFechaEntrega = Nz(Me.FechaEntrega, "")
    strTipoAccion = Nz(Me.ComboTipoAccion, "")
    strFechaInicial = Nz(Me.FechaIncio, "")
    strGarantia = Nz(Me.ComboGarantia, "")
    If IsDate(strFechaInicial) Then
        strFechaInicial = Format(strFechaInicial, "dd/mm/yyyy")
    End If
    If IsDate(strFechaEntrega) Then
        strFechaEntrega = Format(strFechaEntrega, "dd/mm/yyyy")
    End If
    If IsNumeric(strCOSTE) Then
        strCOSTE = CStr(CDbl(strCOSTE))
    End If
    If IsNumeric(strCosteAlInicio) Then
        strCosteAlInicio = CStr(CDbl(strCosteAlInicio))
    End If
    If strMaterial <> strMaterialAlInicio Or strPN <> strPNAlInicio Or strNS <> strNSAlinicio Or strCOSTE <> strCosteAlInicio Or _
        strFechaEntrega <> strFechaEntregaAlInicio Or strReparadoPor <> strReparadoPorAlInicio Or strTipoAccion <> strTipoAccionAlInicio Or _
        strDESCRIPCION <> strDescripcionAlInicio Or strFechaInicial <> strFechaInicialAlInicio Or strGarantia <> strGarantiaAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Private Sub cmdSalir_Click()
    Dim strTextoError As String, strNombreFormulario As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    '---------------------------------------------
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '---------------------------------------------
    If HaHabidoCambios() And strEscritura = "Sí" Then
        pregunta = MsgBox("¿Desea guardar los cambios producidos?", vbExclamation + vbYesNo + vbDefaultButton1, "Guardar Cambios")
        If pregunta = vbYes Then
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            '-------------------------------------------------------------------
            '   -Devuelve:
            '       Registrar = Descriptivo
            '       Registrar = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = Registrar()
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
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
        strTextoError = "No se ha podido determinar el IDMaterial"
        Err.Raise 1000
    End If
    Set m_Material = Constructor.getMaterial(strIDMaterial, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
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

Private Function Registrar() As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 10/04/13
    '   -Modificaciones:
    '       -11/09/2014: Se incluye la FechaInicial y la Garantía
    '       -06/10/2014: Se incluye un campo del tbMaterial -->EsReparacion.
    '   -Funcionamiento:
    '       -va a editar un registro en la tabla TbMaterial
    
    '       -strMaterial<>"" y strMaterial<>"#"
    '       -Si no hay coste ha de ser 0
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------------
    Dim strMaterial As String, strPN As String, strNS As String, strCOSTE As String, strReparadoPor As String, _
        strTipoAccion As String, strDESCRIPCION As String, strFechaInicial As String, strGarantia As String
    Dim strEsReparacion As String, strFechaEntrega As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    strMaterial = Nz(Me.ComboMaterial, "")
    strPN = Nz(Me.PN, "#")
    strNS = Nz(Me.NS, "#")
    strCOSTE = Nz(Me.COSTE, "#")
    strFechaInicial = Nz(Me.FechaIncio, "#")
    strGarantia = Nz(Me.ComboGarantia, "#")
    strTipoAccion = Nz(Me.ComboTipoAccion, "#")
    strReparadoPor = Nz(Me.ComboReparadoPor, "#")
    strEsReparacion = Nz(Me.ComboEsReparacion, "#")
    strDESCRIPCION = Nz(Me.Descripcion, "#")
    strFechaEntrega = Nz(Me.FechaEntrega, "#")
    If strEsDeActividad = "Sí" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       EdicionMaterialEnActividad = Descriptivo
        '       EdicionMaterialEnActividad = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EdicionMaterialEnActividad(strIDMaterial, strMaterial, strCOSTE, strTipoAccion, _
                                                    strPN, strNS, strReparadoPor, strDESCRIPCION, strFechaInicial, strGarantia, strEsReparacion, strFechaEntrega)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EdicionMaterialEnActividad ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        '----------------------------------------------------------------------
        '   -Devuelve:
        '       EdicionMaterialSinActividad = Descriptivo
        '       EdicionMaterialSinActividad = "#ERR" & "|" & strTextoError
        '----------------------------------------------------------------------
        flag = EdicionMaterialSinActividad(strIDMaterial, strMaterial, strCOSTE, strTipoAccion, _
                                                        strPN, strNS, strReparadoPor, strDESCRIPCION, strFechaInicial, strGarantia, strEsReparacion, strFechaEntrega)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EdicionMaterialSinActividad ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    Registrar = "Material editado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoRegistrar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '-----------------------------------------------
    If Not HaHabidoCambios() Then
        strTextoError = "No se ha registrado ningún cambio"
        Err.Raise 1000
    End If
    '----------------------------------------------------------------------
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto un error: " & vbNewLine & dato(1)
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


Private Function EstablecerDatos() As String
    Dim m_Material As Material
    
    Dim rcdDatos As DAO.Recordset, fso As New FileSystemObject, strNodo As String, strBUI As String, strSUBSISTEMA As String, _
        strEquipo As String, strIDEvento As String, strIDActividad As String, strIDEquipo As String, ctl As Control
    Dim strEsReparacion As String, strFechaEntrega As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterial.IDEquipo, TbMaterial.IDActividad,TbMaterial.EsReparacion,TbMaterial.FechaEntrega " & _
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
        strEsReparacion = Nz(.Fields("EsReparacion"), "")
        strFechaEntrega = Nz(.Fields("FechaEntrega"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If intNumeroBarras = 0 Then
        
        If strIDActividad <> "" Then
            '-----------------------------------------------
            '   -Devuelve:
            '       EstaFranqueado = strValorFranqueado
            '       EstaFranqueado = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = EstaFranqueado("MA", strIDMaterial)
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método EstaFranqueado ha devuelto un error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            strFranqueado = flag
        Else
            strFranqueado = "No"
        End If
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstaFacturado = strEstaFacturado
        '       EstaFacturado = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = EstaFacturado("MA", strIDMaterial)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFacturado ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEstaFacturado = flag
        Set m_Material = Constructor.getMaterial(strIDMaterial, strTextoError)
        If strTextoError <> "" Then
            Err.Raise 1000
        End If
        strURLAnexo = m_Material.URLAnexos
        
    End If
    
    
    
    If strEstaFacturado = "Sí" Then
        Me.lblFacturado.Visible = True
    Else
        Me.lblFacturado.Visible = False
    End If
    If strIDActividad <> "" Then
        strEsDeActividad = "Sí"
    Else
        strEsDeActividad = "No"
    End If
    If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        strEsReparacion = "No"
    End If
    Me.ComboEsReparacion = strEsReparacion
    If IsDate(strFechaEntrega) Then
        Me.FechaEntrega = Format(strFechaEntrega, "dd/mm/yyyy")
    End If
    If strEsDeActividad = "Sí" Then
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
                strNodo = Nz(.Fields("NODO"), "")
                strBUI = Nz(.Fields("BUI"), "")
                strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
                strEquipo = Nz(.Fields("Equipo"), "")
                strIDEvento = Nz(.Fields("IDEvento"), "")
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        If strNodo <> "" Then
            Me.NODO = strNodo
        End If
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
    Else
        Me.ComandoVerEvento.Enabled = False
        m_SQL = "SELECT TbEquipos.NODO, TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo " & _
                "FROM TbEquipos " & _
                "WHERE (((TbEquipos.IDEquipo)=" & strIDEquipo & "));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "El Equipo del que depende este material: " & strIDEquipo & " no está registrado en la tabla equipos"
                Err.Raise 1000
            End If
            strNodo = Nz(.Fields("NODO"), "")
            strBUI = Nz(.Fields("BUI"), "")
            strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
            strEquipo = Nz(.Fields("Equipo"), "")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        If strNodo <> "" Then
            Me.NODO = strNodo
        End If
        If strBUI <> "" Then
            Me.BUI = strBUI
        End If
        If strSUBSISTEMA <> "" Then
            Me.SubSistema = strSUBSISTEMA
        End If
        If strEquipo <> "" Then
            Me.Equipo = strEquipo
        End If
    End If
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
            strMaterialAlInicio = Nz(.Fields("Material"), "")
            strPNAlInicio = Nz(.Fields("PN"), "")
            strNSAlinicio = Nz(.Fields("NS"), "")
            strCosteAlInicio = Nz(.Fields("COSTE"), "")
            strFechaEntregaAlInicio = Nz(.Fields("FechaEntrega"), "")
            strTipoAccionAlInicio = Nz(.Fields("TipoAccion"), "")
            strReparadoPorAlInicio = Nz(.Fields("ReparadoPor"), "")
            strDescripcionAlInicio = Nz(.Fields("Descripcion"), "")
            strFechaInicialAlInicio = Nz(.Fields("FechaIncio"), "")
            strGarantiaAlInicio = Nz(.Fields("Garantia"), "")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        If strMaterialAlInicio <> "" Then
            Me.ComboMaterial = strMaterialAlInicio
        End If
        If strPNAlInicio <> "" Then
            Me.PN = strPNAlInicio
        End If
        If strNSAlinicio <> "" Then
            Me.NS = strNSAlinicio
        End If
        If IsNumeric(strCosteAlInicio) Then
            Me.COSTE = Format(strCosteAlInicio, "#,##0.00 €")
            strCosteAlInicio = CStr(CDbl(strCosteAlInicio))
        End If
        If strTipoAccionAlInicio <> "" Then
            Me.ComboTipoAccion = strTipoAccionAlInicio
        End If
        If IsDate(strFechaEntregaAlInicio) Then
            strFechaEntregaAlInicio = Format(strFechaEntregaAlInicio, "dd/mm/yyyy")
            Me.FechaEntrega = strFechaEntregaAlInicio
        End If
        If strReparadoPorAlInicio <> "" Then
            Me.ComboReparadoPor = strReparadoPorAlInicio
        End If
        If strDescripcionAlInicio <> "" Then
            Me.Descripcion = strDescripcionAlInicio
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
        If Nz(Me.Equipo, "") <> "" Then
            m_SQL = "SELECT DISTINCT TbMaterial.Material " & _
                    "FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
                    "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) " & _
                    "INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                    "WHERE (((TbEquipos.Equipo)='" & Me.Equipo & "')) " & _
                    "ORDER BY TbMaterial.Material;"
            
        Else
            m_SQL = "SELECT DISTINCT TbMaterial.Material " & _
                    "FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
                    "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) " & _
                    "INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                    "WHERE (((TbEquipos.Equipo)='-1')) " & _
                    "ORDER BY TbMaterial.Material;"
        End If
        Me.ComboMaterial.RowSource = m_SQL
        Me.ComboMaterial.Requery
    Me.cmdSalir.SetFocus
    If strFranqueado = "Sí" Or strEstaFacturado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    If strEscritura = "No" Then
        Me.lblTitulo1.Caption = "DETALLE DE MATERIAL.REPARACIÓN"
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
        Me.lblTitulo1.Caption = "EDICIÓN DE MATERIAL.REPARACIÓN"
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
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
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
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    'strArgumento=strIDMaterial  & "|" & strFranqueado & "|" & strEstaFacturado & "|" & strURLAnexo
    'strArgumento=strIDMaterial
    strArgumento = Nz(Me.OpenArgs(), "")
    If strArgumento = "" Then
        strTextoError = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    If InStr(1, strArgumento, "|") = 0 Then
        intNumeroBarras = 0
        strIDMaterial = strArgumento
    Else
        dato = Split(strArgumento, "|")
        intNumeroBarras = UBound(dato)
        If intNumeroBarras <> 3 Then
            strTextoError = "El formulario se ha abierto con parámetros insuficientes"
            Err.Raise 1000
        End If
        strIDMaterial = dato(0)
        strFranqueado = dato(1)
        strEstaFacturado = dato(2)
        strURLAnexo = dato(3)
    End If
    
    
    If strIDMaterial = "" Then
        strTextoError = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    '-----------------------------------------------
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
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ImagenAnexoMaterial_Click()
    Dim strTextoError As String
    On Error GoTo errores
    DoCmd.Hourglass True
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
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub







