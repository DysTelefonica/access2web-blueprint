Option Compare Database
Option Explicit
Dim strIDIntervencion As String, strEstaFacturado As String, strFranqueado As String, strIDMaterial As String, strEscritura As String
Dim strFechaFinIntervencionAlInicio As String, strFechaInicioIntervencionAlInicio As String, _
        strMotivoAlInicio As String, strResultadoAlInicio As String, strDestinoAlInicio As String, strComboUltimaIntervencionAlInicio As String
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
            '-------------------------------------------------------------------
            '       FormularioAbierto = true or false
            '       FormularioAbierto =False-->Error
            '--------------------------------------------
            If FormularioAbierto("FormEventoGestion") Then
                Forms("FormEventoGestion").Controls("ListaActividades").Requery
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
        End If
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Function HaHabidoCambios() As Boolean
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/10/2013
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a ver si ha habido cambios desde el momento que se abrió el formulario
    '   -Llamada desde:
    
    '   -Devuelve:
    '       HaHabidoCambios = True or false
    
    '           strEstado="1"---->True or false
    '           strEstado="-1"--->True
    '-------------------------------------------------------------------
    Dim strFechaInicioIntervencion As String, strFechaFinIntervencion As String, strMotivo As String, strResultado As String, strDestino As String, _
        strComboUltimaIntervencion As String
    On Error GoTo errores
    strFechaInicioIntervencion = Nz(Me.FechaInicioIntervencion, "")
    strFechaFinIntervencion = Nz(Me.FechaFinIntervencion, "")
    strMotivo = Nz(Me.MOTIVO, "")
    strResultado = Nz(Me.RESULTADO, "")
    strDestino = Nz(Me.DESTINO, "")
    strComboUltimaIntervencion = Nz(Me.ComboUltimaIntervencion, "")
    If IsDate(strFechaInicioIntervencion) Then
        strFechaInicioIntervencion = Format(strFechaInicioIntervencion, "dd/mm/yyyy")
    End If
    If IsDate(strFechaFinIntervencion) Then
        strFechaInicioIntervencion = Format(strFechaFinIntervencion, "dd/mm/yyyy")
    End If
    If IsDate(strFechaInicioIntervencionAlInicio) Then
        strFechaInicioIntervencionAlInicio = Format(strFechaInicioIntervencionAlInicio, "dd/mm/yyyy")
    End If
    If IsDate(strFechaFinIntervencionAlInicio) Then
        strFechaFinIntervencionAlInicio = Format(strFechaFinIntervencionAlInicio, "dd/mm/yyyy")
    End If
    If strFechaInicioIntervencionAlInicio <> strFechaInicioIntervencion Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strFechaFinIntervencionAlInicio <> strFechaFinIntervencion Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strMotivoAlInicio <> strMotivo Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strResultadoAlInicio <> strResultado Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strDestinoAlInicio <> strDestino Then
        HaHabidoCambios = True
        Exit Function
    End If
    If strComboUltimaIntervencion <> strComboUltimaIntervencionAlInicio Then
        HaHabidoCambios = True
        Exit Function
    End If
    
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Private Function Registrar() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/10/2013
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a emplear la función Material.IntervencionEdicion
    '   -Llamada desde:
    '       -me.ComandoRegistrar_Click
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strFechaInicioIntervencion As String, strFechaFinIntervencion As String, strMotivo As String, strResultado As String, _
        strDestino As String, strComboUltimaIntervencion As String, strTextoError As String
    On Error GoTo errores
    If strIDIntervencion = "" Then
        strTextoError = "No se ha podido determinar cuál es la intervención de material"
        Err.Raise 1000
    End If
    strFechaInicioIntervencion = Nz(Me.FechaInicioIntervencion, "#")
    strFechaFinIntervencion = Nz(Me.FechaFinIntervencion, "#")
    strMotivo = Nz(Me.MOTIVO, "#")
    strResultado = Nz(Me.RESULTADO, "#")
    strDestino = Nz(Me.DESTINO, "#")
    strComboUltimaIntervencion = Nz(Me.ComboUltimaIntervencion, "#")
    '-----------------------------------------------
    '   -Devuelve:
    '       IntervencionEdicion = Descriptivo
    '       IntervencionEdicion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = IntervencionEdicion(strIDIntervencion, strIDMaterial, strFechaInicioIntervencion, strDestino, _
                                            strMotivo, strResultado, strFechaFinIntervencion, strComboUltimaIntervencion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método IntervencionEdicion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------
    If FormularioAbierto("FormMaterialSeguimiento") Then
        Form_FormMaterialSeguimiento.EstablecerDatos
    End If
    Registrar = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El Registrar se ha devuelto el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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

Private Sub Form_Open(Cancel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    strIDIntervencion = Nz(Me.OpenArgs, "")
    If strIDIntervencion = "" Then
        strTextoError = "Se ha abierto con parámetros inadecuados"
        Err.Raise 1000
    End If
    
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
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
        strTextoError = "Al Abrir el formulario se ha devuelto el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, strProvieneDeEvento As String, strIDActividad As String, _
        strProvieneDeActividad As String, ctl As Control, strIDEvento As String, strEquipo As String, strMaterial As String, strTextoError As String
    On Error GoTo errores
    If strIDIntervencion = "" Then
        strTextoError = "Se ha abierto con parámetros inadecuados"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterialSeguimiento.*,  TbFacturaMaterialesInvolucrados.IDFactura,TbMaterial.IDActividad " & _
            "FROM (TbMaterialSeguimiento INNER JOIN TbMaterial ON " & _
            "TbMaterialSeguimiento.IDMaterial = TbMaterial.IDMaterial) " & _
            "LEFT JOIN TbFacturaMaterialesInvolucrados ON " & _
            "TbMaterialSeguimiento.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial " & _
            "WHERE (((TbMaterialSeguimiento.IDSeguimiento )='" & strIDIntervencion & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "Se ha abierto con parámetros inadecuados"
            Err.Raise 1000
        End If
        strFechaInicioIntervencionAlInicio = Nz(.Fields("FechaInicioIntervencion"), "")
        strFechaFinIntervencionAlInicio = Nz(.Fields("FechaFinIntervencion"), "")
        strMotivoAlInicio = Nz(.Fields("Motivo"), "")
        strResultadoAlInicio = Nz(.Fields("Resultado"), "")
        strDestinoAlInicio = Nz(.Fields("Destino"), "")
        strComboUltimaIntervencionAlInicio = Nz(.Fields("UltimaIntervencion"), "")
        strIDMaterial = Nz(.Fields("IDMaterial"), "")
        
        strIDActividad = Nz(.Fields("IDActividad"), "")
        If strIDActividad <> "" Then
            strProvieneDeActividad = "Sí"
            If .Fields("FechaInicioIntervencion") = True Then
                strFranqueado = "Sí"
            Else
                strFranqueado = "No"
            End If
        Else
            strProvieneDeActividad = "No"
            strFranqueado = "No"
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If IsDate(strFechaInicioIntervencionAlInicio) Then
        Me.FechaInicioIntervencion = strFechaInicioIntervencionAlInicio
    End If
    If IsDate(strFechaFinIntervencionAlInicio) Then
        Me.FechaFinIntervencion = strFechaFinIntervencionAlInicio
    End If
    If strMotivoAlInicio <> "" Then
        Me.MOTIVO = strMotivoAlInicio
    End If
    If strResultadoAlInicio <> "" Then
        Me.RESULTADO = strResultadoAlInicio
    End If
    If strDestinoAlInicio <> "" Then
        Me.DESTINO = strDestinoAlInicio
    End If
    If strComboUltimaIntervencionAlInicio <> "" Then
        Me.ComboUltimaIntervencion = strComboUltimaIntervencionAlInicio
    End If
    '---------------
    ' OBTENCIÓN DE VALORES DE LOS DATOS PREVIOS DEL MATERIAL
    '---------------
    If strIDMaterial = "" Then
        strTextoError = "No se ha podido obtener el IDMaterial a partir del IDSeguimiento: " & strIDIntervencion
        Err.Raise 1000
    End If
    If strProvieneDeActividad = "Sí" Then
        m_SQL = "SELECT TbEquipos.Equipo, TbEventos.BUI, TbEventos.SUBSISTEMA, TbMaterial.Material, TbActividades.IDEvento " & _
                "FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
                "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) " & _
                "INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "Se ha abierto con parámetros inadecuados"
                Err.Raise 1000
            End If
            strEquipo = Nz(.Fields("Equipo"), "")
            strIDEvento = Nz(.Fields("IDEvento"), "")
            strMaterial = Nz(.Fields("Material"), "")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    Else
        m_SQL = "SELECT TbEquipos.Equipo, TbMaterial.Material, TbEquipos.BUI, TbEquipos.SUBSISTEMA " & _
                    "FROM TbMaterial INNER JOIN TbEquipos ON TbMaterial.IDEquipo = TbEquipos.IDEquipo " & _
                    "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "Se ha abierto con parámetros inadecuados"
                Err.Raise 1000
            End If
            strEquipo = Nz(.Fields("Equipo"), "")
            strMaterial = Nz(.Fields("Material"), "")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
    If strMaterial <> "" Then
        Me.Material = strMaterial
    End If
    If strIDEvento <> "" Then
        Me.IDEVENTO = strIDEvento
        Me.ComandoVerEvento.Enabled = True
    Else
        Me.ComandoVerEvento.Enabled = False
    End If
    If strFranqueado = "Sí" Then
        Me.lblFranqueado.Visible = True
    Else
        Me.lblFranqueado.Visible = False
    End If
    
    If strEstaFacturado = "Sí" Then
        Me.lblFacturado.Visible = True
    Else
        Me.lblFacturado.Visible = False
    End If
    If strFranqueado = "Sí" Or strEstaFacturado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    If strEscritura = "No" Then
        Me.ComandoRegistrar.Enabled = False
        Me.lblTitulo1.Caption = "Detalle de Intervención de Material"
        For Each ctl In Me.Controls
            If ctl.Tag <> "FIJOS" Then
                If (ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox) Then
                    ctl.BackColor = lngColorCampoNoEditable
                    ctl.Enabled = True
                    ctl.Locked = True
                End If
            End If
        Next
    Else
        Me.ComandoRegistrar.Enabled = True
        Me.lblTitulo1.Caption = "Edición de Intervención de Material"
        For Each ctl In Me.Controls
            If ctl.Tag <> "FIJOS" Then
                If (ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox) Then
                    ctl.BackColor = lngColorCampoEditable
                    ctl.Enabled = True
                    ctl.Locked = False
                End If
            End If
        Next
    End If
    Me.Caption = Me.lblTitulo1.Caption
    EstablecerDatos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al EstablecerDatos se ha devuelto el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstablecerDatos = "#ERR" & "|" & strTextoError
End Function
Public Function IntervencionEdicion( _
                                    strIDIntervencion As String, _
                                    Optional strIDMaterial As String, _
                                    Optional strFechaInicioIntervencion As String, _
                                    Optional strDestino As String, _
                                    Optional strMotivo As String, _
                                    Optional strResultado As String, _
                                    Optional strFechaFinIntervencion As String, _
                                    Optional strUltimaIntervencion As String _
                                    ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/09/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Obtenemos la posición de la intervención dentro del seguimiento.
    '           A) Es la única intervención
    '               -Cotejamos strFechaInicioIntervencion>=TbMaterial.FechaInicio
    '               -Cotejamos strFechaFinIntervencion>strFechaInicioIntervencion
    '           B) Más de una Intervención
    '               -Ocupa el primer Lugar
    '                   -Cotejamos strFechaInicioIntervencion>=TbMaterial.FechaInicio
    '                   -Cotejamos strFechaFinIntervencion>strFechaInicioIntervencion
    '               -Ocupa un lugar intermedio
    '                   -Cotejamos strFechaInicioIntervencion>=strFechaFinIntervencionAnterior
    '                   -strFechaFinIntervencion Obligada
    '                   -strFechaFinIntervencion>=strFechaInicioIntervencion AND strFechaFinIntervencion<=strFechaInicioIntervencionSiguiente
    '                   -strUltimaIntervencion No puede ser Sí
    '               -Ocupa el último lugar
    '                   -Cotejamos strFechaInicioIntervencion>=strFechaInicioIntervencionAnterior
    '                   -Si strFechaFinIntervencion es fecha-->strFechaFinIntervencion>=strFechaInicioIntervencion
    '       -Grabar en la tabla TbMaterialSeguimiento
    '   -Llamada desde
    
    '   -Devuelve:
    '       IntervencionEdicion = Descriptivo
    '       IntervencionEdicion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaFinIntervencionAnterior As String, strFechaInicionIntervencionPosterior As String, _
        strFechaInicioIntervencionAnterior As String, strFechaFinalMaterial As String, strResultadoFuncion As String, strSeguimientoCerrado As String, _
        strUltimaFechaFinSeguimiento As String, strOrdenIntervencionActual As String, strNumeroIntervencionesEnSeguimiento As String, _
         blnEsUnicaIntervencion As Boolean, strLugarQueOcupa As String, strTextoError As String
    On Error GoTo errores
    
    If strIDIntervencion = "" Then
        strTextoError = "No se sabe qué intervención se desea editar"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaInicioIntervencion) Then
        strTextoError = "Es obligatoria la fecha de inicio de intervención"
        Err.Raise 1000
    End If
    If strUltimaIntervencion <> "Sí" And strUltimaIntervencion <> "No" Then
        strUltimaIntervencion = "No"
    End If
    If strDestino = "" Or strDestino = "#" Then
        strTextoError = "Destino obligatorio"
        Err.Raise 1000
    End If
     If strMotivo = "" Or strMotivo = "#" Then
        strTextoError = "Motivo obligatorio"
        Err.Raise 1000
    End If
    If IsDate(strFechaFinIntervencion) Then
        If CDate(strFechaFinIntervencion) < CDate(strFechaInicioIntervencion) Then
             strTextoError = "La fecha fin de intervención no puede ser anterior a la del inicio del mismo."
            Err.Raise 1000
        End If
    End If
    If IsDate(strFechaFinIntervencion) Then
        If strResultado = "" Or strResultado = "#" Then
            strTextoError = "La fecha fin de intervención implica indicar el resultado de la misma."
            Err.Raise 1000
        End If
    End If
    If strResultado <> "" And strResultado <> "#" Then
         If Not IsDate(strFechaFinIntervencion) Then
            strTextoError = "El resultado de la intervención implica el fin de la misma"
            Err.Raise 1000
         End If
    End If
    If strUltimaIntervencion = "Sí" Then
        If Not IsDate(strFechaFinIntervencion) Then
            strTextoError = "Si se marca como intervención final, se ha de indicar la fecha de fin."
            Err.Raise 1000
        End If
    End If
    If strIDMaterial = "" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       Dame = strDescriptivo
        '       Dame = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = Dame("TbMaterialSeguimiento", "IDMaterial", "IDSeguimiento", strIDIntervencion)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDMaterial = flag
    End If
    If strIDMaterial = "" Then
        strTextoError = "No se ha podido determinar de qué material es la intervención"
        Err.Raise 1000
    End If
'    '-----------------
'    ' VEMOS SI ES LA ÚNICA INTERVENCIÓN
'    '--------------
'    m_SQL = "SELECT TbMaterialSeguimiento.IDSeguimiento " & _
'                "FROM TbMaterialSeguimiento " & _
'                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "')) " & _
'                "ORDER BY TbMaterialSeguimiento.IDSeguimiento;"
'    '-------------------------------------------------------------
'    '   -Devuelve:
'    '       DameNumeroRegistrosPorSQL = CStr(lngRegistros)
'    '       DameNumeroRegistrosPorSQL = "#ERR" & "|" & strTextoError
'    '-------------------------------------------------------------------
'    flag = DameNumeroRegistrosPorSQL(m_SQL)
'    If InStr(1, flag, "|") <> 0 Then
'        dato = Split(flag, "|")
'        strTextoError = "El método DameNumeroRegistrosPorSQL ha devuelto el error: " & vbNewLine & dato(1)
'        Err.Raise 1000
'    End If
'    strNumeroIntervencionesEnSeguimiento = flag
'    If strNumeroIntervencionesEnSeguimiento = "1" Then
'        blnEsUnicaIntervencion = True
'    ElseIf strNumeroIntervencionesEnSeguimiento = "0" Then
'        strTextoError = "No está registrada la intervención"
'        Err.Raise 1000
'    Else
'        blnEsUnicaIntervencion = False
'    End If
'    If blnEsUnicaIntervencion = False Then
'        '-----------------------------------------------
'        '   -Devuelve:
'        '           strCodigoFinal =strNumeroIntervencionesEnSeguimiento;strOrdenIntervencionActual; _
'                        strFechaFinIntervencionAnterior;strFechaInicionIntervencionPosterior
'        '           strOrdenIntervencionActual="Primera o Intermedia o Última"
'        '       IntervencionDatosDefechasIntermedias = strCodigoFinal
'        '       IntervencionDatosDefechasIntermedias = "#ERR" & "|" & strTextoError
'        '-----------------------------------------------
'        flag = IntervencionDatosDefechasIntermedias(strIDIntervencion, strIDMaterial, strNumeroIntervencionesEnSeguimiento)
'        If InStr(1, flag, "|") <> 0 Then
'            dato = Split(flag, "|")
'            strTextoError = "El método IntervencionDatosDefechasIntermedias ha devuelto un error: " & vbNewLine & dato(1)
'            Err.Raise 1000
'        End If
'        If InStr(1, flag, ";") = 0 Then
'            strTextoError = "El método IntervencionDatosDefechasIntermedias ha devuelto un resultado con formato desconocido"
'            Err.Raise 1000
'        End If
'        dato = Split(flag, ";")
'        If UBound(dato) <> 3 Then
'            strTextoError = "El método IntervencionDatosDefechasIntermedias ha devuelto un resultado con formato desconocido"
'            Err.Raise 1000
'        End If
'        strNumeroIntervencionesEnSeguimiento = CStr(dato(0))
'        strOrdenIntervencionActual = CStr(dato(1))
'        strFechaFinIntervencionAnterior = CStr(dato(2))
'        strFechaInicionIntervencionPosterior = CStr(dato(3))
'        If IsDate(strFechaFinIntervencionAnterior) Then
'            If CDate(strFechaInicioIntervencion) < CDate(strFechaFinIntervencionAnterior) Then
'                strTextoError = "La fecha de inicio de la intervención es anterior al fin del proceso anterior"
'                Err.Raise 1000
'            End If
'            If IsDate(strFechaFinIntervencion) Then
'                If CDate(strFechaFinIntervencion) < CDate(strFechaInicioIntervencion) Then
'                    strTextoError = "La fecha final de la intervención es anterior a la inicial"
'                    Err.Raise 1000
'                Else
'                    If IsDate(strFechaInicionIntervencionPosterior) Then
'                        If CDate(strFechaFinIntervencion) > CDate(strFechaInicionIntervencionPosterior) Then
'                            strTextoError = "La fecha final de la intervención es anterior a la inicial de la intervención posterior"
'                            Err.Raise 1000
'                        End If
'                    End If
'                End If
'            End If
'        Else
'            If strFechaFinIntervencionAnterior = "#" Then
'                If Not strOrdenIntervencionActual = "Última" Then
'                    strTextoError = "Existe una intervención posterior, por lo que no puede anular la fecha final de éste."
'                    Err.Raise 1000
'                End If
'            End If
'        End If
'        If strUltimaIntervencion = "Sí" And Not strOrdenIntervencionActual = "Última" Then
'            strTextoError = "El cierre de un seguimiento sólo puede realizarse con la última intervención"
'            Err.Raise 1000
'        End If
'        If strResultado = "#" And Not strOrdenIntervencionActual = "Última" Then
'             strTextoError = "El Resultado de un seguimiento sólo puede borrarse en la última intervención"
'            Err.Raise 1000
'        End If
'    End If
    m_SQL = "SELECT TbMaterialSeguimiento.* " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDSeguimiento)='" & strIDIntervencion & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "La intervención con ID: " & strIDIntervencion & " no está registrada en la base de datos"
            Err.Raise 1000
        End If
        .Edit
            .Fields("FechaInicioIntervencion") = strFechaInicioIntervencion
            .Fields("Destino") = strDestino
            .Fields("Motivo") = strMotivo
            If IsDate(strFechaFinIntervencion) Then
                .Fields("FechaFinIntervencion") = strFechaFinIntervencion
            Else
                If strFechaFinIntervencion = "#" Then
                    .Fields("FechaFinIntervencion") = Null
                End If
            End If
            If strResultado <> "" And strResultado <> "#" Then
                .Fields("Resultado") = strResultado
            Else
                If strResultado = "#" Then
                    .Fields("Resultado") = Null
                End If
            End If
            .Fields("UltimaIntervencion") = strUltimaIntervencion
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoCerrado = "Sí/No"
    '       SeguimientoCerrado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoCerrado(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoCerrado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       UltimaFechaFinSeguimiento = strDescriptivo
        '       UltimaFechaFinSeguimiento = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = UltimaFechaFinSeguimiento(strIDMaterial)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método UltimaFechaFinSeguimiento ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strUltimaFechaFinSeguimiento = flag
        If IsDate(strUltimaFechaFinSeguimiento) Then
           m_SQL = "UPDATE TbMaterial SET TbMaterial.FechaEntrega = #" & Format(strUltimaFechaFinSeguimiento, "mm/dd/yyyy") & "# " & _
                   "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
           DoCmd.SetWarnings False
           DoCmd.RunSQL m_SQL
           DoCmd.SetWarnings True
        End If
    End If
    IntervencionEdicion = "Intervención editada correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método IntervencionEdicion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.SetWarnings True
    IntervencionEdicion = "#ERR" & "|" & strTextoError
End Function



