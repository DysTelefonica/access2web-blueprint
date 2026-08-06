Option Compare Database
Option Explicit
Dim strIDMaterial As String, strURLAnexo As String, strNumeroIntervencionesEnSeguimiento As String, strIDEventoAlInicio As String
Dim strFacturado As String, strEventoFranqueado As String, strBloqueado As String
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
Private Sub ComandoAdjuntos_Click()

     Dim m_MaterialSeguimiento As MaterialSeguimiento
    Dim strFormularioQueAbre As String
    Dim strIDIntervencionSeleccionado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDIntervencionSeleccionado = Nz(Me.ListaIntervenciones.Column(0), "")
    If strIDIntervencionSeleccionado = "" Then
        strTextoError = "Seleccione una intervención de la lista superior"
        Err.Raise 1000
    End If
    Set m_MaterialSeguimiento = Constructor.getMaterialSeguimiento(strIDIntervencionSeleccionado, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    Set m_ObjEntidadParaAnexoActiva = m_MaterialSeguimiento
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario de anexos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAlta_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
     '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoCerrado = "Sí/"No"
    '       SeguimientoCerrado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoCerrado(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoCerrado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "El proceso ya estaba finalizado previamente"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       MaterialSeguimientoConItervencionAbierta = strDescriptivo
    '       MaterialSeguimientoConItervencionAbierta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = MaterialSeguimientoConItervencionAbierta(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método MaterialSeguimientoConItervencionAbierta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "Hay una intervención previa sin cerrar"
        Err.Raise 1000
    End If
    '-------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormMaterialIntervencionAlta") Then
        DoCmd.Close acForm, "FormMaterialIntervencionAlta"
    End If
    DoCmd.OpenForm "FormMaterialIntervencionAlta", , , , , , strIDMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta una intervención se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoEdicion_Click()
    Dim strTextoError As String, strIDIntervencionSeleccionada As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDIntervencionSeleccionada = Nz(Me.ListaIntervenciones.Column(0), "")
     If strIDIntervencionSeleccionada = "" Then
        strTextoError = "No se ha seleccionado ninguna intervención"
        Err.Raise 1000
    End If
    '-------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormMaterialIntervencionEdicion") Then
        DoCmd.Close acForm, "FormMaterialIntervencionEdicion"
    End If
    DoCmd.OpenForm "FormMaterialIntervencionEdicion", , , , , , strIDIntervencionSeleccionada
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar una intervención se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminar_Click()
    Dim strIDIntervencionSeleccionada As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDIntervencionSeleccionada = Nz(Me.ListaIntervenciones.Column(0), "")
    If strIDIntervencionSeleccionada = "" Then
        strTextoError = "No se ha seleccionado ninguna intervención"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea eliminar la intervención seleccionada: " & strIDIntervencionSeleccionada & " ?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar intervención")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoMaterialEliminar = strDescriptivo
    '       SeguimientoMaterialEliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoMaterialEliminar(strIDIntervencionSeleccionada, "no", strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoMaterialEliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoMaterialEliminar = strDescriptivo
    '       SeguimientoMaterialEliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoMaterialEliminar(strIDIntervencionSeleccionada, "Sí", strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoMaterialEliminar ha devuelto un error: " & vbNewLine & dato(1)
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
        strTextoError = "El método SeguimientoMaterialEliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar una intervención se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoExportarExcel_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_SQL = Me.ListaIntervenciones.RowSource
    GenerarConsultasSQL m_SQL, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar exportar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Canel As Integer)
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento = strIDMaterialSeleccionado & "|" & strNumeroSeguimientos & "|" & strIDEvento  & "|" & strFacturado & "|" & strEventoFranqueado
    strArgumento = Nz(Me.OpenArgs(), "")
    If InStr(1, strArgumento, "|") = 0 Then
        strIDMaterial = strArgumento
    Else
        dato = Split(strArgumento, "|")
        If UBound(dato) <> 4 Then
            strTextoError = "No se ha abierto el formulario con los datos necesarios"
            Err.Raise 1000
        End If
        strIDMaterial = dato(0)
        strNumeroIntervencionesEnSeguimiento = dato(1)
        strIDEventoAlInicio = dato(2)
        strFacturado = dato(3)
        strEventoFranqueado = dato(4)
    End If
    If strIDMaterial = "" Then
        strTextoError = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstablecerDatos
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
Public Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, strIDEvento As String, strIDActividad As String, strEquipo As String, strMaterial As String, strBUI As String, _
        strSUBSISTEMA As String, strNodo As String, strGarantia As String, strResultadoUltimasFechas As String, strUltimaFechaInicial As String, _
        strUltimaFechaFinal As String, strTextoError As String, strUltimoDestino As String
    On Error GoTo errores
    
    If strFacturado = "Sí" Then
        Me.lblFacturado.Visible = True
    Else
        Me.lblFacturado.Visible = False
    End If
    If strEventoFranqueado = "Sí" Then
        Me.lblFranqueado.Visible = True
    Else
        Me.lblFranqueado.Visible = False
    End If
    If strFacturado = "Sí" Or strEventoFranqueado = "Sí" Then
        strBloqueado = "Sí"
    Else
        strBloqueado = "No"
    End If
    If strIDEventoAlInicio <> "" Then
        m_SQL = "SELECT TbMaterial.Garantia, TbMaterial.Material, TbActividades.IDEvento, TbActividades.IDActividad, TbEquipos.NODO, " & _
                    "TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo " & _
                    "FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
                    "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) " & _
                    "INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                    "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "No se ha obtenido ningún dato con el IDMaterial:" & strIDMaterial
                Err.Raise 1000
            End If
            strIDEvento = Nz(.Fields("IDEvento"), "")
            strIDActividad = Nz(.Fields("IDActividad"), "")
            strEquipo = Nz(.Fields("Equipo"), "")
            strMaterial = Nz(.Fields("Material"), "")
            strNodo = Nz(.Fields("NODO"), "")
            strBUI = Nz(.Fields("BUI"), "")
            strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
            strGarantia = Nz(.Fields("Garantia"), "No")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    Else
        m_SQL = "SELECT TbMaterial.Garantia, TbMaterial.Material, TbEquipos.NODO, TbEquipos.BUI, TbEquipos.SUBSISTEMA, " & _
                "TbEquipos.Equipo " & _
                "FROM TbMaterial INNER JOIN TbEquipos ON TbMaterial.IDEquipo = TbEquipos.IDEquipo " & _
                "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "No se ha obtenido ningún dato con el IDMaterial:" & strIDMaterial
                Err.Raise 1000
            End If
            strIDEvento = ""
            strIDActividad = ""
            strEquipo = Nz(.Fields("Equipo"), "")
            strMaterial = Nz(.Fields("Material"), "")
            strNodo = Nz(.Fields("NODO"), "")
            strBUI = Nz(.Fields("BUI"), "")
            strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
            strGarantia = Nz(.Fields("Garantia"), "No")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
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
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
    If strMaterial <> "" Then
        Me.Material = strMaterial
    End If
    If strNodo <> "" Then
        Me.NODO = strNodo
    End If
    If strBUI <> "" Then
        Me.BUI = strBUI
    End If
    If strSUBSISTEMA <> "" Then
        Me.SubSistema = strSUBSISTEMA
    End If
    If strGarantia <> "Sí" And strGarantia <> "No" Then
        Me.GARANTÍA = "No"
    Else
        Me.GARANTÍA = strGarantia
    End If
    Me.lblTitulo1.Caption = "SEGUIMIENTO DEL MATERIAL " & strIDMaterial
    m_SQL = "SELECT TbMaterialSeguimiento.IDSeguimiento, TbMaterialSeguimiento.FechaInicioIntervencion AS [F Inicio], " & _
            "TbMaterialSeguimiento.Destino, TbMaterialSeguimiento.FechaFinIntervencion AS [F Fin], TbMaterialSeguimiento.UltimaIntervencion AS Última " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "')) " & _
            "ORDER BY TbMaterialSeguimiento.FechaInicioIntervencion DESC;"
    Me.ListaIntervenciones.RowSource = m_SQL
    '-----------------------------------------------
    '       IntervencionUltimosDatos =strCodigoFinal
    '       IntervencionUltimosDatos ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = IntervencionUltimosDatos(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If InStr(1, flag, ";") <> 0 Then
        dato = Split(flag, ";")
        strUltimaFechaInicial = dato(1)
        strUltimaFechaFinal = dato(2)
        strUltimoDestino = dato(3)
    End If
    If IsDate(strUltimaFechaInicial) Then
        Me.UltimaFechaInicial = strUltimaFechaInicial
    Else
        Me.UltimaFechaInicial = Null
    End If
    If IsDate(strUltimaFechaFinal) Then
        Me.UltimaFechaFinal = strUltimaFechaFinal
    Else
        Me.UltimaFechaFinal = Null
    End If
    If strUltimoDestino <> "" Then
        Me.UltimoDestino = strUltimoDestino
    Else
        Me.UltimoDestino = Null
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoCerrado = "Sí/"No"
    '       SeguimientoCerrado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoCerrado(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoCerrado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "Sí" And flag <> "No" Then
        Me.FINTERVENCION = "No"
        If strBloqueado = "No" Then
            Me.ComandoAlta.Enabled = True
        Else
            Me.ComandoAlta.Enabled = False
        End If
    Else
        Me.FINTERVENCION = flag
        If flag = "Sí" Then
            
            Me.ComandoAlta.Enabled = False
        Else
            Me.ComandoAlta.Enabled = True
        End If
        
    End If
    
    EstablecerDatos = "ok"
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
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Sub ListaIntervenciones_Click()
    Dim strIDSeguimientoSeleccionado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDSeguimientoSeleccionado = Nz(Me.ListaIntervenciones.Column(0), "")
    If strIDSeguimientoSeleccionado = "" Then
        strURLAnexo = ""
        Me.ComandoEdicion.Enabled = False
        Me.ComandoEliminar.Enabled = False
        Me.ComandoAdjuntos.Enabled = False
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoEdicion.Enabled = True
    If strBloqueado = "No" Then
        Me.ComandoEliminar.Enabled = True
    Else
        Me.ComandoEliminar.Enabled = False
    End If
    Me.ComandoAdjuntos.Enabled = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar hacer click en la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function IntervencionUltimosDatos(strIDMaterial As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/09/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -El codigo que devuelve es:
    '           strCodigoFinal =strNumeroIntervencionesEnSeguimiento;strFechaInicioUltima; _
                strFechaFinUltima;strDestinoUltimo;strSeguimientoFinalizado
    
  
    '   -Llamada desde
    '       -Form_FormMaterialSeguimiento.EstablecerDatos
    '   -Devuelve:
    '        strCodigoFinal =strNumeroIntervencionesEnSeguimiento;strFechaInicioUltima; _
                strFechaFinUltima;strDestinoUltimo;strSeguimientoFinalizado
    '       IntervencionUltimosDatos =strCodigoFinal
    '       IntervencionUltimosDatos ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaInicioUltima As String, strFechaFinUltima As String, _
        strDestinoUltimo As String, strSeguimientoFinalizado As String, strCodigoFinal As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "Intervención no registrada"
        Err.Raise 1000
    End If
    '--------------
    ' Primero strFechaInicioUltima,strFechaFinUltima,strDestinoUltimo
    '--------------
    m_SQL = "SELECT TbMaterialSeguimiento.FechaInicioIntervencion, TbMaterialSeguimiento.Destino, TbMaterialSeguimiento.FechaFinIntervencion " & _
                "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "')) " & _
            "ORDER BY TbMaterialSeguimiento.FechaInicioIntervencion DESC;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            strFechaInicioUltima = Nz(.Fields("FechaInicioIntervencion"), "")
            strDestinoUltimo = Nz(.Fields("Destino"), "")
            strFechaFinUltima = Nz(.Fields("FechaFinIntervencion"), "")
            If Not IsDate(strFechaFinUltima) And CLng(strNumeroIntervencionesEnSeguimiento) > 1 Then
                .MoveNext
                strFechaFinUltima = Nz(.Fields("FechaFinIntervencion"), "")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If IsDate(strFechaFinUltima) Then
        '-------------
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
        strSeguimientoFinalizado = flag
        strCodigoFinal = strNumeroIntervencionesEnSeguimiento & ";" & strFechaInicioUltima & ";" & strFechaFinUltima & ";" & strDestinoUltimo & ";" & strSeguimientoFinalizado
    End If
    IntervencionUltimosDatos = strCodigoFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Material.IntervencionUltimosDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    IntervencionUltimosDatos = "#ERR" & "|" & strTextoError
End Function

Public Function SeguimientoMaterialEliminar( _
                                            strIDIntervencion As String, _
                                            Optional strEliminando As String = "No", _
                                            Optional strIDMaterial As String _
                                            ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/9/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Se obtiene me.IntervencionDatosDefechasIntermedias-->strNumeroIntervencionesEnSeguimiento,strOrdenIntervencionActual
    '               strOrdenIntervencionActual="Primera o Intermedia o Última"
    '           -Si strNumeroIntervencionesEnSeguimiento="1"
    '               Se puede eliminar
    '           -si clng(strNumeroIntervencionesEnSeguimiento)>1
    '               se puede eliminar si strOrdenIntervencionActual=" Última""
    '   -Llamada desde
    '       -Form_FormMaterialSeguimiento.ComandoEliminar_Click
    '   -Devuelve:
    '       SeguimientoMaterialEliminar = Descriptivo
    '       SeguimientoMaterialEliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim m_Anexo As Anexo
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_MaterialSeguimiento As MaterialSeguimiento
    Dim rcdDatos As DAO.Recordset, fso As New FileSystemObject, strNombreArchivo As String, strURLAnexo As String, strIDActividadEnMaterial As String, _
        colAnexosAEliminar As New Collection, VarItem As Variant, strTextoError As String
    On Error GoTo errores
    If strIDIntervencion = "" Then
        strTextoError = "El IDIntervencion es obligatorio"
        Err.Raise 1000
    End If
    If strIDMaterial = "" Then
             '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbMaterialSeguimiento", "IDMaterial", "IDSeguimiento", strIDIntervencion)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDMaterial = flag
    End If
    If strIDMaterial = "" Then
        strTextoError = "El IDIntervencion: " & strIDIntervencion & " no se ha registrado ningún IDMaterial"
        Err.Raise 1000
    End If
    
    '-----------------------------------------------
    '   -Devuelve:
    '       EstaFranqueado = strValorFranqueado
    '       EstaFranqueado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstaFranqueado("MA", strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Evento.EstaFranqueado ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "Proviene de un evento franqueado"
        Err.Raise 1000
    End If
    '---------------------------------------------------------
    '   -Devuelve:
    '       EstaFacturado = strEstaFacturado
    '       EstaFacturado = "#ERR" & "|" & strTextoError
    '---------------------------------------------------------
    flag = EstaFacturado("MA", strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstaFacturado ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "Proviene de un material ya facturado"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '           strCodigoFinal =strNumeroIntervencionesEnSeguimiento;strOrdenIntervencionActual; _
                    strFechaFinIntervencionAnterior;strFechaInicionIntervencionPosterior
    '           strOrdenIntervencionActual="Primera o Intermedia o Última"
    '       IntervencionDatosDefechasIntermedias = strCodigoFinal
    '       IntervencionDatosDefechasIntermedias = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = IntervencionDatosDefechasIntermedias(strIDIntervencion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método IntervencionDatosDefechasIntermedias ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If InStr(1, flag, ";") = 0 Then
        strTextoError = "El método IntervencionDatosDefechasIntermedias ha devuelto un resultado con formato desconocido"
        Err.Raise 1000
    End If
    dato = Split(flag, ";")
    If UBound(dato) <> 3 Then
        strTextoError = "El método IntervencionDatosDefechasIntermedias ha devuelto un resultado con formato desconocido"
        Err.Raise 1000
    End If
    If Not IsNumeric(dato(0)) Then
        strTextoError = "La función IntervencionDatosDefechasIntermedias no ha devuelto un valor de Número de intervenciones con formato conocido"
        Err.Raise 1000
    End If
    If CInt(dato(0)) = 0 Then
        strTextoError = "Este material no tiene intervenciones"
        Err.Raise 1000
    End If
    If CLng(dato(0)) > 1 Then
        If dato(1) <> "Última" Then
            strTextoError = "No se puede eliminar una intervención intermedia sin haber borrado la/as posteriores"
            Err.Raise 1000
        End If
    End If
    Set m_MaterialSeguimiento = Constructor.getMaterialSeguimiento(strIDIntervencion, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    If strEliminando = "Sí" Then
        Set m_Col = m_MaterialSeguimiento.ColAnexos
        strTextoError = m_MaterialSeguimiento.Error
        If strTextoError <> "" Then
            Err.Raise 1000
        End If
        If Not m_Col Is Nothing Then
            For Each m_ID In m_Col
                Set m_Anexo = m_Col(m_ID)
                m_Anexo.Eliminar strTextoError
                If strTextoError <> "" Then
                    Err.Raise 1000
                End If
                Set m_Anexo = Nothing
            Next
        End If
    End If
    
    m_SQL = "SELECT TbMaterialSeguimiento.* " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDSeguimiento)='" & strIDIntervencion & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El IDIntervencion " & strIDIntervencion & " no está registrado en la base de datos"
            Err.Raise 1000
        End If
        If strEliminando = "Sí" Then
            .Delete
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing

    '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoCerrado = "Sí/"No"
    '       SeguimientoCerrado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoCerrado(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoCerrado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "No" And strEliminando = "Sí" Then
        m_SQL = "UPDATE TbMaterial SET TbMaterial.FechaEntrega = Null " & _
                "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL m_SQL
        DoCmd.SetWarnings True
    End If
    If strEliminando = "Sí" Then
        If colAnexosAEliminar.count > 0 Then
            For Each VarItem In colAnexosAEliminar
                strURLAnexo = CStr(VarItem)
                If fso.FileExists(strURLAnexo) Then
                    fso.DeleteFile strURLAnexo, True
                End If
            Next
        End If
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    SeguimientoMaterialEliminar = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método SeguimientoMaterialEliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    SeguimientoMaterialEliminar = "#ERR" & "|" & strTextoError
End Function