Option Compare Database
Option Explicit
Dim strValorAnteriorNodo As String, strValorAnteriorBUI As String
Dim strIDEquipo As String
Dim strNodoAlInicio As String, strBUIAlInicio As String, strSubSistemaAlInicio As String, strEquipoAlInicio As String, _
    strDescripcionAlInicio As String, strFechaObsoletoAlInicio As String
Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  28/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '
    '           strEstado="1"----> True Or False
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, strDESCRIPCION As String, _
        strFechaObsoleto As String
    On Error GoTo errores
    strNodo = Nz(Me.ComboNodo, "")
    strBUI = Nz(Me.ComboBui, "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    strEquipo = Nz(Me.Equipo, "")
    strDESCRIPCION = Nz(Me.Descripcion, "")
    strFechaObsoleto = Nz(Me.FechaObsoleto, "")
    If IsDate(strFechaObsoleto) Then
        strFechaObsoleto = Format(strFechaObsoleto, "ddd/mm/yyyy")
    End If
    If strNodo <> strNodoAlInicio Or strBUI <> strBUIAlInicio Or strSUBSISTEMA <> strSubSistemaAlInicio Or strEquipo <> strEquipoAlInicio Or _
        strDESCRIPCION <> strDescripcionAlInicio Or strFechaObsoleto <> strFechaObsoletoAlInicio Then
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
    If HaHabidoCambios() Then
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
            If FormularioAbierto("FormIDEquiposGestion") Then
                Forms("FormIDEquiposGestion").Controls("ListaFiltrados").Requery
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
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    If Not HaHabidoCambios() Then
        strTextoError = "Ningún cambio que grabar"
        Err.Raise 1000
    End If
     '-----------------------------------------------
    '   -Devuelve:
    '       Registrar = strEstaFacturado
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
    If FormularioAbierto("FormIDEquiposGestion") Then
        Forms("FormIDEquiposGestion").Controls("ListaFiltrados").Requery
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar la Actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboBui_Exit(Cancel As Integer)
    Dim strValorBUI As String, strValorActualSubSistema As String, strSQLSubSistema As String
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValorBUI = Nz(Me.ComboBui, "")
    If strValorBUI <> strValorAnteriorBUI Then
        Me.ComboSubsistema = Null
    End If
    If strValorBUI = "" Then
        strSQLSubSistema = "SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA " & _
                            "FROM TbSubsistemaBui " & _
                            "WHERE (((TbSubsistemaBui.BUI) Like '*')) " & _
                            "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    Else
        
        strSQLSubSistema = "SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA " & _
                            "FROM TbSubsistemaBui " & _
                            "WHERE (((TbSubsistemaBui.BUI) = '" & strValorBUI & "')) " & _
                            "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    End If
    Me.ComboSubsistema.RowSource = strSQLSubSistema
    Me.ComboSubsistema.Requery
    strValorAnteriorBUI = strValorBUI
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComboNodo_Exit(Cancel As Integer)
    Dim strValorNodo As String, strValorActualBUI As String, strSQLBUI As String, strValorActualSubSistema As String
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValorNodo = Nz(Me.ComboNodo, "")
    If strValorNodo <> strValorAnteriorNodo Then
        Me.ComboBui = Null
        Me.ComboSubsistema = Null
    End If
    If strValorNodo = "" Then
        strSQLBUI = "SELECT distinct TbNodoBUI.BUI " & _
                    "FROM TbNodoBUI " & _
                    "WHERE (((TbNodoBUI.NODO) Like '*')) " & _
                    "ORDER BY TbNodoBUI.BUI;"
    Else
       strSQLBUI = "SELECT distinct TbNodoBUI.BUI " & _
                    "FROM  TbNodoBUI " & _
                    "WHERE (((TbNodoBUI.NODO)='" & strValorNodo & "')) " & _
                    "ORDER BY TbNodoBUI.BUI;"
    End If
    Me.ComboBui.RowSource = strSQLBUI
    Me.ComboBui.Requery
    strValorAnteriorNodo = strValorNodo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Function Registrar() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, strDESCRIPCION As String, strFechaObsoleto As String
    Dim strTextoError As String
    On Error GoTo errores
    If strIDEquipo = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    strNodo = Nz(Me.ComboNodo, "")
    strBUI = Nz(Me.ComboBui, "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    strEquipo = Nz(Me.Equipo, "")
    strDESCRIPCION = Nz(Me.Descripcion, "#")
    strFechaObsoleto = Nz(Me.FechaObsoleto, "#")
    '-----------------------------------------------
    '   -Devuelve:
    '       EditarIDEquipo = Descriptivo
    '       EditarIDEquipo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EditarIDEquipo(strIDEquipo, strNodo, strBUI, strSUBSISTEMA, strEquipo, strDESCRIPCION, , strFechaObsoleto)
    dato = Split(flag, "|")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método GenerarConsultas ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub Form_Close()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    strIDEquipo = Nz(Me.OpenArgs, "")
    If strIDEquipo = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = Descriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método GenerarConsultas ha devuelto un error: " & vbNewLine & dato(1)
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
    
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquipos.* " & _
            "FROM TbEquipos " & _
            "WHERE (((TbEquipos.IDEquipo)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strNodoAlInicio = Nz(.Fields("NODO"), "")
        strBUIAlInicio = Nz(.Fields("BUI"), "")
        strSubSistemaAlInicio = Nz(.Fields("SUBSISTEMA"), "")
        strEquipoAlInicio = Nz(.Fields("Equipo"), "")
        strDescripcionAlInicio = Nz(.Fields("DESCRIPCION"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strNodoAlInicio <> "" Then
        Me.ComboNodo = strNodoAlInicio
         m_SQL = "SELECT DISTINCT TbNodoBUI.BUI " & _
                "FROM TbNodoBUI " & _
                "WHERE (((TbNodoBUI.NODO)='" & strNodoAlInicio & "')) " & _
                "ORDER BY TbNodoBUI.BUI;"
        Me.ComboBui.RowSource = m_SQL
        Me.ComboBui.Requery
        strValorAnteriorNodo = strNodoAlInicio
    End If
    If strBUIAlInicio <> "" Then
        m_SQL = "SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA " & _
                    "FROM TbSubsistemaBui " & _
                    "WHERE (((TbSubsistemaBui.BUI)='" & strBUIAlInicio & "')) " & _
                    "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
        Me.ComboSubsistema.RowSource = m_SQL
        Me.ComboSubsistema.Requery
        Me.ComboBui = strBUIAlInicio
        strValorAnteriorBUI = strBUIAlInicio
    End If
    If strSubSistemaAlInicio <> "" Then
        Me.ComboSubsistema = strSubSistemaAlInicio
    End If
    If strEquipoAlInicio <> "" Then
        Me.Equipo = strEquipoAlInicio
    End If
    If strDescripcionAlInicio <> "" Then
        Me.Descripcion = strDescripcionAlInicio
    End If
    Me.Equipo.SetFocus
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = Descriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbEquipos", "FechaObsoleto", "IDEquipo", strIDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaObsoletoAlInicio = flag
    If IsDate(strFechaObsoletoAlInicio) Then
        Me.FechaObsoleto = strFechaObsoletoAlInicio
    End If
    EstablecerDatos = "Datos del Equipo establecidos correctamente"
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
Public Function EditarIDEquipo( _
                                strIDEquipo As String, _
                                strNodo As String, _
                                strBUI As String, _
                                strSUBSISTEMA As String, _
                                strEquipo As String, _
                                Optional strDESCRIPCION As String, _
                                Optional strAlias As String, _
                                Optional strFechaObsoleto As String _
                                ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 25/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a editar los datos que entra y los va a meter TbEquipo--->
    '       -strEquipo existente
    '       -ver si para ese strEquipo ha ya registradas algún material, en ese caso no se puede editar si el evento del que depende está franqueado
    '       -sólo puede ser # strDescripcion
    '   -Llamada desde
    '       -Form_FormEquipoEdicion.Editar
    '   -Devuelve:
    '       EditarIDEquipo = Descriptivo
    '       EditarIDEquipo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strEventoFranqueado As String
    Dim strIDMaterial As String, blnComprobarNODOBUISUBS As Boolean, strTextoError As String
    On Error GoTo errores
    
    If strIDEquipo = "" Or strIDEquipo = "#" Then
        strTextoError = "Es obligatorio indicar el IDEquipo "
        Err.Raise 1000
    End If
    If strNodo = "" Or strNodo = "#" Then
        strTextoError = "Es obligatorio indicar el Nodo"
        Err.Raise 1000
    End If
    If strBUI = "" Or strBUI = "#" Then
        strTextoError = "Es obligatorio indicar el BUI"
        Err.Raise 1000
    End If
    If strSUBSISTEMA = "" Or strSUBSISTEMA = "#" Then
        strTextoError = "Es obligatorio indicar el Subsistema"
        Err.Raise 1000
    End If
    If strEquipo = "" Or strEquipo = "#" Then
        strTextoError = "Es obligatorio indicar el Equipo"
        Err.Raise 1000
    End If
    If strFechaObsoleto <> "" Then
        If strFechaObsoleto = "#" Then
            strFechaObsoleto = ""
        End If
        '-----------------------------------------------
        '       MarcarDesMarcarEquipoComoObsoleto =Descriptivo
        '       MarcarDesMarcarEquipoComoObsoleto ="#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = MarcarDesMarcarEquipoComoObsoleto(strIDEquipo, strFechaObsoleto)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método MarcarDesMarcarEquipoComoObsoleto ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    If strNodo <> strNodoAlInicio Or strBUI <> strBUIAlInicio Or strSUBSISTEMA <> strSubSistemaAlInicio Then
        '-------------
        ' veo si está franqueado el evento del que depende,para ello obtengo el IDActividad
        '-------------
        '   -Devuelve:
        '       EstaFranqueado = strValorFranqueado
        '       EstaFranqueado = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EstaFranqueado("EQ", strIDEquipo)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFranqueado ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEventoFranqueado = flag
        If strEventoFranqueado = "Sí" Then
            strTextoError = "No se puede cambiar el nodo, bui o subsistema de un equipo que pertenece a un evento franqueado"
            Err.Raise 1000
        End If
        blnComprobarNODOBUISUBS = True
    Else
        blnComprobarNODOBUISUBS = False
    End If
    If blnComprobarNODOBUISUBS = True Then
         '-----------
        ' AHORA SE VA VER SI ES COMPATIBLE EL NODO,BUI,SUBSISTEMA
        '--------------
        m_SQL = "SELECT TbNodoBUI.NODO, TbNodoBUI.BUI, TbSubsistemaBui.SUBSISTEMA " & _
                "FROM TbNodoBUI INNER JOIN TbSubsistemaBui ON TbNodoBUI.BUI = TbSubsistemaBui.BUI " & _
                "WHERE (((TbNodoBUI.NODO)='" & strNodo & _
                "') AND ((TbNodoBUI.BUI)='" & strBUI & _
                "') AND ((TbSubsistemaBui.SUBSISTEMA)='" & strSUBSISTEMA & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "El subsistema " & strSUBSISTEMA & " no está registrado para el NODO " & strNodo & " y el BUI " & strBUI
                Err.Raise 1000
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    '----------
    ' edición propiamente dicha
    '------------
    m_SQL = "SELECT TbEquipos.* " & _
            "FROM TbEquipos " & _
            "WHERE (((TbEquipos.IDEquipo)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No está registrado un equipo con ID: " & strEquipo
            Err.Raise 1000
        End If
        .Edit
            If strEventoFranqueado = "No" Then
                .Fields("NODO") = strNodo
                .Fields("BUI") = strBUI
                .Fields("SUBSISTEMA") = strSUBSISTEMA
                .Fields("IDEquipo") = strIDEquipo
            End If
            If strDESCRIPCION <> "" And strDESCRIPCION <> "#" Then
                .Fields("Descripcion") = strDESCRIPCION
            Else
                If strDESCRIPCION = "#" Then
                    .Fields("Descripcion") = Null
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strAlias <> "" Then
        m_SQL = "SELECT TbMantenimientoPreventivoEquipos.AliasParaExcel " & _
                "FROM TbMantenimientoPreventivoEquipos " & _
                "WHERE (((TbMantenimientoPreventivoEquipos.IDEquipoProgramable)=" & strEquipo & "));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                .Edit
                    If strAlias <> "#" Then
                        .Fields("AliasParaExcel") = strAlias
                    Else
                        .Fields("AliasParaExcel") = Null
                    End If
                .Update
            End If
         End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EditarIDEquipo = "Edición correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Material.EditarIDEquipo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EditarIDEquipo = "#ERR" & "|" & strTextoError
End Function
Public Function MarcarDesMarcarEquipoComoObsoleto( _
                                                    strIDEquipo As String, _
                                                    Optional strFechaObsoleto As String _
                                                    ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 24/03/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a rellenar la tabla TbEquipos con la fecha que viene para el IDEquipo que viene si viene fecha
    '       -si no viene fecha se quita la obsolescencia
    '   -Llamada desde
    
    '   -Devuelve:
    '       MarcarDesMarcarEquipoComoObsoleto =Descriptivo
    '       MarcarDesMarcarEquipoComoObsoleto ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "Se ha de indicar un equipo id con formato válido"
        Err.Raise 1000
    End If
    If strFechaObsoleto <> "" Then
        If Not IsDate(strFechaObsoleto) Then
            strTextoError = "Para marcar un equipo como obsoleto se ha de indicar una fecha válida"
            Err.Raise 1000
        End If
    End If
    If IsDate(strFechaObsoleto) Then
        m_SQL = "SELECT TbEquipos.* " & _
                    "FROM TbEquipos " & _
                    "WHERE (((TbEquipos.IDEquipo)=" & strIDEquipo & "));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                .Edit
                    .Fields("FechaObsoleto") = strFechaObsoleto
                .Update
            
            End If
        End With
    Else
        m_SQL = "SELECT TbEquipos.* " & _
                    "FROM TbEquipos " & _
                    "WHERE (((TbEquipos.IDEquipo)=" & strIDEquipo & "));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                .Edit
                    .Fields("FechaObsoleto") = Null
                .Update
            
            End If
        End With
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    MarcarDesMarcarEquipoComoObsoleto = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método  MarcarDesMarcarEquipoComoObsoleto ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    MarcarDesMarcarEquipoComoObsoleto = "#ERR" & "|" & strTextoError
End Function





