Option Compare Database
Option Explicit
Private m_Error As String
Private fso As New FileSystemObject
Dim strIDActividadSeleccionada As String, strURLAnexo As String
Dim strTecnicoAnterior As String, strFechaInicialAnterior As String, strFechaFinalAnterior As String, strIDActividadAnterior As String, strFormularioQueAbre As String
Public Function CopiarActividadNoDeEvento( _
                                            strIDActividad As String, _
                                            Optional strFecha As String, _
                                            Optional p_Horas As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim rcdDatos As DAO.Recordset, rcdDatosDestino As DAO.Recordset, fld As DAO.Field, strIDEvento As String, _
        strIDActividadNueva As String, strAliasTecnico As String
    On Error GoTo errores
    If strIDActividad = "" Then
        p_Error = "Se ha de indicar el IDActividad"
        Err.Raise 1000
    End If
    If Not IsDate(strFecha) Then
        strFecha = Now()
    End If
    If Not IsNumeric(p_Horas) Then
        p_Horas = "7,5"
    End If
    m_SQL = "TbActividades"
    Set rcdDatosDestino = CurrentDb().OpenRecordset(m_SQL)
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDActividad)='" & strIDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "Actividad No registrada"
            Err.Raise 1000
        End If
        strAliasTecnico = Nz(.Fields("AliasTecnico"), "")
        If strAliasTecnico = "" Then
            p_Error = "La actividad original, no tiene registrado el alias del técnico"
            Err.Raise 1000
        End If
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       DameIDActividad = strIDActividad
        '       DameIDActividad = "ERR" & "|" & p_Error
        '-------------------------------------------------------------------
        flag = DameIDActividad(strAliasTecnico, strFecha)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            p_Error = "El método DameIDActividad ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDActividadNueva = flag
        rcdDatosDestino.AddNew
            For Each fld In rcdDatos.Fields
                If fld.Name <> "IDActividad" And fld.Name <> "CodImportacion" And fld.Name <> "IDFacturacion" And fld.Name <> "IDExportacion" And fld.Name <> "IDEVENTOGENERADO" And fld.Name <> "FechaAlta" And fld.Name <> "HorasLaborables" Then
                    rcdDatosDestino(fld.Name) = rcdDatos(fld.Name)
                End If
                
            Next
            rcdDatosDestino("IDActividad") = strIDActividadNueva
            rcdDatosDestino("FechaAlta") = Format(strFecha, "dd/mm/yyyy")
            rcdDatosDestino("HorasLaborables") = p_Horas
        rcdDatosDestino.Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not rcdDatosDestino Is Nothing Then
        rcdDatosDestino.Close
        Set rcdDatosDestino = Nothing
    End If
    CopiarActividadNoDeEvento = strIDActividadNueva
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarActividadNoDeEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Not rcdDatosDestino Is Nothing Then
        rcdDatosDestino.Close
        Set rcdDatosDestino = Nothing
    End If
    
End Function
Public Function HorasRegistrables( _
                                    m_Alias As String, _
                                    m_Fecha As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    Dim Dia As Long
    Dim m_HorasRegistradas As Double
    Dim m_HorasRestantes As Double
   
    
    On Error GoTo errores
    If m_Alias = "" Then
        p_Error = "Falta el alias"
        Err.Raise 1000
    End If
    If Not IsDate(m_Fecha) Then
        p_Error = "Falta la fecha"
        Err.Raise 1000
    End If
    Dia = VBA.Weekday(m_Fecha, vbMonday)
    If Dia = 6 Or Dia = 7 Then
        HorasRegistrables = "0"
        Exit Function
    End If
    m_SQL = "SELECT TbTecnicosAusencias.FechaLibranza " & _
                "FROM TbTecnicosAusencias " & _
                "WHERE (((TbTecnicosAusencias.Alias)='" & m_Alias & "') " & _
                "AND ((TbTecnicosAusencias.FechaLibranza)=#" & Format(m_Fecha, "mm/dd/yyyy") & "#));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            HorasRegistrables = "0"
            Exit Function
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    m_SQL = "SELECT TbTecnicosFiestas.FechaFiesta " & _
                "FROM TbTecnicosFiestas " & _
                "WHERE (((TbTecnicosFiestas.FechaFiesta)=#" & Format(m_Fecha, "mm/dd/yyyy") & "#));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            HorasRegistrables = "0"
            Exit Function
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    m_SQL = "SELECT TbActividades.HorasLaborables " & _
                "FROM TbActividades " & _
                "WHERE (((TbActividades.ALIASTECNICO)='" & m_Alias & _
                "') AND ((TbActividades.FechaAlta)=#" & Format(m_Fecha, "mm/dd/yyyy") & "#));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("HorasLaborables"), "")) Then
                m_HorasRegistradas = CDbl(.Fields("HorasLaborables"))
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_HorasRestantes = CDbl(7.5) - m_HorasRegistradas
    HorasRegistrables = CStr(m_HorasRestantes)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HorasRegistrables ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function





Public Sub LimpiarUltimos()
    strTecnicoAnterior = ""
    strFechaInicialAnterior = ""
    strFechaFinalAnterior = ""
    strIDActividadAnterior = ""
End Sub

Private Sub ComandoAdjuntosActividad_Click()
    Dim m_Actividad As Actividad
    Dim strEscritura As String, strFacturado As String
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDActividadSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDActividadSeleccionada = "" Then
        m_Error = "Se ha de Seleccionar una actividad primero"
        Err.Raise 1000
    End If
    strFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    
    
    If strFacturado <> "" Then
        strFacturado = "Sí"
    Else
        strFacturado = "No"
    End If
    If strFacturado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    Set m_Actividad = Constructor.getActividad(strIDActividadSeleccionada, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '
    '           strEstado="1"---->true or false
    '           strEstado="-1"--->false
    '-------------------------------------------------------------------
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    strFormularioQueAbre = ""
    Set m_ObjEntidadParaAnexoActiva = m_Actividad
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoAltaActividad_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    '---------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormActividadAlta") Then
        DoCmd.Close acForm, "FormActividadAlta", acSaveNo
    End If
    'strArgumento=strIDEvento & "|" & strIDActividadOriginal & "|" & strEventoFranqueado
    'strArgumento=""
    DoCmd.OpenForm "FormActividadAlta"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar dar de alta una actividad se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoCopiarActividad_Click()
    Dim strIDEvento As String, strIDActividadACopiar As String, strIDActividadNueva As String, strEventoFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadACopiar = Nz(Me.ListaFiltrados.Column(0), "")
    strIDEvento = Nz(Me.ListaFiltrados.Column(3), "")
    If strIDEvento <> "" Then
        m_Error = "Está pensado para copiar una actividad NO DE EVENTO"
        Err.Raise 1000
    End If
    strIDActividadNueva = CopiarActividadNoDeEvento(strIDActividadACopiar, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    pregunta = MsgBox("Generada la nueva Actividad a partir de la seleccionada", vbInformation, "Copia de Actividad")
    ActualizarLista
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCopiarActividad_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoCopiarActividadMultiple_Click()
    Dim strIDEvento As String, strIDActividadACopiar As String
    Dim m_FechaInicial As String
    Dim m_FechaFinal As String
    Dim m_HorasRegistrables As String
    Dim m_Alias As String
    Dim m_Horas As String
    Dim Dias As Long
    Dim i As Long
    Dim m_Fecha As String
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDActividadACopiar = Nz(Me.ListaFiltrados.Column(0), "")
    strIDEvento = Nz(Me.ListaFiltrados.Column(3), "")
    m_Alias = Nz(Me.ListaFiltrados.Column(7), "")
    
    If strIDEvento <> "" Then
        m_Error = "Está pensado para copiar una actividad NO DE EVENTO"
        Err.Raise 1000
    End If
    m_FechaInicial = Nz(InputBox("Introduzca fecha inicial", "Fecha1"), "")
    If Not IsDate(m_FechaInicial) Then
        m_Error = "Fecha no válida"
        Err.Raise 1000
    End If
    m_FechaFinal = Nz(InputBox("Introduzca fecha Final", "Fecha2"), "")
    If Not IsDate(m_FechaFinal) Then
        m_Error = "Fecha no válida"
        Err.Raise 1000
    End If
    If CDate(m_FechaFinal) < CDate(m_FechaInicial) Then
        m_Error = "Periodo de fechas incorrecto"
        Err.Raise 1000
    End If
    If CDate(m_FechaFinal) = CDate(m_FechaInicial) Then
        m_HorasRegistrables = HorasRegistrables(m_Alias, m_FechaInicial, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
        m_FechaInicial = m_FechaInicial & " " & Format(Hour(Now()), "00") & ":" & Format(Minute(Now()), "00") & ":" & Format(Second(Now()), "00")
        flag = CopiarActividadNoDeEvento(strIDActividadACopiar, m_FechaInicial)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método CopiarActividadNoDeEvento ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        
        ActualizarLista
        DoCmd.Hourglass False
        Exit Sub
    End If
    Dias = DateDiff("d", CDate(m_FechaFinal), CDate(m_FechaInicial))
    If Dias < 1 Then
        Dias = Dias * -1
    End If
    
    m_HorasRegistrables = HorasRegistrables(m_Alias, m_FechaInicial, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HorasRegistrables <> "0" Then
        m_FechaInicial = m_FechaInicial & " " & Format(Hour(Now()), "00") & ":" & Format(Minute(Now()), "00") & ":" & Format(Second(Now()), "00")
        flag = CopiarActividadNoDeEvento(strIDActividadACopiar, m_FechaInicial, m_HorasRegistrables)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método CopiarActividadNoDeEvento ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    For i = 1 To Dias
       
        m_HorasRegistrables = "0"
        m_Fecha = DateAdd("d", CDbl(i), CDate(Format(m_FechaInicial, "dd/mm/yyyy"))) & " " & Format(Hour(Now()), "00") & ":" & Format(Minute(Now()), "00") & ":" & Format(Second(Now()), "00")
        
        m_HorasRegistrables = HorasRegistrables(m_Alias, Format(m_Fecha, "dd/mm/yyyy"), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_HorasRegistrables <> "0" Then
            flag = CopiarActividadNoDeEvento(strIDActividadACopiar, m_Fecha, m_HorasRegistrables)
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                m_Error = "El método CopiarActividadNoDeEvento ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
        End If
    Next
   
    ActualizarLista
    DoCmd.Hourglass False
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCopiarActividadMultiple_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEditarActividad_Click()
    Dim strIDActividadSeleccionada As String, strEstaFacturado As String, _
        strFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDActividadSeleccionada = "" Then
        m_Error = "Seleccione la actividad que pretende editar"
        Err.Raise 1000
    End If
    strEstaFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    If strEstaFacturado = "" Then
        strEstaFacturado = "No"
    Else
        strEstaFacturado = "Sí"
    End If
    
    strFranqueado = Nz(Me.ListaFiltrados.Column(6), "")
    '---------------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------------
    If FormularioAbierto("FormActividadEdicion") Then
        DoCmd.Close acForm, "FormActividadEdicion", acSaveNo
    End If
    'strArgumento=strIDActividad & & "|" & strFranqueado & "|" & strEstaFacturado & "|" & strURLAnexoActividad
    DoCmd.OpenForm "FormActividadEdicion", , , , , , strIDActividadSeleccionada & "|" & strFranqueado & "|" & _
                    strEstaFacturado & "|" & strURLAnexo
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar editar la actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoEliminarActividad_Click()
    Dim strActividadDeEventoFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDActividadSeleccionada = "" Then
        m_Error = "Seleccione la actividad que pretende eliminar"
        Err.Raise 1000
    End If
    strActividadDeEventoFranqueado = Nz(Me.ListaFiltrados.Column(6), "")
    If strActividadDeEventoFranqueado = "Sí" Then
        m_Error = "Pretende eliminar una actividad que depende de un evento que ya está franqueado. Ha de quitar el franqueo antes de eliminar la actividad"
        Err.Raise 1000
    End If
    DoCmd.Hourglass False
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '
    '           strEstado="1"---->true or false
    '           strEstado="-1"--->false
    '-------------------------------------------------------------------
    pregunta = MsgBox("¿Desea eliminar la actividad con código: " & strIDActividadSeleccionada & _
                " junto con los materiales que dependan de ella y todos los anexos implicados?", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar actividad")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       ActividadEliminar = cstr(intNumeroMaterialesEliminados) & ";" & strCadenaMaterialesNoEliminados
    '       ActividadEliminar = "#ERR" & "|" & m_Error
    '-----------------------------------------------
    flag = ActividadEliminar(strIDActividadSeleccionada, "Sí", "Sí", strActividadDeEventoFranqueado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método ActividadEliminar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.ListaFiltrados.Requery
    strIDActividadSeleccionada = ""
    ActualizarLista
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar eliminar la actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoExportarExcelActividades_Click()
    Dim strSQLInicial As String, strParteWhere As String, intPosicionWhere As Integer, _
        strSQLLista As String
    On Error GoTo errores
    
    m_Error = ""
    If Me.ListaFiltrados.ListCount - 1 = 0 Then
        Exit Sub
    End If
    strSQLLista = Me.ListaFiltrados.RowSource
    intPosicionWhere = InStr(1, strSQLLista, "Where")
    If intPosicionWhere = 0 Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strSQLInicial = "SELECT TbActividades.IDActividad, TbActividades.DESCRIPCION AS DESCRIPCIÓN, " & _
                        "TbActividades.HorasLaborables AS HLab, TbActividades.IDEvento, TbActividades.FechaAlta AS Alta, " & _
                        "TbFacturaActividadesInvolucradas.IDFactura, TbEventos.Franqueado, " & _
                        "TbActividades.ALIASTECNICO " & _
                        "FROM (TbActividades LEFT JOIN TbFacturaActividadesInvolucradas ON " & _
                        "TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad) " & _
                        "LEFT JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento "
    dato = Split(strSQLLista, "Where")
    strParteWhere = "Where" & dato(1)
    m_SQL = strSQLInicial & strParteWhere
    
    GenerarConsultasSQL m_SQL, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoInforme_Click()
    Dim m_ObjTecnicoSeleccionado As Tecnico
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    m_Error = ""
    If m_ObjTecnicoSeleccionado Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(7), "") <> "" Then
            Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaFiltrados.Column(7), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        
    End If
    If FormularioAbierto("FormTecnicoConsultaHoras") Then
        DoCmd.Close acForm, "FormTecnicoConsultaHoras", acSaveNo
    End If
    Set m_ObjTecnicoActivo = m_ObjTecnicoSeleccionado
    DoCmd.OpenForm "FormTecnicoConsultaHoras"


   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoInforme_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoLimpiarFechas_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.FechaInicialBusqueda, "") = "" And Nz(Me.FechaFinalBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.FechaInicialBusqueda = Null
    Me.FechaFinalBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarIDActividad_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboIDActividadBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboIDActividadBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarTecnico_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboTécnicoBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboTécnicoBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Public Sub ComboIDActividadBusqueda_Change()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboIDActividadBusqueda, "")
    If strValor <> strIDActividadAnterior Then
        ActualizarLista
    End If
    strIDActividadAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboIDActividadBusqueda_Exit(Cancel As Integer)
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboIDActividadBusqueda, "")
    If strValor <> strIDActividadAnterior Then
        ActualizarLista
    End If
    strIDActividadAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboTécnicoBusqueda_Change()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboTécnicoBusqueda, "")
    If strValor <> strTecnicoAnterior Then
        ActualizarLista
    End If
    strTecnicoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboTécnicoBusqueda_Exit(Cancel As Integer)
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    strValor = Nz(Me.ComboTécnicoBusqueda, "")
    If strValor <> strTecnicoAnterior Then
        ActualizarLista
    End If
    strTecnicoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub FechaFinalBusqueda_Exit(Cancel As Integer)
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.FechaFinalBusqueda, "")
    If strValor <> strFechaFinalAnterior Then
        ActualizarLista
    End If
    strFechaFinalAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub FechaInicialBusqueda_Exit(Cancel As Integer)
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.FechaInicialBusqueda, "")
    If strValor <> strFechaInicialAnterior Then
        ActualizarLista
    End If
    strFechaInicialAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub Form_Open(Canel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
'    Me.InsideHeight = 6945
'    Me.InsideWidth = 14730
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    RellenarComboTecnicos Me.ComboTécnicoBusqueda, EnumSino.Sí, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ActualizarLista
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ImagenAnexoActividad_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strURLAnexo = "" Then
        m_Error = "En estos momentos el anexo seleccionado no es alcanzable por la herramienta. Revise que ubicación de la documentación remota esté bien establecida"
        Err.Raise 1000
    End If
    If InStr(1, strURLAnexo, ";") = 0 Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AbrirEnLocal = strURLAnexoLocal
        '       AbrirEnLocal = "#ERR" & "|" & m_Error
        '-------------------------------------------------------------------
        flag = AbrirEnLocal(strURLAnexo, Me.hWnd)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        ComandoAdjuntosActividad_Click
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el archivo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ListaFiltrados_Click()
    Dim m_Actividad As Actividad
    Dim strIDEvento As String, strFacturado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    strIDEvento = Nz(Me.ListaFiltrados.Column(3), "")
    If strIDActividadSeleccionada = "" Then
        Me.ComandoEditarActividad.Enabled = False
        Me.ComandoEliminarActividad.Enabled = False
        Me.ComandoCopiarActividad.Enabled = False
        Me.ComandoCopiarActividadMultiple.Enabled = False
        Me.ImagenAnexoActividad.Visible = False
        Me.ComandoAdjuntosActividad.Enabled = False
        Me.lblFacturado.Visible = False
       
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoEditarActividad.Enabled = True
    Me.ComandoAdjuntosActividad.Enabled = True
    strFacturado = Nz(Me.ListaFiltrados.Column(5), "")
    If strFacturado = "" Then
        strFacturado = "No"
        Me.ComandoAdjuntosActividad.Enabled = True
        Me.ComandoEliminarActividad.Enabled = True
        Me.lblFacturado.Visible = False
    Else
        strFacturado = "Sí"
        
        Me.ComandoEliminarActividad.Enabled = False
        Me.lblFacturado.Visible = True
    End If
    If strIDEvento <> "" Then
        Me.ComandoCopiarActividad.Enabled = False
        Me.ComandoCopiarActividadMultiple.Enabled = False
    Else
        Me.ComandoCopiarActividad.Enabled = True
        Me.ComandoCopiarActividadMultiple.Enabled = True
    End If
    Set m_Actividad = Constructor.getActividad(strIDActividadSeleccionada, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    strURLAnexo = m_Actividad.URLAnexos
    
    If strURLAnexo = "" Then
        Me.ImagenAnexoActividad.Visible = False
    Else
        Me.ImagenAnexoActividad.Visible = True
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar hacer click sobre la lista de actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Function Filtrar( _
                            Optional strIDActividad As String, _
                            Optional strAliasTecnico As String, _
                            Optional strFechaInicial As String, _
                            Optional strFechaFinal As String, _
                            Optional ByRef p_Error As String _
                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    '       -me.ActualizarLista
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &p_Error
    '-------------------------------------------------------------------
    Dim strParteWhere As String, strSQLParteInicial As String, strFechaInicialParaConsulta As String, strFechaFinalParaConsulta As String, strSQLFinal As String
    Dim strParteWhereIDActividad As String, strParteWhereAliasTecnico As String, strParteWhereFechas As String
    
    On Error GoTo errores
    strSQLParteInicial = "SELECT TbActividades.IDActividad, TbActividades.DESCRIPCION AS DESCRIPCIÓN, " & _
                        "TbActividades.HorasLaborables AS HLab, TbActividades.IDEvento, TbActividades.FechaAlta AS Alta, " & _
                        "TbFacturaActividadesInvolucradas.IDFactura, TbEventos.Franqueado, " & _
                        "TbActividades.ALIASTECNICO " & _
                        "FROM (TbActividades LEFT JOIN TbFacturaActividadesInvolucradas ON " & _
                        "TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad) " & _
                        "LEFT JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento "
    If strIDActividad = "" And strAliasTecnico = "" And strFechaInicial = "" And strFechaFinal = "" Then
        strParteWhere = "WHERE (((TbActividades.IDActividad)='-1'));"
        strSQLFinal = strSQLParteInicial & strParteWhere
        Filtrar = strSQLFinal
        Exit Function
    End If
    If Not IsDate(strFechaInicial) Then
        strFechaInicialParaConsulta = "01/01/1900"
    Else
        strFechaInicialParaConsulta = Format(strFechaInicial, "mm/dd/yyyy")
    End If
    If Not IsDate(strFechaFinal) Then
        strFechaFinalParaConsulta = "01/01/2100"
    Else
        strFechaFinalParaConsulta = Format(strFechaFinal, "mm/dd/yyyy")
    End If
    
    If strIDActividad = "" Then
        strParteWhereIDActividad = "((TbActividades.IDActividad) Like '*' Or (TbActividades.IDActividad) Is Null)"
    Else
        strParteWhereIDActividad = "((TbActividades.IDActividad) = '" & strIDActividad & "')"
    End If
    If strAliasTecnico = "" Then
        strParteWhereAliasTecnico = "((TbActividades.ALIASTECNICO) Like '*' Or (TbActividades.ALIASTECNICO) Is Null)"
    Else
        strParteWhereAliasTecnico = "((TbActividades.ALIASTECNICO) = '" & strAliasTecnico & "')"
    End If
    strParteWhereFechas = "((TbActividades.FechaAlta) Between #" & strFechaInicialParaConsulta & "# And #" & strFechaFinalParaConsulta & "#)"
    strParteWhere = "WHERE (" & _
                            strParteWhereIDActividad & _
                            " AND " & strParteWhereAliasTecnico & _
                            " AND " & strParteWhereFechas & _
                            ") ORDER BY  TbActividades.FechaAlta DESC;"
    strSQLFinal = strSQLParteInicial & strParteWhere
    Filtrar = strSQLFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Sub ActualizarLista()
    Dim intElementos As Integer, lst As ListBox, VarItem As Variant
    Dim strIDActividadSeleccionadaActualmente As String, blnEncontrada As Boolean, i As Integer
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = ""
    Set lst = Me.ListaFiltrados
    strIDActividadSeleccionadaActualmente = Nz(lst.Column(0), "")
    Me.ComandoEditarActividad.Enabled = False
    Me.ComandoEliminarActividad.Enabled = False
    Me.ComandoCopiarActividad.Enabled = False
    Me.ComandoCopiarActividadMultiple.Enabled = False
   
    Me.ComandoAdjuntosActividad.Enabled = False
    Me.ImagenAnexoActividad.Visible = False
    Me.ComandoExportarExcelActividades.Enabled = False
    
    Me.lblFacturado.Visible = False
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &m_Error
    '-------------------------------------------------------------------
    m_SQL = Filtrar( _
                    Nz(Me.ComboIDActividadBusqueda, ""), Nz(Me.ComboTécnicoBusqueda, ""), _
                    Nz(Me.FechaInicialBusqueda, ""), Nz(Me.FechaFinalBusqueda, ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 0 Then
        Me.ComandoExportarExcelActividades.Enabled = False
    Else
        If intElementos = 1 Then
            strIDActividadSeleccionadaActualmente = lst.Column(1, 1)
            lst.Selected(1) = True
            ListaFiltrados_Click
        Else
            For i = 1 To lst.ListCount - 1
                If lst.Selected(i) = True Then
                    strIDActividadSeleccionadaActualmente = lst.Column(1, i)
                    ListaFiltrados_Click
                    Exit For
                End If
            Next
        End If
        Me.ComandoExportarExcelActividades.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ActualizarLista la lista de Actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub






