Option Compare Database
Option Explicit
Dim m_ValorAnteriorBUI As String, m_ValorAnteriorSubSistema As String, m_ValorAnteriorIDEquipo As String
Private m_Equipo As EquipoMedida

Dim m_Error As String



Private Sub BUI_Change()
    Dim m_ValorActualBUI As String
    Dim m_SQLSubSistema As String
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_ValorActualBUI = Nz(Me.BUI, "")
    Me.SubSistema = Null
    Me.IDEquipo = Null
    Me.Material = Null
    
'    If m_ValorActualBUI = m_ValorAnteriorBUI Then
'        VBA.DoEvents
'        DoCmd.Hourglass False
'        VBA.DoEvents
'        Exit Sub
'    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_SQLSubSistema = EstablecerComboSubSistema(m_ValorActualBUI)
    Me.SubSistema.RowSource = m_SQLSubSistema
    Me.SubSistema.Requery
    
    m_ValorAnteriorBUI = m_ValorActualBUI
   
    RellenarComboContacto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al BUI_Change se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAlta_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    
    pregunta = MsgBox("Se van a dar de alta tantos eventos,actividades y materiales como Números de serie haya metido" & vbNewLine & _
                    "¿Desea proseguir con la operación?", vbExclamation + vbYesNo + vbDefaultButton2, "Alta por Lotes")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AltaLote m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
   
    If Not FormularioAbierto("FormEventoGestion") Then
        DoCmd.OpenForm "FormEventoGestion"
    End If
    Form_FormEventoGestion.ComandoActualizarTodo_Click
    
    Forms("FormEventoGestion").Controls("FechaInicialBusqueda") = Me.FECHAALTAEVENTO
    Forms("FormEventoGestion").Controls("FechaFinalBusqueda") = Me.FECHAALTAEVENTO
    Forms("FormEventoGestion").Controls("ComboBUIBusqueda") = Me.BUI
    Form_FormEventoGestion.ComboBUIBusqueda_AfterUpdate
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Se han dado de alta tantos eventos como números de serie haya incluido ", vbInformation, "Alta de Eventos por Lotes")
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub








Private Sub IDEquipo_Change()
    
    Dim m_ValorActualIDEquipo As String
    Dim m_SQLMaterial As String
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    m_ValorActualIDEquipo = Nz(Me.IDEquipo.Column(0), "")
    
    If m_ValorActualIDEquipo = m_ValorAnteriorIDEquipo Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    Me.Material = Null
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_SQLMaterial = EstablecerComboMaterial(m_ValorActualIDEquipo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.Material.RowSource = m_SQLMaterial
    Me.Material.Requery
    m_ValorAnteriorIDEquipo = m_ValorActualIDEquipo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al IDEquipo_Change se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub IDEquipo_Enter()
    On Error Resume Next
    Me.IDEquipo.Requery
End Sub





Private Sub Material_Exit(Cancel As Integer)
    m_Error = ""
    If Nz(Me.Descripcion, "") <> "" Then
        Exit Sub
    End If
    If Nz(Me.Material, "") = "" Or Nz(Me.IDEquipo, "") = "" Then
        Exit Sub
    End If
    Me.Descripcion = Me.Material & " " & Me.IDEquipo.Column(1)
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Material_Exit se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Originador_Enter()
    On Error Resume Next
    Me.Originador.Requery
End Sub

Private Sub SubSistema_Change()
    Dim m_ValorActualSubSistema As String
    Dim m_ValorBUI As String
    Dim m_SQLIDEquipo As String
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    m_ValorBUI = Nz(Me.BUI, "")
    m_ValorActualSubSistema = Nz(Me.SubSistema, "")
    
    If m_ValorBUI = "" Or m_ValorActualSubSistema = "" Then
        Me.IDEquipo = Null
        Me.Material = Null
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ValorActualSubSistema = m_ValorAnteriorSubSistema Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.IDEquipo = Null
    Me.Material = Null
        VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_SQLIDEquipo = EstablecerComboEquipo(m_ValorBUI, m_ValorActualSubSistema)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método EstablecerComboEquipo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.IDEquipo.RowSource = m_SQLIDEquipo
    Me.IDEquipo.Requery
    m_ValorAnteriorSubSistema = m_ValorActualSubSistema
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al SubSistema_Change se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Subsistema_Enter()
    On Error Resume Next
    Me.SubSistema.Requery
End Sub
Private Sub TipoEvento_Enter()
    On Error Resume Next
    Me.TipoEvento.Requery
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Contacto_Enter()
    Me.Requery
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
        
    m_Error = ""
    AjustarTamaño Me
    RellenarComboTecnicos Me.ALIASTECNICO, EnumSino.No, EnumSino.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.Criticidad = "5"
    Me.ALIASTECNICO = "FSP"
    Me.Originador = "ENVIO A REPARAR"
    Me.TipoEvento = "INCIDENCIA/AVERÍA"
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Public Function AltaLote(Optional ByRef p_Error As String) As String
    
    
    Dim m_IDEquipo As String
    Dim m_CadenaNS As String
    Dim m_DescripcionUsuario As String
    Dim m_IDEvento As String
    Dim m_ColNS As Scripting.Dictionary
    Dim ctl As Control
    Dim rcdDatos As DAO.Recordset
    Dim m_IDEventoEnBD As String
    Dim m_IDEventoResultante As String
    Dim m_Evento As Evento
    Dim m_Actividad As Actividad
    Dim m_Material As Material
    Dim dato
    Dim NS As Variant
    On Error GoTo errores
    
    For Each ctl In Me.Controls
        If InStr(1, ctl.Tag, "OBLIGATORIO") <> 0 Then
            If Nz(ctl.Value, "") = "" Then
                p_Error = "El campo " & ctl.Name & " no está relleno"
                Err.Raise 1000
            End If
        End If
        
    Next
    m_CadenaNS = Me.CadenaNS
    m_DescripcionUsuario = Me.Descripcion
    
   
    If InStr(1, m_DescripcionUsuario, "S/N") <> 0 Then
        p_Error = "No ponga nada de Número de Serie en el título del evento. " & _
            "El Sistema generará una descripción para cada evento de cada uno de los N/S y se mostrára en ella"
        Err.Raise 1000
    End If
    Set m_Evento = New Evento
   
    With m_Evento
        For Each ctl In Me.Controls
            If InStr(1, ctl.Tag, "DATOS") <> 0 Then
                .SetPropiedad ctl.Name, Nz(ctl.Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        Next
    End With
    'relleno de la colección de N/S
    dato = Split(m_CadenaNS, ";")
    For Each NS In dato
        If m_ColNS Is Nothing Then
            Set m_ColNS = New Scripting.Dictionary
            m_ColNS.CompareMode = TextCompare
        End If
        If Not m_ColNS.Exists(NS) Then
            m_ColNS.Add NS, NS
        End If
    Next
    For Each NS In m_ColNS
        'BATERIA HARRYS 7800 GUIA 12/23 S/N xxxxxx
        m_Evento.Descripcion = m_DescripcionUsuario & " " & "S/N" & " " & NS
        m_Evento.BUI = Me.BUI
        m_Evento.NODO = Dame("TbNodoBUI", "NODO", "BUI", Me.BUI)
        m_IDEvento = AltaEvento(m_Evento, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Evento.IDEVENTO = m_IDEvento
       
        Set m_Actividad = New Actividad
        With m_Actividad
            .IDEVENTO = m_IDEvento
            .HorasLaborables = "0,2"
            .Actividad = "LG003"
            .Ubicacion = "EMPLAZAMIENTO HABITUAL"
            .ALIASTECNICO = m_Evento.ALIASTECNICO
            .Descripcion = "Entrada"
            AltaActividad m_Evento, m_Actividad, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End With
        Set m_Material = New Material
        With m_Material
            .IDActividad = m_Actividad.IDActividad
            .Material = Me.Material
            .FechaIncio = m_Evento.FECHAALTAEVENTO
            .NS = CStr(NS)
            .Descripcion = "Reparación"
            .EsReparacion = "Sí"
            AltaMaterial m_Material, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End With
        
        Set m_Material = Nothing
        Set m_Actividad = Nothing
        
    Next
    AltaLote = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaLote ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function AltaMaterial( _
                                p_Material As Material, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_IDMaterial As String
    
    On Error GoTo errores
    m_SQL = "SELECT * " & _
            "FROM TbMaterial " & _
            "WHERE NS='" & p_Material.NS & "' AND IDActividad='" & p_Material.IDActividad & "';"
    
   
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Material.IDMaterial = DameValorSiguienteAlMaximoDeUnCampo("TbMaterial", "IDMaterial", p_Material.IDActividad & "_", 2)
            .AddNew
                .Fields("IDMaterial") = p_Material.IDMaterial
                .Fields("IDActividad") = p_Material.IDActividad
                .Fields("Material") = p_Material.Material
                .Fields("FechaIncio") = p_Material.FechaIncio
                .Fields("NS") = p_Material.NS
                
                .Fields("Descripcion") = p_Material.Descripcion
                .Fields("EsReparacion") = p_Material.EsReparacion
                
            .Update
        End If
    End With
    rcdDatos.Close
    
    
    Set rcdDatos = Nothing
    AltaMaterial = p_Material.IDMaterial
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaMaterial ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function



Public Function AltaActividad( _
                                p_Evento As Evento, _
                                p_Actividad As Actividad, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    Dim rcdDatos As DAO.Recordset
    Dim m_FechaAhora As String
    Dim m_Año As String
    Dim m_Mes As String
    Dim m_Dia As String
    Dim m_HoraMinuto As String
    Dim dato
    Dim m_Hora As String
    Dim m_Minuto As String
    Dim m_ParteFija As String
    
    On Error GoTo errores
    
    
    m_FechaAhora = Now()
    m_Año = Format(CDate(p_Evento.FECHAALTAEVENTO), "yy")
    m_Mes = Format(CDate(p_Evento.FECHAALTAEVENTO), "mm")
    m_Dia = Format(CDate(p_Evento.FECHAALTAEVENTO), "dd")
    m_HoraMinuto = Format(CDate(m_FechaAhora), "hh:mm")
    dato = Split(m_HoraMinuto, ":")
    m_Hora = Nz(dato(0), "")
    m_Minuto = Nz(dato(1), "")
    
    m_ParteFija = p_Evento.ALIASTECNICO & "_" & m_Año & m_Mes & m_Dia & m_Hora & m_Minuto
    
    p_Actividad.IDActividad = DameValorSiguienteAlMaximoDeUnCampo("TbActividades", "IDActividad", m_ParteFija, 2)
    
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDActividad)='" & p_Actividad.IDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            p_Error = "La función DameIDActividad ha devuelto un valor de IDActividad ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDActividad") = p_Actividad.IDActividad
            .Fields("IDEvento") = p_Evento.IDEVENTO
            .Fields("ALIASTecnico") = p_Evento.ALIASTECNICO
            .Fields("FechaAlta") = p_Evento.FECHAALTAEVENTO
            If IsNumeric(p_Actividad.HorasLaborables) Then
                .Fields("HorasLaborables") = CDbl(p_Actividad.HorasLaborables)
            End If
            If IsNumeric(p_Actividad.HorasExtras) Then
                .Fields("HorasExtras") = CDbl(p_Actividad.HorasExtras)
            End If
            If p_Actividad.Actividad <> "" Then
                .Fields("Actividad") = p_Actividad.Actividad
            End If
            If p_Actividad.Ubicacion <> "" Then
                .Fields("Ubicacion") = p_Actividad.Ubicacion
            End If
            
            If p_Actividad.Descripcion <> "" Then
                .Fields("DESCRIPCION") = p_Actividad.Descripcion
            End If
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    AltaActividad = p_Actividad.IDActividad
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaActividad ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    AltaActividad = "#ERR" & "|" & p_Error
End Function

Public Function AltaEvento( _
                            p_Evento As Evento, _
                            Optional ByRef p_Error As String _
                            ) As String
     
    Dim rcdDatos As DAO.Recordset
    Dim m_IDEventoResultante As String
    On Error GoTo errores
    
    
    m_IDEventoResultante = DameIDEvento(p_Evento.BUI, p_Evento.FECHAALTAEVENTO)
    If m_IDEventoResultante = "" Then
        p_Error = "No se ha podido obtener el IDEvento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE (((TbEventos.IDEvento)='" & m_IDEventoResultante & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            p_Error = "La función Evento.DameIDEvento ha devuelto un IDEvento ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDEvento") = m_IDEventoResultante
            .Fields("NODO") = p_Evento.NODO
            .Fields("BUI") = p_Evento.BUI
            .Fields("SUBSISTEMA") = p_Evento.SubSistema
            If p_Evento.PMPR <> "" Then
                .Fields("PMPR") = p_Evento.PMPR
            End If
            .Fields("TIPOEVENTO") = p_Evento.TipoEvento
            .Fields("CRITICIDAD") = p_Evento.Criticidad
            .Fields("ALIASTECNICO") = p_Evento.ALIASTECNICO
            .Fields("ORIGINADOR") = p_Evento.Originador
            .Fields("FECHAALTAEVENTO") = p_Evento.FECHAALTAEVENTO
            .Fields("HORAINICIALEVENTO") = "7:30"
            .Fields("TIEMPORESPUESTAEVENTO") = "00:00"
            .Fields("IDEquipo") = p_Evento.IDEquipo
            .Fields("DESCRIPCION") = p_Evento.Descripcion
            If p_Evento.Contacto <> "" Then
                .Fields("CONTACTO") = p_Evento.Contacto
            End If
            .Fields("FechaRegistroAlta") = Now()
            .Fields("FechaRegistroModificacion") = Now()
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    AltaEvento = m_IDEventoResultante
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function RellenarComboContacto(Optional ByRef p_Error As String) As String
    
    Dim m_BUI As String
    On Error GoTo errores
    
    m_BUI = Nz(Me.BUI, "")
    
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

