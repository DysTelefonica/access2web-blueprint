Option Compare Database
Option Explicit

Private m_Error As String
Private m_ObjTecnicoSeleccionado As Tecnico
Private m_ColTecnicos As Scripting.Dictionary
Private m_ObjActividad As Actividad


Private Function RellenarCole(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ID As String
    Dim i As Integer
    Dim m_ObjTecnico As Tecnico
    
    On Error GoTo errores
    
    Set m_ColTecnicos = Nothing
    Set lst = Me.ListaTecnicos
    If lst.ListCount = 1 Then
        Set lst = Nothing
        Exit Function
    End If
    For i = 1 To lst.ListCount - 1
        m_ID = Nz(lst.Column(0, i), "")
        Set m_ObjTecnico = Constructor.getTecnico(CStr(m_ID), p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ColTecnicos Is Nothing Then
            Set m_ColTecnicos = New Scripting.Dictionary
            m_ColTecnicos.CompareMode = TextCompare
            If Not m_ColTecnicos.Exists(m_ID) Then
                m_ColTecnicos.Add m_ID, m_ColTecnicos
            End If
            Set m_ColTecnicos = Nothing
        End If
    Next
    Set lst = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCole ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function AddTecnico(p_Alias As String, Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ID As String
    Dim i As Integer
    Dim m_ObjTecnico As Tecnico
   
    
    On Error GoTo errores
    If p_Alias = "" Then
        Exit Function
    End If
    
    Set lst = Me.ListaTecnicos
    
    
    
    If lst.ListCount > 1 Then
        For i = 1 To lst.ListCount - 1
            m_ID = Nz(lst.Column(0, i), "")
            If m_ID = p_Alias Then
                Exit Function
            End If
        Next
    End If
    Set m_ObjTecnico = Constructor.getTecnico(p_Alias, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjTecnico Is Nothing Then
        p_Error = "No existe un técnico con ese alias"
        Err.Raise 1000
    End If
   
    With m_ObjTecnico
        lst.AddItem .Alias & ";" & .Nombre & ";" & .Tipo
        If m_ColTecnicos Is Nothing Then
            Set m_ColTecnicos = New Scripting.Dictionary
            m_ColTecnicos.CompareMode = TextCompare
        End If
        If Not m_ColTecnicos.Exists(.Alias) Then
            m_ColTecnicos.Add .Alias, m_ObjTecnico
        End If
    End With
    Set m_ObjTecnico = Nothing
    Set lst = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AddTecnico ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function DeleteTecnico(p_Alias As String, Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ID As String
    Dim i As Integer
    Dim VarItem  As Variant
    
    On Error GoTo errores
    If p_Alias = "" Then
        Exit Function
    End If
    
    Set lst = Me.ListaTecnicos
    
    
    If lst.ListCount = 1 Then
        Set lst = Nothing
        Exit Function
    End If
    If lst.ItemsSelected.count = 0 Then
        Set lst = Nothing
        Exit Function
    End If
    If lst.ListCount > 1 Then
        For i = 1 To lst.ListCount - 1
            m_ID = Nz(lst.Column(0, i), "")
            If m_ID = p_Alias Then
                If Not m_ColTecnicos Is Nothing Then
                    If m_ColTecnicos.Exists(p_Alias) Then
                        m_ColTecnicos.Remove p_Alias
                        If m_ColTecnicos.count = 0 Then
                            Set m_ColTecnicos = Nothing
                        End If
                    End If
                End If
            End If
        Next
    End If
    For Each VarItem In lst.ItemsSelected
        lst.RemoveItem VarItem
    Next
    Set lst = Nothing
    
        
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DeleteTecnico ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function


Private Sub ComandoAltaTecnico_Click()
    
    On Error GoTo errores
    m_Error = ""
    If Nz(Me.ComboTecnico.Column(0), "") = "" Then
        Exit Sub
    End If
    AddTecnico Me.ComboTecnico.Column(0), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaTecnico_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarTecnico_Click()
    
    On Error GoTo errores
    m_Error = ""
    If Nz(Me.ListaTecnicos.Column(0), "") = "" Then
        Exit Sub
    End If
    DeleteTecnico Me.ListaTecnicos.Column(0), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarTecnico_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrar_Click()
        
    Dim m_Actividad As String
    Dim m_Ubicacion As String
    Dim m_HorasTotales As String
    Dim m_FechaPrimeraActividad As String
    Dim m_DESCRIPCION As String
    Dim m_ColActividadesPorLotes As Scripting.Dictionary
    
    Dim lst As ListBox
    
        
    On Error GoTo errores
    
    
    m_Error = ""
    m_Actividad = Nz(Me.ComboActividad, "")
    m_Ubicacion = Nz(Me.ComboUbicacion, "")
    m_HorasTotales = Nz(Me.HorasTotales, "")
    
    m_FechaPrimeraActividad = Nz(Me.FechaPrimeraActividad, "")
    m_DESCRIPCION = Nz(Me.Descripcion, "")
    If m_DESCRIPCION = "" Then
        m_Error = "Ha de indicar la Descripción"
        Err.Raise 1000
    End If
    If m_Actividad = "" Then
        m_Error = "Ha de indicar la actividad"
        Err.Raise 1000
    End If
    If m_Ubicacion = "" Then
        m_Error = "Ha de indicar la ubicación"
        Err.Raise 1000
    End If
    If Not IsNumeric(m_HorasTotales) Then
        m_Error = "Ha de indicar las horas totales"
        Err.Raise 1000
    End If
    If CDbl(m_HorasTotales) = 0 Then
        m_Error = "Ha de indicar las horas totales"
        Err.Raise 1000
    End If
    If m_ColTecnicos Is Nothing Then
        m_Error = "Ha de indicar al menos un técnico"
        Err.Raise 1000
    End If
    If Not IsDate(m_FechaPrimeraActividad) Then
        m_Error = "Ha de indicar la fecha del primer día"
        Err.Raise 1000
    End If
    Set m_ObjActividad = New Actividad
    With m_ObjActividad
        .Actividad = m_Actividad
        .Ubicacion = m_Ubicacion
        .Descripcion = m_DESCRIPCION
        .Observaciones = "De Lote"
    End With
    If m_ColTecnicos.count = 1 Then
        pregunta = MsgBox("Desea generar las actividades que cubran las " & m_HorasTotales & " para el técnico seleccionado?", vbExclamation + vbYesNo + vbDefaultButton1, _
                        "Actividades por Lotes de Evento")
    Else
        pregunta = MsgBox("Desea generar las actividades para repartir las " & m_HorasTotales & " entre los " & _
                        m_ColTecnicos.count & " técnicos?", vbExclamation + vbYesNo + vbDefaultButton1, _
                        "Actividades por Lotes de Evento")
    End If
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_ObjEventoActivo.GenerarActividadesPorLotes m_ColTecnicos, m_ObjActividad, m_HorasTotales, m_FechaPrimeraActividad, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    '---------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        Form_FormEventoGestion.ListaEventos_Click
    End If
    '---------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormActividadGestion") Then
        Forms("FormActividadGestion").Controls("ListaFiltrados").Requery
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComboTecnico_AfterUpdate()
    If Nz(Me.ComboTecnico, "") = "" Then
        Me.ComandoAltaTecnico.Enabled = False
    Else
        Me.ComandoAltaTecnico.Enabled = True
    End If
End Sub

Private Sub Detalle_DblClick(Cancel As Integer)
    Application.Echo False
    Me.ComboActividad = "CR001"
    Me.ComboUbicacion = "DESPLAZAMIENTO NACIONAL"
    Me.HorasTotales = 100
    Me.Descripcion = "Prueba Masiva"
    Me.FechaPrimeraActividad = "02/12/2022"
    Me.ComboTecnico = "JGA"
    ComandoAltaTecnico_Click
    Me.ComboTecnico = Null
    Application.Echo True
End Sub

Private Sub Form_Load()
    
    Dim m_Titulo As String
    Dim m_SQL As String
    Dim cbo As ComboBox
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    m_Error = ""
    If m_ObjEventoActivo Is Nothing Then
        m_Error = "No se sabe el evento activo"
        Err.Raise 1000
    End If
    If m_ObjEventoActivo.FechaFacturado <> "" Then
        m_Error = "El Evento aparece como facturado el " & m_ObjEventoActivo.FechaFacturado
        Err.Raise 1000
    End If
    If m_ObjEventoActivo.Franqueado = True Then
        m_Error = "El Evento aparece como Franquedo"
        Err.Raise 1000
    End If
    m_Titulo = "ALTA DE ACTIVIDADES POR LOTES DE " & UCase(m_ObjEventoActivo.IDEVENTO)
    Me.lblTitulo1.Caption = m_Titulo
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, EnumSino.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbCodActividad.CODIGO, TbCodActividad.ACTIVIDADNOPROGRAMADA FROM TbCodActividad ORDER BY TbCodActividad.CODIGO;"
    Set cbo = Me.ComboActividad
    RellenaCombo cbo, m_SQL
    m_SQL = "SELECT TbUbicacion.UBICACION FROM TbUbicacion ORDER BY TbUbicacion.UBICACION; "
    Set cbo = Me.ComboUbicacion
    RellenaCombo cbo, m_SQL
    Me.ListaTecnicos.RowSource = "ALIAS;NOMBRE"
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

