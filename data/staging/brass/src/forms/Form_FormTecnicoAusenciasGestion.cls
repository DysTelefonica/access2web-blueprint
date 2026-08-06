Option Compare Database
Option Explicit

Private m_Error As String
Private m_ObjLibranzaSeleccionada As LIbranza
Private WithEvents m_Form As Form_FormTecnicoAusencias


Private Function RellenarCombos(Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim cmb As ComboBox
    On Error GoTo errores
    
    Set cmb = Me.ComboAño
    cmb.RowSource = ""
    m_SQL = "SELECT DISTINCT Year([FechaLibranza]) AS Año FROM TbTecnicosAusencias;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            cmb.AddItem Nz(.Fields("Año"), "")
            .MoveNext
        Loop
    End With
    Set cmb = Me.ComboTecnico
    cmb.RowSource = ""
    m_SQL = "SELECT DISTINCT TbTecnicosAusencias.Alias FROM TbTecnicosAusencias ORDER BY TbTecnicosAusencias.Alias; "
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            cmb.AddItem Nz(.Fields("Alias"), "")
            .MoveNext
        Loop
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCombos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function


Private Sub ComandoAlta_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    
    Set m_ObjLibranzaActiva = Nothing
    Set m_ObjTecnicoActivo = Constructor.getTecnico(Nz(Me.ListaAusencias.Column(1), ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjTecnicoActivo Is Nothing Then
        Set m_ObjTecnicoActivo = Constructor.getTecnico(Nz(Me.ComboTecnico, ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_ObjTecnicoActivo Is Nothing Then
        m_Error = "No se conoce el técnico activo"
        Err.Raise 1000
    End If
    If FormularioAbierto("FormTecnicoAusencias") Then
        DoCmd.Close acForm, "FormTecnicoAusencias", acSaveNo
    End If
    DoCmd.OpenForm "FormTecnicoAusencias"
    On Error Resume Next
    Set m_Form = Forms("FormTecnicoAusencias")
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEdicion_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjLibranzaSeleccionada Is Nothing Then
        If Nz(Me.ListaAusencias.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjLibranzaSeleccionada = Constructor.getLibranza(Me.ListaAusencias.Column(0), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjLibranzaSeleccionada Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjLibranzaActiva = m_ObjLibranzaSeleccionada
    
    If FormularioAbierto("FormTecnicoAusencias") Then
        DoCmd.Close acForm, "FormTecnicoAusencias", acSaveNo
    End If
    DoCmd.OpenForm "FormTecnicoAusencias"
    On Error Resume Next
    Set m_Form = Forms("FormTecnicoAusencias")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEdicion_Click se ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoExportarExcel_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    
    flag = GenerarConsultas("FiltroAusencias", Me.hWnd, , Me.ListaAusencias, "No")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método GenerarConsultas ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoExportarExcel_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function Filtrar( _
                        Optional p_Año As String, _
                        Optional p_Tecnico As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim m_ObjLibranza As LIbranza
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim i As Integer
    Dim lst As ListBox
    On Error GoTo errores
    
    Set lst = Me.ListaAusencias
    Me.ComandoEdicion.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    
    lst.RowSource = "IDAusencia;Alias;Ausencia;Motivo"
    Set m_Col = Constructor.getLibranzas(p_Año, p_Tecnico, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjLibranza = m_Col(m_ID)
        lst.AddItem m_ObjLibranza.IDAusencia & ";" & m_ObjLibranza.Alias & ";" & _
                    m_ObjLibranza.FechaLibranza & ";" & m_ObjLibranza.MOTIVO
        Set m_ObjLibranza = Nothing
    Next
    
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaAusencias_Click
    Else
        If lst.ItemsSelected.count > 0 Then
            For i = 1 To lst.ListCount - 1
                If lst.Selected(i) = True Then
                    ListaAusencias_Click
                    Exit For
                End If
            Next
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Sub ComandoEliminar_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjLibranzaSeleccionada Is Nothing Then
        If Nz(Me.ListaAusencias.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjLibranzaSeleccionada = Constructor.getLibranza(Me.ListaAusencias.Column(0), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjLibranzaSeleccionada Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente eliminar la libranza seleccionada?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de libranza")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjLibranzaSeleccionada.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Filtrar Nz(Me.ComboAño, ""), Nz(Me.ComboTecnico, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoLimpiarAño_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboAño = Null
    Filtrar Nz(Me.ComboAño.Column(0), ""), Nz(Me.ComboTecnico, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiarAño_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoLimpiarTecnico_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboTecnico = Null
    Filtrar Nz(Me.ComboAño.Column(0), ""), Nz(Me.ComboTecnico, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiarTecnico_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboAño_AfterUpdate()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Filtrar Nz(Me.ComboAño.Column(0), ""), Nz(Me.ComboTecnico, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboAño_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboAño_Enter()
    On Error Resume Next
    Me.ComboAño.Requery
End Sub

Public Sub ComboTecnico_AfterUpdate()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Filtrar Nz(Me.ComboAño.Column(0), ""), Nz(Me.ComboTecnico, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboTecnico_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboTecnico_Enter()
    On Error Resume Next
    Me.ComboTecnico.Requery
End Sub
Private Sub Form_Open(Canel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me

    Me.Caption = Me.lblTitulo1.Caption
    RellenarCombos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjTecnicoActivo Is Nothing Then
        Me.ComboTecnico = m_ObjTecnicoActivo.Alias
        Me.ComboAño = Year(Date)
        Filtrar Nz(Me.ComboAño.Column(0), ""), Nz(Me.ComboTecnico, ""), m_Error
    Else
        Filtrar , , m_Error
    End If
    
    If m_Error <> "" Then
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
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ListaAusencias_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    Me.ComandoEdicion.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    If Me.ListaAusencias.ItemsSelected.count = 0 Then
        Set m_ObjLibranzaSeleccionada = Nothing
        Exit Sub
    End If
    If Me.ListaAusencias.ItemsSelected.count > 1 Then
        Set m_ObjLibranzaSeleccionada = Nothing
        Exit Sub
    End If
    If Nz(Me.ListaAusencias.Column(0), "") = "" Then
        Set m_ObjLibranzaSeleccionada = Nothing
        Exit Sub
    End If
    If m_ObjLibranzaSeleccionada Is Nothing Then
        Set m_ObjLibranzaSeleccionada = Constructor.getLibranza(Me.ListaAusencias.Column(0), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaAusencias.Column(0) <> m_ObjLibranzaSeleccionada.IDAusencia Then
            Set m_ObjLibranzaSeleccionada = Constructor.getLibranza(Me.ListaAusencias.Column(0), , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjLibranzaSeleccionada Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 100
    End If
    Me.ComandoEdicion.Enabled = True
    Me.ComandoEliminar.Enabled = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaAusencias_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub m_Form_Alta()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Filtrar Nz(Me.ComboAño.Column(0), ""), Nz(Me.ComboTecnico, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Alta se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub