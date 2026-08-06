Option Compare Database
Option Explicit

Public Event Eliminado()

Private m_Error As String
Private m_ObjTecnicoSeleccionado As Tecnico
Private WithEvents m_Form As Form_FormTecnico





Private Sub ComandoAlta_Click()
    On Error GoTo errores
    m_Error = ""
    Set m_ObjTecnicoActivo = Nothing
    If FormularioAbierto("FormTecnico") Then
        DoCmd.Close acForm, "FormTecnico", acSaveNo
    End If
    DoCmd.OpenForm "FormTecnico"
    Set m_Form = Forms("FormTecnico")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Sub




Private Sub ComandoActividades_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjTecnicoSeleccionado Is Nothing Then
        If Nz(Me.ListaTecnicos.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaTecnicos.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTecnicoSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    If Not FormularioAbierto("FormActividadGestion") Then
        DoCmd.OpenForm "FormActividadGestion"
    End If
    Forms("FormActividadGestion").Controls("ComboTécnicoBusqueda") = m_ObjTecnicoSeleccionado.Alias
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActividades_Click se ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoEdicion_Click()
   
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjTecnicoSeleccionado Is Nothing Then
        If Nz(Me.ListaTecnicos.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaTecnicos.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTecnicoSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjTecnicoActivo = m_ObjTecnicoSeleccionado
    
    If FormularioAbierto("FormTecnico") Then
        DoCmd.Close acForm, "FormTecnico", acSaveNo
    End If
    DoCmd.OpenForm "FormTecnico"
    Set m_Form = Forms("FormTecnico")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEdicion_Click se ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoFiestas_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    DoCmd.OpenForm "FormTecnicoFiestasGestion"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoFiestas_Click se ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoInformeEnExcel_Click()
  
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
     m_Error = ""
    
    flag = GenerarConsultas("Lista de Personal", Me.hWnd, , Me.ListaTecnicos)
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
        m_Error = "Al ComandoInformeEnExcel_Click se ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjTecnicoSeleccionado Is Nothing Then
        If Nz(Me.ListaTecnicos.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaTecnicos.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTecnicoSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    
    pregunta = MsgBox("¿Desea realmente eliminar al técnico seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de técnico")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    DoCmd.Hourglass True
    m_ObjTecnicoSeleccionado.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjTecnicoSeleccionado = Nothing
    Set m_ObjTecnicoActivo = Nothing
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Filtrar Nz(Me.ComboTecnico, ""), Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoInformeHoras_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_ObjTecnicoSeleccionado Is Nothing Then
        If Nz(Me.ListaTecnicos.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaTecnicos.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTecnicoSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    If FormularioAbierto("FormTecnicoConsultaHoras") Then
        DoCmd.Close acForm, "FormTecnicoConsultaHoras", acSaveNo
    End If
    Set m_ObjTecnicoActivo = m_ObjTecnicoSeleccionado
    DoCmd.OpenForm "FormTecnicoConsultaHoras"
    
    
'    lbl.Visible = False
'    Set lbl = Nothing
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoInformeHoras_Click se ha producido el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoLibranzas_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjTecnicoSeleccionado Is Nothing Then
        If Nz(Me.ListaTecnicos.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaTecnicos.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTecnicoSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjTecnicoActivo = m_ObjTecnicoSeleccionado
    If FormularioAbierto("FormTecnicoAusenciasGestion") Then
        DoCmd.Close acForm, "FormTecnicoAusenciasGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormTecnicoAusenciasGestion"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLibranzas_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoLimpiarEnActivo_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEnActivo = Null
    Filtrar Nz(Me.ComboTecnico.Column(0), ""), Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiarEnActivo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
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
    Filtrar Nz(Me.ComboTecnico.Column(0), ""), Nz(Me.ComboEnActivo, ""), m_Error
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

Private Sub ComandoTipoTecnico_Click()
    On Error GoTo errores
    
    m_Error = ""
    DoCmd.OpenForm "FormTipoTecnicoPreciosGestion"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoTipoTecnico_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboEnActivo_AfterUpdate()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Filtrar Nz(Me.ComboTecnico.Column(0), ""), Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboEnActivo_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboTecnico_AfterUpdate()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Filtrar Nz(Me.ComboTecnico.Column(0), ""), Nz(Me.ComboEnActivo, ""), m_Error
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
'    Me.InsideHeight = 8700
'    Me.InsideWidth = 9810
    
   
    
    Me.Caption = Me.lblTitulo1.Caption
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Filtrar , , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
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
Private Sub ListaTecnicos_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    Me.ComandoEdicion.Enabled = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoAsignar.Enabled = False
    Me.ComandoLibranzas.Enabled = False
    Me.ComandoInformeHoras.Enabled = False
    Me.ComandoActividades.Enabled = False
    
    If Nz(Me.ListaTecnicos.Column(0), "") = "" Then
        Set m_ObjTecnicoSeleccionado = Nothing
        Exit Sub
    End If
    If m_ObjTecnicoSeleccionado Is Nothing Then
        Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaTecnicos.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaTecnicos.Column(0) <> m_ObjTecnicoSeleccionado.Alias Then
            Set m_ObjTecnicoSeleccionado = Constructor.getTecnico(Me.ListaTecnicos.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjTecnicoSeleccionado Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 100
    End If
    Me.ComandoEdicion.Enabled = True
    Me.ComandoEliminar.Enabled = True
    Me.ComandoAsignar.Enabled = True
    Me.ComandoLibranzas.Enabled = True
    Me.ComandoInformeHoras.Enabled = True
    Me.ComandoActividades.Enabled = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaTecnicos_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function Filtrar( _
                        Optional p_Alias As String, _
                        Optional p_EnActivo As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    Dim m_Valor As String
    Dim m_Where As String
    Dim m_WhereNombre As String
    Dim m_WhereActivo As String
    Dim m_SQLInicio As String
    Dim lst As ListBox
    Dim i As Integer
    
    On Error GoTo errores
    
    Set m_ObjTecnicoActivo = Nothing
    Set lst = Me.ListaTecnicos
    Me.ComandoEdicion.Enabled = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoAsignar.Enabled = False
    Me.ComandoLibranzas.Enabled = False
    Me.ComandoInformeHoras.Enabled = False
    Me.ComandoActividades.Enabled = False
    
    lst.RowSource = "ALIAS;NOMBRE;TIPO;FECHAALTA;FECHABAJA"
    m_SQLInicio = "SELECT TbTecnicos.* " & _
                        "FROM TbTecnicos "
    If p_Alias = "" Then
        m_WhereNombre = "(ALIAS Like '*' Or ALIAS Is Null)"
    Else
        m_WhereNombre = "ALIAS Like '*" & p_Alias & "*'"
    End If
    If p_EnActivo = "" Then
        m_WhereActivo = "(FECHABAJA Like '*' Or FECHABAJA Is Null)"
    Else
        If p_EnActivo = "Sí" Then
            m_WhereActivo = "FECHABAJA Is Null"
        Else
            m_WhereActivo = "Not FECHABAJA Is Null"
        End If
        
    End If
    
    m_Where = "WHERE " & _
                            m_WhereNombre & _
                            " AND " & m_WhereActivo & _
                            ";"
    m_SQL = m_SQLInicio & m_Where
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                lst.AddItem Nz(.Fields("ALIAS"), "") & ";" & Nz(.Fields("NOMBRE"), "") & ";" & _
                            Nz(.Fields("TIPO"), "") & ";" & Nz(.Fields("FECHAALTA"), "") & ";" & _
                            Nz(.Fields("FECHABAJA"), "")
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaTecnicos_Click
    Else
        If lst.ItemsSelected.count > 0 Then
            For i = 1 To lst.ListCount - 1
                If lst.Selected(i) = True Then
                    ListaTecnicos_Click
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

Private Sub m_Form_Alta(m_ObjTecnico As Tecnico)
    
    On Error GoTo errores
    
    m_Error = ""
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Filtrar Nz(Me.ComboTecnico, ""), Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Alta se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Form_Borrado()
    On Error GoTo errores
    
    m_Error = ""
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Filtrar Nz(Me.ComboTecnico, ""), Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Borrado se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Form_Editado(m_ObjTecnico As Tecnico)
    On Error GoTo errores
    
    m_Error = ""
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Filtrar Nz(Me.ComboTecnico, ""), Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Editado se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub