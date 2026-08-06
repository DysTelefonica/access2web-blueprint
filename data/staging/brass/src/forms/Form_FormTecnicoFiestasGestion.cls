Option Compare Database
Option Explicit

Private m_Error As String
Private m_ObjFiestaSeleccionada As Fiesta
Private WithEvents m_Form As Form_FormTecnicoFiestas

Private Function RellenarCombo(Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim cmb As ComboBox
    On Error GoTo errores
    
    Set cmb = Me.ComboAño
    cmb.RowSource = ""
    m_SQL = "SELECT DISTINCT Year([FechaFiesta]) AS Año FROM TbTecnicosFiestas;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
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
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCombo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function Filtrar( _
                        Optional p_Año As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim m_ObjFiesta As Fiesta
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim i As Integer
    Dim lst As ListBox
    On Error GoTo errores
    
    Set lst = Me.ListaFiestas
    
    Me.ComandoEliminar.Enabled = False
    
    
    lst.RowSource = "Fecha;Observaciones"
    Set m_Col = Constructor.getFiestas(p_Año, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjFiesta = m_Col(m_ID)
        lst.AddItem m_ObjFiesta.FechaFiesta & ";" & m_ObjFiesta.Observaciones
        Set m_ObjFiesta = Nothing
    Next
    
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaFiestas_Click
    Else
        If lst.ItemsSelected.count > 0 Then
            For i = 1 To lst.ListCount - 1
                If lst.Selected(i) = True Then
                    ListaFiestas_Click
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

Private Sub ComandoAlta_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormTecnicoFiestas"
    Set m_Form = Forms("FormTecnicoFiestas")
    
End Sub

Private Sub ComandoEliminar_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjFiestaSeleccionada Is Nothing Then
        If Nz(Me.ListaFiestas.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjFiestaSeleccionada = Constructor.getFiesta(Me.ListaFiestas.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjFiestaSeleccionada Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    
    pregunta = MsgBox("¿Desea realmente eliminar la fiesta seleccionada?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de Fiesta")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    DoCmd.Hourglass True
    m_ObjFiestaSeleccionada.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjFiestaSeleccionada = Nothing
    
    
    Filtrar Nz(Me.ComboAño, ""), m_Error
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
Private Sub ComandoExportarExcel_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    
    flag = GenerarConsultas("FiltroFiestas", Me.hWnd, , Me.ListaFiestas, "No")
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


Private Sub ComandoLimpiarAño_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboAño = Null
    Filtrar Nz(Me.ComboAño.Column(0), ""), m_Error
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

Private Sub ComboAño_AfterUpdate()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Filtrar Nz(Me.ComboAño.Column(0), ""), m_Error
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
Private Sub Form_Open(Canel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me

    Me.Caption = Me.lblTitulo1.Caption
    RellenarCombo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Filtrar , m_Error
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
Private Sub ListaFiestas_Click()
    On Error GoTo errores
    
    m_Error = ""
    Me.ComandoEliminar.Enabled = False
    
    
    
    If Nz(Me.ListaFiestas.Column(0), "") = "" Then
        Set m_ObjFiestaSeleccionada = Nothing
        Exit Sub
    End If
    If m_ObjFiestaSeleccionada Is Nothing Then
        Set m_ObjFiestaSeleccionada = Constructor.getFiesta(Me.ListaFiestas.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaFiestas.Column(0) <> m_ObjFiestaSeleccionada.FechaFiesta Then
            Set m_ObjFiestaSeleccionada = Constructor.getFiesta(Me.ListaFiestas.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjFiestaSeleccionada Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 100
    End If
    Me.ComandoEliminar.Enabled = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiestas_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Form_Alta()
    
    On Error GoTo errores
    
    m_Error = ""
   
    Filtrar Nz(Me.ComboAño, ""), m_Error
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