Option Compare Database
Option Explicit


Private m_Error As String
Private m_ObjTipoTecnicoPrecioSeleccionado As TipoTecnicoPrecio
Private WithEvents m_Form As Form_FormTipoTecnicoPrecio
Private WithEvents m_FormTipoTecnico As Form_FormTecnicoTipo

Private Sub ComandoAltaPerfil_Click()
    
    On Error GoTo errores
    m_Error = ""
    Set m_ObjTipoTecnicoActivo = Nothing
    If FormularioAbierto("FormTecnicoTipo") Then
        DoCmd.Close acForm, "FormTecnicoTipo", acSaveNo
    End If
    DoCmd.OpenForm "FormTecnicoTipo"
    Set m_FormTipoTecnico = Forms("FormTecnicoTipo")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaPerfil_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Sub

Private Sub ComandoAltaPerfilHora_Click()
    
    On Error GoTo errores
    m_Error = ""
    
    If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
        If Nz(Me.ListaPreciosPerfil.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTipoTecnicoPrecioSeleccionado = Constructor.getTecnico(Me.ListaPreciosPerfil.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    
    Set m_ObjTipoTecnicoPrecioActivo = m_ObjTipoTecnicoPrecioSeleccionado
    m_ObjTipoTecnicoPrecioActivo.ParaAlta = EnumSino.Sí
    If FormularioAbierto("FormTipoTecnicoPrecio") Then
        DoCmd.Close acForm, "FormTipoTecnicoPrecio", acSaveNo
    End If
    DoCmd.OpenForm "FormTipoTecnicoPrecio"
    Set m_Form = Forms("FormTipoTecnicoPrecio")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaPerfilHora_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Sub

Private Sub ComandoEdicionPerfilHora_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
        If Nz(Me.ListaPreciosPerfil.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTipoTecnicoPrecioSeleccionado = Constructor.getTecnico(Me.ListaPreciosPerfil.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
   
    Set m_ObjTipoTecnicoPrecioActivo = m_ObjTipoTecnicoPrecioSeleccionado
    If m_ObjTipoTecnicoPrecioSeleccionado.EsUltima = EnumSino.No Then
        m_Error = "Solo se puede editar el último perfil del seleccionado"
        Err.Raise 1000
    End If
    m_ObjTipoTecnicoPrecioActivo.ParaAlta = EnumSino.No
    
    If FormularioAbierto("FormTipoTecnicoPrecio") Then
        DoCmd.Close acForm, "FormTipoTecnicoPrecio", acSaveNo
    End If
    DoCmd.OpenForm "FormTipoTecnicoPrecio"
    Set m_Form = Forms("FormTipoTecnicoPrecio")
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEdicionPerfilHora_Click se ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub





Private Sub ComandoEditarPerfil_Click()
    
End Sub


Private Sub ComandoEliminarPerfilHora_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
        If Nz(Me.ListaPreciosPerfil.Column(0), "") = "" Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
        Set m_ObjTipoTecnicoPrecioSeleccionado = Constructor.getTipoTecnicoPrecio(Me.ListaPreciosPerfil.Column(0), , , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
            m_Error = "Ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    
    pregunta = MsgBox("¿Desea realmente eliminar el Perfil/Precio seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de Perfil/Técnico")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    DoCmd.Hourglass True
    m_ObjTipoTecnicoPrecioSeleccionado.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjTipoTecnicoPrecioSeleccionado = Nothing
    Set m_ObjTipoTecnicoPrecioActivo = Nothing
    
    
    Filtrar Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarPerfilHora_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    Filtrar Nz(Me.ComboEnActivo, ""), m_Error
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

Private Sub ComboEnActivo_AfterUpdate()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Filtrar Nz(Me.ComboEnActivo, ""), m_Error
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



Private Sub Form_Open(Canel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me
'    Me.InsideHeight = 8700
'    Me.InsideWidth = 9810
    Me.Caption = Me.lblTitulo1.Caption
    Filtrar , m_Error
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
Private Sub ListaPreciosPerfil_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComandoAltaPerfilHora.Enabled = False
    Me.ComandoEdicionPerfilHora.Enabled = False
    Me.ComandoEliminarPerfilHora.Enabled = False
    Me.ComandoAltaPerfil.Enabled = True
    
    If Nz(Me.ListaPreciosPerfil.Column(0), "") = "" Then
        Set m_ObjTipoTecnicoPrecioSeleccionado = Nothing
        Exit Sub
    End If
    If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
        Set m_ObjTipoTecnicoPrecioSeleccionado = Constructor.getTipoTecnicoPrecio(Me.ListaPreciosPerfil.Column(0), , , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaPreciosPerfil.Column(0) <> m_ObjTipoTecnicoPrecioSeleccionado.IDTipoTecnicoFecha Then
            Set m_ObjTipoTecnicoPrecioSeleccionado = Constructor.getTipoTecnicoPrecio(Me.ListaPreciosPerfil.Column(0), , , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjTipoTecnicoPrecioSeleccionado Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 100
    End If
    Me.ComandoAltaPerfilHora.Enabled = True
    Me.ComandoEdicionPerfilHora.Enabled = True
    
    Me.ComandoEliminarPerfilHora.Enabled = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaPreciosPerfil_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function Filtrar( _
                        Optional p_EnActivo As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    Dim m_Valor As String
    Dim m_Where As String
    Dim i As Integer
    Dim m_WhereActivo As String
    Dim m_SQLInicio As String
    Dim lst As ListBox
    On Error GoTo errores
    
    Set m_ObjTipoTecnicoPrecioActivo = Nothing
    Set lst = Me.ListaPreciosPerfil
    Me.ComandoAltaPerfilHora.Enabled = False
    Me.ComandoEdicionPerfilHora.Enabled = False
    
    Me.ComandoEliminarPerfilHora.Enabled = False
    Me.ComandoAltaPerfil.Enabled = True
   
    
    lst.RowSource = "ID;PERFIL;FInicial;FFinal;PrecioHLab;PrecioHExtra"
    m_SQLInicio = "SELECT TbTipoTecnicoPrecios.* " & _
                    "FROM TbTipoTecnicoPrecios "
    
    If p_EnActivo = "" Then
        m_WhereActivo = "(FechaFinal Like '*' Or FechaFinal Is Null)"
    Else
        If p_EnActivo = "Sí" Then
            m_WhereActivo = "FechaFinal Is Null"
        Else
            m_WhereActivo = "Not FechaFinal Is Null"
        End If
        
    End If
    
    m_Where = "WHERE " & _
                            m_WhereActivo & _
                            " ORDER BY TbTipoTecnicoPrecios.TIPOTECNICO;"
    m_SQL = m_SQLInicio & m_Where
    'TODO: Cambiar currentdb por getdb
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                lst.AddItem Nz(.Fields("IDTipoTecnicoFecha"), "") & ";" & Nz(.Fields("TIPOTECNICO"), "") & ";" & _
                            Nz(.Fields("FechaInicial"), "") & ";" & Nz(.Fields("FechaFinal"), "") & ";" & _
                            Format(Nz(.Fields("PrecioHoraLab"), "0"), "#,##0.00 €") & ";" & _
                            Format(Nz(.Fields("PrecioHoraExtra"), "0"), "#,##0.00 €")
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaPreciosPerfil_Click
    Else
        If lst.ItemsSelected.count > 0 Then
            For i = 1 To lst.ListCount - 1
                If lst.Selected(i) = True Then
                    ListaPreciosPerfil_Click
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



Private Sub m_Form_Alta(m_ObjTipoTecnico As TipoTecnicoPrecio)
    
    On Error GoTo errores
    
    m_Error = ""
    
    Filtrar Nz(Me.ComboEnActivo, ""), m_Error
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



Private Sub m_Form_Edicion(m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio)
    On Error GoTo errores
    
    m_Error = ""
    
    Filtrar Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Edicion se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormTipoTecnico_Alta(m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio)
    
    On Error GoTo errores
    
    m_Error = ""
    
    Filtrar Nz(Me.ComboEnActivo, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormTipoTecnico_Alta se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub