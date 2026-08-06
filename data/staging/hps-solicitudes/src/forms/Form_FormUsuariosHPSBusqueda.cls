Option Compare Database
Option Explicit


Public m_UsuarioSeleccionado As UsuarioHPS

Public Event Seleccionado(p_Usuario As UsuarioHPS)
Private m_Tipo As String
Private m_Error As String



Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
'
Public Function Filtrar(Optional ByRef p_Error As String) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_UsuarioHPS As UsuarioHPS
    Dim m_ID As Variant
    Dim i As Long


    Dim lst As ListBox


    On Error GoTo errores

    p_Error = ""
    Set m_UsuarioSeleccionado = Nothing

    If Me.ListaFiltrados.ListCount > 1 Then
        For i = 1 To Me.ListaFiltrados.ListCount - 1
            If Me.ListaFiltrados.Selected(i) = True Then
                Me.ListaFiltrados.Selected(i) = False
            End If
        Next
    End If

    Set lst = Me.ListaFiltrados
    lst.RowSource = "ID;DNI;Nombre"
    Me.lblListaDeResultados.Caption = "Registros filtrados (0)"
    If Nz(Me.PalabraClave, "") = "" Then
        Set m_Col = m_ObjEntorno.colUsuariosHPS
    Else
        Set m_Col = constructor.getUsuariosHPSPorPalabraClave(p_PalabraClave:=Me.PalabraClave, p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
   
    
    If m_Col Is Nothing Then
        Exit Function
    End If

    'ID;DNI;Nombre
    For Each m_ID In m_Col
        Set m_UsuarioHPS = m_Col(m_ID)
        With m_UsuarioHPS

            lst.AddItem .id & ";" & .DNI & ";" & .Nombre & " " & .Apellido_1 & " " & .Apellido_2
        End With
siguiente:
        Set m_UsuarioHPS = Nothing
    Next
    Me.lblListaDeResultados.Caption = "Registros filtrados (" & lst.ListCount - 1 & ")"
    If lst.ListCount = 1 Then
        Exit Function
    End If
    lst.Selected(1) = True
    ListaFiltrados_Click


    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub ComandoActualizar_Click()

    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.PalabraClave = ""
    Set m_ObjEntorno.colUsuariosHPS = Nothing

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoSeleccionar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_UsuarioSeleccionado Is Nothing Then
        m_Error = "Seleccione una elemento de la lista"
        Err.Raise 1000
    End If
    RaiseEvent Seleccionado(m_UsuarioSeleccionado)
    'Set m_ObjUsuarioHPSActivo = constructor.getUsuarioHPS(p_IDUsuario:=m_UsuarioSeleccionado.ID)
    Set m_ObjUsuarioHPSActivo = m_UsuarioSeleccionado
    If FormularioAbierto("FormSolicitudRenovacion") Then
        DoCmd.Close acForm, "FormSolicitudRenovacion", acSaveNo
    End If
    Set m_ObjSolicitudActiva = Nothing
    Set m_ObjSolicitudAlInicio = Nothing
    DoCmd.OpenForm "FormSolicitudRenovacion", OpenArgs:=m_Tipo
    DoCmd.Close acForm, Me.Name
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSeleccionar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Me.AllowEdits = False
    If Me.ComandoSeleccionar.Enabled = True Then
        ComandoSeleccionar_Click
    End If

    Me.AllowEdits = True
End Sub

Private Sub Form_Open(Cancel As Integer)

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    Ajustar Me
    Me.Caption = m_TituloFormulario
    Me.AllowEdits = False

    Me.Caption = Me.lblTitulo.Caption
    m_Tipo = Nz(Me.OpenArgs(), "")
    If m_Tipo = "" Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    'On Error Resume Next
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaFiltrados_Click()

    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""


    Me.ComandoSeleccionar.Enabled = False
    

    Me.AllowEdits = False
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Set m_UsuarioSeleccionado = Nothing

        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If m_UsuarioSeleccionado Is Nothing Then
        Set m_UsuarioSeleccionado = constructor.getUsuarioHPS(p_IDUsuario:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaFiltrados.Column(0) <> m_UsuarioSeleccionado.id Then
            Set m_UsuarioSeleccionado = constructor.getUsuarioHPS(p_IDUsuario:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If

        End If
    End If


    If m_UsuarioSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    Me.ComandoSeleccionar.Enabled = True
    

    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub


Private Sub PalabraClave_BeforeUpdate(Cancel As Integer)
    Filtrar
End Sub