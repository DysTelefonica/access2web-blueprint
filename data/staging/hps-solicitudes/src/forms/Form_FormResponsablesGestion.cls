Option Compare Database
Option Explicit

Private WithEvents m_Form As Form_FormResponsable
Public m_ResponsableSeleccionado As Responsable


Private m_Error As String



Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
'
Public Function Filtrar(Optional ByRef p_Error As String) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_Responsable As Responsable
    Dim m_ID As Variant
    Dim i As Long
    
    
    Dim lst As ListBox
    

    On Error GoTo errores

    p_Error = ""
    Set m_ResponsableSeleccionado = Nothing
    
    If Me.ListaFiltrados.ListCount > 1 Then
        For i = 1 To Me.ListaFiltrados.ListCount - 1
            If Me.ListaFiltrados.Selected(i) = True Then
                Me.ListaFiltrados.Selected(i) = False
            End If
        Next
    End If
    
    Set lst = Me.ListaFiltrados
    lst.RowSource = "IDResponsable;Nombre;Correo;Descripción"
    Me.lblListaDeResultados.Caption = "Registros filtrados (0)"

    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    Set m_Col = constructor.getResponsables(p_Nombre:=Nz(Me.Nombre, ""), p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If

    'lst.RowSource = "IDResponsable;Nombre;Correo;Descripción"
    For Each m_ID In m_Col
        Set m_Responsable = m_Col(m_ID)
        With m_Responsable
            
            lst.AddItem .IDResponsable & ";" & .Nombre & ";" & .Correo & ";" & .Descripcion
        End With
siguiente:
        Set m_Responsable = Nothing
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
   
    Filtrar
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


Private Sub ComandoAlta_Click()
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    m_Error = ""
   If FormularioAbierto("FormResponsable") Then
        DoCmd.Close acForm, "FormResponsable", acSaveNo
    End If
    Set m_ObjResposableActivo = Nothing
    DoCmd.OpenForm "FormResponsable"
    If FormularioAbierto("FormResponsable") Then
        Set m_Form = Forms("FormResponsable")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEditar_Click()
        
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    If m_ResponsableSeleccionado Is Nothing Then
        m_Error = "Seleccione un Responsable de la lista"
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormResponsable") Then
        DoCmd.Close acForm, "FormResponsable", acSaveNo
    End If
    Set m_ObjResposableActivo = m_ResponsableSeleccionado
    DoCmd.OpenForm "FormResponsable"
    If FormularioAbierto("FormResponsable") Then
        Set m_Form = Forms("FormResponsable")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
     
    Dim m_ResponsableOp As ResponsableOperaciones
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
   
    If m_ResponsableSeleccionado Is Nothing Then
        m_Error = "Seleccione un Responsable de la lista"
        Err.Raise 1000
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox("¿Desea realmente eliminar definitivamente el responsable Seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar responsable")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Set m_ResponsableOp = New ResponsableOperaciones
    With m_ResponsableOp
        Set .Responsable = m_ResponsableSeleccionado
        .Eliminar p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    ComandoActualizar_Click
    Filtrar p_Error:=m_Error
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
        m_Error = "Al ComandoEliminar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    'On Error Resume Next
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Me.AllowEdits = False
    If Me.ComandoEditar.Enabled = True Then
        ComandoEditar_Click
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

    Dim m_SOl As solicitud
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""


    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    Me.AllowEdits = False
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Set m_ResponsableSeleccionado = Nothing
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If m_ResponsableSeleccionado Is Nothing Then
        Set m_ResponsableSeleccionado = constructor.getResponsable(p_IDResponsable:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaFiltrados.Column(0) <> m_ResponsableSeleccionado.IDResponsable Then
            Set m_ResponsableSeleccionado = constructor.getResponsable(p_IDResponsable:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If

        End If
    End If
    
    
    If m_ResponsableSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    Me.ComandoEditar.Enabled = True
    Me.ComandoEliminar.Enabled = True
   
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

Private Sub m_Form_Alta()
    
    Filtrar
End Sub


Private Sub m_Form_Editado()
    Filtrar
End Sub

Private Sub Nombre_BeforeUpdate(Cancel As Integer)
    Filtrar
End Sub

