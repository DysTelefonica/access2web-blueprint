Option Compare Database
Option Explicit
Private m_ObjAplicacionSeleccionada As Aplicacion
Public Event AplicacionSeleccionada(Aplicacion As Aplicacion)
Public Event QuitarAplicacionSeleccionada()
Private m_Error As String
Private Sub ComandoLimpiarNombre_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.NombreAplicacion = Null
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiarNombre_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function Filtrar( _
                        Optional ByRef p_Error As String _
                            ) As String
    
    Dim lst As ListBox
    Dim m_ID As Variant
    Dim m_objAplicacion As Aplicacion
    Dim m_ObjListaAplicaciones As Scripting.Dictionary
    
    On Error GoTo errores
    
    Set m_ObjAplicacionActiva = Nothing
    Me.AplicacionSeleccionada = Null
    Set lst = Me.ListaAplicaciones
    lst.RowSource = "ID;Nombre"
    
    Set m_ObjListaAplicaciones = Constructor.getListaAplicacionesPorNombre(Nz(Me.NombreAplicacion, ""), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjListaAplicaciones Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjListaAplicaciones
        Set m_objAplicacion = m_ObjListaAplicaciones(m_ID)
         With m_objAplicacion
            lst.AddItem .IDAplicacion & ";" & .NombreAplicacion
        End With
        Set m_ObjAplicacionSeleccionada = m_objAplicacion

    Next
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaAplicaciones_Click
        Set m_ObjAplicacionActiva = m_ObjAplicacionSeleccionada
        
        Me.AplicacionSeleccionada = m_ObjAplicacionActiva.NombreAplicacion
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Sub ComandoQuitarAplicacionActiva_Click()
    Me.AplicacionSeleccionada = Null
    Set m_ObjAplicacionActiva = Nothing
    RaiseEvent QuitarAplicacionSeleccionada
End Sub

Private Sub Form_Load()
    Me.ListaAplicaciones.RowSource = "ID;Nombre"
    If Not m_ObjAplicacionActiva Is Nothing Then
        Me.AplicacionSeleccionada = m_ObjAplicacionActiva.NombreAplicacion
    End If
End Sub

Private Sub ListaAplicaciones_Click()
    Dim m_ID As String
    On Error GoTo errores
    m_Error = ""
    Me.UsuariosConectados = Null
    m_ID = Nz(Me.ListaAplicaciones.Column(0), "")
    If Not IsNumeric(m_ID) Then
        Set m_ObjAplicacionSeleccionada = Nothing
        Set m_ObjAplicacionActiva = Nothing
        Exit Sub
    End If
    Set m_ObjAplicacionSeleccionada = Constructor.getAplicacion(m_ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjAplicacionActiva = m_ObjAplicacionSeleccionada
    Me.AplicacionSeleccionada = m_ObjAplicacionActiva.NombreAplicacion
    
    Me.UsuariosConectados = m_ObjAplicacionSeleccionada.CadenaCorreoUsuariosConectados
    RaiseEvent AplicacionSeleccionada(m_ObjAplicacionActiva)
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaAplicaciones_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub NombreAplicacion_AfterUpdate()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Nombre_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

