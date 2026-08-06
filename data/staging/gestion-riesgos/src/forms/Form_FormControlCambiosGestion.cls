Option Compare Database
Option Explicit
Private m_VersionSeleccionada As CCVersion
Private m_Error As String

Private Function RellenarComboVersiones(Optional ByRef p_Error As String) As String
    Dim cmb As ComboBox
    Dim p_Version As CCVersion
    Dim m_ID As Variant
    
    On Error GoTo errores
    Set cmb = Me.Version
    cmb.RowSource = ""
    If m_ObjEntorno.VersionesAplicacion Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjEntorno.VersionesAplicacion
        Set p_Version = m_ObjEntorno.VersionesAplicacion(m_ID)
        cmb.AddItem p_Version.Version
        Set p_Version = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método RellenarComboVersiones ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoLimpiarVersionPrincipal_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Nz(Me.VersionPrincipal, "") <> "" Then
        Me.VersionPrincipal = Null
        Filtrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiarVersionPrincipal_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoLimpiaVersion_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Nz(Me.Version, "") <> "" Then
        Me.Version = Null
        Filtrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiaVersion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoLimpiarCambio_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Nz(Me.Cambio, "") <> "" Then
        Me.Cambio = Null
        Filtrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiarCambio_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub




Private Sub ComandoInforme_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_VersionSeleccionada Is Nothing Then
        Set m_VersionSeleccionada = Constructor.getCCVersion(Nz(Me.ListaFiltrados.Column(0), ""), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_VersionSeleccionada Is Nothing Then
            m_Error = "Se ha de seleccionar un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    MostrarControlCambiosVersion m_VersionSeleccionada, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoInforme_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub VersionPrincipal_BeforeUpdate(Cancel As Integer)
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
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al VersionPrincipal_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Version_BeforeUpdate(Cancel As Integer)
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
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Version_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If Me.ComandoInforme.Enabled = True Then
        ComandoInforme_Click
    End If
    Application.Echo True
End Sub

Public Sub Cambio_Exit(Cancel As Integer)
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
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Cambio_Exit se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me
    m_Error = ""
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    RellenarComboVersiones m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub



Public Function Filtrar( _
                        Optional p_Error As String _
                        ) As String
    
    
    
    Dim lst As ListBox
    Dim m_Versiones As Scripting.Dictionary
    Dim m_IDVersion As Variant
    Dim m_Version As CCVersion
    Dim m_IDCambio As Variant
    Dim m_Cambio As CCCambio
    Dim m_Cambios As Scripting.Dictionary
    On Error GoTo errores
    
    m_Error = ""
    
    Set m_VersionSeleccionada = Nothing
    Set lst = Me.ListaFiltrados
    
    lst.RowSource = "IDVersion;Version;Descripcion"
    Set m_Versiones = m_ObjEntorno.ColVersiones
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Versiones Is Nothing Then
        Exit Function
    End If
    
    For Each m_IDVersion In m_Versiones
        Set m_Version = m_Versiones(m_IDVersion)
        With m_Version
            If Nz(Me.VersionPrincipal, "") <> "" Then
                If .VersionPrincipal <> Me.VersionPrincipal Then
                    GoTo siguiente
                End If
            End If
            If Nz(Me.Version, "") <> "" Then
                If .Version <> Me.Version Then
                    GoTo siguiente
                End If
            End If
            If Nz(Me.Cambio, "") <> "" Then
                Set m_Cambios = .Cambios
                If m_Cambios Is Nothing Then
                    GoTo siguiente
                End If
                For Each m_IDCambio In m_Cambios
                    Set m_Cambio = m_Cambios(m_IDCambio)
                    If InStr(1, m_Cambio.NombreCambio, Me.Cambio) <> 0 Then
                        Set m_Cambio = Nothing
                        GoTo escribe
                    End If
                    
                    Set m_Cambio = Nothing
                Next
                GoTo siguiente
            End If
escribe:
            lst.AddItem .IDVersion & ";" & .Version & ";" & .Descripcion
        End With
siguiente:
        
    

    Next
    If lst.ListCount > 1 Then
        lst.Selected(1) = True
        Set m_VersionSeleccionada = m_Version
        Me.ComandoInforme.Enabled = True
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Sub ListaFiltrados_Click()
    
    
    On Error GoTo errores
    
    m_Error = ""
    Application.Echo False
    Me.ComandoInforme.Enabled = False
    If Nz(Me.ListaFiltrados, "") = "" Then
        Set m_VersionSeleccionada = Nothing
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_VersionSeleccionada Is Nothing Then
        Set m_VersionSeleccionada = Constructor.getCCVersion(Nz(Me.ListaFiltrados.Column(0), ""), , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_VersionSeleccionada.IDVersion <> Me.ListaFiltrados.Column(0) Then
            Set m_VersionSeleccionada = Constructor.getCCVersion(Nz(Me.ListaFiltrados.Column(0), ""), , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_VersionSeleccionada Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoInforme.Enabled = True
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Application.Echo True
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    Application.Echo True
End Sub
