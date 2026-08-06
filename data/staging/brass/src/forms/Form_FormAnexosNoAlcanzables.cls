Option Compare Database
Option Explicit

Private fso As New FileSystemObject
Private m_AnexoSeleccionado As Anexo

Private m_Error As String

Public Function Filtrar(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Anexo As Anexo
    Dim m_Evento As Evento
    Dim m_Actividad As Actividad
    Dim m_Material As Material
    Dim m_BUI As String
    
    On Error GoTo errores
    Set lst = Me.ListaFiltrados
    lst.RowSource = "IDAnexo;BUI;IDEvento;IDMaterial;Archivo"
    Set m_Col = Constructor.getAnexosNoAlcanzablesBusqueda( _
                        Nz(Me.ComboIDEventoBusqueda.Column(0), ""), _
                        Nz(Me.ComboBUIBusqueda, ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Anexo = m_Col(m_ID)
        Set m_Evento = m_Anexo.Evento
        If m_Anexo.IDEVENTO <> "" Then
            m_BUI = m_Anexo.Evento.BUI
        Else
            If m_Anexo.IDActividad <> "" Then
                m_BUI = m_Anexo.Actividad.Evento.BUI
            Else
                If m_Anexo.IDMaterial <> "" Then
                    m_BUI = m_Anexo.Material.Actividad.Evento.BUI
                End If
            End If
        End If
        lst.AddItem m_Anexo.IDAnexo & ";" & m_BUI & ";" & m_Anexo.IDEVENTO & ";" & m_Anexo.IDMaterial & ";" & m_Anexo.NombreArchivo
         
        m_BUI = ""
siguiente:
        Set m_Anexo = Nothing
    Next
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function



Private Sub ComandoFiltrar_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
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
        m_Error = "Al ComboIDEventoBusqueda_Change se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoResubir_Click()
    
    Dim m_URL As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_AnexoSeleccionado Is Nothing Then
        If Nz(Me.ListaFiltrados.Column(0), "") <> "" Then
            Set m_AnexoSeleccionado = Constructor.getAnexo(Me.ListaFiltrados.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If fso.FileExists(m_AnexoSeleccionado) Then
        m_Error = "El archivo ya es alcanzable"
        Err.Raise 1000
    End If
    m_URL = Seleccionar(True, "Seleccione el archivo a Resubir")
    If Not fso.FileExists(m_URL) Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_AnexoSeleccionado.Resubir m_URL, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoResubir_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    
    Me.lblEstado.Visible = False
    Me.Caption = Me.lblTitulo1.Caption
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

Private Sub ListaFiltrados_Click()
    
    Dim m_ID As String
   
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComandoResubir.Enabled = False
    
    
    m_ID = Nz(Me.ListaFiltrados.Column(0), "")
    If m_ID = "" Then
        Set m_AnexoSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_AnexoSeleccionado Is Nothing Then
        Set m_AnexoSeleccionado = Constructor.getAnexo(m_ID, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ID <> m_AnexoSeleccionado.IDAnexo Then
            Set m_AnexoSeleccionado = Constructor.getAnexo(m_ID, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_AnexoSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoResubir.Enabled = True
    
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Clicka se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

