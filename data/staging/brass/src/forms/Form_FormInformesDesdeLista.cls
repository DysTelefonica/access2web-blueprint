Option Compare Database
Option Explicit
Private m_Error As String
Private fso As New FileSystemObject
Private m_ObjInforme As Informe
Private m_ParaRegularizar As Boolean
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub ComandoAbrirLista_Click()
    DoCmd.OpenTable "TbAuxEventosParaInforme"
End Sub


Private Sub ComandoDetalleEvento_Click()
    Dim m_frm As Form
    On Error Resume Next
    If Me.ListaFiltrados.ItemsSelected.count = 0 Then
        Exit Sub
    End If
    If Me.ListaFiltrados.ItemsSelected.count > 1 Then
        Exit Sub
    End If
    
    If Not FormularioAbierto("FormEventoGestion") Then
        DoCmd.OpenForm "FormEventoGestion"
    End If
    Set m_frm = Forms("FormEventoGestion")
    With m_frm
        .Controls("FechaInicialBusqueda") = Null
        .Controls("FechaFinalBusqueda") = Null
        .Controls("ComboTipoEventoBusqueda") = Null
        .Controls("ComboBUIBusqueda") = Null
        .Controls("ComboSubsistemaBusqueda") = Null
        .Controls("ComboIDEquipoBusqueda") = Null
        .Controls("ComboIDEventoBusqueda") = Nz(Me.ListaFiltrados.Column(0), "")
    End With
    Form_FormEventoGestion.ActualizarEventos
    Forms("FormEventoGestion").SetFocus
    Set m_frm = Nothing
End Sub

Private Sub ComandoEliminar_Click()
    
    Dim m_IDEventoSeleccionado As Variant
    
    On Error GoTo errores
    If m_ColEventosParaInformeDeRAC Is Nothing Then
        Exit Sub
    End If
    If Me.ListaFiltrados.ItemsSelected.count = 0 Then
        m_Error = "Seleccione al menos un registro"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea sacar el/los Parte/es de trabajo del informe?", vbExclamation + vbYesNo + vbDefaultButton2, "Quitar Parte de Trabajo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    For Each m_IDEventoSeleccionado In Me.ListaFiltrados.ItemsSelected
        m_IDEventoSeleccionado = Me.ListaFiltrados.Column(0, m_IDEventoSeleccionado)
        If m_ColEventosParaInformeDeRAC.Exists(m_IDEventoSeleccionado) Then
            m_ColEventosParaInformeDeRAC.Remove (m_IDEventoSeleccionado)
        End If
    Next
    RellenarLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar quitar los partes de trabajo del informe se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoExportarExcel_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    '------------------------------------------
    '   -Devuelve:
    '       GenerarConsultas = Descriptivo
    '       GenerarConsultas = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = GenerarConsultas("Eventos para Informe", Me.hWnd, , Me.ListaFiltrados, "No")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método GenerarConsultas ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


Public Function RellenarLista(Optional ByRef p_Error As String) As String
    
    Dim m_list As ListBox
    Dim m_IDEvento As Variant
    Dim m_ObjEvento As Evento
    On Error GoTo errores
    
    Set m_list = Me.ListaFiltrados
    m_list.RowSource = "P.TRABAJO;DESCRIPCIÓN;FINICIAL;FFINAL"
    Me.ComandoExportarExcel.Enabled = False
    Me.ComandoGenerarInforme.Enabled = False
    
    If m_ColEventosParaInformeDeRAC Is Nothing Then
        Exit Function
    End If
    If m_ColEventosParaInformeDeRAC.count = 0 Then
        Exit Function
    End If
    For Each m_IDEvento In m_ColEventosParaInformeDeRAC.Keys
        Set m_ObjEvento = m_ColEventosParaInformeDeRAC(m_IDEvento)
        If Not m_ObjEvento Is Nothing Then
            m_list.AddItem m_IDEvento & ";" & m_ObjEvento.Descripcion & ";" & m_ObjEvento.FECHAALTAEVENTO & ";" & m_ObjEvento.FechaFinal
        Else
           ' Stop
        End If
        Set m_ObjEvento = Nothing
    Next
    m_list.Requery
    
    Me.ComandoExportarExcel.Enabled = True
    Me.ComandoGenerarInforme.Enabled = True
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarLista ha devuelto el error: " & Err.Description
    End If
End Function

Private Sub ComandoGenerarInforme_Click()
    
    
    Dim m_URLInforme As String
    Dim m_Alcanzable As Boolean
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Set m_ObjInforme = New Informe
    m_ObjInforme.EventosRAC m_ColEventosParaInformeDeRAC, m_URLInforme, True, m_Error
    Set m_ObjInforme = Nothing
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLInforme) Then
        m_Error = "No se ha podido obtener la URL del informe"
        Err.Raise 1000
    End If
   
    Ejecutar Me.hWnd, "open", m_URLInforme, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar generar el informe se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoLlenarDesdeLista_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComandoEliminar.Enabled = False
    Me.ComandoDetalleEvento.Enabled = False
    Me.ComandoExportarExcel.Enabled = False
    
    
    Set m_ColEventosParaInformeDeRAC = getListaEventosDeTabla(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    RellenarLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar obtener los eventos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me



    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    Me.SetFocus
    VBA.DoEvents
    Me.Caption = Me.lblTitulo1.Caption
    Set m_ColEventosParaInformeDeRAC = New Scripting.Dictionary
    Me.ListaFiltrados.RowSource = "P.TRABAJO;DESCRIPCIÓN;FINICIAL;FFINAL"
    Me.lblEstado.Visible = False
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ListaFiltrados_Click()
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Me.ComandoEliminar.Enabled = False
        Me.ComandoDetalleEvento.Enabled = False
        
    Else
        Me.ComandoEliminar.Enabled = True
        Me.ComandoDetalleEvento.Enabled = True
        
    End If
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If Me.ComandoDetalleEvento.Enabled Then
        ComandoDetalleEvento_Click
    End If
    Application.Echo True
End Sub
