Option Compare Database
Option Explicit

Private m_Error As String
Private fso As New FileSystemObject
Private m_TipoAnexo As EnumTipoAnexo
Private m_ID As String
Private m_URLAnexo As String
Private m_URLParteEvento As String
Private m_IDAnexoElegido As String

Private lst As ListBox
Private m_Escritura As String
Private m_FormularioQueAbre As String
Private m_ObjAnexo As Anexo
Private m_ObjAnexoSeleccionado As Anexo


Private Sub ComandoAnexar_Click()
    
    Dim m_URL As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    If Nz(Me.TituloDoc, "") = "" Then
        m_Error = "El título es obligatorio"
        Err.Raise 1000
    End If
    Set m_ObjAnexo = New Anexo
    AsignarTipoAnexo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_ObjAnexo.Titulo = Nz(Me.TituloDoc, "")
    Me.ComandoEliminarDocumento.Enabled = False
    Me.ComandoAnexo.Enabled = False
    Me.ComandoParte.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    m_URL = Seleccionar(True, "Seleccione el archivo a anexar")
    If Not fso.FileExists(m_URL) Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjAnexo.Registrar m_URL, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerLista EnumSino.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormEventoGestion") Then
        If m_ObjAnexo.TipoAnexo = EnumTipoAnexo.Evento Then
            Form_FormEventoGestion.ListaEventos_Click
        ElseIf m_ObjAnexo.TipoAnexo = EnumTipoAnexo.Actividad Then
            Form_FormEventoGestion.ListaActividades_Click
        ElseIf m_ObjAnexo.TipoAnexo = EnumTipoAnexo.Material Then
            Form_FormEventoGestion.ListaMateriales_Click
        End If
    End If
        
    If FormularioAbierto("FormActividadGestion") Then
        Form_FormActividadGestion.ListaFiltrados_Click
    End If
    Me.TituloDoc = Null
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub cmdSalir_Click()
   
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If FormularioAbierto(m_FormularioQueAbre) Then
        DoCmd.Close acForm, m_FormularioQueAbre, acSaveNo
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoAnexo_Click()
    
    Dim m_URL As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_ObjAnexoSeleccionado Is Nothing Then
        If Nz(Me.ListaDocumentos.Column(0), "") <> "" Then
            Set m_ObjAnexoSeleccionado = Constructor.getAnexo(Me.ListaDocumentos.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    m_URL = m_ObjAnexoSeleccionado.URLAnexo
    m_Error = m_ObjAnexoSeleccionado.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URL) Then
        m_Error = "Parece que el archivo no se pudo subir en su momento. Tiene que resubirlo"
        Err.Raise 1000
    End If
    flag = AbrirEnLocal(m_URL, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarDocumento_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    
    m_Error = ""
    If m_ObjAnexoSeleccionado Is Nothing Then
        If Nz(Me.ListaDocumentos.Column(0), "") <> "" Then
            Set m_ObjAnexoSeleccionado = Constructor.getAnexo(Me.ListaDocumentos.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_ObjAnexoSeleccionado Is Nothing Then
            m_Error = "Seleccione un documento de la lista"
            Err.Raise 1000
        End If
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente eliminar el anexo llamado : " & m_ObjAnexoSeleccionado.NombreArchivo & _
                        " definitivamente, tanto del disco duro como del registro de la base de datos?", _
                        vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de anexo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjAnexoSeleccionado.Eliminar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjAnexoSeleccionado = Nothing
    EstablecerLista EnumSino.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormEventoGestion") Then
        If m_ObjAnexo.TipoAnexo = EnumTipoAnexo.Evento Then
            Form_FormEventoGestion.ListaEventos_Click
        ElseIf m_ObjAnexo.TipoAnexo = EnumTipoAnexo.Actividad Then
            Form_FormEventoGestion.ListaActividades_Click
        ElseIf m_ObjAnexo.TipoAnexo = EnumTipoAnexo.Material Then
            Form_FormEventoGestion.ListaMateriales_Click
        End If
    End If
        
    If FormularioAbierto("FormActividadGestion") Then
        Form_FormActividadGestion.ListaFiltrados_Click
    End If
    Me.TituloDoc = Null
    Me.ListaDocumentos.Requery
    Me.ComandoAnexo.Enabled = False
    Me.ComandoParte.Enabled = False
    Me.ComandoEliminarDocumento.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarDocumento_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoParte_Click()
    
    Dim m_Parte As Parte
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not TypeOf m_ObjEntidadParaAnexoActiva Is Evento Then
        m_Error = "La entidad del anexo no es un evento"
        Err.Raise 1000
    End If
    Set m_Parte = m_ObjEntidadParaAnexoActiva.Parte
    m_Error = m_ObjEntidadParaAnexoActiva.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Parte Is Nothing Then
        m_Error = "El evento no tiene parte registrado"
        Err.Raise 1000
    End If
    
    flag = AbrirEnLocal(m_Parte.URLParte, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoParte_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If m_ObjAnexoSeleccionado Is Nothing Then
        If Nz(Me.ListaDocumentos.Column(0), "") <> "" Then
            Set m_ObjAnexoSeleccionado = Constructor.getAnexo(Me.ListaDocumentos.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    m_URL = Seleccionar(True, "Seleccione el archivo a Resubir")
    If Not fso.FileExists(m_URL) Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjAnexoSeleccionado.Resubir m_URL, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerLista EnumSino.Sí, m_Error
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

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    
    AjustarTamaño Me
    m_Error = ""
    
    m_FormularioQueAbre = Nz(OpenArgs, "")
    Me.Caption = Me.lblTitulo1.Caption
    If m_ObjEntidadParaAnexoActiva Is Nothing Then
        m_Error = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
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
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_Parte As Parte
    On Error GoTo errores
    
    p_Error = ""
    Me.ComandoAnexo.Enabled = False
    Me.ComandoParte.Enabled = False
    Me.ComandoEliminarDocumento.Enabled = False
    Me.IDActividad.Enabled = True
    Me.IDActividad.Locked = True
    Me.IDEVENTO.Enabled = True
    Me.IDEVENTO.Locked = True
    Me.IDMaterial.Enabled = True
    Me.IDMaterial.Locked = True
    Me.IDSubContratacion.Enabled = True
    Me.IDSubContratacion.Locked = True
    Me.IDGasto.Enabled = True
    Me.IDGasto.Locked = True
    Me.IDMaterialSeguimiento.Enabled = True
    Me.IDMaterialSeguimiento.Locked = True
    
    Set m_ObjAnexo = New Anexo
    If TypeOf m_ObjEntidadParaAnexoActiva Is Evento Then
        Me.IDEVENTO = m_ObjEntidadParaAnexoActiva.IDEVENTO
        m_ObjAnexo.IDEVENTO = m_ObjEntidadParaAnexoActiva.IDEVENTO
        Set m_ObjAnexo.Evento = m_ObjEntidadParaAnexoActiva
        Set m_Parte = m_ObjEntidadParaAnexoActiva.Parte
        p_Error = m_ObjEntidadParaAnexoActiva.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_Parte Is Nothing Then
            If m_Parte.URLParte <> "" Then
                Me.ComandoParte.Enabled = True
            End If
        End If
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Actividad Then
        Me.IDActividad = m_ObjEntidadParaAnexoActiva.IDActividad
        m_ObjAnexo.IDActividad = m_ObjEntidadParaAnexoActiva.IDActividad
        Set m_ObjAnexo.Actividad = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Material Then
        Me.IDMaterial = m_ObjEntidadParaAnexoActiva.IDMaterial
        m_ObjAnexo.IDMaterial = m_ObjEntidadParaAnexoActiva.IDMaterial
        Set m_ObjAnexo.Material = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is MaterialSeguimiento Then
        Me.IDMaterialSeguimiento = m_ObjEntidadParaAnexoActiva.IDMaterialSeguimiento
        m_ObjAnexo.IDMaterialSeguimiento = m_ObjEntidadParaAnexoActiva.IDMaterialSeguimiento
        Set m_ObjAnexo.MaterialSeg = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Subcontratacion Then
        Me.IDSubContratacion = m_ObjEntidadParaAnexoActiva.IDSubContratacion
        m_ObjAnexo.IDSubContratacion = m_ObjEntidadParaAnexoActiva.IDSubContratacion
        Set m_ObjAnexo.Subcontratacion = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Gasto Then
        Me.IDGasto = m_ObjEntidadParaAnexoActiva.IDGasto
        m_ObjAnexo.IDGasto = m_ObjEntidadParaAnexoActiva.IDGasto
        Set m_ObjAnexo.Gasto = m_ObjEntidadParaAnexoActiva
    Else
        p_Error = "Tipo de Anexo no reconocido"
        Err.Raise 1000
    End If
    
    
    EstablecerLista EnumSino.No, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    EstablecerDatos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function AsignarTipoAnexo(Optional ByRef p_Error As String) As String
    
    Dim m_Parte As Parte
    On Error GoTo errores
    
    p_Error = ""
   
    
    Set m_ObjAnexo = New Anexo
    If TypeOf m_ObjEntidadParaAnexoActiva Is Evento Then
        Me.IDEVENTO = m_ObjEntidadParaAnexoActiva.IDEVENTO
        m_ObjAnexo.IDEVENTO = m_ObjEntidadParaAnexoActiva.IDEVENTO
        Set m_ObjAnexo.Evento = m_ObjEntidadParaAnexoActiva
        
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Actividad Then
        Me.IDActividad = m_ObjEntidadParaAnexoActiva.IDActividad
        m_ObjAnexo.IDActividad = m_ObjEntidadParaAnexoActiva.IDActividad
        Set m_ObjAnexo.Actividad = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Material Then
        Me.IDMaterial = m_ObjEntidadParaAnexoActiva.IDMaterial
        m_ObjAnexo.IDMaterial = m_ObjEntidadParaAnexoActiva.IDMaterial
        Set m_ObjAnexo.Material = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is MaterialSeguimiento Then
        Me.IDMaterialSeguimiento = m_ObjEntidadParaAnexoActiva.IDMaterialSeguimiento
        m_ObjAnexo.IDMaterialSeguimiento = m_ObjEntidadParaAnexoActiva.IDMaterialSeguimiento
        Set m_ObjAnexo.MaterialSeg = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Subcontratacion Then
        Me.IDSubContratacion = m_ObjEntidadParaAnexoActiva.IDSubContratacion
        m_ObjAnexo.IDSubContratacion = m_ObjEntidadParaAnexoActiva.IDSubContratacion
        Set m_ObjAnexo.Subcontratacion = m_ObjEntidadParaAnexoActiva
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Gasto Then
        Me.IDGasto = m_ObjEntidadParaAnexoActiva.IDGasto
        m_ObjAnexo.IDGasto = m_ObjEntidadParaAnexoActiva.IDGasto
        Set m_ObjAnexo.Gasto = m_ObjEntidadParaAnexoActiva
    Else
        p_Error = "Tipo de Anexo no reconocido"
        Err.Raise 1000
    End If
    
    
    AsignarTipoAnexo = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AsignarTipoAnexo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Sub ListaDocumentos_Click()
    
    Dim m_ID As String
    On Error GoTo errores
    
    m_Error = ""
    Me.ComandoResubir.Enabled = False
    Me.ComandoEliminarDocumento.Enabled = False
    Me.ComandoAnexo.Enabled = False
    m_ID = Nz(Me.ListaDocumentos.Column(0), "")
    If m_ID = "" Then
        Set m_ObjAnexoSeleccionado = Nothing
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjAnexoSeleccionado Is Nothing Then
        Set m_ObjAnexoSeleccionado = Constructor.getAnexo(m_ID, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjAnexoSeleccionado.IDAnexo <> m_ID Then
            Set m_ObjAnexoSeleccionado = Constructor.getAnexo(m_ID, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjAnexoSeleccionado Is Nothing Then
        Exit Sub
    End If
    Me.ComandoEliminarDocumento.Enabled = True
    Me.ComandoAnexo.Enabled = True
    If Not fso.FileExists(m_ObjAnexoSeleccionado.URLAnexo) Then
        Me.ComandoResubir.Enabled = True
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaDocumentos_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Function ActualizarEntidad( _
                                    Optional ByRef p_Error As String) As String
    
   
    
    On Error GoTo errores
    
   If TypeOf m_ObjEntidadParaAnexoActiva Is Evento Then
        Set m_ObjEntidadParaAnexoActiva = Constructor.getEvento(m_ObjEntidadParaAnexoActiva.IDEVENTO, p_Error)
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Actividad Then
        Set m_ObjEntidadParaAnexoActiva = Constructor.getActividad(m_ObjEntidadParaAnexoActiva.IDActividad, p_Error)
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Material Then
        Set m_ObjEntidadParaAnexoActiva = Constructor.getMaterial(m_ObjEntidadParaAnexoActiva.IDMaterial, p_Error)
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is MaterialSeguimiento Then
        Set m_ObjEntidadParaAnexoActiva = Constructor.getMaterialSeguimiento(m_ObjEntidadParaAnexoActiva.IDMaterialSeguimiento, p_Error)
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Subcontratacion Then
        Set m_ObjEntidadParaAnexoActiva = Constructor.getSubContratacion(m_ObjEntidadParaAnexoActiva.IDSubContratacion, p_Error)
    ElseIf TypeOf m_ObjEntidadParaAnexoActiva Is Gasto Then
        Set m_ObjEntidadParaAnexoActiva = Constructor.getGasto(m_ObjEntidadParaAnexoActiva.IDGasto, p_Error)
    End If
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarEntidad se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Function EstablecerLista( _
                                    Optional p_Actualizando As EnumSino = EnumSino.No, _
                                    Optional ByRef p_Error As String) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim lst As ListBox
    Dim m_ObjAnexoEnLista As Anexo
    
    On Error GoTo errores
    
    p_Error = ""
    If p_Actualizando = EnumSino.Sí Then
        ActualizarEntidad p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set lst = Me.ListaDocumentos
    lst.RowSource = "ID;IDAnexo;Título;Archivo"
    Set m_Col = m_ObjEntidadParaAnexoActiva.ColAnexos
    p_Error = m_ObjEntidadParaAnexoActiva.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjAnexoEnLista = m_Col(m_ID)
        With m_ObjAnexoEnLista
            lst.AddItem .IDAnexo & ";" & .Titulo & ";" & .NombreArchivo
        End With
        Set m_ObjAnexoEnLista = Nothing
    Next
    
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaDocumentos_Click
    
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


