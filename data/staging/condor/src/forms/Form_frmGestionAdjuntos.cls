

' ==========================================================================
' MÓDULO: Form_frmGestionAdjuntos.cls
' ACCIÓN: Refactorizado para permitir Auto-Carga desde TempVars
' ==========================================================================
Option Compare Database
Option Explicit

Private m_idSolicitud As Long
Private m_adjServ As AdjuntosServicio
Private m_colAdjuntos As Collection

Private Sub Form_Load()
    Dim solServ As SolicitudServicio
    Dim sol As Solicitud
    
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    Ajustar Me
    Set m_adjServ = New AdjuntosServicio
    
    m_idSolicitud = CLng(Nz(Me.OpenArgs, 0))
    If m_idSolicitud <= 0 Then
        Err.Raise 513, "Form_Load", "No se proporcionó un ID de Solicitud válido."
    End If
    
    Set solServ = New SolicitudServicio
    Set sol = solServ.getSolicitudPorID(m_idSolicitud)
    If Not sol Is Nothing Then
        Me.Caption = "Visor de Adjuntos para: " & sol.codigoSolicitud
    Else
        Me.Caption = "Visor de Adjuntos para Solicitud ID " & m_idSolicitud
    End If
    
    With Me.lstAdjuntos
        .RowSourceType = "Value List"
        .ColumnCount = 6
        .ColumnHeads = True
        .ColumnWidths = "0cm;0cm;7,5cm;3cm;5cm;6cm"
    End With
    
    Me.cmdAnadirAdjunto.Enabled = (rolUsuario <> rol.Tecnico)
    Call CargarLista

    ' --- NUEVA LÓGICA: DETECCIÓN DE AUTO-CARGA ---
    ' Si existe la variable temporal, activamos el Timer para lanzar el diálogo
    ' un momento después de que el formulario se pinte.
    If Nz(Application.TempVars("AutoUploadEtapa"), "") <> "" Then
        Me.TimerInterval = 300 ' 300ms para asegurar que la UI cargó
    End If

LimpiarYSalir:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    DoCmd.Close acForm, Me.name
End Sub

Private Sub Form_Timer()
    ' Este evento solo se dispara si se configuró la auto-carga en el Load
    On Error Resume Next
    Me.TimerInterval = 0 ' Desactivar timer para que no se repita
    
    Dim etapaAuto As String
    etapaAuto = Application.TempVars("AutoUploadEtapa")
    
    ' Limpiamos la variable para futuras llamadas
    Application.TempVars.Remove "AutoUploadEtapa"
    
    If etapaAuto <> "" Then
        ' Lanzamos directamente el flujo de subida saltando la selección de etapa
        Call EjecutarFlujoSubida(etapaAuto, etapaAuto)
    End If
End Sub

Private Sub cmdAnadirAdjunto_Click()
    On Error GoTo Errores
    
    ' 1. Selección de Etapa (Modo Manual)
    DoCmd.OpenForm "frmElegirEtapaAdjunto", , , , , acDialog, CStr(m_idSolicitud)
    
    If Forms("frmElegirEtapaAdjunto").Tag = "OK" Then
        Dim etapa As String, nombreEtapa As String, descripcion As String
        etapa = Forms("frmElegirEtapaAdjunto").EtapaSeleccionada
        nombreEtapa = Forms("frmElegirEtapaAdjunto").EtapaNombreSeleccionado
        descripcion = Forms("frmElegirEtapaAdjunto").DescripcionAdjunto
        DoCmd.Close acForm, "frmElegirEtapaAdjunto"
        
        ' Llamamos al procedimiento común
        Call EjecutarFlujoSubida(etapa, nombreEtapa, descripcion)
    Else
        DoCmd.Close acForm, "frmElegirEtapaAdjunto"
    End If
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdAnadirAdjunto_Click", Me)
End Sub

' EN: Form_frmGestionAdjuntos.cls

Private Sub EjecutarFlujoSubida(ByVal etapa As String, ByVal nombreEtapa As String, Optional ByVal descripcion As String = "")
    Dim fDialog As FileDialog
    Dim filePath As String
    
    ' Variables para validación de negocio
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim estaAprobadoEnBD As Boolean
    
    On Error GoTo Errores
    
    ' GAP 26.3: Se permite adjuntar el Documento Final Firmado sin exigir decisión previa
    ' ----------------------------------------
    
    ' 1. Selección de Archivo
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.Title = "Seleccione el documento para '" & nombreEtapa & "'"
    fDialog.AllowMultiSelect = False
    fDialog.Filters.Clear
    
    If nombreEtapa = "Cierre" Then
        fDialog.Filters.Add "Documentos PDF", "*.pdf"
    Else
        fDialog.Filters.Add "Todos los archivos", "*.*"
        fDialog.Filters.Add "Documentos PDF", "*.pdf"
        fDialog.Filters.Add "Documentos Word", "*.docx; *.doc"
    End If
    fDialog.FilterIndex = 1
    
    If fDialog.Show = -1 Then
        filePath = fDialog.SelectedItems(1)
        Me.lblBloqueo.Visible = True
        DoCmd.Hourglass True
        
        ' --- CONTROL DE UNICIDAD Y SUSTITUCIÓN ---
        If nombreEtapa = "Documento Final Firmado" Then
            Dim colExistentes As Collection
            Dim adjExistente As AdjuntoViewModel
            Dim idExistente As Long
            
            Set colExistentes = m_adjServ.getAdjuntosViewModelPorSolicitud(m_idSolicitud)
            idExistente = 0
            
            If Not colExistentes Is Nothing Then
                For Each adjExistente In colExistentes
                    If adjExistente.etapaNombre = "Documento Final Firmado" Then
                        idExistente = adjExistente.idAdjunto
                        Exit For
                    End If
                Next adjExistente
            End If
            
            If idExistente > 0 Then
                Me.lblBloqueo.Visible = False
                If MsgBox("Ya existe un 'Documento Final Firmado'. ¿Desea sustituirlo?", vbExclamation + vbYesNo + vbDefaultButton2, "CONDOR") = vbNo Then
                    GoTo LimpiarYSalir
                End If
                Me.lblBloqueo.Visible = True
                Call m_adjServ.EliminarAdjunto(idExistente)
            End If
        End If
        
        ' 2. Subir adjunto + callback atómica de aprobación si es "Documento Final Firmado"
        ' Slice A5 (2026-06-15): la callback atómica se extrajo a AdjuntosServicio.SubirYCerrar.
        Dim idAdjuntoOut As Long
        Call m_adjServ.SubirYCerrar(idSolicitud:=m_idSolicitud, etapaNombre:=etapa, rutaArchivoOrigen:=filePath, tipoAccion:="Adjuntado", descripcion:=descripcion, idAdjuntoOut:=idAdjuntoOut)
        
        ' 3. UI feedback
        If nombreEtapa = "Documento Final Firmado" Then
            MsgBox "Documento final adjuntado y solicitud aprobada.", vbInformation, "Condor"
            If FormularioAbierto("frmGestionSolicitud") Then
                Forms("frmGestionSolicitud").NotificarCambioExterno
            End If
        Else
            MsgBox "Archivo adjuntado correctamente.", vbInformation, "Condor"
        End If
        
        Call CargarLista
    End If

LimpiarYSalir:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set fDialog = Nothing
    Set solServ = Nothing: Set pcServ = Nothing: Set cdcaServ = Nothing
    Exit Sub
Errores:
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Call ManejadorErroresFormulario("EjecutarFlujoSubida", Me)
End Sub

Private Sub CargarLista()
    On Error GoTo Errores
    Set m_colAdjuntos = m_adjServ.getAdjuntosViewModelPorSolicitud(m_idSolicitud)
    Dim rowSource As String
    rowSource = "idAdjunto;Ruta;Nombre;Tipo;Etapa;Descripción"
    Dim vm As AdjuntoViewModel
    If Not m_colAdjuntos Is Nothing Then
        For Each vm In m_colAdjuntos
            rowSource = rowSource & ";" & vm.idAdjunto & ";" & m_ObjEntorno.URLDirectorioDocumentacion & _
                vm.nombreArchivo & ";" & vm.nombreArchivo & ";" & Nz(vm.tipoAccion, "N/A") & ";" & vm.etapaNombre & ";" & vm.descripcion
        Next vm
    End If
    Me.lstAdjuntos.rowSource = rowSource
    Me.cmdAbrirAdjunto.Enabled = False
    Me.cmdEliminarAdjunto.Enabled = False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("CargarLista", Me)
End Sub

Private Sub lstAdjuntos_Click()
    Dim haySeleccion As Boolean
    Dim tipoAccion As String
    
    If IsNumeric(Me.lstAdjuntos.Column(0)) Then haySeleccion = True
    Me.cmdAbrirAdjunto.Enabled = haySeleccion
    
    If haySeleccion Then
        tipoAccion = Nz(Me.lstAdjuntos.Column(3), "")
    Else
        tipoAccion = ""
    End If
    
    Me.cmdEliminarAdjunto.Enabled = (haySeleccion And (rolUsuario <> rol.Tecnico) And (tipoAccion <> "Generado"))
End Sub

' EN: Form_frmGestionAdjuntos.cls

Private Sub cmdEliminarAdjunto_Click()
    Dim idAdjunto As Long
    Dim etapaAdjunto As String
    Dim sol As Solicitud
    Dim solServ As New SolicitudServicio
    Dim wfServ As New WorkflowServicio
    Dim pcServ As DatosPCServicio
    Dim cdcaServ As DatosCDCAServicio
    Dim mensajeConfirmacion As String
    Dim esDocumentoCierre As Boolean
    
    On Error GoTo Errores
    
    ' 1. Validar selección
    If Me.lstAdjuntos.ListIndex < 0 Then Exit Sub
    idAdjunto = CLng(Nz(Me.lstAdjuntos.Column(0), 0))
    etapaAdjunto = Nz(Me.lstAdjuntos.Column(4), "") ' Asumiendo que la columna 4 es la Etapa
    
    If idAdjunto = 0 Then Exit Sub
    
    ' 2. Obtener estado actual de la solicitud
    Set sol = solServ.getSolicitudPorID(m_idSolicitud)
    
    ' Detectar si es el documento crítico
    esDocumentoCierre = (etapaAdjunto = "Documento Final Firmado" And sol.idEstadoInterno = estadoAprobada)
    
    ' 3. Configurar el mensaje de advertencia
    If esDocumentoCierre Then
        mensajeConfirmacion = "?? ADVERTENCIA DE SEGURIDAD ??" & vbCrLf & vbCrLf & _
                              "Va a eliminar el documento que certifica el cierre de esta solicitud." & vbCrLf & vbCrLf & _
                              "CONSECUENCIAS:" & vbCrLf & _
                              "1. Se borrarán los datos de la pestaña 'Decisión Final' (Autoridad Competente)." & vbCrLf & _
                              "2. La solicitud dejará de estar APROBADA." & vbCrLf & _
                              "3. El estado retrocederá a 'FORMALIZACIÓN'." & vbCrLf & vbCrLf & _
                              "¿Está seguro de que desea continuar?"
    Else
        mensajeConfirmacion = "¿Está seguro de que desea eliminar el archivo adjunto seleccionado?"
    End If
    
    ' 4. Confirmación del usuario
    If MsgBox(mensajeConfirmacion, vbExclamation + vbYesNo + vbDefaultButton2, "CONDOR") = vbNo Then Exit Sub
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 5. Eliminar el adjunto (Físico y BD)
    Call m_adjServ.EliminarAdjunto(idAdjunto)
    
    ' 6. Acciones colaterales (Solo si era el documento de cierre y estaba aprobada)
    If esDocumentoCierre Then
        
        ' A. Borrar datos de la decisión final según el tipo
        Select Case sol.tipoSolicitud
            Case "PC"
                Set pcServ = New DatosPCServicio
                Call pcServ.EliminarDecisionFinal(m_idSolicitud)
            Case "CD_CA", "CD_CA_SUB"
                Set cdcaServ = New DatosCDCAServicio
                Call cdcaServ.EliminarDecisionFinal(m_idSolicitud)
            Case "PC_SUB"
                Dim pcsubServDel As DatosPCSUBServicio
                Set pcsubServDel = New DatosPCSUBServicio
                Call pcsubServDel.EliminarDecisionFinal(m_idSolicitud)
        End Select
        
        ' B. Revertir el estado a FORMALIZACIÓN
        ' Usamos el método de reversión que ya gestiona logs y notificaciones
        Call wfServ.RevertirAFaseAnterior_FromUI(sol, estadoFormalizacion, m_ObjUsuarioActivo, "Eliminación del Documento Final Firmado")
        
        MsgBox "El adjunto se ha eliminado. La solicitud ha vuelto al estado 'FORMALIZACIÓN'.", vbInformation, "CONDOR"
        
        ' C. Notificar al formulario padre para que se repinte (bloquee/desbloquee botones)
        If FormularioAbierto("frmGestionSolicitud") Then
            Forms("frmGestionSolicitud").NotificarCambioExterno
        End If
        
    End If
    
    ' 7. Refrescar la lista de adjuntos
    Call CargarLista
    
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set solServ = Nothing
    Set wfServ = Nothing
    Set pcServ = Nothing
    Set cdcaServ = Nothing
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("cmdEliminarAdjunto_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdAbrirAdjunto_Click()
    Dim ruta As String
    ruta = Nz(Me.lstAdjuntos.Column(1), "")
    If ruta <> "" Then Call AbrirEnLocal(ruta)
End Sub

Private Sub lstAdjuntos_DblClick(Cancel As Integer)
    If Me.cmdAbrirAdjunto.Enabled Then cmdAbrirAdjunto_Click
End Sub

Private Sub cmdCerrar_Click()
    DoCmd.Close acForm, Me.name
End Sub



