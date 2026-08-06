Attribute VB_Name = "Form_FormUsuarioHistorico02Edicion"
Option Compare Database
Option Explicit
Private m_Error As String
Private m_ObjObservacionSeleccionada As ObservacionHistorica
Public Sub Form_Load()
    
    Dim ctl As Control
    On Error GoTo errores
    
    m_Error = ""
    
    ' --- CAMBIO CRÍTICO ---
    ' Ponemos AllowEdits = True para que Access permita hacer CLIC en las listas.
    ' La protección de datos la haremos abajo bloqueando control a control.
    Me.AllowEdits = True
    
    If m_ObjUsuarioHistoricoActivo Is Nothing Then
        m_Error = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    ' --- BLOQUEO DE SEGURIDAD PARA HISTÓRICO ---
    ' Recorremos los controles para bloquear la escritura, pero dejamos las listas vivas.
    For Each ctl In Me.Controls
        ' Bloqueamos la escritura en campos de texto, combos y casillas
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
            ctl.Locked = True
        End If
        
        ' Aseguramos que las listas SÍ se puedan tocar
        If TypeOf ctl Is ListBox Then
            ctl.Locked = False
            ctl.Enabled = True
        End If
    Next
    
    ' Asegúrate de desactivar los botones de acción al inicio si los tienes en este formulario
    ' (Ejemplo: Si tienes botones de editar observación, inícialos en False aquí)
    ' Me.ComandoObservacionEditar.Enabled = False
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim m_Dato As String
    Dim m_NombreCampo As String
    Dim m_Lugar As String
    
    On Error GoTo errores
    
    p_Error = ""
    
    Me.ListaObservaciones.RowSource = ""
    Me.ListaObservacionesCurso.RowSource = ""
    
    If m_ObjUsuarioHistoricoActivo Is Nothing Then
        Exit Function
    End If
    
    DatosDelObjetoAlFormulario m_ObjUsuarioHistoricoActivo, Me, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' NOTA: El bucle de "INTOCABLE" que tenías aquí ya no es estrictamente necesario
    ' porque el Form_Load va a bloquear (Locked=True) TODO al final.
    ' Pero lo dejamos por si acaso tienes lógica específica.
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            If Me.AllowEdits = True Then
                If InStr(1, ctl.Tag, "DATOS;INTOCABLE") <> 0 Then
                    ctl.Enabled = False
                End If
            End If
        End If
    Next
    
    If Not m_ObjUsuarioHistoricoActivo.SuministradorContratoObj Is Nothing Then
        Me.EmpresaUsuario = m_ObjUsuarioHistoricoActivo.SuministradorContratoObj.NombreConCif
    End If
    If Not m_ObjUsuarioHistoricoActivo.SuministradorTramitadoraObj Is Nothing Then
        Me.EmpresaTramitadora = m_ObjUsuarioHistoricoActivo.SuministradorTramitadoraObj.NombreConCif
    End If
   
    If Not m_ObjUsuarioHistoricoActivo.Expediente Is Nothing Then
        Me.Titulo = m_ObjUsuarioHistoricoActivo.Expediente.CodExp & " " & m_ObjUsuarioHistoricoActivo.Expediente.Titulo
    End If
    
    Me.Titulo.ForeColor = Me.Correo_e.ForeColor
    
    ' Cargar las listas (Ahora funcionarán los clics gracias al AllowEdits=True del Load)
    EstablecerListasObservaciones , p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & Err.Description
    End If
End Function
Private Function EstablecerListasObservaciones( _
                                                Optional p_Tipo As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_IDObservacion As Variant
    Dim m_ObjObServacion As ObservacionHistorica
    Dim lst As ListBox
    On Error GoTo errores
    
    p_Error = ""
    m_Error = ""
    
    
    
    If p_Tipo = "Curso" Then
        GoTo Curso
    End If
    Set lst = Me.ListaObservaciones
    lst.RowSource = ""
    Set m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricas = Nothing
    If Not m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricas Is Nothing Then
        p_Error = m_ObjUsuarioHistoricoActivo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDObservacion In m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricas.Keys
            If Nz(m_IDObservacion, "") <> "" Then
                Set m_ObjObServacion = m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricas(m_IDObservacion)
                With m_ObjObServacion
                    lst.AddItem .IDObservacion & ";" & .Fecha & ";" & .Observacion
                End With
                Set m_ObjObServacion = Nothing
            End If
        Next
        
    End If
Curso:
    Set lst = Me.ListaObservacionesCurso
    lst.RowSource = ""
    Set m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricasCurso = Nothing
    If Not m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricasCurso Is Nothing Then
        p_Error = m_ObjUsuarioHistoricoActivo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDObservacion In m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricasCurso.Keys
            If Nz(m_IDObservacion, "") <> "" Then
                Set m_ObjObServacion = m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricasCurso(m_IDObservacion)
                With m_ObjObServacion
                    lst.AddItem .IDObservacion & ";" & .Fecha & ";" & .Observacion
                End With
                Set m_ObjObServacion = Nothing
            End If
        Next
        
    End If

    EstablecerListasObservaciones = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerListasObservaciones ha devuelto el error: " & Err.Description
    End If
End Function



Private Sub ListaObservaciones_DblClick(Cancel As Integer)
    Dim m_IDObservacionSeleccionada As String
    On Error GoTo errores
    
    m_Error = ""
    
    m_IDObservacionSeleccionada = Nz(Me.ListaObservaciones.Column(0), "")
    If m_IDObservacionSeleccionada = "" Then
        
        
        Exit Sub
    End If
    
    If FormularioAbierto("FormUsuario02DatosObservacionHistorico") Then
        DoCmd.Close acForm, "FormUsuario02DatosObservacionHistorico", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario02DatosObservacionHistorico", OpenArgs:=m_IDObservacionSeleccionada
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en la lista se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ListaObservacionesCurso_DblClick(Cancel As Integer)
     Dim m_IDObservacionSeleccionada As String
    On Error GoTo errores
    
    m_Error = ""
    
    m_IDObservacionSeleccionada = Nz(Me.ListaObservacionesCurso.Column(0), "")
    If m_IDObservacionSeleccionada = "" Then
        
        
        Exit Sub
    End If
    
    If FormularioAbierto("FormUsuario02DatosObservacionHistorico") Then
        DoCmd.Close acForm, "FormUsuario02DatosObservacionHistorico", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario02DatosObservacionHistorico", OpenArgs:=m_IDObservacionSeleccionada
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en la lista se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub