Option Compare Database
Option Explicit
Private m_Error As String
Private m_Edicion As Edicion

Private Sub ComandoActualizarFechaMaxPublicacion_Click()
    Dim m_Proyecto As Proyecto
    Dim m_Dias As Long
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_Proyecto = Constructor.getProyecto(p_IDProyecto:=m_Edicion.IDProyecto, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Proyecto.FechaMaxProximaPublicacionGrabar p_Error:=m_Error
    Set m_Edicion = Constructor.getEdicion(p_IDEdicion:=m_Edicion.IDEdicion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_Proyecto.FechaMaxProximaPublicacion) Then
        Me.FechaProgPublicar = m_Proyecto.FechaMaxProximaPublicacion
        m_Dias = DateDiff("d", Date, CDate(m_Proyecto.FechaMaxProximaPublicacion))
        Me.DiasParaPublicar = m_Dias
        If m_Dias > 0 Then
            Me.DiasParaPublicar.ForeColor = Me.FechaProgPublicar.ForeColor
        Else
            Me.DiasParaPublicar.ForeColor = RGB(255, 0, 0)
        End If
    Else
        Me.FechaProgPublicar = Null
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    AvanceCerrar
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizarFechaMaxPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGenerarInforme_Click()
    Dim url As String
  
    
    On Error GoTo errores
    
    If m_Edicion.TodosLosRiesgosPriorizados = EnumSiNo.No Then
        pregunta = MsgBox("Para poder generar el informe ha de tener todos sus riesgos priorizados", vbExclamation, "Error")
        Exit Sub
    End If
    
    ' 1. Leemos el valor del combo (asumiendo HTML si está vacío)
   
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    ' 2. Llamamos a la función inyectando el parámetro PDF
    url = GenerarInformePublicacion(m_Edicion, EnumTipoInformePublicacion.html, "", "", , m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' Dejo esta línea comentada porque GenerarInformeEdicionHTML ya se encarga de abrirlo
    ' Ejecutar Me.hWnd, "open", url, "", "", 1
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    AvanceCerrar
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarInforme_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub




'
Private Sub ComandoInformePublicabilidad_Click()

    Dim m_URL As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
     m_URL = GenerarInformePublicabilidadEdicionInteractivoHTML(p_Edicion:=m_Edicion, p_hWnd:=Me.hWnd, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLHTMLActivo = m_URL

    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    AvanceCerrar
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoInformePublicabilidad_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If

End Sub
'
Private Sub ComandoIrAPublicar_Click()

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    Set m_ObjProyectoActivo = m_Edicion.Proyecto
    Set m_ObjEdicionActiva = m_Edicion
    
    DoCmd.OpenForm "FormPublicacionCalidad"


    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIrAPublicar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If

End Sub
'
Public Sub ComandoVerDetalle_Click()

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjProyectoActivo = m_Edicion.Proyecto
    Set m_ObjEdicionActiva = m_Edicion
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.OpenForm "FormRiesgosGestion"
     VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerDetalle_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If

End Sub

Private Sub ComandoVerProyecto_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjProyectoActivo = m_Edicion.Proyecto
    
    Set m_ObjProyectoAlInicio = Nothing
    m_EsAlta = EnumSiNo.No
    If FormularioAbierto("FormGestionRiesgos") Then
        DoCmd.Close acForm, "FormGestionRiesgos", acSaveNo
    End If
    DoCmd.OpenForm "FormGestionRiesgos"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

'
Private Sub Form_Open(Cancel As Integer)


    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Form_FormCalidadTareas.m_EdicionSeleccionada Is Nothing Then
        m_Error = "No se conoce la edición activa"
        Err.Raise 1000
    End If
    Set m_Edicion = Form_FormCalidadTareas.m_EdicionSeleccionada
    EstablecerDatos m_Error
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
'
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String

    Dim m_Proyecto As Proyecto
    Dim m_Dias As Long
    On Error GoTo errores


    Me.ComandoVerDetalle.Enabled = False
    If EsAdministrador = EnumSiNo.Sí Or EsAdministradorConectadoInicialmente = EnumSiNo.Sí Then
        Me.lblIDExpediente.Visible = True
        Me.IDExpediente.Visible = True
        Me.lblIDProyecto.Visible = True
        Me.IDProyecto.Visible = True
        Me.lblIDEdicion.Visible = True
        Me.IDEdicion.Visible = True
    Else
        Me.lblIDExpediente.Visible = False
        Me.IDExpediente.Visible = False
        Me.lblIDProyecto.Visible = False
        Me.IDProyecto.Visible = False
        Me.lblIDEdicion.Visible = False
        Me.IDEdicion.Visible = False
    End If
    Set m_Proyecto = m_Edicion.Proyecto
    If Not m_Edicion.EdicionAnterior Is Nothing Then
        Me.EdicionAnterior = m_Edicion.EdicionAnterior.Edicion
        If IsDate(m_Edicion.EdicionAnterior.FechaPublicacion) Then
            Me.FechaPublicacionEdicionAnterior = m_Edicion.EdicionAnterior.FechaPublicacion
        End If
    End If
     Me.Proyecto = m_Proyecto.Proyecto
    Me.NombreProyecto = m_Proyecto.NombreProyecto
    Me.Supervisor = m_Proyecto.NombreUsuarioCalidad
    Me.IDExpediente = m_Proyecto.IDExpediente
    Me.IDProyecto = m_Proyecto.IDProyecto
    Me.IDEdicion = m_Edicion.IDEdicion
    Me.EdicionActual = m_Edicion.Edicion
    If IsDate(m_Proyecto.FechaMaxProximaPublicacion) Then
        Me.FechaProgPublicar = m_Proyecto.FechaMaxProximaPublicacion
        m_Dias = DateDiff("d", Date, CDate(m_Proyecto.FechaMaxProximaPublicacion))
        Me.DiasParaPublicar = m_Dias
        If m_Dias > 0 Then
            Me.DiasParaPublicar.ForeColor = Me.FechaProgPublicar.ForeColor
        Else
            Me.DiasParaPublicar.ForeColor = RGB(255, 0, 0)
        End If
    End If
    If IsDate(m_Edicion.FechaPreparadaParaPublicar) Then
        Me.FechaPreparadaParaPublicar = m_Edicion.FechaPreparadaParaPublicar
    End If
    If IsDate(m_Proyecto.FechaCierre) Then
        Me.FechaCierre = m_Proyecto.FechaCierre
    End If
    If Not m_Proyecto.Expediente Is Nothing Then
            If IsDate(m_Proyecto.Expediente.FechaFinContrato) Then
                Me.FechaFinContrato = m_Proyecto.Expediente.FechaFinContrato
                
                
                If Me.FechaCierre <> "" Then
                    Me.FechaFinContrato.ForeColor = Me.FechaProgPublicar.ForeColor
                Else
                    If CDate(m_Proyecto.Expediente.FechaFinContrato) < CDate(Date) Then
                        Me.FechaFinContrato.ForeColor = RGB(255, 0, 0)
                    Else
                        Me.FechaFinContrato.ForeColor = Me.FechaProgPublicar.ForeColor
                    End If
                End If
                
            End If
        End If
    
    
    Me.ComandoVerDetalle.Enabled = True
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If

End Function
'
'
