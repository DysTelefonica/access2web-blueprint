Option Compare Database
Option Explicit
Private m_RiesgoMaterializado As RiesgoMaterializacion
Private m_Riesgo As riesgo
Private m_Edicion As Edicion
Private m_Error As String

Private Sub ComandoIrFormularioCalidad_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormCalidadRiesgoMaterializaciones") Then
        DoCmd.Close acForm, "FormCalidadRiesgoMaterializaciones", acSaveNo
    End If
    Set m_ObjRiesgoActivo = m_Riesgo
    DoCmd.OpenForm "FormCalidadRiesgoMaterializaciones"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIrFormularioCalidad_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVerRiesgo_Click()
    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormRiesgo") Then
        DoCmd.Close acForm, "FormRiesgo", acSaveNo
    End If
    m_EsAlta = EnumSiNo.No
    If m_Riesgo Is Nothing Then
        If Form_FormCalidadTareas.m_RiesgoMaterializadoSeleccionado Is Nothing Then
            m_Error = "No se sabe el riesgo seleccionado"
            Err.Raise 1000
        End If
        
        If m_Riesgo Is Nothing Then
            If Nz(Me.IDRiesgo, "") = "" Then
                m_Error = "No se sabe el riesgo seleccionado"
                Err.Raise 1000
            End If
            
        End If
    End If
    Set m_ObjRiesgoActivo = m_Riesgo
    DoCmd.OpenForm "FormRiesgo"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRiesgo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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

    m_Error = ""
    If Form_FormCalidadTareas.m_RiesgoMaterializadoSeleccionado Is Nothing Then
        m_Error = "No se ha seleccionado ningún riesgo materializado"
        Err.Raise 1000
    End If
    Set m_RiesgoMaterializado = Form_FormCalidadTareas.m_RiesgoMaterializadoSeleccionado
    If Not m_RiesgoMaterializado.RiesgoEnEdicionActiva Is Nothing Then
        Set m_Riesgo = m_RiesgoMaterializado.RiesgoEnEdicionActiva
    Else
        Set m_Riesgo = m_RiesgoMaterializado.riesgo
    End If
    
    If m_Riesgo Is Nothing Then
        m_Error = "No se ha seleccionado ningún riesgo materializado"
        Err.Raise 1000
    End If
    Set m_Edicion = m_RiesgoMaterializado.Edicion

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
Public Function EstablecerDatos(Optional ByRef p_Error As String) As String

    Dim m_Titulo As String
    Dim m_ParticulaActiva As String

    On Error GoTo errores

    p_Error = ""
    If EsAdministrador = EnumSiNo.Sí Or EsAdministradorConectadoInicialmente = EnumSiNo.Sí Then
        Me.lblIDExpediente.Visible = True
        Me.IDExpediente.Visible = True
        Me.lblIDProyecto.Visible = True
        Me.IDProyecto.Visible = True
        Me.lblIDEdicion.Visible = True
        Me.IDEdicion.Visible = True
        Me.lblIDRiesgo.Visible = True
        Me.IDRiesgo.Visible = True
    Else
        Me.lblIDExpediente.Visible = False
        Me.IDExpediente.Visible = False
        Me.lblIDProyecto.Visible = False
        Me.IDProyecto.Visible = False
        Me.lblIDEdicion.Visible = False
        Me.IDEdicion.Visible = False
        Me.lblIDRiesgo.Visible = False
        Me.IDRiesgo.Visible = False
    End If
    Me.ComandoVerRiesgo.Enabled = True
    With m_RiesgoMaterializado
       
        With m_Riesgo
            Me.IDExpediente = .Edicion.Proyecto.IDExpediente
            Me.IDProyecto = .Edicion.IDProyecto
            Me.IDEdicion = .IDEdicion
            Me.IDRiesgo = .IDRiesgo
        End With
        Me.DescripcionRiesgo = "(" & m_Riesgo.CodigoRiesgo & ") " & m_Riesgo.Descripcion
        Me.Proyecto = .Proyecto.Proyecto
        Me.Supervisor = .Edicion.Proyecto.NombreUsuarioCalidad
        
        If m_Edicion.EsActivo = EnumSiNo.Sí Then
            m_ParticulaActiva = "activa"
        Else
            m_ParticulaActiva = "no Activa"
        End If
        Me.DetalleEdicion = "Edición " & .Edicion.Edicion & " (" & m_ParticulaActiva & " )"
        Me.ComandoVerRiesgo.Enabled = True
        If IsDate(m_Riesgo.FechaMaterializado) Then
            Me.FechaMaterializado = m_Riesgo.FechaMaterializado
        End If
    End With
    

    m_Titulo = "Riesgo materializado NC por decidir"
    Me.lblTitulo.Caption = m_Titulo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function