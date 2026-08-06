 Option Compare Database
Option Explicit

Public Event Alta()
Public Event Editado()
Private m_GradoClasificacionAlInicio As GradoClasificacion
Private m_Error As String

Private Sub cmdSalir_Click()
     On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

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
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrar_Click()
    
    Dim m_HaHabidoCambios As Boolean
    Dim m_GradoClasificacionOp As GradoClasificacionOperaciones
    Dim parsed As Object
    Dim payload As Object
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    MostrarPopupProgreso "Procesando...", "Guardando..."
    m_Error = ""
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set parsed = JsonConverter.ParseJson(modGradoClasificacionFormHelper.GradoClasificacionForm_RegisterDecision(m_HaHabidoCambios, m_GradoClasificacionAlInicio Is Nothing, m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    If m_ObjGradoClasificacionActivo Is Nothing Then
        Set m_ObjGradoClasificacionActivo = New GradoClasificacion
    End If
    Dim m_ValoresForm As Scripting.Dictionary
    Set m_ValoresForm = modGradoClasificacionFormHelper.GradoClasificacionForm_BuildValues(Nz(Me.GradoClasificacion, ""), Nz(Me.DESCRIPCION, ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Helper_EntidadCRUD.CopiarCamposAObjeto m_ObjGradoClasificacionActivo, m_ValoresForm, Array("GradoClasificacion", "DESCRIPCION"), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_GradoClasificacionOp = New GradoClasificacionOperaciones
    With m_GradoClasificacionOp
        Set .GradoClasificacion = m_ObjGradoClasificacionActivo
        .Registrar m_GradoClasificacionAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    If CStr(payload("eventName")) = "Alta" Then
        RaiseEvent Alta
    Else
        RaiseEvent Editado
    End If
    VBA.DoEvents
    CerrarPopupProgreso
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    Dim parsed As Object
    Dim payload As Object
    Dim m_ValoresForm As Scripting.Dictionary

    On Error GoTo errores
    Set m_ValoresForm = modGradoClasificacionFormHelper.GradoClasificacionForm_BuildValues(Nz(Me.GradoClasificacion, ""), Nz(Me.DESCRIPCION, ""), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Set parsed = JsonConverter.ParseJson(modGradoClasificacionFormHelper.GradoClasificacionForm_HasChanges(m_GradoClasificacionAlInicio, m_ValoresForm, p_Error))
    If p_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    HaHabidoCambios = CBool(payload("hasChanges"))
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error num " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores
    
    Set parsed = JsonConverter.ParseJson(modGradoClasificacionFormHelper.GradoClasificacionForm_FormOpenState(EsAdministrador = EnumSiNo.Sí, Not (m_ObjGradoClasificacionActivo Is Nothing), p_Error))
    If p_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    If m_ObjGradoClasificacionActivo Is Nothing Then
        Me.lblTitulo.Caption = CStr(payload("title"))
        
        Exit Function
    End If
    Me.lblTitulo.Caption = CStr(payload("title"))
   
    Set m_GradoClasificacionAlInicio = constructor.getGradoClasificacion(p_IdGradoClasificacion:=m_ObjGradoClasificacionActivo.IdGradoClasificacion, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_GradoClasificacionAlInicio Is Nothing Then
        p_Error = "No se ha podido encontrar el Grado Clasificación registrado"
        Err.Raise 1000
    End If
    Me.lblTitulo.Caption = CStr(payload("title"))
    Me.Caption = m_TituloFormulario
    With m_GradoClasificacionAlInicio
        If .GradoClasificacion <> "" Then
            Me.GradoClasificacion = .GradoClasificacion
        End If
        If .DESCRIPCION <> "" Then
            Me.DESCRIPCION = .DESCRIPCION
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Private Sub Form_Open(Cancel As Integer)
    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Ajustar Me
    
    Set parsed = JsonConverter.ParseJson(modGradoClasificacionFormHelper.GradoClasificacionForm_FormOpenState(EsAdministrador = EnumSiNo.Sí, Not (m_ObjGradoClasificacionActivo Is Nothing), m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    Me.ComandoEditar.Enabled = CBool(payload("editarEnabled"))
    Me.Caption = m_TituloFormulario
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
