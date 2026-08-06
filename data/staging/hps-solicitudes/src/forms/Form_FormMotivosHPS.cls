Option Compare Database
Option Explicit

Private WithEvents m_Form As Form_FormMotivoHPS
Private m_MotivoSeleccionado As MotivoHPS
Private m_Error As String

Private Sub cmdSalir_Click()
    
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Motivo As MotivoHPS
    
    On Error GoTo errores
    Me.ComandoEditar.Enabled = True
    Me.ComandoEliminar.Enabled = True
    
    Set lst = Me.ListaFiltrados
    lst.RowSource = ""
    Set m_Col = m_ObjEntorno.ColMotivos
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_ID In m_Col
        Set m_Motivo = m_Col(m_ID)
        
        lst.AddItem m_Motivo.MotivoHPS
        
        Set m_Motivo = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha devuelto el error: " & Err.Description
    End If
End Function



Private Sub ComandoAlta_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If FormularioAbierto("FormMotivoHPS") Then
        DoCmd.Close acForm, "FormMotivoHPS", acSaveNo
    End If
    Set m_ObjMotivoActivo = Nothing
    DoCmd.OpenForm "FormMotivoHPS"
    If FormularioAbierto("FormMotivoHPS") Then
        Set m_Form = Forms("FormMotivoHPS")
    End If
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEditar_Click()
    
  
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_MotivoSeleccionado Is Nothing Then
        m_Error = "Seleccione una elemento de la lista"
        Err.Raise 1000
    End If
    If FormularioAbierto("FormMotivoHPS") Then
        DoCmd.Close acForm, "FormMotivoHPS", acSaveNo
    End If
    Set m_ObjMotivoActivo = m_MotivoSeleccionado
    DoCmd.OpenForm "FormMotivoHPS"
    If FormularioAbierto("FormMotivoHPS") Then
        Set m_Form = Forms("FormMotivoHPS")
    End If
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
    On Error GoTo errores
    
    
    m_Error = ""
    If m_MotivoSeleccionado Is Nothing Then
        m_Error = "Seleccione una elemento de la lista"
        Err.Raise 1000
    End If
    If m_EnOficina = EnumSiNo.no Then
        m_Error = "Sólo puede hacerse en la oficina"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea eliminar el motivo seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar motivo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_MotivoSeleccionado.EliminarMotivo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Motivo borrado correctamente", vbInformation, "Motivo eliminado")
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_ObjEntorno.ColMotivos = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me
    
    Me.Caption = Me.lblTitulo1.Caption
    
    
    
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaFiltrados_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    If Nz(Me.ListaFiltrados, "") = "" Then
        Set m_MotivoSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_MotivoSeleccionado Is Nothing Then
        Set m_MotivoSeleccionado = constructor.getMotivo(p_Motivo:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaFiltrados.Column(0) <> m_MotivoSeleccionado.MotivoHPS Then
            Set m_MotivoSeleccionado = constructor.getMotivo(p_Motivo:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If

        End If
    End If
    
    
    If m_MotivoSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        
        Exit Sub
    End If
    Me.ComandoEditar.Enabled = True
    Me.ComandoEliminar.Enabled = True
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaMotivos_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    
    On Error Resume Next
    Application.Echo False
    If Me.ComandoEditar.Enabled = True Then
        ComandoEditar_Click
    End If
    Application.Echo True
End Sub

Private Sub m_Form_Alta()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjEntorno.ColMotivos = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Alta se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Form_Editado()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjEntorno.ColMotivos = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Editado se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub