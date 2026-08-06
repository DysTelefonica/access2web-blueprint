Option Compare Database
Option Explicit
Private m_Error As String



Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjProyectoActivo Is Nothing Then
        m_Error = "No se tienen datos del proyecto activo"
        Err.Raise 1000
    End If
    
    
    
    
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
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    
    Dim m_Responsable As ExpedienteResponsable
    Dim m_ID As Variant
    Dim lst As ListBox
    On Error GoTo errores
    
    Set lst = Me.ListaAutorizados
    lst.RowSource = "Autorizado;Con Correo"
    If m_ObjProyectoActivo.Expediente Is Nothing Then
        Exit Function
    End If
    If m_ObjProyectoActivo.Expediente.Responsables Is Nothing Then
        Exit Function
    End If
    
   
    For Each m_ID In m_ObjProyectoActivo.Expediente.Responsables
        
        Set m_Responsable = m_ObjProyectoActivo.Expediente.Responsables(m_ID)
        lst.AddItem m_Responsable.Usuario.Nombre & ";" & m_Responsable.CorreoSiempre
        Set m_Responsable = Nothing
    Next
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.description
        
    End If
    
End Function