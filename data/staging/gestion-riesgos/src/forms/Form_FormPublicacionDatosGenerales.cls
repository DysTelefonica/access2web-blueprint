Option Compare Database
Option Explicit
Private m_Error As String

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me

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
    
End Sub


Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
       
    Dim m_CorreoRAC As String
    On Error Resume Next
    
    With m_ObjEdicionActiva
       
        Me.FechaEdicion = Format(.FechaEdicion, "dd/mm/yyyy")
        Me.CodigoDocumento = .Proyecto.CodigoDocumento
        Me.Edicion = .Edicion
        If .FechaPublicacion <> "" Then
            
            Me.Elaborado = .Elaborado
            Me.Revisado = .Revisado
            Me.Aprobado = .Aprobado
        Else
           
            
            Me.Elaborado = .Proyecto.Elaborado
            Me.Revisado = .Proyecto.Revisado
            Me.Aprobado = .Proyecto.Aprobado
        End If
        
        m_CorreoRAC = .Proyecto.CorreoRACCalculado
        If m_CorreoRAC <> "" Then
            Me.CorreoRAC = m_CorreoRAC
        End If
        
    End With

End Function
