 Option Compare Database
Option Explicit

Public Event Alta()
Public Event Editado()
Private m_ResponsableAlInicio As Responsable
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
    Dim m_ResponsableOp As ResponsableOperaciones
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = False Then
        m_Error = "No hay cambios que guardar"
        Err.Raise 1000
    End If
    If m_ObjResposableActivo Is Nothing Then
        Set m_ObjResposableActivo = New Responsable
    End If
    With m_ObjResposableActivo
        
        .Nombre = Nz(Me.Nombre, "")
        .Correo = Nz(Me.Correo, "")
        .Descripcion = Nz(Me.Descripcion, "")
        
    End With
    
    Set m_ResponsableOp = New ResponsableOperaciones
    With m_ResponsableOp
        Set .Responsable = m_ObjResposableActivo
        .Registrar m_ResponsableAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    If m_ResponsableAlInicio Is Nothing Then
        RaiseEvent Alta
    Else
        RaiseEvent Editado
    End If
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

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    On Error GoTo errores
    If m_ResponsableAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_ResponsableAlInicio
        
        If .Nombre <> Nz(Me.Nombre, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Correo <> Nz(Me.Correo, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Descripcion <> Nz(Me.Descripcion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error num " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_Responsable As Responsable
    On Error GoTo errores
   
    If m_ObjResposableActivo Is Nothing Then
        Me.lblTitulo.Caption = "ALTA DE RESPONSABLE"
        
        Exit Function
    End If
    Me.lblTitulo.Caption = "EDICIÓN DE RESPONSABLE"
    
    Set m_ResponsableAlInicio = constructor.getResponsable(p_IDResponsable:=m_ObjResposableActivo.IDResponsable, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ResponsableAlInicio Is Nothing Then
        p_Error = "No se ha podido encontrar el Responsable registrado"
        Err.Raise 1000
    End If
   
    
    With m_ResponsableAlInicio
        If .Nombre <> "" Then
            Me.Nombre = .Nombre
        End If
        If .Correo <> "" Then
            Me.Correo = .Correo
        End If
        If .Descripcion <> "" Then
            Me.Descripcion = .Descripcion
        End If
        
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Ajustar Me
    Me.ComandoRegistrar.Enabled = False
    Me.AllowEdits = False
    Me.Caption = m_TituloFormulario
    If EsTecnico <> EnumSiNo.Sí Then
        Me.ComandoRegistrar.Enabled = True
        Me.AllowEdits = True
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub