Attribute VB_Name = "Form_FormUsuario00Principal"

Option Compare Database
Option Explicit
Private m_Error As String


Public Sub CmdSalir_Click()
    On Error Resume Next
    
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub Form_Load()
    
    Dim m_Titulo As String
    On Error GoTo errores
    
   
'    Me.InsideHeight = 9915
'    Me.InsideWidth = 17760
    AjustarTamaño Me
   ' DoCmd.Maximize
'    Debug.Print "ME.InsideHeight=" & Me.InsideHeight
'    Debug.Print "ME.InsideWidth=" & Me.InsideWidth
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "No se puede determinar el usuario activo"
        Err.Raise 1000
    End If
    m_ObjUsuarioActivo.ParaAlta = EnumSiNo.No
    If Me.AllowEdits = True Then
        m_Titulo = "EDICIÓN DE " & UCase(m_ObjUsuarioActivo.NombreCompleto)
    Else
        m_Titulo = "DETALLE DE " & UCase(m_ObjUsuarioActivo.NombreCompleto)
    End If
    Me.lblTitulo.Caption = m_Titulo
    Me.Caption = m_Titulo
    
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Me.SubFormLateral.SourceObject <> "FormUsuario01Lateral" Then
        Me.SubFormLateral.SourceObject = "FormUsuario01Lateral"
    End If
    CopiarEstiloBoton Me.cmdSalir
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

