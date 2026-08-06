Attribute VB_Name = "Form_FormUsuarioHistorico00Principal"

Option Compare Database
Option Explicit
Private m_Error As String


Public Sub CmdSalir_Click()
    On Error Resume Next
    
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub ComandoPasarAActual_Click()
    Dim m_Servicio As UsuarioServicio
    On Error GoTo errores

    m_Error = ""
    pregunta = MsgBox("¿Desea realmente pasar este usuario a Actual?", vbExclamation + vbYesNo + vbDefaultButton2, "Pasar a Actual")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Set m_Servicio = New UsuarioServicio

    ' Reloj de arena durante el move (puede tardar con muchos anexos / observación)
    DoCmd.Hourglass True
    m_Servicio.PasarDeHistoricoAActual m_ObjUsuarioHistoricoActivo.ID, m_Error
    If m_Error <> "" Then
        DoCmd.Hourglass False
        Err.Raise 1000
    End If
    DoCmd.Hourglass False

    If m_Error = "" Then
        Indicadores_NotificarMovimiento "h2c", CStr(m_ObjUsuarioHistoricoActivo.ID)
    End If

    MsgBox "El usuario ha sido pasado a Actual correctamente.", vbInformation, "Operación completada"

    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPasarAActual_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    Dim m_Titulo As String
    On Error GoTo errores
    
   
'    Me.InsideHeight = 9915
'    Me.InsideWidth = 17760
    AjustarTamaño Me
'    Debug.Print "ME.InsideHeight=" & Me.InsideHeight
'    Debug.Print "ME.InsideWidth=" & Me.InsideWidth
    Me.AllowEdits = False
    If m_ObjUsuarioHistoricoActivo Is Nothing Then
        m_Error = "No se puede determinar el usuario activo"
        Err.Raise 1000
    End If
    
    m_Titulo = "HISTÓRICO DETALLE DE " & UCase(m_ObjUsuarioHistoricoActivo.NombreCompleto)
    Me.lblTitulo.Caption = m_Titulo
    Me.Caption = m_Titulo
    
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Me.SubFormLateral.SourceObject <> "FormUsuarioHistorico01Lateral" Then
        Me.SubFormLateral.SourceObject = "FormUsuarioHistorico01Lateral"
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.ComandoPasarAActual.Enabled = False
    Else
        Me.ComandoPasarAActual.Enabled = False
    End If
    CopiarEstiloBoton Me.ComandoPasarAActual
    CopiarEstiloBoton Me.cmdSalir
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
