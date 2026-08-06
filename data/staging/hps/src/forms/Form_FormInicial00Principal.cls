Attribute VB_Name = "Form_FormInicial00Principal"
Option Compare Database
Option Explicit
Private m_Error As String

Public Event ActualizadoUsuarios()
Public Event AvanceActualizacionUsuarios()


Private Sub CmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub ComandoActualizar_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set lbl = Me.lblAvance
    SincronizarCachesLocalesFrontend p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerContadores
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoBuscar_Click()
    Dim Palabra As String
    Dim frm As Form
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    Palabra = Nz(Me.PalabraClave, "")
    
    If Not FormularioAbierto("FormInicial06BusquedaGlobal") Then
        DoCmd.OpenForm "FormInicial06BusquedaGlobal"
    End If
    If FormularioAbierto("FormInicial06BusquedaGlobal") Then
        Set frm = Forms("FormInicial06BusquedaGlobal")
        If Palabra <> "" Then
            frm.Controls("PalabraClave") = Palabra
            Form_FormInicial06BusquedaGlobal.PalabraClave_Exit False
        End If
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al buscar se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Public Sub Form_Load()
    
    On Error GoTo errores
    
    
    'DoCmd.Maximize
    Application.Echo False
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    Application.TempVars("EnDesarrollo") = EnumSiNo.No
    'Application.TempVars("EnDesarrollo") = EnumSiNo.Sí
    
    Application.Echo True
    If m_ObjEntorno Is Nothing Then
        EVE m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If

    End If
    
    

   
    
    Me.EncabezadoDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor
    Me.PieDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor

    CopiarEstiloBoton Me.ComandoBuscar, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoActualizar, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = "Usuario conectado: " & m_ObjUsuarioConectado.Nombre
    Me.SubFormLateral.SourceObject = "FormInicial01Lateral"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Quit acQuitSaveNone
End Sub


Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' Código que deseas ejecutar cuando se pulse Enter
        If Me.ComandoBuscar.Enabled = True Then
            Me.ComandoBuscar.SetFocus
            ComandoBuscar_Click
        End If
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    End If
End Sub