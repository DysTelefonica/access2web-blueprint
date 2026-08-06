Attribute VB_Name = "Form_FormUsuario01Lateral"
Option Compare Database
Option Explicit
Private m_Error As String


Private Sub Form_Load()
    
    On Error GoTo errores
    m_Error = ""
   
    
    Me.lblVersion.Caption = "Versión: " & m_ObjEntorno.VersionAplicacion
    Me.lblDatosEn.Caption = "Datos en: " & m_ObjEntorno.DatosEn
    Me.lblModoEscritura.Caption = "Modo Escritura: " & m_ObjEntorno.ModoEscritura
    SeleccionarEnlace , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    lblDatosGenerales_Click
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Public Sub lblAnexos_Click()
    On Error GoTo errores
    m_Error = ""
    If Me.Parent.SubFormCentral.SourceObject = "FormUsuario02Anexos" Then
        Exit Sub
    End If
    SeleccionarEnlace "ANEXO", m_Error
    Me.lblAnexos.Tag = "Elegibles;Seleccionado"
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerControlesformlateral Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuario02Anexos"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub lblDatosGenerales_Click()
    
    On Error GoTo errores
    m_Error = ""
    
    If Not Me.Parent.SubFormCentral.SourceObject = "FormUsuario02Edicion" Then
        Me.Parent.SubFormCentral.SourceObject = "FormUsuario02Edicion"
    End If
    
    SeleccionarEnlace "", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.lblDatosGenerales.Tag = "Elegibles;Seleccionado"
    EstablecerControlesformlateral Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al establecer los controles se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub lblHPS_Click()
    On Error GoTo errores
    m_Error = ""
    If Me.Parent.SubFormCentral.SourceObject = "FormUsuario02HPS" Then
        Exit Sub
    End If
    
    SeleccionarEnlace "TODOS", m_Error
    Me.lblHPS.Tag = "Elegibles;Seleccionado"
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerControlesformlateral Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuario02HPS"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Function SeleccionarEnlace(Optional p_TipoHPS As String, Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim m_NombreControl As String
    Dim m_nombreControlSeleccionado As String
    On Error GoTo errores
    If p_TipoHPS = "" Then
        If m_ObjUsuarioActivo Is Nothing Then
            Exit Function
        End If
        For Each ctl In Me.Controls
            If TypeOf ctl Is Label Then
                
                If InStr(1, ctl.Tag, ";") <> 0 Then
                    ctl.Tag = "Elegibles;NoSeleccionado"
                End If
            End If
        Next
        
        Exit Function
    ElseIf p_TipoHPS = "ANEXO" Then
        m_nombreControlSeleccionado = Me.lblAnexos.Name
        Me.lblAnexos.Tag = "Elegibles;Seleccionado"
        For Each ctl In Me.Controls
            If TypeOf ctl Is Label Then
                m_NombreControl = ctl.Name
                If m_NombreControl <> m_nombreControlSeleccionado Then
                    If InStr(1, Nz(ctl.Tag, ";"), ";") <> 0 Then
                        ctl.Tag = "Elegibles;NoSeleccionado"
                    End If
                End If
                
            End If
        Next
        Exit Function
    End If
    Me.lblHPS.Tag = "Elegibles;Seleccionado"
    For Each ctl In Me.Controls
        If TypeOf ctl Is Label Then
            m_NombreControl = ctl.Name
            If m_NombreControl <> m_nombreControlSeleccionado Then
                If InStr(1, Nz(ctl.Tag, ";"), ";") <> 0 Then
                    ctl.Tag = "Elegibles;NoSeleccionado"
                End If
            End If
            
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Al SeleccionarEnlace se ha producido el error: " & Err.Description
    End If
End Function
