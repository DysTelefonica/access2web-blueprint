Attribute VB_Name = "Form_FormInicial05Configuracion"
Option Compare Database
Option Explicit

Private m_Error As String

Private Sub ComandoImportarDatosOficinaALocal_Click()
    Dim URLCarpetaLocal As String
    
    On Error GoTo errores
    m_Error = ""
'    Seleccionar False, "Seleccione la carpeta para exportar los datos", URLCarpetaLocal, m_Error
'    If m_Error <> "" Then
'        Err.Raise 1000
'    End If
'    If URLCarpetaLocal = "" Then
'        Exit Sub
'    End If
    DoCmd.Hourglass True
    
    URLCarpetaLocal = Application.CurrentProject.path
    CopiarDatosDeOficinaACarpetaLocal URLCarpetaLocal, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox("Datos transferidos al aplicativo local correctamente", vbInformation, "Importación de Datos")
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar importar los datos se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    
    On Error GoTo errores
    m_Error = ""
    Me.EstadoConexion = m_EstadoConexion
    
    If m_EnOficina = EnumSiNo.Sí Then
        Me.DatosDeLaOficina = "Sí"
    Else
        Me.DatosDeLaOficina = "No"
    End If
    
    If m_EnOficina = EnumSiNo.Sí Then
        Me.EstoyEnLaOficina = "Sí"
    Else
        Me.EstoyEnLaOficina = "No"
    End If
    
    Me.ComandoImportarDatosOficinaALocal.Enabled = False
    
    
    If Me.EstoyEnLaOficina = "Sí" And Me.DatosDeLaOficina = "Sí" Then
        If InStr(1, Application.CurrentProject.path, "AppData\Roaming") = 0 Then
            Me.ComandoImportarDatosOficinaALocal.Enabled = True
            
        End If
        
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Form_Load ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

