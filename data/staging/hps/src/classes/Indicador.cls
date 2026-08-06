VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Indicador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_objColUsuariosPendientesCursoHPS As Scripting.Dictionary
    Private m_objColUsuariosPtesPrimeraConvocatoriaCurso As Scripting.Dictionary
    Private m_objColUsuariosPtesSegundaConvocatoriaCurso As Scripting.Dictionary
    Private m_objColUsuariosPtesEnvioCorreoJefeSeguridadCurso As Scripting.Dictionary
Private m_objColUsuariosHPSconHPSAPuntoDeCaducar As Scripting.Dictionary
Private m_objColUsuariosHPSconHPSACaducadas As Scripting.Dictionary
Private m_objColUsuariosHPSconHPSASolicitandose As Scripting.Dictionary

Private m_iNPendientesCursoHPS As Integer

Private m_iNHPSAPuntoDeCaducar As Integer
Private m_iNHPSACaducadas As Integer
Private m_iNHPSASolicitandose As Integer

Private m_iNUsuariosPtesPrimeraConvocatoriaCurso As Integer
Private m_iNUsuariosPtesSegundaConvocatoriaCurso As Integer
Private m_iNUsuariosPtesEnvioCorreoJefeSeguridadCurso As Integer
Public Error As String
Private m_Error As String


Public Property Get ColUsuariosPendientesCursoHPS() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosPendientesCursoHPS Is Nothing Then
        Set ColUsuariosPendientesCursoHPS = m_objColUsuariosPendientesCursoHPS
        Exit Property
    End If
    Set m_objColUsuariosPendientesCursoHPS = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PendientesCursoHPS, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosPendientesCursoHPS = m_objColUsuariosPendientesCursoHPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.ColUsuariosPendientesCursoHPS ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property


Public Property Get ColUsuariosPtesPrimeraConvocatoriaCurso() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosPtesPrimeraConvocatoriaCurso Is Nothing Then
        Set ColUsuariosPtesPrimeraConvocatoriaCurso = m_objColUsuariosPtesPrimeraConvocatoriaCurso
        Exit Property
    End If
    Set m_objColUsuariosPtesPrimeraConvocatoriaCurso = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PtesPrimeraConvocatoriaCurso, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosPtesPrimeraConvocatoriaCurso = m_objColUsuariosPtesPrimeraConvocatoriaCurso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.ColUsuariosPtesPrimeraConvocatoriaCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property


Public Property Get ColUsuariosPtesSegundaConvocatoriaCurso() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosPtesSegundaConvocatoriaCurso Is Nothing Then
        Set ColUsuariosPtesSegundaConvocatoriaCurso = m_objColUsuariosPtesSegundaConvocatoriaCurso
        Exit Property
    End If
    Set m_objColUsuariosPtesSegundaConvocatoriaCurso = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PtesSegundaConvocatoriaCurso, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosPtesSegundaConvocatoriaCurso = m_objColUsuariosPtesSegundaConvocatoriaCurso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.ColUsuariosPtesSegundaConvocatoriaCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ColUsuariosPtesEnvioCorreoJefeSeguridadCurso() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosPtesEnvioCorreoJefeSeguridadCurso Is Nothing Then
        Set ColUsuariosPtesEnvioCorreoJefeSeguridadCurso = m_objColUsuariosPtesEnvioCorreoJefeSeguridadCurso
        Exit Property
    End If
    Set m_objColUsuariosPtesEnvioCorreoJefeSeguridadCurso = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PtesEnvioCorreoJefeSeguridadCurso, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosPtesEnvioCorreoJefeSeguridadCurso = m_objColUsuariosPtesEnvioCorreoJefeSeguridadCurso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.ColUsuariosPtesEnvioCorreoJefeSeguridadCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ColUsuariosHPSconHPSAPuntoDeCaducar() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosHPSconHPSAPuntoDeCaducar Is Nothing Then
        Set ColUsuariosHPSconHPSAPuntoDeCaducar = m_objColUsuariosHPSconHPSAPuntoDeCaducar
        Exit Property
    End If
    Set m_objColUsuariosHPSconHPSAPuntoDeCaducar = Constructor.getUsuariosParaIndicadores( _
                                                    EnumTipoIndicador.HPSAPuntoDeCaducar, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosHPSconHPSAPuntoDeCaducar = m_objColUsuariosHPSconHPSAPuntoDeCaducar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.ColUsuariosHPSconHPSAPuntoDeCaducar ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Set ColUsuariosHPSconHPSAPuntoDeCaducar(ByVal objNewValue As Variant)

    Set m_objColUsuariosHPSconHPSAPuntoDeCaducar = objNewValue

End Property

Public Property Get ColUsuariosHPSconHPSACaducadas() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosHPSconHPSACaducadas Is Nothing Then
        Set ColUsuariosHPSconHPSACaducadas = m_objColUsuariosHPSconHPSACaducadas
        Exit Property
    End If
    Set m_objColUsuariosHPSconHPSACaducadas = Constructor.getUsuariosParaIndicadores( _
                                                    EnumTipoIndicador.HPSACaducadas, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosHPSconHPSACaducadas = m_objColUsuariosHPSconHPSACaducadas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.ColUsuariosHPSconHPSACaducadas ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Set ColUsuariosHPSconHPSACaducadas(ByVal objNewValue As Variant)

    Set m_objColUsuariosHPSconHPSACaducadas = objNewValue

End Property

Public Property Get ColUsuariosHPSconHPSASolicitandose() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosHPSconHPSASolicitandose Is Nothing Then
        Set ColUsuariosHPSconHPSASolicitandose = m_objColUsuariosHPSconHPSASolicitandose
        Exit Property
    End If
    Set m_objColUsuariosHPSconHPSASolicitandose = Constructor.getUsuariosParaIndicadores( _
                                                    EnumTipoIndicador.HPSASolicitandose, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosHPSconHPSASolicitandose = m_objColUsuariosHPSconHPSASolicitandose
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.ColUsuariosHPSconHPSASolicitandose ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Set ColUsuariosHPSconHPSASolicitandose(ByVal objNewValue As Variant)

    Set m_objColUsuariosHPSconHPSASolicitandose = objNewValue
    

End Property

Public Property Get NPendientesCursoHPS() As Integer
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    If Me.ColUsuariosPendientesCursoHPS Is Nothing Then
        m_iNPendientesCursoHPS = 0
    Else
        m_iNPendientesCursoHPS = Me.ColUsuariosPendientesCursoHPS.Count
    End If
    NPendientesCursoHPS = m_iNPendientesCursoHPS
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.NPendientesCursoHPS ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property



Public Property Get NHPSAPuntoDeCaducar() As Integer

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
   
    If Me.ColUsuariosHPSconHPSAPuntoDeCaducar Is Nothing Then
        m_iNHPSAPuntoDeCaducar = 0
    Else
        m_iNHPSAPuntoDeCaducar = Me.ColUsuariosHPSconHPSAPuntoDeCaducar.Count
    End If
     NHPSAPuntoDeCaducar = m_iNHPSAPuntoDeCaducar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.NHPSAPuntoDeCaducar ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get NHPSACaducadas() As Integer

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
   
    If Me.ColUsuariosHPSconHPSACaducadas Is Nothing Then
        m_iNHPSACaducadas = 0
    Else
        m_iNHPSACaducadas = Me.ColUsuariosHPSconHPSACaducadas.Count
    End If
     NHPSACaducadas = m_iNHPSACaducadas
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.NHPSACaducadas ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get NHPSASolicitandose() As Integer

   On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
   
    If Me.ColUsuariosHPSconHPSASolicitandose Is Nothing Then
        m_iNHPSASolicitandose = 0
    Else
        m_iNHPSASolicitandose = Me.ColUsuariosHPSconHPSASolicitandose.Count
    End If
     NHPSASolicitandose = m_iNHPSASolicitandose
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.NHPSASolicitandose ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get NUsuariosPtesPrimeraConvocatoriaCurso() As Integer

   On Error GoTo errores

    Me.Error = ""
    m_Error = ""


    If Me.ColUsuariosPtesPrimeraConvocatoriaCurso Is Nothing Then
        m_iNUsuariosPtesPrimeraConvocatoriaCurso = 0
    Else
        m_iNUsuariosPtesPrimeraConvocatoriaCurso = Me.ColUsuariosPtesPrimeraConvocatoriaCurso.Count
    End If
     NUsuariosPtesPrimeraConvocatoriaCurso = m_iNUsuariosPtesPrimeraConvocatoriaCurso
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.NUsuariosPtesPrimeraConvocatoriaCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get NUsuariosPtesSegundaConvocatoriaCurso() As Integer

   On Error GoTo errores

    Me.Error = ""
    m_Error = ""


    If Me.ColUsuariosPtesSegundaConvocatoriaCurso Is Nothing Then
        m_iNUsuariosPtesSegundaConvocatoriaCurso = 0
    Else
        m_iNUsuariosPtesSegundaConvocatoriaCurso = Me.ColUsuariosPtesSegundaConvocatoriaCurso.Count
    End If
     NUsuariosPtesSegundaConvocatoriaCurso = m_iNUsuariosPtesSegundaConvocatoriaCurso
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.NUsuariosPtesSegundaConvocatoriaCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get NUsuariosPtesEnvioCorreoJefeSeguridadCurso() As Integer

   On Error GoTo errores

    Me.Error = ""
    m_Error = ""


    If Me.ColUsuariosPtesEnvioCorreoJefeSeguridadCurso Is Nothing Then
        m_iNUsuariosPtesEnvioCorreoJefeSeguridadCurso = 0
    Else
        m_iNUsuariosPtesEnvioCorreoJefeSeguridadCurso = Me.ColUsuariosPtesEnvioCorreoJefeSeguridadCurso.Count
    End If
     NUsuariosPtesEnvioCorreoJefeSeguridadCurso = m_iNUsuariosPtesEnvioCorreoJefeSeguridadCurso
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Indicador.NUsuariosPtesEnvioCorreoJefeSeguridadCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

