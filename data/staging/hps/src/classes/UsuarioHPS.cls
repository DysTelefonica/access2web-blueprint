VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioHPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public ID As String
Public DNI As String
Public Nombre As String
Public Apellido_1 As String
Public Apellido_2 As String
Public Telefono As String
Public Correo_e As String
Public F_Nacimiento As String
Public LugarNacimiento As String
Public Motivo_HPS As String
Public IDExpediente As String
Public F_Curso As String
Public CursoEnVigor As String
Public Requiere_Curso As String
Public FechaPrimeraConvocatoria As String
Public FechaSegundaConvocatoria As String
Public FechaCorreoNoCurso As String
Public F_Baja As String
Public LugarPrestacionServicio As String
Public IDEmpresaUsuario As String
Public IDEmpresaHPS As String
Public IDJuridicaContrato As String
Public FAvisoConcesion As String
Public RequiereComunicacionConcesion As String
Public FechaHPSConcesionMinima As String
Public CadenaContratistas As String

Public ParaAlta As EnumSiNo
Private m_objEmpresaUsuario As Suministrador
Private m_objEmpresaHPS As Suministrador

Private m_ObjExpediente As Expediente
Private m_ObjColAnexosActuales As Scripting.Dictionary
Private m_ObjColAnexosHistoricos As Scripting.Dictionary
Private m_ObjHPSNAC As HPS
Private m_ObjHPSUE As HPS
Private m_ObjHPSOTAN As HPS
Private m_ObjHPSESA As HPS
Private m_sNombreCompleto As String
Private m_sCurso_Realizado As String
Private m_ObjColObservaciones As Scripting.Dictionary
Private m_objColObservacionesCurso As Scripting.Dictionary
Private m_ObjColCamposParaCambios As Scripting.Dictionary
Private m_sIDNuevo As String
Private m_sF_BajaCalculada As String

Private m_sURLCarpetaAnexosActuales As String
Private m_sURLCarpetaAnexosHistoricos As String
Private m_eRequiereCursoCalculado As EnumSiNo
Private m_sRequiereCursoCalculadoTexto As String
Private m_eRequiereComunicacionConcesionCalculado As EnumSiNo
Private m_sRequiereComunicacionConcesionCalculadoTexto As String
Private m_sRequiere_PrimeraConvocatoriaCurso As String
Private m_sRequiere_SegundaConvocatoriaCurso As String
Private m_sRequiere_CorreoJefeSeguridadCurso As String
Private m_sObservaciones As String
Private m_sHPS_NAC_SIN_DATOS As String
Private m_sFechaHPSConcesionMinimaCalculada As String
Private m_objUsuarioSICA As UsuarioSICA
Private m_sCadenaContratistasCalculados As String


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get CadenaContratistasCalculados() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    
    m_sCadenaContratistasCalculados = Me.Expediente.CadenaContratistas
    m_Error = Me.Expediente.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    CadenaContratistasCalculados = m_sCadenaContratistasCalculados
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.CadenaContratistasCalculados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get FechaHPSConcesionMinimaCalculada() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sFechaHPSConcesionMinimaCalculada = getFechaHPSConcesionMinima(p_ID:=Me.ID, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    FechaHPSConcesionMinimaCalculada = m_sFechaHPSConcesionMinimaCalculada
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.FechaHPSConcesionMinimaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get HPS_NAC_SIN_DATOS() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sHPS_NAC_SIN_DATOS <> "" Then
        HPS_NAC_SIN_DATOS = m_sHPS_NAC_SIN_DATOS
        Exit Property
    End If
   
    HPS_NAC_SIN_DATOS = m_sHPS_NAC_SIN_DATOS
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.HPS_NAC_SIN_DATOS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get Observaciones() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sObservaciones <> "" Then
        Observaciones = m_sObservaciones
        Exit Property
    End If
    m_sObservaciones = getListaObservacionesParaExcel(Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Observaciones = m_sObservaciones
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.Observaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get EmpresaUsuario() As Suministrador

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objEmpresaUsuario Is Nothing Then
        Set EmpresaUsuario = m_objEmpresaUsuario
        Exit Property
    End If
    If Me.IDEmpresaUsuario = "" Then
        Exit Property
    End If
    Set m_objEmpresaUsuario = Constructor.getSuministrador(Me.IDEmpresaUsuario, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EmpresaUsuario = m_objEmpresaUsuario
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.EmpresaUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EmpresaHPS() As Suministrador

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objEmpresaHPS Is Nothing Then
        Set EmpresaHPS = m_objEmpresaHPS
        Exit Property
    End If
    If Me.IDEmpresaHPS = "" Then
        Exit Property
    End If
    Set m_objEmpresaHPS = Constructor.getSuministrador(Me.IDEmpresaHPS, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EmpresaHPS = m_objEmpresaHPS
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.EmpresaHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get F_BajaCalculada() As String

    On Error GoTo errores
    
    Me.Error = ""
    If IsDate(m_sF_BajaCalculada) Then
        F_BajaCalculada = m_sF_BajaCalculada
        Exit Property
    End If
    If IsDate(Me.F_Baja) Then
        m_sF_BajaCalculada = Me.F_Baja
        F_BajaCalculada = m_sF_BajaCalculada
        Exit Property
    End If
    
    m_sF_BajaCalculada = getFBajaHPS(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    F_BajaCalculada = m_sF_BajaCalculada
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.F_BajaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get URLCarpetaAnexosActuales() As String
    Dim m_URLCarpetaAnexosGlobales As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_sURLCarpetaAnexosActuales <> "" Then
        URLCarpetaAnexosActuales = m_sURLCarpetaAnexosActuales
        Exit Property
    End If
    If Me.DNI = "" Then
        Exit Property
    End If
    m_URLCarpetaAnexosGlobales = m_ObjEntorno.URLCarpetaAnexos
    If m_URLCarpetaAnexosGlobales = "" Then
        Exit Property
    End If
    If Right(m_URLCarpetaAnexosGlobales, 1) <> "\" Then
        m_URLCarpetaAnexosGlobales = m_URLCarpetaAnexosGlobales & "\"
    End If
    ' [anexos subcarpetas] Para usuarios activos, los anexos actuales viven
    ' DENTRO del <DNI>\ como subcarpeta Actual\ (misma convencion que
    ' URLCarpetaAnexosHistoricos usa para HISTORICO\). Sin este sufijo,
    ' el per-anexo move (PasarAnexo) buscaba el archivo directamente en
    ' <DNI>\ y no en <DNI>\Actual\ donde estaba realmente.
    m_sURLCarpetaAnexosActuales = m_URLCarpetaAnexosGlobales & Me.DNI & "\Actual\"
    URLCarpetaAnexosActuales = m_sURLCarpetaAnexosActuales

    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.URLCarpetaAnexosActuales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLCarpetaAnexosHistoricos() As String
    Dim m_URLCarpetaAnexosGlobales As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_sURLCarpetaAnexosHistoricos <> "" Then
        URLCarpetaAnexosHistoricos = m_sURLCarpetaAnexosHistoricos
        Exit Property
    End If
    If Me.DNI = "" Then
        Exit Property
    End If
    ' [anexos subcarpetas] Para usuarios activos, los anexos historicos viven
    ' en <DNI>\HISTORICO\ como SIBLING de <DNI>\Actual\ (no anidado). Cada
    ' anexo va a una u otra subcarpeta segun su EsHistorico. El root de red
    ' ANEXOS_HISTORICO_PATH aplica solo a usuarios completamente historicos
    ' (TbUsuariosHistoricos).
    m_URLCarpetaAnexosGlobales = m_ObjEntorno.URLCarpetaAnexos
    If m_URLCarpetaAnexosGlobales = "" Then
        Exit Property
    End If
    If Right(m_URLCarpetaAnexosGlobales, 1) <> "\" Then
        m_URLCarpetaAnexosGlobales = m_URLCarpetaAnexosGlobales & "\"
    End If
    m_sURLCarpetaAnexosHistoricos = m_URLCarpetaAnexosGlobales & Me.DNI & "\HISTORICO\"
    URLCarpetaAnexosHistoricos = m_sURLCarpetaAnexosHistoricos

    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.URLCarpetaAnexosHistoricos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property



Public Property Get Expediente() As Expediente
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjExpediente Is Nothing Then
        Set Expediente = m_ObjExpediente
        Exit Property
    End If
    If Me.IDExpediente = "" Then
        
        Exit Property
    End If
    Set m_ObjExpediente = Constructor.getExpediente(Me.IDExpediente, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Expediente = m_ObjExpediente
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.expediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set Expediente(ByVal objNewValue As Expediente)

    Set m_ObjExpediente = objNewValue

End Property


Public Property Get colAnexosActuales() As Scripting.Dictionary
    

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAnexosActuales Is Nothing Then
        Set colAnexosActuales = m_ObjColAnexosActuales
        Exit Property
    End If
    Set m_ObjColAnexosActuales = Constructor.getAnexosUsuarioHPS(Me.ID, EnumSiNo.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colAnexosActuales = m_ObjColAnexosActuales
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.colAnexosActuales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set colAnexosActuales(ByVal objNewValue As Variant)

    Set m_ObjColAnexosActuales = objNewValue

End Property

Public Property Get colAnexosHistoricos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAnexosHistoricos Is Nothing Then
        Set colAnexosHistoricos = m_ObjColAnexosHistoricos
        Exit Property
    End If
    Set m_ObjColAnexosHistoricos = Constructor.getAnexosUsuarioHPS(Me.ID, EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colAnexosHistoricos = m_ObjColAnexosHistoricos
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.colAnexosHistoricos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colAnexosHistoricos(ByVal objNewValue As Variant)

    Set m_ObjColAnexosHistoricos = objNewValue

End Property

Public Property Get HPSNAC() As HPS
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjHPSNAC Is Nothing Then
        Set HPSNAC = m_ObjHPSNAC
        Exit Property
    End If
    Set m_ObjHPSNAC = getHPS("Nacional", Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set HPSNAC = m_ObjHPSNAC
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.HPSNAC ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set HPSNAC(ByVal sNewValue As HPS)
    
    Set m_ObjHPSNAC = sNewValue
End Property
Public Property Get HPSUE() As HPS

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjHPSUE Is Nothing Then
        Set HPSUE = m_ObjHPSUE
        Exit Property
    End If
    Set m_ObjHPSUE = getHPS("UE", Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set HPSUE = m_ObjHPSUE
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.HPSUE ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set HPSUE(ByVal sNewValue As HPS)
    
    Set m_ObjHPSUE = sNewValue
End Property
Public Property Get HPSOTAN() As HPS

    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjHPSOTAN Is Nothing Then
        Set HPSOTAN = m_ObjHPSOTAN
        Exit Property
    End If
    Set m_ObjHPSOTAN = getHPS("OTAN", Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set HPSOTAN = m_ObjHPSOTAN
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.HPSOTAN ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set HPSOTAN(ByVal sNewValue As HPS)
    
    Set m_ObjHPSOTAN = sNewValue
End Property
Public Property Get HPSESA() As HPS

    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjHPSESA Is Nothing Then
        Set HPSESA = m_ObjHPSESA
        Exit Property
    End If
    Set m_ObjHPSESA = getHPS("ESA", Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set HPSESA = m_ObjHPSESA
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.HPSESA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set HPSESA(ByVal sNewValue As HPS)
    
    Set m_ObjHPSESA = sNewValue
End Property


Public Property Get NombreCompleto() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreCompleto <> "" Then
        NombreCompleto = m_sNombreCompleto
        Exit Property
    End If
    If Me.Nombre = "" And Me.Apellido_1 = "" And Me.Apellido_2 = "" Then
        Exit Property
    End If
    m_sNombreCompleto = Me.Nombre & " " & Me.Apellido_1 & " " & Me.Apellido_2
    NombreCompleto = m_sNombreCompleto
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.NombreCompleto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColObservaciones() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColObservaciones Is Nothing Then
        Set ColObservaciones = m_ObjColObservaciones
        Exit Property
    End If
    
    Set m_ObjColObservaciones = Constructor.getObservaciones(Me.ID, "General", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColObservaciones = m_ObjColObservaciones
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.ColObservaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColObservaciones(ByVal objNewValue As Variant)
    
    Set m_ObjColObservaciones = objNewValue

End Property
Public Property Get ColObservacionesCurso() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColObservacionesCurso Is Nothing Then
        Set ColObservacionesCurso = m_objColObservacionesCurso
        Exit Property
    End If
    
    Set m_objColObservacionesCurso = Constructor.getObservaciones(Me.ID, "Curso", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColObservacionesCurso = m_objColObservacionesCurso
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.ColObservacionesCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColObservacionesCurso(ByVal objNewValue As Variant)
    
    Set m_objColObservacionesCurso = objNewValue

End Property
Public Property Get Curso_Realizado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCurso_Realizado <> "" Then
        Curso_Realizado = m_sCurso_Realizado
        Exit Property
    End If
    If IsDate(Me.F_Curso) Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        m_sCurso_Realizado = "Sí"
    Else
        m_sCurso_Realizado = "No"
    End If
    Curso_Realizado = m_sCurso_Realizado
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.Curso_Realizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get RequiereCursoCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eRequiereCursoCalculado <> Empty Then
        RequiereCursoCalculado = m_eRequiereCursoCalculado
        Exit Property
    End If
    If IsDate(Me.FechaHPSConcesionMinima) Then
        If Me.Curso_Realizado <> "Sí" Then
            m_eRequiereCursoCalculado = EnumSiNo.Sí
        Else
            m_eRequiereCursoCalculado = EnumSiNo.No
        End If
    Else
        m_eRequiereCursoCalculado = EnumSiNo.No
    End If
    
    RequiereCursoCalculado = m_eRequiereCursoCalculado
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.RequiereCursoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get RequiereCursoCalculadoTexto() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.RequiereCursoCalculado = EnumSiNo.Sí Then
        m_sRequiereCursoCalculadoTexto = "Sí"
    ElseIf Me.RequiereCursoCalculado = EnumSiNo.No Then
        m_sRequiereCursoCalculadoTexto = "No"
    End If
    
    RequiereCursoCalculadoTexto = m_sRequiereCursoCalculadoTexto
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.RequiereCursoCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get Requiere_PrimeraConvocatoriaCurso() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sRequiere_PrimeraConvocatoriaCurso <> "" Then
        Requiere_PrimeraConvocatoriaCurso = m_sRequiere_PrimeraConvocatoriaCurso
        Exit Property
    End If
    If Me.RequiereCursoCalculadoTexto = "Sí" Then
        If IsDate(Me.FechaPrimeraConvocatoria) Then
            m_sRequiere_PrimeraConvocatoriaCurso = "No"
           
        Else
            m_sRequiere_PrimeraConvocatoriaCurso = "Sí"
            
        End If
        
    Else
        m_sRequiere_PrimeraConvocatoriaCurso = "No"
        
    End If
    Requiere_PrimeraConvocatoriaCurso = m_sRequiere_PrimeraConvocatoriaCurso
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.Requiere_PrimeraConvocatoriaCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get Requiere_SegundaConvocatoriaCurso() As String
    
    Dim m_DiasEnPrimeraConvocatoriaSinCurso As Long
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sRequiere_SegundaConvocatoriaCurso <> "" Then
        Requiere_SegundaConvocatoriaCurso = m_sRequiere_SegundaConvocatoriaCurso
        Exit Property
    End If
    If Me.RequiereCursoCalculadoTexto = "Sí" Then
        If IsDate(Me.FechaPrimeraConvocatoria) Then
            
            If Not IsDate(Me.FechaSegundaConvocatoria) Then
                
                
                m_DiasEnPrimeraConvocatoriaSinCurso = CLng(DateDiff("d", CDate(Now()), CDate(Me.FechaPrimeraConvocatoria)))
                If m_DiasEnPrimeraConvocatoriaSinCurso < 0 Then
                    m_DiasEnPrimeraConvocatoriaSinCurso = m_DiasEnPrimeraConvocatoriaSinCurso * -1
                End If
                If m_DiasEnPrimeraConvocatoriaSinCurso >= m_ObjEntorno.DiasEnPrimeraConvocatoriaSinCurso Then
                    m_sRequiere_SegundaConvocatoriaCurso = "Sí"
                Else
                    m_sRequiere_SegundaConvocatoriaCurso = "No"
                End If
                
            Else
                m_sRequiere_SegundaConvocatoriaCurso = "No"
                
            End If
        Else
            
            m_sRequiere_SegundaConvocatoriaCurso = "No"
           
        End If
        
    Else
       
        m_sRequiere_SegundaConvocatoriaCurso = "No"
        
    End If
    Requiere_SegundaConvocatoriaCurso = m_sRequiere_SegundaConvocatoriaCurso
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.Requiere_SegundaConvocatoriaCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get Requiere_CorreoJefeSeguridadCurso() As String
    
    Dim m_DiasEnPrimeraConvocatoriaSinCurso As Long
    Dim m_DiasEnSegundaConvocatoriaSinCurso As Long
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sRequiere_CorreoJefeSeguridadCurso <> "" Then
        Requiere_CorreoJefeSeguridadCurso = m_sRequiere_CorreoJefeSeguridadCurso
        Exit Property
    End If
    If Me.RequiereCursoCalculadoTexto = "Sí" Then
        If IsDate(Me.FechaPrimeraConvocatoria) Then
            
            If Not IsDate(Me.FechaSegundaConvocatoria) Then
                m_sRequiere_CorreoJefeSeguridadCurso = "No"
                
                
            Else
                
                If Not IsDate(Me.FechaCorreoNoCurso) Then
                    m_DiasEnSegundaConvocatoriaSinCurso = CLng(DateDiff("d", CDate(Now()), CDate(Me.FechaSegundaConvocatoria)))
                    
                    
                    If m_DiasEnSegundaConvocatoriaSinCurso < 0 Then
                        m_DiasEnSegundaConvocatoriaSinCurso = m_DiasEnSegundaConvocatoriaSinCurso * -1
                    End If
                    If m_DiasEnSegundaConvocatoriaSinCurso >= m_ObjEntorno.DiasEnSegundaConvocatoriaSinCurso Then
                        m_sRequiere_CorreoJefeSeguridadCurso = "Sí"
                    Else
                        m_sRequiere_CorreoJefeSeguridadCurso = "No"
                    End If
                Else
                    m_sRequiere_CorreoJefeSeguridadCurso = "No"
                End If
            End If
        Else
           
            m_sRequiere_CorreoJefeSeguridadCurso = "No"
        End If
        
    Else
       
        m_sRequiere_CorreoJefeSeguridadCurso = "No"
    End If
    Requiere_CorreoJefeSeguridadCurso = m_sRequiere_CorreoJefeSeguridadCurso
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.Requiere_CorreoJefeSeguridadCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get RequiereComunicacionConcesionCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eRequiereComunicacionConcesionCalculado <> Empty Then
        RequiereComunicacionConcesionCalculado = m_eRequiereComunicacionConcesionCalculado
        Exit Property
    End If
    
    If IsDate(Me.FechaHPSConcesionMinima) Then
        If IsDate(Me.FAvisoConcesion) Then
            m_eRequiereComunicacionConcesionCalculado = EnumSiNo.No
        Else
            m_eRequiereComunicacionConcesionCalculado = EnumSiNo.Sí
        End If
        
    Else
        m_eRequiereComunicacionConcesionCalculado = EnumSiNo.No
    End If
    
    RequiereComunicacionConcesionCalculado = m_eRequiereComunicacionConcesionCalculado
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.RequiereComunicacionConcesionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get RequiereComunicacionConcesionCalculadoTexto() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.RequiereComunicacionConcesionCalculado = EnumSiNo.Sí Then
        m_sRequiereComunicacionConcesionCalculadoTexto = "Sí"
    ElseIf Me.RequiereComunicacionConcesionCalculado = EnumSiNo.No Then
        m_sRequiereComunicacionConcesionCalculadoTexto = "No"
    End If
    RequiereComunicacionConcesionCalculadoTexto = m_sRequiereComunicacionConcesionCalculadoTexto
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.RequiereComunicacionConcesionCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property



Public Property Get IDNuevo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    m_sIDNuevo = getID("TbUsuarios", "ID", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDNuevo = m_sIDNuevo
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.IDNuevo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioSICA() As UsuarioSICA
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
     
    If Not m_objUsuarioSICA Is Nothing Then
       Set UsuarioSICA = m_objUsuarioSICA
       Exit Property
    End If
    If Me.ID = "" Then
        Exit Property
    End If
    Set m_objUsuarioSICA = Constructor.getUsuarioSICA(Me.ID, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioSICA = m_objUsuarioSICA
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.UsuarioSICA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set UsuarioSICA(ByVal objNewValue As UsuarioSICA)

    Set m_objUsuarioSICA = objNewValue

End Property

Private Function getFBajaHPS( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    
    Dim rcdDatos As DAO.Recordset
    
    Dim m_SQL As String
    
    On Error GoTo errores
    
     m_SQL = "SELECT TbHPS.F_Baja " & _
            "FROM TbHPS " & _
            "WHERE ((Not (TbHPS.F_Baja) Is Null) And ((TbHPS.IDUsuario) =" & Me.ID & ")) " & _
            "ORDER BY TbHPS.F_Baja;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            getFBajaHPS = .Fields("F_Baja")
            rcdDatos.Close
        Set rcdDatos = Nothing
            Exit Function
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    m_SQL = "SELECT TbHPS.F_Caducidad " & _
            "FROM TbHPS " & _
            "WHERE (((TbHPS.Renovacion)='No') AND ((TbHPS.IDUsuario)=" & Me.ID & ") AND ((TbHPS.F_Caducidad)<Now()));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            getFBajaHPS = .Fields("F_Caducidad")
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFBajaHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




Public Function MotivoNoOK( _
                                Optional ByRef p_ObjUsuarioHPSAlInicio As UsuarioHPS, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim blnValorRepetido As Boolean
    
    On Error GoTo errores
        
    With Me
        If .ParaAlta = Empty Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
            
        End If
        If .ParaAlta = EnumSiNo.No Then
            Set p_ObjUsuarioHPSAlInicio = Constructor.getUsuarioHPS(.ID, , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If p_ObjUsuarioHPSAlInicio Is Nothing Then
                MotivoNoOK = "Es una edición de usuario y no se ha encontrado registrado"
                Exit Function
            End If
        End If
        If .DNI = "" Then
            MotivoNoOK = "El DNI es obligatorio"
            Exit Function
        End If
        If .ParaAlta = EnumSiNo.No Then
            If .DNI <> p_ObjUsuarioHPSAlInicio.DNI Then
                blnValorRepetido = ValorRepetido("TbUsuarios", "DNI", .DNI, getdb(), p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If blnValorRepetido Then
                    MotivoNoOK = "El DNI ha sido registrado para otro usuario"
                    Exit Function
                End If
            End If
        End If
        If .Nombre = "" Then
            MotivoNoOK = "El Nombre es obligatorio"
            Exit Function
        End If
        If .Apellido_1 = "" Then
            MotivoNoOK = "El Apellido_1 es obligatorio"
            Exit Function
        End If
        If .Telefono = "" Then
            MotivoNoOK = "El Telefono es obligatorio"
            Exit Function
        End If
'        If .EmpresaUsuario Is Nothing Then
'            p_Error = .Error
'            If p_Error <> "" Then
'                Err.Raise 1000
'            End If
'            MotivoNoOK = "La empresa de usuario es obligatoria"
'            Exit Function
'        Else
'            p_Error = .Error
'            If p_Error <> "" Then
'                Err.Raise 1000
'            End If
'        End If
        If .EmpresaHPS Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "La empresa de Tramitación del HPS es obligatoria"
            Exit Function
        Else
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        If InStr(1, .Correo_e, "@") = 0 Then
            MotivoNoOK = "Falta el Correo electrónico o no tiene la forma adecuada"
            Exit Function
        End If
        If .ParaAlta = EnumSiNo.No Then
            If .Correo_e <> p_ObjUsuarioHPSAlInicio.Correo_e Then
                blnValorRepetido = ValorRepetido("TbUsuarios", "Correo_e", .Correo_e, getdb(), p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If blnValorRepetido Then
                    MotivoNoOK = "El Correo_e ha sido registrado para otro usuario"
                    Exit Function
                End If
            End If
        End If
        If .Motivo_HPS = "" Then
            MotivoNoOK = "El Motivo_HPS es obligatorio"
            Exit Function
        End If
        If .Expediente Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "El expediente es obligatorio"
            Exit Function
        Else
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .Expediente.CadenaContratistas = "" Then
            MotivoNoOK = "El expediente ha de tener al menos un contratista"
            Exit Function
        End If
        If .CursoEnVigor = "Sí" Then
            If Not IsDate(.F_Curso) Then
                MotivoNoOK = "No Hay Fecha de curso y se ha indicado que está vigente"
                Exit Function
            End If
        End If
        
        If Not IsDate(.F_Nacimiento) Then
            MotivoNoOK = "La F_Nacimiento es obligatoria"
            Exit Function
        End If
        m_eRequiereCursoCalculado = Empty
        If IsDate(Me.F_Curso) Then
            If IsDate(.FechaPrimeraConvocatoria) Then
                MotivoNoOK = "Tiene registrada la fecha de curso, por lo que no se puede registrar la fecha de primera convocatoria"
                Exit Function
            End If
            If IsDate(.FechaSegundaConvocatoria) Then
                MotivoNoOK = "Tiene registrada la fecha de curso, por lo que no se puede registrar la fecha de segunda convocatoria"
                Exit Function
            End If
            If IsDate(.FechaCorreoNoCurso) Then
                MotivoNoOK = "Tiene registrada la fecha de curso, por lo que no se puede registrar la fecha de correo al Jefe de Seguridad"
                Exit Function
            End If
        End If
        
        If IsDate(.FechaCorreoNoCurso) Then
            If Not IsDate(.FechaSegundaConvocatoria) Or Not IsDate(.FechaPrimeraConvocatoria) Then
                MotivoNoOK = "No se puede rellenar la fecha de correo no curso si no se ha rellenado antes las de primera y segunda convocatoria"
                Exit Function
            End If
        End If
        If IsDate(.FechaSegundaConvocatoria) Then
            If Not IsDate(.FechaPrimeraConvocatoria) Then
                MotivoNoOK = "No se puede rellenar la fecha de segunda convocatoria si no se ha rellenado antes la de primera"
                Exit Function
            End If
        End If
        If IsDate(.FechaCorreoNoCurso) Then
            If CDate(.FechaCorreoNoCurso) < CDate(.FechaSegundaConvocatoria) Then
                MotivoNoOK = "La fecha del correo al Jefe de Seguridad no puede ser anterior a la de segunda convocatoria"
                Exit Function
            End If
        End If
        If IsDate(.FechaSegundaConvocatoria) Then
            If CDate(.FechaSegundaConvocatoria) < CDate(.FechaPrimeraConvocatoria) Then
                MotivoNoOK = "La fecha de la segunda convocatoria al curso no puede ser anterior a la de la primera"
                Exit Function
            End If
        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    Dim m_ObjAnexoUsuarioHPS As AnexoUsuarioHPS
    Dim m_ObjAnexoUsuarioHistorico As AnexoUsuarioHistorico
    Dim m_IDAnexo As Variant
    Dim m_URLAnexo As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    If Not Me.colAnexosActuales Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDAnexo In Me.colAnexosActuales
            If m_IDAnexo = Empty Then GoTo siguiente
            Set m_ObjAnexoUsuarioHPS = Me.colAnexosActuales(m_IDAnexo)
            m_URLAnexo = m_ObjAnexoUsuarioHPS.URLAnexo
            If fso.FileExists(m_URLAnexo) Then
                If FicheroAbierto(m_URLAnexo) Then
                    p_Error = "Hay al menos un anexo que está abierto o es de sólo lectura"
                    Err.Raise 1000
                End If
                
            End If
            Set m_ObjAnexoUsuarioHPS = Nothing
siguiente:
        Next
    Else
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If Not Me.colAnexosHistoricos Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDAnexo In Me.colAnexosHistoricos
            If m_IDAnexo = Empty Then GoTo siguiente1
            Set m_ObjAnexoUsuarioHPS = Me.colAnexosHistoricos(m_IDAnexo)
            m_URLAnexo = m_ObjAnexoUsuarioHPS.URLAnexo
            If fso.FileExists(m_URLAnexo) Then
                If FicheroAbierto(m_URLAnexo) Then
                    p_Error = "Hay al menos un anexo que está abierto o es de sólo lectura"
                    Err.Raise 1000
                End If
                
            End If
            Set m_ObjAnexoUsuarioHPS = Nothing
siguiente1:
        Next
    Else
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
        
        
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                            Optional ByRef p_ObjUsuarioHPSAlInicio As UsuarioHPS, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_RequiereCurso As String
    Dim m_RequiereComunicacion As String
    Dim m_UsuarioSICA As UsuarioSICA
    Dim m_FechaBajaCalculada As String
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If Me.ParaAlta = 0 Then
        p_Error = "No se sabe si es alta o edición"
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSiNo.No Then
        If p_ObjUsuarioHPSAlInicio Is Nothing Then
            Set p_ObjUsuarioHPSAlInicio = Constructor.getUsuarioHPS(Me.ID, , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    m_MotivoNoOK = MotivoNoOK(p_ObjUsuarioHPSAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSiNo.Sí Then
        m_SQL = "TbUsuarios"
        If Me.ID = "" Then
             Me.ID = Me.IDNuevo
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
       
    Else
        m_SQL = "SELECT * FROM TbUsuarios " & _
            "WHERE ID=" & Me.ID & ";"
        
    End If
    If IsDate(Me.F_Curso) And Me.CursoEnVigor = "Sí" Then
        Me.FechaPrimeraConvocatoria = ""
        Me.FechaSegundaConvocatoria = ""
        Me.FechaCorreoNoCurso = ""
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Me.ParaAlta = EnumSiNo.No Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If Me.ParaAlta = EnumSiNo.Sí Then
            .AddNew
                .Fields("ID") = Me.ID
                .Fields("DNI") = Me.DNI
                .Fields("Nombre") = Me.Nombre
                .Fields("Apellido_1") = Me.Apellido_1
                If Me.Apellido_2 <> "" Then
                    .Fields("Apellido_2") = Me.Apellido_2
                End If
                .Fields("Telefono") = Me.Telefono
                .Fields("Correo_e") = Me.Correo_e
                .Fields("F_Nacimiento") = Me.F_Nacimiento
                If Me.LugarNacimiento <> "" Then
                    .Fields("LugarNacimiento") = Me.LugarNacimiento
                End If
                .Fields("Motivo_HPS") = Me.Motivo_HPS
                .Fields("IDExpediente") = Me.IDExpediente
                If IsDate(Me.F_Curso) Then
                    .Fields("F_Curso") = Me.F_Curso
                End If
                If Me.CursoEnVigor = "Sí" Or Me.CursoEnVigor = "No" Then
                    .Fields("CursoEnVigor") = Me.CursoEnVigor
                End If
                m_RequiereCurso = Me.RequiereCursoCalculadoTexto
                If m_RequiereCurso = "Sí" Or m_RequiereCurso = "No" Then
                    .Fields("Requiere_Curso") = m_RequiereCurso
                End If
                If IsDate(Me.F_BajaCalculada) Then
                    .Fields("F_Baja") = Me.F_BajaCalculada
                End If
                If Me.LugarPrestacionServicio <> "" Then
                    .Fields("LugarPrestacionServicio") = Me.LugarPrestacionServicio
                End If
                .Fields("IDEmpresaUsuario") = Me.IDEmpresaUsuario
                .Fields("IDEmpresaHPS") = Me.IDEmpresaHPS
                Me.CadenaContratistas = Me.CadenaContratistasCalculados
                .Fields("CadenaContratistas") = Me.CadenaContratistas
                
                If IsDate(Me.FAvisoConcesion) Then
                    .Fields("FAvisoConcesion") = Me.FAvisoConcesion
                End If
                m_RequiereComunicacion = Me.RequiereComunicacionConcesionCalculadoTexto
                If m_RequiereComunicacion = "Sí" Or m_RequiereComunicacion = "No" Then
                    .Fields("RequiereComunicacionConcesion") = m_RequiereComunicacion
                End If
                If IsDate(Me.FechaPrimeraConvocatoria) Then
                    .Fields("FechaPrimeraConvocatoria") = Me.FechaPrimeraConvocatoria
                End If
                If IsDate(Me.FechaSegundaConvocatoria) Then
                    .Fields("FechaSegundaConvocatoria") = Me.FechaSegundaConvocatoria
                End If
                If IsDate(Me.FechaCorreoNoCurso) Then
                    .Fields("FechaCorreoNoCurso") = Me.FechaCorreoNoCurso
                End If
            .Update
        Else
            .Edit
                .Fields("DNI") = Me.DNI
                .Fields("Nombre") = Me.Nombre
                .Fields("Apellido_1") = Me.Apellido_1
                If Me.Apellido_2 <> "" Then
                    .Fields("Apellido_2") = Me.Apellido_2
                Else
                    .Fields("Apellido_2") = Null
                End If
                .Fields("Telefono") = Me.Telefono
                .Fields("Correo_e") = Me.Correo_e
                .Fields("F_Nacimiento") = Me.F_Nacimiento
                If Me.LugarNacimiento <> "" Then
                    .Fields("LugarNacimiento") = Me.LugarNacimiento
                Else
                    .Fields("LugarNacimiento") = Null
                End If
                .Fields("Motivo_HPS") = Me.Motivo_HPS
                .Fields("IDExpediente") = Me.IDExpediente
                If IsDate(Me.F_Curso) Then
                    .Fields("F_Curso") = Me.F_Curso
                Else
                    .Fields("F_Curso") = Null
                End If
                If Me.CursoEnVigor = "Sí" Or Me.CursoEnVigor = "No" Then
                    .Fields("CursoEnVigor") = Me.CursoEnVigor
                Else
                    .Fields("CursoEnVigor") = Null
                End If
                m_RequiereCurso = Me.RequiereCursoCalculadoTexto
                If m_RequiereCurso = "Sí" Or m_RequiereCurso = "No" Then
                    .Fields("Requiere_Curso") = m_RequiereCurso
                Else
                    .Fields("Requiere_Curso") = Null
                End If
                m_FechaBajaCalculada = Me.F_BajaCalculada
                If IsDate(m_FechaBajaCalculada) Then
                    .Fields("F_Baja") = m_FechaBajaCalculada
                Else
                    .Fields("F_Baja") = Null
                End If
                If IsDate(m_FechaBajaCalculada) And Not IsDate(p_ObjUsuarioHPSAlInicio.F_Baja) Then
                    Set m_UsuarioSICA = Me.UsuarioSICA
                    p_Error = Me.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_UsuarioSICA Is Nothing Then
                        m_SQL = "UPDATE TbUsuariosSICA SET TbUsuariosSICA.FBajaTdE = #" & Format(m_FechaBajaCalculada, "mm/dd/yyyy") & "# " & _
                                "WHERE (((TbUsuariosSICA.ID)='" & m_UsuarioSICA.ID & _
                                "') AND (Not (TbUsuariosSICA.FAltaTdE) Is Null) " & _
                                "AND ((TbUsuariosSICA.FBajaTdE) Is Null));"
                        getdb().Execute m_SQL
                        
                        m_SQL = "UPDATE TbUsuariosSICA SET TbUsuariosSICA.FbajaTSOL = #" & Format(m_FechaBajaCalculada, "mm/dd/yyyy") & "# " & _
                                "WHERE (((TbUsuariosSICA.ID)='" & m_UsuarioSICA.ID & _
                                "') AND (Not (TbUsuariosSICA.FAltaTSOL) Is Null) " & _
                                "AND ((TbUsuariosSICA.FBajaTdE) Is Null));"
                        getdb().Execute m_SQL
                        
                    End If
                End If
                If Me.LugarPrestacionServicio <> "" Then
                    .Fields("LugarPrestacionServicio") = Me.LugarPrestacionServicio
                Else
                    .Fields("LugarPrestacionServicio") = Null
                End If
                .Fields("IDEmpresaUsuario") = Me.IDEmpresaUsuario
                .Fields("IDEmpresaHPS") = Me.IDEmpresaHPS
                Me.CadenaContratistas = Me.CadenaContratistasCalculados
                .Fields("CadenaContratistas") = Me.CadenaContratistas
                If IsDate(Me.FAvisoConcesion) Then
                    .Fields("FAvisoConcesion") = Me.FAvisoConcesion
                Else
                    .Fields("FAvisoConcesion") = Null
                End If
                m_RequiereComunicacion = Me.RequiereComunicacionConcesionCalculadoTexto
                If m_RequiereComunicacion = "Sí" Or m_RequiereComunicacion = "No" Then
                    .Fields("RequiereComunicacionConcesion") = m_RequiereComunicacion
                Else
                    .Fields("RequiereComunicacionConcesion") = Null
                End If
                If IsDate(Me.FechaPrimeraConvocatoria) Then
                    .Fields("FechaPrimeraConvocatoria") = Me.FechaPrimeraConvocatoria
                Else
                    .Fields("FechaPrimeraConvocatoria") = Null
                End If
                If IsDate(Me.FechaSegundaConvocatoria) Then
                    .Fields("FechaSegundaConvocatoria") = Me.FechaSegundaConvocatoria
                Else
                    .Fields("FechaSegundaConvocatoria") = Null
                End If
                If IsDate(Me.FechaCorreoNoCurso) Then
                    .Fields("FechaCorreoNoCurso") = Me.FechaCorreoNoCurso
                Else
                    .Fields("FechaCorreoNoCurso") = Null
                End If
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    cache_usuario_regenerar p_ID:=Me.ID, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.registrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function FechaHPSConcesionMinimaRegistrar( _
                                                    Optional ByRef p_Error As String _
                                                 ) As String
                                         
     
    On Error GoTo errores
     
    Me.Error = ""
    p_Error = ""
    RefreshHpsUserCacheAfterMutation p_IDUsuario:=Me.ID, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
     
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.FechaHPSConcesionMinimaRegistrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function BorrarCurso( _
                            Optional ByRef p_Error As String _
                            ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    If Me.ID = "" Then
        Exit Function
    End If
    m_SQL = "UPDATE TbUsuarios SET TbUsuarios.F_Curso = Null,TbUsuarios.CursoEnVigor = Null " & _
            "WHERE (((TbUsuarios.ID)=" & Me.ID & "));"
    getdb().Execute m_SQL
    cache_usuario_regenerar p_ID:=Me.ID, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerContadores p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.BorrarCurso ha devuelto el error: " & Err.Description
    End If
End Function
Public Function CadenaContratistasGrabar( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    
    Dim m_SQL As String
   
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    Me.CadenaContratistas = Me.CadenaContratistasCalculados
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.CadenaContratistas <> "" Then
        m_SQL = "UPDATE TbUsuarios SET CadenaContratistas = '" & Me.CadenaContratistas & "' " & _
                "WHERE ID=" & Me.ID & ";"
    Else
        m_SQL = "UPDATE TbUsuarios SET CadenaContratistas = Null " & _
                "WHERE ID=" & Me.ID & ";"
    End If
    getdb().Execute m_SQL
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.CadenaContratistasGrabar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
    Dim m_ID As String
    Dim m_ObjColURL As Collection
    Dim m_URL As Variant
    Dim m_ObjColObservaciones As Scripting.Dictionary
    Dim m_IDObs As Variant
    Dim m_ObjObServacion As Observacion
    Dim m_ColAnexos As Scripting.Dictionary
    Dim m_Anexo As AnexoUsuarioHPS
    Dim m_IDAnexo As Variant
    Dim m_URLDirDNI As String
    Dim m_SQL As String
    Dim m_URLDirAnexosActuales As String
    Dim m_URlDirAnexosHistoricos As String
    On Error GoTo errores
    
    m_ID = Me.ID
    
    
    
    m_SQL = "DELETE * FROM TbUsuarios " & _
            "WHERE ID=" & Me.ID & ";"
    getdb().Execute m_SQL
    
    Set m_ObjColObservaciones = Me.ColObservacionesCurso
    If Not m_ObjColObservaciones Is Nothing Then
        For Each m_IDObs In m_ObjColObservaciones
            Set m_ObjObServacion = m_ObjColObservaciones(m_IDObs)
            m_ObjObServacion.Borrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjObServacion = Nothing
        Next
    End If
    Set m_ObjColObservaciones = Me.ColObservaciones
    
    If Not m_ObjColObservaciones Is Nothing Then
        For Each m_IDObs In m_ObjColObservaciones
            Set m_ObjObServacion = m_ObjColObservaciones(m_IDObs)
            m_ObjObServacion.Borrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjObServacion = Nothing
        Next
    End If
    Set Me.colAnexosActuales = Nothing
    Set m_ColAnexos = Me.colAnexosActuales
    If Not m_ColAnexos Is Nothing Then
        For Each m_IDAnexo In m_ColAnexos
            Set m_Anexo = m_ColAnexos(m_IDAnexo)
            m_Anexo.Borrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_Anexo = Nothing
        Next
    End If
    Set Me.colAnexosHistoricos = Nothing
    Set m_ColAnexos = Me.colAnexosHistoricos
    If Not m_ColAnexos Is Nothing Then
        For Each m_IDAnexo In m_ColAnexos
            Set m_Anexo = m_ColAnexos(m_IDAnexo)
            m_Anexo.Borrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_Anexo = Nothing
        Next
    End If
    m_URLDirAnexosActuales = Me.URLCarpetaAnexosActuales
    m_URlDirAnexosHistoricos = Me.URLCarpetaAnexosHistoricos
    If fso.FolderExists(m_URLDirAnexosActuales) Then
        m_URLDirDNI = fso.GetParentFolderName(URLCarpetaAnexosActuales)
        If Right(m_URLDirAnexosActuales, 1) = "\" Then
           m_URLDirAnexosActuales = Left(m_URLDirAnexosActuales, Len(m_URLDirAnexosActuales) - 1)
        End If
        fso.DeleteFolder m_URLDirAnexosActuales, True
        If fso.FolderExists(m_URlDirAnexosHistoricos) Then
            If Right(m_URlDirAnexosHistoricos, 1) = "\" Then
               m_URlDirAnexosHistoricos = Left(m_URlDirAnexosHistoricos, Len(m_URlDirAnexosHistoricos) - 1)
            End If
            fso.DeleteFolder m_URlDirAnexosHistoricos, True
        End If
        If Right(m_URLDirDNI, 1) = "\" Then
           m_URLDirDNI = Left(m_URLDirDNI, Len(m_URLDirDNI) - 1)
        End If
        fso.DeleteFolder m_URLDirDNI, True
    Else
        If fso.FolderExists(m_URlDirAnexosHistoricos) Then
            m_URLDirDNI = fso.GetParentFolderName(URLCarpetaAnexosActuales)
            If Right(m_URlDirAnexosHistoricos, 1) = "\" Then
               m_URlDirAnexosHistoricos = Left(m_URlDirAnexosHistoricos, Len(m_URlDirAnexosHistoricos) - 1)
            End If
            fso.DeleteFolder m_URlDirAnexosHistoricos, True
            If Right(m_URLDirDNI, 1) = "\" Then
               m_URLDirDNI = Left(m_URLDirDNI, Len(m_URLDirDNI) - 1)
            End If
            fso.DeleteFolder m_URLDirDNI, True
        End If
    End If
    cache_usuario_borrar m_ID, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function PasarAHistorico( _
                                Optional ByRef p_Error As String _
                                ) As String
                                    
    
    
    Dim m_ObjColObservaciones As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjObServacion As Observacion
    Dim m_ObjObServacionHistorica As ObservacionHistorica
    Dim m_ObjUsuarioHistorico As UsuarioHistorico
    Dim m_UsuarioSICA As UsuarioSICA
    Dim m_Campo As Variant
    
    
    On Error GoTo errores
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    If Not IsDate(Me.F_BajaCalculada) Then
        p_Error = "El usuario no aparece como estando de baja"
        Err.Raise 1000
    End If
    
    Set m_ObjUsuarioHistorico = Constructor.getUsuarioHistorico(Me.ID, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjUsuarioHistorico Is Nothing Then
        m_ObjUsuarioHistorico.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set m_ObjUsuarioHistorico = New UsuarioHistorico
    With m_ObjUsuarioHistorico
        For Each m_Campo In Me.ColCampos
            .SetPropiedad m_Campo, Me.getPropiedad(m_Campo), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
        .Registrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_ObjUsuarioHistorico = Constructor.getUsuarioHistorico(Me.ID, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioHistorico Is Nothing Then
        p_Error = "No se ha podido grabar el usuario en la tabla de históricos"
        Err.Raise 1000
    End If
    Set m_ObjColObservaciones = Me.ColObservaciones
    If Not m_ObjColObservaciones Is Nothing Then
        For Each m_ID In m_ObjColObservaciones
            Set m_ObjObServacion = m_ObjColObservaciones(m_ID)
            Set m_ObjObServacionHistorica = New ObservacionHistorica
            With m_ObjObServacionHistorica
                .ID = m_ObjObServacion.ID
                .Fecha = m_ObjObServacion.Fecha
                .IDObservacion = m_ObjObServacion.IDObservacion
                .Observacion = m_ObjObServacion.Observacion
                .TIPO = m_ObjObServacion.TIPO
                .ParaAlta = EnumSiNo.Sí
                .Registrar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjObServacion = Nothing
        Next
    End If
    Set m_ObjColObservaciones = Me.ColObservacionesCurso
    If Not m_ObjColObservaciones Is Nothing Then
        For Each m_ID In m_ObjColObservaciones
            Set m_ObjObServacion = m_ObjColObservaciones(m_ID)
            Set m_ObjObServacionHistorica = New ObservacionHistorica
            
            With m_ObjObServacionHistorica
                .ID = m_ObjObServacion.ID
                .Fecha = m_ObjObServacion.Fecha
                .IDObservacion = m_ObjObServacion.IDObservacion
                .Observacion = m_ObjObServacion.Observacion
                .TIPO = m_ObjObServacion.TIPO
                .ParaAlta = EnumSiNo.Sí
                .Registrar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjObServacion = Nothing
        Next
    End If
    
    PasarAnexosAHistoricos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_UsuarioSICA = Me.UsuarioSICA
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_UsuarioSICA Is Nothing Then
        m_UsuarioSICA.PasarAHistorico m_ObjUsuarioHistorico.ID, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Borrar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.ValidacionPasarAHistorico ha devuelto el error: " & Err.Description
    End If
End Function

Private Function PasarAnexosAHistoricos( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                            
    '-----------------
    ' hemos de pasar los anexos actuales a la carpeta que tenga el usuario histórico una vez creado
    ' de sus anexos
    
    Dim m_ObjColAnexosActuales As Scripting.Dictionary
    Dim m_ObjAnexoUsuarioHPS As AnexoUsuarioHPS
    Dim m_ObjAnexoUsuarioHistorico As AnexoUsuarioHistorico
    Dim m_UsuarioHistorico As UsuarioHistorico
    
    Dim m_ID As Variant
    Dim m_URLDirActual As String
    Dim m_URLDirHistorico As String
    Dim m_URLEnHistorico As String
    On Error GoTo errores
    
    Set Me.colAnexosActuales = Nothing
    Set m_ObjColAnexosActuales = Me.colAnexosActuales
    If m_ObjColAnexosActuales Is Nothing Then
        Exit Function
    End If
    Set m_UsuarioHistorico = getUsuarioHistorico(Me.ID, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_UsuarioHistorico Is Nothing Then
        p_Error = "No se pueden trasladar los anexos actuales de usuario HPS si no se ha creado antes el UsuarioHistorico"
        Err.Raise 1000
    End If
    m_URLDirActual = Me.URLCarpetaAnexosActuales
    m_URLDirHistorico = m_UsuarioHistorico.URLCarpetaAnexos
    
    For Each m_ID In m_ObjColAnexosActuales
        If m_ID = Empty Then GoTo siguiente
        'Debug.Print m_ID
        'If CStr(m_ID) = "1812" Then Stop
        Set m_ObjAnexoUsuarioHPS = m_ObjColAnexosActuales(m_ID)
'        If m_ObjAnexoUsuarioHPS Is Nothing Then
'            Stop
'        End If
        If fso.FileExists(m_ObjAnexoUsuarioHPS.URLAnexo) Then
            If FicheroAbierto(m_ObjAnexoUsuarioHPS.URLAnexo) Then
                p_Error = "Está abierto el anexo: " & vbNewLine & m_ObjAnexoUsuarioHPS.URLAnexo
                Err.Raise 1000
            End If
        End If
        m_URLEnHistorico = m_URLDirHistorico & fso.GetFileName(m_ObjAnexoUsuarioHPS.URLAnexo)
        If fso.FileExists(m_URLEnHistorico) Then
            If FicheroAbierto(m_URLEnHistorico) Then
                p_Error = "Está abierto el anexo: " & vbNewLine & m_URLEnHistorico
                Err.Raise 1000
            End If
        End If
        Set m_ObjAnexoUsuarioHPS = Nothing
siguiente:
    Next
    If Not fso.FolderExists(m_URLDirHistorico) Then
        fso.CreateFolder m_URLDirHistorico
    End If
    
    For Each m_ID In m_ObjColAnexosActuales
        If m_ID = Empty Then GoTo siguiente1
        Set m_ObjAnexoUsuarioHPS = m_ObjColAnexosActuales(m_ID)
        If Not fso.FileExists(m_ObjAnexoUsuarioHPS.URLAnexo) Then
            GoTo siguiente1
        End If
        Set m_ObjAnexoUsuarioHistorico = New AnexoUsuarioHistorico
        With m_ObjAnexoUsuarioHistorico
            .IDAnexo = m_ObjAnexoUsuarioHPS.IDAnexo
            .idUsuario = m_UsuarioHistorico.ID
            .NombreAnexo = fso.GetFileName(m_ObjAnexoUsuarioHPS.URLAnexo)
            .URLLocal = m_ObjAnexoUsuarioHPS.URLAnexo
            .Registrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End With
        m_ObjAnexoUsuarioHPS.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjAnexoUsuarioHistorico = Nothing
        Set m_ObjAnexoUsuarioHPS = Nothing
siguiente1:
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.PasarAnexosAHistoricos ha devuelto el error: " & Err.Description
    End If
    
    
End Function



Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "ID"
        .Add "DNI"
        .Add "Nombre"
        .Add "Apellido_1"
        .Add "Apellido_2"
        .Add "Telefono"
        .Add "Correo_e"
        .Add "F_Nacimiento"
        .Add "LugarNacimiento"
        .Add "Motivo_HPS"
        .Add "IDExpediente"
        .Add "F_Curso"
        .Add "CursoEnVigor"
        .Add "Requiere_Curso"
        .Add "FechaPrimeraConvocatoria"
        .Add "FechaSegundaConvocatoria"
        .Add "FechaCorreoNoCurso"
        .Add "F_Baja"
        .Add "LugarPrestacionServicio"
        .Add "IDEmpresaUsuario"
        .Add "IDEmpresaHPS"
        .Add "IDJuridicaContrato"
        .Add "FAvisoConcesion"
        .Add "RequiereComunicacionConcesion"
        .Add "FechaHPSConcesionMinima"
        .Add "CadenaContratistas"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "DNI" Then
        getPropiedad = Me.DNI
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_1" Then
        getPropiedad = Me.Apellido_1
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_2" Then
        getPropiedad = Me.Apellido_2
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        getPropiedad = Me.Telefono
        Exit Function
    End If
    If m_NombrePropiedad = "Correo_e" Then
        getPropiedad = Me.Correo_e
        Exit Function
    End If
    If m_NombrePropiedad = "F_Nacimiento" Then
        getPropiedad = Me.F_Nacimiento
        Exit Function
    End If
    If m_NombrePropiedad = "LugarNacimiento" Then
        getPropiedad = Me.LugarNacimiento
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo_HPS" Then
        getPropiedad = Me.Motivo_HPS
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        getPropiedad = Me.IDExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "F_Curso" Then
        getPropiedad = Me.F_Curso
        Exit Function
    End If
    If m_NombrePropiedad = "CursoEnVigor" Then
        getPropiedad = Me.CursoEnVigor
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_Curso" Then
        getPropiedad = Me.Requiere_Curso
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrimeraConvocatoria" Then
        getPropiedad = Me.FechaPrimeraConvocatoria
        Exit Function
    End If
    If m_NombrePropiedad = "FechaSegundaConvocatoria" Then
        getPropiedad = Me.FechaSegundaConvocatoria
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoNoCurso" Then
        getPropiedad = Me.FechaCorreoNoCurso
        Exit Function
    End If
    
    If m_NombrePropiedad = "F_Baja" Then
        getPropiedad = Me.F_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "LugarPrestacionServicio" Then
        getPropiedad = Me.LugarPrestacionServicio
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        getPropiedad = Me.IDEmpresaUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaHPS" Then
        getPropiedad = Me.IDEmpresaHPS
        Exit Function
    End If
    If m_NombrePropiedad = "IDJuridicaContrato" Then
        getPropiedad = Me.IDJuridicaContrato
        Exit Function
    End If
    If m_NombrePropiedad = "FAvisoConcesion" Then
        getPropiedad = Me.FAvisoConcesion
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereComunicacionConcesion" Then
        getPropiedad = Me.RequiereComunicacionConcesion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaHPSConcesionMinima" Then
        getPropiedad = Me.FechaHPSConcesionMinima
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaContratistas" Then
        getPropiedad = Me.CadenaContratistas
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Usuario.getPropiedad"
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Usuario.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DNI" Then
        Me.DNI = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_1" Then
        Me.Apellido_1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_2" Then
        Me.Apellido_2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        Me.Telefono = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Correo_e" Then
        Me.Correo_e = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Nacimiento" Then
        Me.F_Nacimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "LugarNacimiento" Then
        Me.LugarNacimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo_HPS" Then
        Me.Motivo_HPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        Me.IDExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Curso" Then
        Me.F_Curso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CursoEnVigor" Then
        Me.CursoEnVigor = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_Curso" Then
        Me.Requiere_Curso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrimeraConvocatoria" Then
        Me.FechaPrimeraConvocatoria = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaSegundaConvocatoria" Then
        Me.FechaSegundaConvocatoria = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoNoCurso" Then
        Me.FechaCorreoNoCurso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Baja" Then
        Me.F_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "LugarPrestacionServicio" Then
        Me.LugarPrestacionServicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        Me.IDEmpresaUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaHPS" Then
        Me.IDEmpresaHPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDJuridicaContrato" Then
        Me.IDJuridicaContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FAvisoConcesion" Then
        Me.FAvisoConcesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereComunicacionConcesion" Then
        Me.RequiereComunicacionConcesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaHPSConcesionMinima" Then
        Me.FechaHPSConcesionMinima = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaContratistas" Then
        Me.CadenaContratistas = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Usuario.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Usuario.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function







