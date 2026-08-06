Attribute VB_Name = "testing"
Option Compare Database
Option Explicit


Public Function test_Fatima_Es_Usuario_Calidad() As String
    
    Debug.Assert EnumSiNo.Sí = UsuarioEsDeCalidad("fmc")
    
End Function

Public Function test_Fernando_Es_Usuario_Admin() As String
    
    Debug.Assert EnumSiNo.Sí = UsuarioEsAdministrador("ds01474")
    
End Function

Public Function test_SGM_Es_Usuario_Calidad_Con_Avisos() As String
    
    Debug.Assert EnumSiNo.Sí = UsuarioEsDeCalidad("sgm")
    
End Function
Public Function test_Riesgo_Origen_Aceptado() As String
    
    Dim m_Cod As String
    Dim m_IDEdicion As String
    Dim m_Riesgo As riesgo
    Dim m_RiesgoNacimiento As riesgo
    
    m_Cod = "R012"
    m_IDEdicion = "274"
    Set m_Riesgo = Constructor.getRiesgo(, m_IDEdicion, m_Cod)
    If m_Riesgo Is Nothing Then
        Exit Function
    End If
    Set m_RiesgoNacimiento = m_Riesgo.RiesgoAceptadoEnNacimiento
    test_Riesgo_Origen_Aceptado = m_RiesgoNacimiento.DiasRespuestaCalidadAceptacion
    
    
    
End Function
Public Function test_Riesgo_Origen_retirado() As String
    
    Dim m_Cod As String
    Dim m_IDEdicion As String
    Dim m_Riesgo As riesgo
    Dim m_RiesgoNacimiento As riesgo
    
    m_Cod = "R003"
    m_IDEdicion = "214"
    Set m_Riesgo = Constructor.getRiesgo(, m_IDEdicion, m_Cod)
    If m_Riesgo Is Nothing Then
        Exit Function
    End If
    Set m_RiesgoNacimiento = m_Riesgo.RiesgoRetiradoEnNacimiento
    test_Riesgo_Origen_retirado = m_RiesgoNacimiento.DiasRespuestaCalidadRetiro
    
    
    
End Function
Public Function test_Riesgo_Dias_Por_Aceptar() As String
    
    
    Dim m_IdRiesgo As String
    Dim m_Riesgo As riesgo
    
    m_IdRiesgo = "1395"
   
    Set m_Riesgo = Constructor.getRiesgo(m_IdRiesgo)
    If m_Riesgo Is Nothing Then
        Exit Function
    End If
    m_Riesgo.RegistrarDiasAceptacionCalidad
    test_Riesgo_Dias_Por_Aceptar = m_Riesgo.DiasSinRespuestaCalidadAceptacion
End Function

'-------------------------------------------
' Nombre: test_Cadena_MAIN_vs_SUB
' Propósito: Verificar cadenas generadas por GetCadenaJerarquicaEmpresas
' Parámetros: ninguno
' Retorno: String (resumen simple)
'-------------------------------------------
Public Function test_Cadena_MAIN_vs_SUB() As String
    On Error GoTo ManejoError
    Dim sMain As String
    Dim sSub As String
    Dim sMsg As String
    Dim sErr As String
    
    sMain = GetCadenaJerarquicaEmpresas(1007, "MAIN", sErr)
    sSub = GetCadenaJerarquicaEmpresas(1007, "SUB", sErr)
    
    sMsg = "MAIN=" & sMain & vbCrLf & "SUB=" & sSub
    test_Cadena_MAIN_vs_SUB = sMsg
    Exit Function
ManejoError:
    test_Cadena_MAIN_vs_SUB = "Error " & Err.Number & ": " & Err.Description
End Function
'-------------------------------------------
' Nombre: test_Validacion_Roles
' Propósito: Verificar helpers EsContratistaPrincipal / EsSubContratista
' Parámetros: ninguno
' Retorno: Boolean (True si ambos tests pasan)
'-------------------------------------------
Public Function test_Validacion_Roles() As Boolean
    On Error GoTo ManejoError
    Dim okMain As Boolean
    Dim okSub As Boolean
    
    okMain = EsContratistaPrincipal(Null, "Sí")
    okSub = EsSubContratista(100, "No")
    
    test_Validacion_Roles = (okMain And okSub)
    Exit Function
ManejoError:
    test_Validacion_Roles = False
End Function
'-------------------------------------------
' Nombre: test_Retrocompatibilidad_Bandera
' Propósito: Verificar elección de lógica según TempVars("CadenaJerarquicaModelo")
' Parámetros: ninguno
' Retorno: String (resumen)
'-------------------------------------------
Public Function test_Retrocompatibilidad_Bandera() As String
    On Error GoTo ManejoError
    Dim sModelo As String
    Dim sOut As String
    Dim sErr As String
    Dim dic As Scripting.Dictionary
    
    sModelo = Trim$(Nz(Application.TempVars("CadenaJerarquicaModelo"), "nuevo"))
    
    Set dic = getExpedienteSuministradores_RC("1007", sErr)
    sOut = "Modelo=" & sModelo & "; Subcontratistas=" & IIf(dic Is Nothing, 0, dic.Count)
    test_Retrocompatibilidad_Bandera = sOut
    Exit Function
ManejoError:
    test_Retrocompatibilidad_Bandera = "Error " & Err.Number & ": " & Err.Description
End Function
Public Function test_correo_alta_proyecto() As String
    Dim p As Proyecto
    Dim c As CORREO
    Dim errTxt As String
    Set p = New Proyecto
    p.NombreProyecto = "Proyecto de prueba"
    p.Proyecto = "EXP-TEST"
    p.Cliente = "Cliente prueba"
    p.CodigoDocumento = "DOC-TEST"
    p.NombreUsuarioCalidad = "Resp Calidad"
    Set c = New CORREO
    c.Asunto = "TEST Alta de gestión de riesgos"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoAltaProyecto(p, errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_alta_proyecto = errTxt
    Else
        test_correo_alta_proyecto = "OK"
    End If
End Function
Public Function test_correo_nueva_publicacion() As String
    Dim p As Proyecto
    Dim e As Edicion
    Dim c As CORREO
    Dim errTxt As String
    Set p = New Proyecto
    p.NombreProyecto = "Proyecto de prueba"
    Set e = New Edicion
    e.Edicion = "01"
    e.FechaPublicacion = Date
    e.Elaborado = "Elabora"
    e.Revisado = "Revisa"
    e.Aprobado = "Aprueba"
    Set c = New CORREO
    c.Asunto = "TEST Edición publicada"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoNuevaPublicacion(p, e, "", errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_nueva_publicacion = errTxt
    Else
        test_correo_nueva_publicacion = "OK"
    End If
End Function
Public Function test_correo_revision_edicion() As String
    Dim e As Edicion
    Dim pr As Proyecto
    Dim c As CORREO
    Dim errTxt As String
    Set pr = New Proyecto
    pr.NombreProyecto = "Proyecto de prueba"
    Set e = New Edicion
    Set e.Proyecto = pr
    e.Edicion = "01"
    e.FechaEdicion = Date
    e.Elaborado = "Elabora"
    e.Revisado = "Revisa"
    e.Aprobado = "Aprueba"
    Set c = New CORREO
    c.Asunto = "TEST Solicitud de revisión"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoRevisionEdicion(e, errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_revision_edicion = errTxt
    Else
        test_correo_revision_edicion = "OK"
    End If
End Function
Public Function test_correo_riesgo_materializado() As String
    Dim r As riesgo
    Dim c As CORREO
    Dim errTxt As String
    Set r = New riesgo
    r.CodigoRiesgo = "R-TEST"
    r.Descripcion = "Riesgo de prueba"
    r.DetectadoPor = "Tester"
    r.ImpactoGlobal = "Alto"
    r.FechaMaterializado = Date
    Set c = New CORREO
    c.Asunto = "TEST Riesgo materializado"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoRiesgoMaterializado(r, errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_riesgo_materializado = errTxt
    Else
        test_correo_riesgo_materializado = "OK"
    End If
End Function
Public Function test_correo_riesgo_aceptado_tecnico() As String
    Dim r As riesgo
    Dim c As CORREO
    Dim errTxt As String
    Set r = New riesgo
    r.CodigoRiesgo = "R-TEST"
    r.Descripcion = "Riesgo de prueba"
    r.DetectadoPor = "Tester"
    r.FechaMitigacionAceptar = Date
    Set c = New CORREO
    c.Asunto = "TEST Riesgo aceptado por técnico"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoTecnicoRiesgoAceptado(r, errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_riesgo_aceptado_tecnico = errTxt
    Else
        test_correo_riesgo_aceptado_tecnico = "OK"
    End If
End Function
Public Function test_correo_riesgo_retirado_tecnico() As String
    Dim r As riesgo
    Dim c As CORREO
    Dim errTxt As String
    Set r = New riesgo
    r.CodigoRiesgo = "R-TEST"
    r.Descripcion = "Riesgo de prueba"
    r.DetectadoPor = "Tester"
    Set c = New CORREO
    c.Asunto = "TEST Riesgo retirado por técnico"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoRiesgoAceptadoRetirado(r, EnumSiNo.No, errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_riesgo_retirado_tecnico = errTxt
    Else
        test_correo_riesgo_retirado_tecnico = "OK"
    End If
End Function
Public Function test_correo_riesgo_requiere_retipificacion() As String
    Dim r As riesgo
    Dim c As CORREO
    Dim errTxt As String
    Set r = New riesgo
    r.CodigoRiesgo = "R-TEST"
    r.Descripcion = "Riesgo de prueba"
    r.DetectadoPor = "Tester"
    r.CausaRaiz = "Causa prueba"
    Set c = New CORREO
    c.Asunto = "TEST Retipificación requerida"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoRiesgoRequiereRetipificacion(r, errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_riesgo_requiere_retipificacion = errTxt
    Else
        test_correo_riesgo_requiere_retipificacion = "OK"
    End If
End Function
Public Function test_correo_riesgo_retipificado() As String
    Dim r As riesgo
    Dim r0 As riesgo
    Dim c As CORREO
    Dim errTxt As String
    Set r0 = New riesgo
    r0.CodRiesgoBiblioteca = "LIB-OLD"
    Set r = New riesgo
    r.CodigoRiesgo = "R-TEST"
    r.CodRiesgoBiblioteca = "LIB-NEW"
    Set c = New CORREO
    c.Asunto = "TEST Riesgo retipificado por calidad"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, GetBodyCorreoRiesgoRetipificado(r, r0, errTxt))
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_riesgo_retipificado = errTxt
    Else
        test_correo_riesgo_retipificado = "OK"
    End If
End Function
Public Function test_correo_error_administrador() As String
    Dim c As CORREO
    Dim body As String
    Dim errTxt As String
    body = GetBodyCorreoErrorAdministrador("FormPrueba", "PC-TEST", "usuario.test", "En Oficina", "v2026-002", "Detalle de error de prueba")
    Set c = New CORREO
    c.Asunto = "TEST Incidencia detectada"
    c.Cuerpo = CuerpoHTMLConEstiloCorporativo(c.Asunto, body)
    c.Destinatarios = "andres.romandelperal@telefonica.com"
    c.FechaGrabacion = Now()
    c.Registrar errTxt
    If errTxt <> "" Then
        test_correo_error_administrador = errTxt
    Else
        test_correo_error_administrador = "OK"
    End If
End Function
Public Function test_correo_runner() As String
    On Error GoTo ManejoError
    Dim s As String
    s = "AltaProyecto=" & test_correo_alta_proyecto & vbCrLf
    s = s & "NuevaPublicacion=" & test_correo_nueva_publicacion & vbCrLf
    s = s & "RevisionEdicion=" & test_correo_revision_edicion & vbCrLf
    s = s & "RiesgoMaterializado=" & test_correo_riesgo_materializado & vbCrLf
    s = s & "RiesgoAceptadoTecnico=" & test_correo_riesgo_aceptado_tecnico & vbCrLf
    s = s & "RiesgoRetiradoTecnico=" & test_correo_riesgo_retirado_tecnico & vbCrLf
    s = s & "RiesgoRequiereRetipificacion=" & test_correo_riesgo_requiere_retipificacion & vbCrLf
    s = s & "RiesgoRetipificado=" & test_correo_riesgo_retipificado & vbCrLf
    s = s & "ErrorAdministrador=" & test_correo_error_administrador
    test_correo_runner = s
    Exit Function
ManejoError:
    test_correo_runner = "Error " & Err.Number & ": " & Err.Description
End Function
