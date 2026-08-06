Attribute VB_Name = "Test"
Option Compare Database
Option Explicit

Public Function Prueba1() As String
    Dim a As Long
    WizHook.key = 51488399
    a = WizHook.GetFileName2(1, "Hola", "", "", "", "", "", 0, 0, 0, True, 0)
    Debug.Print a
End Function

Public Function test_Alta_Solicitud(Optional ByRef p_Error As String) As String
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Solicitud As solicitud
    On Error GoTo errores
    
    Set m_Solicitud = New solicitud
    With m_Solicitud
        .Apellido1 = "Román"
        .Apellido2 = "del Peral"
        .DNI = "02248439M"
        .email = "andres.romandelperal@telefonica.com"
        .emailResponsable = "ardelperal@gmail.com"
        .IDEmpresaTramitadora = "48"
        .IDEmpresaUsuario = "26"
        .IDExpediente = "400"
        .Nombre = "Andrés"
        .Telefono = "627834200"
        .TIPO = "Alta"
        .Gestor = "mtr"
        .FNacimiento = "02/04/1973"
    End With
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_Solicitud
        .RegistrarDatos p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_Solicitud = constructor.getSolicitud(p_IDSolicitud:=m_Solicitud.IDSolicitud, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Solicitud Is Nothing Then
        Err.Raise 1000
    End If
    test_Alta_Solicitud = "1"
    Exit Function
errores:
    test_Alta_Solicitud = "-1"
End Function


Public Function test_Borrar_Solicitud(p_IDSolicitud As String, Optional ByRef p_Error As String) As String
    Dim m_Solicitud As solicitud
    Dim m_SolicitudOp As SolicitudOperaciones
    On Error GoTo errores
    Set m_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Solicitud Is Nothing Then
        Err.Raise 1000
    End If
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_Solicitud
        .Eliminar p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Solicitud Is Nothing Then
        Err.Raise 1000
    End If
    test_Borrar_Solicitud = "1"
    
    
    Exit Function
errores:
    test_Borrar_Solicitud = "-1"
End Function

Public Function test_Correo_Solicitud_EnvioExcel(p_IDSolicitud As String, Optional ByRef p_Error As String) As String
    Dim m_Solicitud As solicitud
    Dim m_Correo As Correo
    Dim m_IDCorreo As String
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Salida As String
    On Error GoTo errores
    
    Set m_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Solicitud Is Nothing Then
        Err.Raise 1000
    End If
    m_Solicitud.HTMLParaCorreo = EnumSiNo.Sí
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_Solicitud
        m_IDCorreo = .RegistroEnvioExcel(p_ConCorreo:=EnumSiNo.Sí, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_Correo = constructor.getCorreo(p_IDCorreo:=m_IDCorreo, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Correo Is Nothing Then
        Err.Raise 1000
    End If
    m_Salida = CorreoAlServidor(m_Correo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    test_Correo_Solicitud_EnvioExcel = "1"
    Exit Function
errores:
    test_Correo_Solicitud_EnvioExcel = "-1"
End Function

Public Function test_Correo_1(p_IDCorreo As String, Optional ByRef p_Error As String) As String
    
    Dim m_Correo As Correo
    Dim m_Salida As String
    On Error GoTo errores
    
    Set m_Correo = constructor.getCorreo(p_IDCorreo:=p_IDCorreo, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Correo Is Nothing Then
        Err.Raise 1000
    End If
    
    m_Salida = CorreoAlServidor(m_Correo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    test_Correo_1 = "1"
    Exit Function
errores:
    test_Correo_1 = "-1"
End Function




Public Function test_GenerarHTMLParametrizado(p_IDSolicitud As String, Optional ByRef p_Error As String) As String
    
    Dim m_URLPlantillaHTML As String
    Dim m_ColVariables As Scripting.Dictionary
    Dim m_RutaDestino As String
    Dim m_Solicitud As solicitud
    
    On Error GoTo errores
'    EVE p_Error
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
    Set m_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Solicitud Is Nothing Then
        Err.Raise 1000
    End If
    With m_Solicitud
        Set m_ColVariables = .ColParametrosParaHTML
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    m_URLPlantillaHTML = m_ObjEntorno.URLPlantillasHTMLExcelSolicitud
    
    m_RutaDestino = m_ObjEntorno.URLDirectorioLocal & fso.GetBaseName(fso.GetTempName()) & ".html"
    
    GenerarHTML , m_URLPlantillaHTML, m_ColVariables, m_RutaDestino, EnumSiNo.no, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLHTMLActivo = m_RutaDestino
    If FormularioAbierto("FormWeb") Then
        DoCmd.Close acForm, "FormWeb", acSaveNo
    End If
    DoCmd.OpenForm "FormWeb"
    Ejecutar 1, "open", m_RutaDestino, "", "", 1
    'fso.DeleteFile m_RutaDestino, True
    Exit Function
errores:
   test_GenerarHTMLParametrizado = "-1"
End Function

Public Function test_SolPtesEnvioExcel(Optional ByRef p_Error As String) As String
    Dim m_Usuario As Usuario
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    Set m_Usuario = constructor.getUsuario(p_UsuarioRed:="eaa", p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Usuario Is Nothing Then
        Err.Raise 1000
    End If
    Set m_Col = constructor.getSolicitudesPendientes(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.RecordatorioExcel1, p_Tramitador:=m_Usuario.Nombre, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        test_SolPtesEnvioExcel = "0"
    Else
        test_SolPtesEnvioExcel = m_Col.Count
    End If
    
    Exit Function
errores:
    test_SolPtesEnvioExcel = "-1"
End Function

Sub TestEsDNIValido()
    Dim resultado As Boolean
    
    ' Prueba 1: DNI válido
    resultado = EsDNIValido("12345678Z")
    If Not resultado Then
        Debug.Print "Prueba 1 fallida"
    Else
        Debug.Print "Prueba 1 pasada"
    End If
    
    ' Prueba 2: DNI inválido (longitud incorrecta)
    resultado = EsDNIValido("12345678")
    If resultado Then
        Debug.Print "Prueba 2 fallida"
    Else
        Debug.Print "Prueba 2 pasada"
    End If
    
    ' Prueba 3: DNI inválido (letra incorrecta)
    resultado = EsDNIValido("12345678A")
    If resultado Then
        Debug.Print "Prueba 3 fallida"
    Else
        Debug.Print "Prueba 3 pasada"
    End If
    
    ' Prueba 4: NIE válido
    resultado = EsDNIValido("X12345678")
    If Not resultado Then
        Debug.Print "Prueba 4 fallida"
    Else
        Debug.Print "Prueba 4 pasada"
    End If
    
    ' Prueba 5: NIE inválido (letra incorrecta)
    resultado = EsDNIValido("X1234567A")
    If resultado Then
        Debug.Print "Prueba 5 fallida"
    Else
        Debug.Print "Prueba 5 pasada"
    End If
    
    ' Prueba 6: NIE inválido (longitud incorrecta)
    resultado = EsDNIValido("X1234567")
    If resultado Then
        Debug.Print "Prueba 6 fallida"
    Else
        Debug.Print "Prueba 6 pasada"
    End If
    
End Sub
Public Sub PruebaIntegracionParametrosParser()
    ' Probar integración con Solicitud
    Dim solicitud As New solicitud
    ' ... configurar solicitud ...
    
    Dim cadena As String
    cadena = solicitud.ColParametrosParaHTMLTexto
    
    ' Verificar que se puede parsear de vuelta
    Dim parser As New ParametrosParser
    Dim diccionario As Scripting.Dictionary
    Set diccionario = parser.ParsearCadenaRecursos(cadena)
    
    Debug.Print "Prueba integración - Elementos: " & diccionario.Count
End Sub
Sub ProbarEnvioGenerico()
    m_ObjEntorno.ModoNuevoEnvioCorreo = True ' Activamos flag temporalmente
    
    ' Simulamos enviar un correo SIN solicitud real
    ' OJO: Asegúrate que la plantilla "ENVIO_EXCEL.html" exista en tu carpeta HTML
    Dim resultado As Correo
    Dim errStr As String
    
    Set resultado = Automiatizacion.EnviarCorreoGenerico( _
        p_DESTINATARIOS:="tu_email@prueba.com", _
        p_asunto:="Prueba de desacoplamiento", _
        p_NombrePlantilla:="ENVIO_EXCEL.html", _
        p_CadenaRecursos:="{||Gestor||:PruebaSistema}", _
        p_Accion:="Prueba Sistema", _
        p_TipoCorreo:=1, _
        p_Error:=errStr)
        
    If errStr <> "" Then
        Debug.Print "Error: " & errStr
    Else
        Debug.Print "Correo generado con ID: " & resultado.IDCORREO & " y IDSolicitud: " & resultado.IDSolicitud
    End If
    
    m_ObjEntorno.ModoNuevoEnvioCorreo = False ' Restauramos flag
End Sub

' En Test.bas

Public Function Test_GenerarCorreoPrueba_DesdeID149()
    Dim m_SolicitudOrigen As solicitud
    Dim m_CorreoGenerado As Correo
    Dim p_Error As String
    
    ' --- PARÁMETROS DE LA PRUEBA (Lo que elegirías en el futuro formulario) ---
    Dim ID_Origen As String: ID_Origen = "149"
    Dim Plantilla_Prueba As String: Plantilla_Prueba = "RECORDATORIO_EXCEL1.html" ' Puedes cambiarla aquí
    Dim Destinatario_Prueba As String: Destinatario_Prueba = "andres.romandelperal@telefonica.com"
    Dim Fecha_Programada As Variant: Fecha_Programada = DateAdd("d", 1, Now()) ' Mañana (para probar programación)
    ' --------------------------------------------------------------------------

    On Error GoTo errores

    ' 1. Activamos el modo desacoplado temporalmente
    m_ObjEntorno.ModoNuevoEnvioCorreo = True
    
    ' 2. Hidratamos datos desde la solicitud real (Contexto)
    Set m_SolicitudOrigen = constructor.getSolicitud(ID_Origen, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , "Error cargando solicitud: " & p_Error
    
    Debug.Print "--- DATOS OBTENIDOS DE SOLICITUD " & ID_Origen & " ---"
    Debug.Print "Solicitante: " & m_SolicitudOrigen.NombreCompleto
    Debug.Print "Cadena Recursos Original: " & Left(m_SolicitudOrigen.ColParametrosParaHTMLTexto, 100) & "..."
    
    ' 3. Llamamos al generador genérico
    ' Fíjate que usamos los datos de m_SolicitudOrigen para rellenar los huecos,
    ' pero sobreescribimos el Destinatario y la Fecha.
    
    Set m_CorreoGenerado = Automiatizacion.EnviarCorreoGenerico( _
        p_DESTINATARIOS:=Destinatario_Prueba, _
        p_asunto:="PRUEBA TEMPLATE: " & m_SolicitudOrigen.AsuntoCalculado(EnumTipoEnvioCorreo.RecordatorioExcel1), _
        p_NombrePlantilla:=Plantilla_Prueba, _
        p_CadenaRecursos:=m_SolicitudOrigen.ColParametrosParaHTMLTexto, _
        p_Accion:="Prueba de Plantilla " & Plantilla_Prueba, _
        p_TipoCorreo:=EnumTipoEnvioCorreo.RecordatorioExcel1, _
        p_DestinatariosCopia:="", _
        p_URLAdjunto:="", _
        p_IDSolicitudVinculada:=CONST_IDSOLICITUD_PRUEBA, _
        p_FechaOrden:=Fecha_Programada, _
        p_Error:=p_Error)
        
    If p_Error <> "" Then Err.Raise 1000
    
    Debug.Print "------------------------------------------------"
    Debug.Print "¡ÉXITO! Correo de prueba generado en la cola."
    Debug.Print "ID Correo: " & m_CorreoGenerado.IDCORREO
    Debug.Print "Estado: " & m_CorreoGenerado.EstadoTexto
    Debug.Print "Fecha Salida: " & m_CorreoGenerado.FechaOrdenEnvio
    Debug.Print "IDSolicitud guardado: " & m_CorreoGenerado.IDSolicitud
    Debug.Print "------------------------------------------------"

    ' Restaurar flag
    m_ObjEntorno.ModoNuevoEnvioCorreo = False
    Exit Function

errores:
    m_ObjEntorno.ModoNuevoEnvioCorreo = False
    Debug.Print "ERROR EN TEST: " & p_Error & " | " & Err.Description
End Function

