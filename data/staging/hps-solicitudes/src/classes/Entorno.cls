VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public URLArchivoUltimo As String

Private m_sURLAchivoIni As String
Private m_sURLDirectorioAplicacion As String
Private m_sURLDirectorioDocumentacion As String
Private m_sURLDirectorioRecursos As String
Private m_sURLDirectorioPlantillas As String
Private m_sURLDirectorioPlantillasHTML As String
Private m_sURLDirectorioPlantillasExcel As String
Private m_sURLDirectorioLocal As String
Private m_sURLDirectorioDocumentacionAyuda As String
Private m_sURLPlantillaAlta As String
Private m_sURLPlantillaTraspaso As String
Private m_sVersionAplicacion As String
Private m_ObjConfiguracion As Configuracion
Private m_objColAdministradores As Scripting.Dictionary
Private m_objColtramitadores As Scripting.Dictionary
Private m_sDiasParaRecordatorioExcel1 As String
Private m_sDiasParaRecordatorioExcel2 As String
Private m_sDiasCancelacionPreMARGA As String
Private m_sDiasParaRecordatorioRellenoMarga1 As String
Private m_sDiasParaRecordatorioRellenoMarga2 As String
Private m_sDiasCancelacionMARGA As String
Private m_sAutocancelacionPreMARGA As String
Private m_sAutocancelacionMARGA As String
Private m_sBuzonSeguridad As String
Private m_sEmailDirectorSeguridad As String
Private m_sCorreodeEnvio As String
Private m_sCorreosAutomaticos As String
Private m_bFaltaAlgunDatoConfiguracion As Boolean
Private m_ObjColEstadosTexto As Scripting.Dictionary
Private m_ObjColTextoEstados As Scripting.Dictionary
Private m_ObjColEstadosParaTitulo As Scripting.Dictionary
Private m_ObjColTiposTexto As Scripting.Dictionary
Private m_ObjColTextoTipos As Scripting.Dictionary

Private m_ObjColNombrePlantillas As Scripting.Dictionary

Private m_objColPendientesAltaSolicitud As Scripting.Dictionary
Private m_objColPendientesEnvioExcel As Scripting.Dictionary
Private m_objColPendientesRecordatorioExcel1 As Scripting.Dictionary
Private m_objColPendientesRecordatorioExcel2 As Scripting.Dictionary
Private m_objColPendientesExcelRelleno As Scripting.Dictionary
Private m_objColPendienteRegistroMarga As Scripting.Dictionary
Private m_objColPendienteRecordatorioEnvioDPS1 As Scripting.Dictionary
Private m_objColPendienteRecordatorioEnvioDPS2 As Scripting.Dictionary
Private m_objColPendienteEnvioDPS As Scripting.Dictionary
Private m_objColPendienteRegistroONS As Scripting.Dictionary
Private m_objColPendienteRegistroEnHPS As Scripting.Dictionary

Private m_objColTramitados As Scripting.Dictionary
Private m_objColDesestimados As Scripting.Dictionary
Private m_objColCancelados As Scripting.Dictionary
Private m_objColPendienteEnvioTraspasoONS As Scripting.Dictionary
Private m_objColEnTramite As Scripting.Dictionary
Private m_objColSolicitudes As Scripting.Dictionary

Private m_objcolUsuariosHPS As Scripting.Dictionary




Private m_objColMotivos As Scripting.Dictionary

Private m_ObjcolExplicacionTipoSeleccionTarea As Scripting.Dictionary

Private m_sURLDirectorioAplicaciones As String
Private m_sURLPlantillasHTMLExcelSolicitud As String
Private m_sURLHTMLRecordatorioExcel1 As String
Private m_sURLHTMLRecordatorioExcel2 As String
Private m_sURLHTMLRecordatorioMarga1 As String
Private m_sURLHTMLRecordatorioMarga2 As String
Private m_sURLHTMLAutocancelacionPREMARGA As String
Private m_sURLHTMLAutocancelacionMARGA As String

Private m_objColTiposCorreos As Scripting.Dictionary

Private m_sCharsetISO As String
Private m_sCharsetUTF As String
Private m_sTextoConfiguracionCorreosAutomaticos As String
Private m_sDatosParaCorreoAutomaticoRegistrados As EnumSiNo
Private m_ObjExpedientes As Scripting.Dictionary
Private m_ObjResponsables As Scripting.Dictionary


Private m_objcolGradosHPS As Scripting.Dictionary
Private m_ObjEntidades As Scripting.Dictionary
Private m_ObjEntidadesFormularios As Scripting.Dictionary
Private m_eConCoreroCopiaGestor As EnumSiNo
Private m_sCSS As String
Private m_sURLArchivoCSS As String

Private m_sURLDirectorioPlantillasHTMLVersionEstablecida As String
Private m_sURLDirectorioPlantillasEXCELConVersion As String
Private m_ObjUltimoCambio As UltimoCambio
Private m_objColPlantillas As Scripting.Dictionary
Private m_ObjSuministradores As Scripting.Dictionary
Private m_ObjSumTramitadoras As Scripting.Dictionary

Private m_lTotalPendientesGlobal As Long
Private m_lTotalPendientesUsuario As Long
Private m_bIndicadoresCalculados As Boolean
Private m_bModoNuevoEnvioCorreo As Boolean
Private m_ColJustificaciones As Collection

Public Error As String
Private m_Error As String
Private m_objColItems As Scripting.Dictionary

Public Property Get Justificaciones() As Collection
    Dim Servicio As New JustificacionServicio
    ' Lazy Loading: Si no existe, la cargamos. Si existe, devolvemos la memoria.
    If m_ColJustificaciones Is Nothing Then
        Set m_ColJustificaciones = Servicio.ObtenerListaParaSeleccion()
    End If
    Set Justificaciones = m_ColJustificaciones
End Property

' Método para forzar recarga si alguien añade una nueva desde administración
Public Sub RecargarJustificaciones()
    Set m_ColJustificaciones = Nothing
End Sub
' Controla si usamos el sistema antiguo (acoplado) o el nuevo (genérico)
Public Property Get ModoNuevoEnvioCorreo() As Boolean
    ' Por defecto False para mantener comportamiento antiguo hasta que decidas cambiarlo
    ' Puedes cambiar esto a True cuando quieras activar el nuevo sistema globalmente
    ModoNuevoEnvioCorreo = m_bModoNuevoEnvioCorreo
End Property

Public Property Let ModoNuevoEnvioCorreo(ByVal valor As Boolean)
    m_bModoNuevoEnvioCorreo = valor
End Property

Public Property Get ColPlantillas() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColPlantillas Is Nothing Then
        Set ColPlantillas = m_objColPlantillas
        Exit Property
    End If
    
    Set m_objColPlantillas = New Scripting.Dictionary
    m_objColPlantillas.CompareMode = TextCompare
    With m_objColPlantillas
        .Add CStr(EnumPlantillas.EnvioExcel), "ENVIO_EXCEL.html"
        .Add CStr(EnumPlantillas.RecordatorioExcel1), "RECORDATORIO_EXCEL1.html"
        .Add CStr(EnumPlantillas.RecordatorioExcel2), "RECORDATORIO_EXCEL2.html"
        .Add CStr(EnumPlantillas.RecordatorioMarga1), "RECORDATORIO_MARGA1.html"
        .Add CStr(EnumPlantillas.RecordatorioMarga2), "RECORDATORIO_MARGA2.html"
        .Add CStr(EnumPlantillas.CancelacionPreMarga), "AUTOCANCELACION_PREMARGA.html"
        .Add CStr(EnumPlantillas.CancelacionMarga), "AUTOCANCELACION_MARGA.html"
    End With
    
    Set ColPlantillas = m_objColPlantillas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPlantillas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Entidades() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEntidades Is Nothing Then
        Set Entidades = m_ObjEntidades
        Exit Property
    End If
    
    Set m_ObjEntidades = New Scripting.Dictionary
    m_ObjEntidades.CompareMode = TextCompare
    With m_ObjEntidades
        .Add CStr(EnumEntidades.Responsables), "Responsables"
        .Add CStr(EnumEntidades.PlantillasHTML), "Plantillas HTML"
        .Add CStr(EnumEntidades.MotivosHPS), "Motivos HPS"
        .Add CStr(EnumEntidades.JustificacionHPS), "Justificaciones HPS"
    End With
    
    Set Entidades = m_ObjEntidades
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Entidades ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EntidadesFormularios() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEntidadesFormularios Is Nothing Then
        Set EntidadesFormularios = m_ObjEntidadesFormularios
        Exit Property
    End If
    
    Set m_ObjEntidadesFormularios = New Scripting.Dictionary
    m_ObjEntidadesFormularios.CompareMode = TextCompare
    With m_ObjEntidadesFormularios
        
        .Add CStr(EnumEntidades.Responsables), "FormResponsablesGestion"
        .Add CStr(EnumEntidades.PlantillasHTML), "FormPlantillasHTML"
        .Add CStr(EnumEntidades.MotivosHPS), "FormMotivosHPS"
        .Add CStr(EnumEntidades.JustificacionHPS), "frmBuscarJustificacion"
    End With
    
    Set EntidadesFormularios = m_ObjEntidadesFormularios
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EntidadesFormularios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get colUsuariosHPS() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objcolUsuariosHPS Is Nothing Then
        Set colUsuariosHPS = m_objcolUsuariosHPS
        Exit Property
    End If
    Set m_objcolUsuariosHPS = constructor.getUsuariosHPS(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colUsuariosHPS = m_objcolUsuariosHPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosHPS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set colUsuariosHPS(ByVal objNewValue As Variant)

    Set m_objcolUsuariosHPS = objNewValue
    
End Property
Public Property Get Suministradores() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjSuministradores Is Nothing Then
        Set Suministradores = m_ObjSuministradores
        Exit Property
    End If
    Set m_ObjSuministradores = constructor.getSuministradores(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Suministradores = m_ObjSuministradores
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Suministradores ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set Suministradores(ByVal objNewValue As Variant)

    Set m_ObjSuministradores = objNewValue
    
End Property

Public Property Get SumTramitadoras() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjSumTramitadoras Is Nothing Then
        Set SumTramitadoras = m_ObjSumTramitadoras
        Exit Property
    End If
    Set m_ObjSumTramitadoras = constructor.getSuministradores(p_SoloTramitadorasHPS:=EnumSiNo.Sí, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set SumTramitadoras = m_ObjSumTramitadoras
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.SumTramitadoras ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set SumTramitadoras(ByVal objNewValue As Variant)

    Set m_ObjSumTramitadoras = objNewValue
    
End Property
Public Property Get URLArchivoCSS() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLArchivoCSS <> "" Then
        URLArchivoCSS = m_sURLArchivoCSS
        Exit Property
    End If
    m_sURLArchivoCSS = Me.URLDirectorioAplicaciones & "CSS.txt"
    If Not fso.FileExists(m_sURLArchivoCSS) Then
        m_Error = "No es alcanzable URLArchivoCSS " & Me.URLArchivoCSS
        Err.Raise 1000
    End If
    
    URLArchivoCSS = m_sURLArchivoCSS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLArchivoCSS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get CSS() As String
    
    Dim ObjDatos As TextStream
    Dim m_URL As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sCSS <> "" Then
        CSS = m_sCSS
        Exit Property
    End If
    m_URL = Me.URLArchivoCSS
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URL) Then
        m_Error = "Ne es alcanzable el archivo URLArchivoCSS" & vbNewLine & Me.URLArchivoCSS
        Err.Raise 1000
    End If
    Set ObjDatos = fso.OpenTextFile(m_URL, ForReading, False, TristateMixed)
    While Not ObjDatos.AtEndOfStream
          m_sCSS = m_sCSS & ObjDatos.ReadAll
    Wend
    Set ObjDatos = Nothing
    

    CSS = m_sCSS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CSS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ConCoreroCopiaGestor() As EnumSiNo
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eConCoreroCopiaGestor <> Empty Then
        ConCoreroCopiaGestor = m_eConCoreroCopiaGestor
        Exit Property
    End If
    If Application.TempVars("ConCorreoCopiaGestor") = "Sí" Then
        m_eConCoreroCopiaGestor = EnumSiNo.Sí
    Else
        m_eConCoreroCopiaGestor = EnumSiNo.no
    End If
    ConCoreroCopiaGestor = m_eConCoreroCopiaGestor
    
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ConCoreroCopiaGestor ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get colGradosHPS() As Scripting.Dictionary
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objcolGradosHPS Is Nothing Then
        Set colGradosHPS = m_objcolGradosHPS
        Exit Property
    End If
    
    Set m_objcolGradosHPS = constructor.getListaColeccionesGradosHPS(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colGradosHPS = m_objcolGradosHPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colGradosHPS ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property



Public Property Get DatosParaCorreoAutomaticoRegistrados() As EnumSiNo
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    With Me
        If Me.BuzonSeguridad = "" Then
            m_sDatosParaCorreoAutomaticoRegistrados = EnumSiNo.no
            DatosParaCorreoAutomaticoRegistrados = m_sDatosParaCorreoAutomaticoRegistrados
            Exit Property
        End If
        If Me.DiasParaRecordatorioExcel1 = "" Then
            m_sDatosParaCorreoAutomaticoRegistrados = EnumSiNo.no
            DatosParaCorreoAutomaticoRegistrados = m_sDatosParaCorreoAutomaticoRegistrados
            Exit Property
        End If
        If Me.DiasParaRecordatorioExcel2 = "" Then
            m_sDatosParaCorreoAutomaticoRegistrados = EnumSiNo.no
            DatosParaCorreoAutomaticoRegistrados = m_sDatosParaCorreoAutomaticoRegistrados
            Exit Property
        End If
        If Me.DiasParaRecordatorioRellenoMarga1 = "" Then
            m_sDatosParaCorreoAutomaticoRegistrados = EnumSiNo.no
            DatosParaCorreoAutomaticoRegistrados = m_sDatosParaCorreoAutomaticoRegistrados
            Exit Property
        End If
        If Me.DiasParaRecordatorioRellenoMarga2 = "" Then
            m_sDatosParaCorreoAutomaticoRegistrados = EnumSiNo.no
            DatosParaCorreoAutomaticoRegistrados = m_sDatosParaCorreoAutomaticoRegistrados
            Exit Property
        End If
        If Me.EmailDirectorSeguridad = "" Then
            m_sDatosParaCorreoAutomaticoRegistrados = EnumSiNo.no
            DatosParaCorreoAutomaticoRegistrados = m_sDatosParaCorreoAutomaticoRegistrados
            Exit Property
        End If
        
    End With
    
    
    
    m_sDatosParaCorreoAutomaticoRegistrados = EnumSiNo.Sí
    DatosParaCorreoAutomaticoRegistrados = m_sDatosParaCorreoAutomaticoRegistrados
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DatosParaCorreoAutomaticoRegistrados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get TextoConfiguracionCorreosAutomaticos() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Application.TempVars("ActivadoCorreoAutomatico") = "No" Then
        m_sTextoConfiguracionCorreosAutomaticos = "Aún no implementado correos automáticos"
    Else
        If Me.CorreosAutomaticos <> "Sí" Then
            m_sTextoConfiguracionCorreosAutomaticos = "Configuración de Correos Automáticos DESACTIVADA"
            TextoConfiguracionCorreosAutomaticos = m_sTextoConfiguracionCorreosAutomaticos
            Exit Property
        End If
        If Me.DatosParaCorreoAutomaticoRegistrados = EnumSiNo.Sí Then
            m_sTextoConfiguracionCorreosAutomaticos = "Configuración de Correos Automáticos ACTIVADA"
        Else
            m_sTextoConfiguracionCorreosAutomaticos = "Configuración de Correos Automáticos DESACTIVADA"
        End If
    End If
    
    
    TextoConfiguracionCorreosAutomaticos = m_sTextoConfiguracionCorreosAutomaticos
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.TextoConfiguracionCorreosAutomaticos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get Expedientes() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjExpedientes Is Nothing Then
        Set Expedientes = m_ObjExpedientes
        Exit Property
    End If
    
    Set m_ObjExpedientes = constructor.getExpedientes(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Expedientes = m_ObjExpedientes
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Expedientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set Expedientes(p_Valor As Variant)

    Set m_ObjExpedientes = p_Valor
    
End Property

Public Property Get Responsables() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjResponsables Is Nothing Then
        Set Responsables = m_ObjResponsables
        Exit Property
    End If
    
    Set m_ObjResponsables = constructor.getResponsables(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Responsables = m_ObjResponsables
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Responsables ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set Responsables(p_Valor As Variant)

    Set m_ObjResponsables = p_Valor
    
End Property

Public Property Get DiasCancelacionMARGA() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sDiasCancelacionMARGA = Me.Configuracion.DiasCancelacionMARGA
    DiasCancelacionMARGA = m_sDiasCancelacionMARGA
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasCancelacionMARGA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Get AutocancelacionPreMARGA() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sAutocancelacionPreMARGA = Me.Configuracion.AutocancelacionPreMARGA
    AutocancelacionPreMARGA = m_sAutocancelacionPreMARGA
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.AutocancelacionPreMARGA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Get AutocancelacionMARGA() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sAutocancelacionMARGA = Me.Configuracion.AutocancelacionMARGA
    AutocancelacionMARGA = m_sAutocancelacionMARGA
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.AutocancelacionMARGA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get CharsetISO() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCharsetISO <> "" Then
        CharsetISO = m_sCharsetISO
        Exit Property
    End If
    m_sCharsetISO = "ISO-8859-1"
    CharsetISO = m_sCharsetISO
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CharsetISO  ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Get CharsetUTF() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCharsetUTF <> "" Then
        CharsetUTF = m_sCharsetUTF
        Exit Property
    End If
    m_sCharsetUTF = "UTF-8"
    CharsetUTF = m_sCharsetUTF
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CharsetUTF  ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property



Public Property Get URLDirectorioPlantillasHTMLVersionEstablecida() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioPlantillasHTMLVersionEstablecida <> "" Then
        URLDirectorioPlantillasHTMLVersionEstablecida = m_sURLDirectorioPlantillasHTMLVersionEstablecida
        Exit Property
    End If
    If Me.Configuracion Is Nothing Then
        m_Error = "No se ha podido establecer m_sURLDirectorioPlantillasHTMLVersionEstablecida"
        Err.Raise 1000
    End If
    If Me.Configuracion.VersionPlantillasHTML = "" Then
        m_Error = "No se ha podido establecer m_sURLDirectorioPlantillasHTMLVersionEstablecida"
        Err.Raise 1000
    End If
    m_sURLDirectorioPlantillasHTMLVersionEstablecida = Me.URLDirectorioPlantillasHTML & Me.Configuracion.VersionPlantillasHTML & "\"
    
    If Not fso.FileExists(m_sURLDirectorioPlantillasHTMLVersionEstablecida) Then
        m_Error = "No es alcanzable m_sURLDirectorioPlantillasHTMLVersionEstablecida " & m_sURLDirectorioPlantillasHTMLVersionEstablecida
        Err.Raise 1000
    End If
    
    URLDirectorioPlantillasHTMLVersionEstablecida = m_sURLDirectorioPlantillasHTMLVersionEstablecida
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioPlantillasHTMLVersionEstablecida ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLDirectorioPlantillasEXCELConVersion() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioPlantillasEXCELConVersion <> "" Then
        URLDirectorioPlantillasEXCELConVersion = m_sURLDirectorioPlantillasEXCELConVersion
        Exit Property
    End If
    If Me.Configuracion Is Nothing Then
        m_Error = "No se ha podido establecer m_sURLDirectorioPlantillasEXCELConVersion"
        Err.Raise 1000
    End If
    If Me.Configuracion.VersionPlantillasEXCEL = "" Then
        m_Error = "No se ha podido establecer m_sURLDirectorioPlantillasEXCELConVersion"
        Err.Raise 1000
    End If
    m_sURLDirectorioPlantillasEXCELConVersion = Me.URLDirectorioPlantillasExcel & Me.Configuracion.VersionPlantillasEXCEL & "\"
    
    If Not fso.FileExists(m_sURLDirectorioPlantillasEXCELConVersion) Then
        m_Error = "No es alcanzable m_sURLDirectorioPlantillasEXCELConVersion " & m_sURLDirectorioPlantillasEXCELConVersion
        Err.Raise 1000
    End If
    
    URLDirectorioPlantillasEXCELConVersion = m_sURLDirectorioPlantillasEXCELConVersion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioPlantillasEXCELConVersion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLPlantillasHTMLExcelSolicitud() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
'    If m_sURLPlantillasHTMLExcelSolicitud <> "" Then
'        URLPlantillasHTMLExcelSolicitud = m_sURLPlantillasHTMLExcelSolicitud
'        Exit Property
'    End If
    m_sURLPlantillasHTMLExcelSolicitud = Me.URLDirectorioPlantillasHTMLVersionEstablecida & Me.ColPlantillas(CStr(EnumPlantillas.EnvioExcel))
    
    If Not fso.FileExists(m_sURLPlantillasHTMLExcelSolicitud) Then
        m_Error = "No es alcanzable m_sURLPlantillasHTMLExcelSolicitud " & m_sURLPlantillasHTMLExcelSolicitud
        Err.Raise 1000
    End If
    
    URLPlantillasHTMLExcelSolicitud = m_sURLPlantillasHTMLExcelSolicitud
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillasHTMLExcelSolicitud ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLHTMLRecordatorioExcel1() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLHTMLRecordatorioExcel1 <> "" Then
        URLHTMLRecordatorioExcel1 = m_sURLHTMLRecordatorioExcel1
        Exit Property
    End If
    m_sURLHTMLRecordatorioExcel1 = Me.URLDirectorioPlantillasHTMLVersionEstablecida & Me.ColPlantillas(CStr(EnumPlantillas.RecordatorioExcel1))
    
    If Not fso.FileExists(m_sURLHTMLRecordatorioExcel1) Then
        m_Error = "No es alcanzable m_sURLHTMLRecordatorioExcel1 " & m_sURLHTMLRecordatorioExcel1
        Err.Raise 1000
    End If
    
    URLHTMLRecordatorioExcel1 = m_sURLHTMLRecordatorioExcel1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLHTMLRecordatorioExcel1 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLHTMLRecordatorioExcel2() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLHTMLRecordatorioExcel2 <> "" Then
        URLHTMLRecordatorioExcel2 = m_sURLHTMLRecordatorioExcel2
        Exit Property
    End If
    m_sURLHTMLRecordatorioExcel2 = Me.URLDirectorioPlantillasHTMLVersionEstablecida & Me.ColPlantillas(CStr(EnumPlantillas.RecordatorioExcel2))
    
    If Not fso.FileExists(m_sURLHTMLRecordatorioExcel2) Then
        m_Error = "No es alcanzable m_sURLHTMLRecordatorioExcel2 " & m_sURLHTMLRecordatorioExcel2
        Err.Raise 1000
    End If
    
    URLHTMLRecordatorioExcel2 = m_sURLHTMLRecordatorioExcel2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLHTMLRecordatorioExcel2 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLHTMLRecordatorioMarga1() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLHTMLRecordatorioMarga1 <> "" Then
        URLHTMLRecordatorioMarga1 = m_sURLHTMLRecordatorioMarga1
        Exit Property
    End If
    m_sURLHTMLRecordatorioMarga1 = Me.URLDirectorioPlantillasHTMLVersionEstablecida & Me.ColPlantillas(CStr(EnumPlantillas.RecordatorioMarga1))
    
    If Not fso.FileExists(m_sURLHTMLRecordatorioMarga1) Then
        m_Error = "No es alcanzable m_sURLHTMLRecordatorioMarga1 " & m_sURLHTMLRecordatorioMarga1
        Err.Raise 1000
    End If
    
    URLHTMLRecordatorioMarga1 = m_sURLHTMLRecordatorioMarga1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLHTMLRecordatorioMarga1 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLHTMLRecordatorioMarga2() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    
    m_sURLHTMLRecordatorioMarga2 = Me.URLDirectorioPlantillasHTMLVersionEstablecida & Me.ColPlantillas(CStr(EnumPlantillas.RecordatorioMarga2))
    
    If Not fso.FileExists(m_sURLHTMLRecordatorioMarga2) Then
        m_Error = "No es alcanzable m_sURLHTMLRecordatorioMarga2 " & m_sURLHTMLRecordatorioMarga2
        Err.Raise 1000
    End If
    
    URLHTMLRecordatorioMarga2 = m_sURLHTMLRecordatorioMarga2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLHTMLRecordatorioMarga2 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLHTMLAutocancelacionPREMARGA() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    
    m_sURLHTMLAutocancelacionPREMARGA = Me.URLDirectorioPlantillasHTMLVersionEstablecida & Me.ColPlantillas(CStr(EnumPlantillas.CancelacionPreMarga))
    
    If Not fso.FileExists(m_sURLHTMLAutocancelacionPREMARGA) Then
        m_Error = "No es alcanzable m_sURLHTMLAutocancelacionPREMARGA " & m_sURLHTMLAutocancelacionPREMARGA
        Err.Raise 1000
    End If
    
    URLHTMLAutocancelacionPREMARGA = m_sURLHTMLAutocancelacionPREMARGA
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLHTMLAutocancelacionPREMARGA ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLHTMLAutocancelacionMARGA() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    
    m_sURLHTMLAutocancelacionMARGA = Me.URLDirectorioPlantillasHTMLVersionEstablecida & Me.ColPlantillas(CStr(EnumPlantillas.CancelacionMarga))
    
    If Not fso.FileExists(m_sURLHTMLAutocancelacionMARGA) Then
        m_Error = "No es alcanzable m_sURLHTMLAutocancelacionMARGA " & m_sURLHTMLAutocancelacionMARGA
        Err.Raise 1000
    End If
    
    URLHTMLAutocancelacionMARGA = m_sURLHTMLAutocancelacionMARGA
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLHTMLAutocancelacionMARGA ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLDirectorioAplicaciones() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioAplicaciones <> "" Then
        URLDirectorioAplicaciones = m_sURLDirectorioAplicaciones
        Exit Property
    End If
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesLocal
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesRemotas
    Else
        m_Error = "No se sabe si se está usando en local o en remoto"
        Err.Raise 1000
    End If
    
   
    If Not fso.FolderExists(m_sURLDirectorioAplicaciones) Then
        m_Error = "No es alcanzable m_sURLDirectorioAplicaciones " & m_sURLDirectorioAplicaciones
        Err.Raise 1000
    End If
    
    URLDirectorioAplicaciones = m_sURLDirectorioAplicaciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAplicaciones ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property




Public Property Get URLPlantillaAlta() As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLPlantillaAlta <> "" Then
        URLPlantillaAlta = m_sURLPlantillaAlta
        Exit Property
    End If
    m_sURLPlantillaAlta = Me.URLDirectorioPlantillasEXCELConVersion & "Solicitud datos - alta.xlsx"
    If Not fso.FileExists(m_sURLPlantillaAlta) Then
        m_Error = "No se puede alcanzar m_sURLPlantillaAlta " & vbNewLine & m_sURLPlantillaAlta
        m_sURLPlantillaAlta = ""
        Err.Raise 1000
    End If
    URLPlantillaAlta = m_sURLPlantillaAlta
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillaAlta ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLPlantillaTraspaso() As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLPlantillaTraspaso <> "" Then
        URLPlantillaTraspaso = m_sURLPlantillaTraspaso
        Exit Property
    End If
    m_sURLPlantillaTraspaso = Me.URLDirectorioPlantillasEXCELConVersion & "Solicitud datos - Traspaso.xlsx"
    If Not fso.FileExists(m_sURLPlantillaTraspaso) Then
        m_Error = "No se puede alcanzar m_sURLPlantillaTraspaso " & vbNewLine & m_sURLPlantillaTraspaso
        m_sURLPlantillaTraspaso = ""
        Err.Raise 1000
    End If
    URLPlantillaTraspaso = m_sURLPlantillaTraspaso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillaTraspaso ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLDirectorioPlantillas() As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioPlantillas <> "" Then
        URLDirectorioPlantillas = m_sURLDirectorioPlantillas
        Exit Property
    End If
    m_sURLDirectorioPlantillas = Me.URLDirectorioRecursos & "plantillas\"
    If Not fso.FolderExists(m_sURLDirectorioPlantillas) Then
        m_Error = "No se puede alcanzar m_sURLDirectorioPlantillas " & vbNewLine & m_sURLDirectorioPlantillas
        m_sURLDirectorioPlantillas = ""
        Err.Raise 1000
    End If
    URLDirectorioPlantillas = m_sURLDirectorioPlantillas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioPlantillas ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLDirectorioPlantillasHTML() As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioPlantillasHTML <> "" Then
        URLDirectorioPlantillasHTML = m_sURLDirectorioPlantillasHTML
        Exit Property
    End If
    m_sURLDirectorioPlantillasHTML = Me.URLDirectorioPlantillas & "HTML\"
    If Not fso.FolderExists(m_sURLDirectorioPlantillasHTML) Then
        m_Error = "No se puede alcanzar m_sURLDirectorioPlantillasHTML " & vbNewLine & m_sURLDirectorioPlantillasHTML
        m_sURLDirectorioPlantillasHTML = ""
        Err.Raise 1000
    End If
    URLDirectorioPlantillasHTML = m_sURLDirectorioPlantillasHTML
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioPlantillasHTML ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLDirectorioPlantillasExcel() As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioPlantillasExcel <> "" Then
        URLDirectorioPlantillasExcel = m_sURLDirectorioPlantillasExcel
        Exit Property
    End If
    m_sURLDirectorioPlantillasExcel = Me.URLDirectorioPlantillas & "EXCEL\"
    If Not fso.FolderExists(m_sURLDirectorioPlantillasExcel) Then
        m_Error = "No se puede alcanzar m_sURLDirectorioPlantillasExcel " & vbNewLine & m_sURLDirectorioPlantillasExcel
        m_sURLDirectorioPlantillasExcel = ""
        Err.Raise 1000
    End If
    URLDirectorioPlantillasExcel = m_sURLDirectorioPlantillasExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioPlantillasExcel ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get ColPendientesEnvioExcel() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendientesEnvioExcel Is Nothing Then
        Set ColPendientesEnvioExcel = m_objColPendientesEnvioExcel
        Exit Property
    End If
    Set m_objColPendientesEnvioExcel = constructor.getSolicitudesPendientesEnvioExcel(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendientesEnvioExcel = m_objColPendientesEnvioExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendientesEnvioExcel ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendientesEnvioExcel(ByVal objNewValue As Variant)

    Set m_objColPendientesEnvioExcel = objNewValue
    
End Property

Public Property Get ColPendientesAltaSolicitud() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendientesAltaSolicitud Is Nothing Then
        Set ColPendientesAltaSolicitud = m_objColPendientesAltaSolicitud
        Exit Property
    End If
    Set m_objColPendientesAltaSolicitud = constructor.getPendientesAltaSolicitud(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendientesAltaSolicitud = m_objColPendientesAltaSolicitud
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendientesAltaSolicitud ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendientesAltaSolicitud(ByVal objNewValue As Variant)

    Set m_objColPendientesAltaSolicitud = objNewValue
    
End Property

Public Property Get ColPendientesRecordatorioExcel1() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendientesRecordatorioExcel1 Is Nothing Then
        Set ColPendientesRecordatorioExcel1 = m_objColPendientesRecordatorioExcel1
        Exit Property
    End If
    Set m_objColPendientesRecordatorioExcel1 = constructor.getSolicitudesPendientesRecordatorioExcel1(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendientesRecordatorioExcel1 = m_objColPendientesRecordatorioExcel1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendientesRecordatorioExcel1 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendientesRecordatorioExcel1(ByVal objNewValue As Variant)

    Set m_objColPendientesRecordatorioExcel1 = objNewValue
    
End Property
Public Property Get ColPendientesRecordatorioExcel2() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendientesRecordatorioExcel2 Is Nothing Then
        Set ColPendientesRecordatorioExcel2 = m_objColPendientesRecordatorioExcel2
        Exit Property
    End If
    Set m_objColPendientesRecordatorioExcel2 = constructor.getSolicitudesPendientesRecordatorioExcel2(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendientesRecordatorioExcel2 = m_objColPendientesRecordatorioExcel2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendientesRecordatorioExcel2 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendientesRecordatorioExcel2(ByVal objNewValue As Variant)

    Set m_objColPendientesRecordatorioExcel2 = objNewValue
    
End Property
Public Property Get ColPendientesExcelRelleno() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendientesExcelRelleno Is Nothing Then
        Set ColPendientesExcelRelleno = m_objColPendientesExcelRelleno
        Exit Property
    End If
    Set m_objColPendientesExcelRelleno = constructor.getSolicitudesPendientesEnvioExcel(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendientesExcelRelleno = m_objColPendientesExcelRelleno
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendientesExcelRelleno ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendientesExcelRelleno(ByVal objNewValue As Variant)

    Set m_objColPendientesExcelRelleno = objNewValue
    
End Property

Public Property Get ColPendienteRegistroMarga() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendienteRegistroMarga Is Nothing Then
        Set ColPendienteRegistroMarga = m_objColPendienteRegistroMarga
        Exit Property
    End If
    Set m_objColPendienteRegistroMarga = constructor.getSolicitudesPendientesAltaMarga(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendienteRegistroMarga = m_objColPendienteRegistroMarga
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendienteRegistroMarga ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendienteRegistroMarga(ByVal objNewValue As Variant)

    Set m_objColPendienteRegistroMarga = objNewValue
    
End Property

Public Property Get ColPendienteRecordatorioEnvioDPS1() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendienteRecordatorioEnvioDPS1 Is Nothing Then
        Set ColPendienteRecordatorioEnvioDPS1 = m_objColPendienteRecordatorioEnvioDPS1
        Exit Property
    End If
    Set m_objColPendienteRecordatorioEnvioDPS1 = constructor.getSolicitudesPendientesRecordatorioEnvioDPS1(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendienteRecordatorioEnvioDPS1 = m_objColPendienteRecordatorioEnvioDPS1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendienteRecordatorioEnvioDPS1 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendienteRecordatorioEnvioDPS1(ByVal objNewValue As Variant)

    Set m_objColPendienteRecordatorioEnvioDPS1 = objNewValue
    
End Property
Public Property Get ColPendienteRecordatorioEnvioDPS2() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendienteRecordatorioEnvioDPS2 Is Nothing Then
        Set ColPendienteRecordatorioEnvioDPS2 = m_objColPendienteRecordatorioEnvioDPS2
        Exit Property
    End If
    Set m_objColPendienteRecordatorioEnvioDPS2 = constructor.getSolicitudesPendientesRecordatorioEnvioDPS2(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendienteRecordatorioEnvioDPS2 = m_objColPendienteRecordatorioEnvioDPS2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendienteRecordatorioEnvioDPS2 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendienteRecordatorioEnvioDPS2(ByVal objNewValue As Variant)

    Set m_objColPendienteRecordatorioEnvioDPS2 = objNewValue
    
End Property
Public Property Get ColPendienteEnvioDPS() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendienteEnvioDPS Is Nothing Then
        Set ColPendienteEnvioDPS = m_objColPendienteEnvioDPS
        Exit Property
    End If
    Set m_objColPendienteEnvioDPS = constructor.getSolicitudesPendientesEnvioDPS(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendienteEnvioDPS = m_objColPendienteEnvioDPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendienteEnvioDPS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendienteEnvioDPS(ByVal objNewValue As Variant)

    Set m_objColPendienteEnvioDPS = objNewValue
    
End Property

Public Property Get ColPendienteRegistroONS() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendienteRegistroONS Is Nothing Then
        Set ColPendienteRegistroONS = m_objColPendienteRegistroONS
        Exit Property
    End If
    Set m_objColPendienteRegistroONS = constructor.getSolicitudesPendientesRegistroONS(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendienteRegistroONS = m_objColPendienteRegistroONS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendienteRegistroONS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Set ColPendienteRegistroONS(ByVal objNewValue As Variant)

    Set m_objColPendienteRegistroONS = objNewValue
    
End Property

Public Property Get ColTramitados() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColTramitados Is Nothing Then
        Set ColTramitados = m_objColTramitados
        Exit Property
    End If
    Set m_objColTramitados = constructor.getSolicitudesTramitadas(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColTramitados = m_objColTramitados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTramitados ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Set ColTramitados(ByVal objNewValue As Variant)

    Set m_objColTramitados = objNewValue
    
End Property

Public Property Get ColDesestimados() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColDesestimados Is Nothing Then
        Set ColDesestimados = m_objColDesestimados
        Exit Property
    End If
    Set m_objColDesestimados = constructor.getSolicitudesDesestimadas(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColDesestimados = m_objColDesestimados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColDesestimados ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColDesestimados(ByVal objNewValue As Variant)

    Set m_objColDesestimados = objNewValue
    
End Property


Public Property Get ColCancelados() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColCancelados Is Nothing Then
        Set ColCancelados = m_objColCancelados
        Exit Property
    End If
    Set m_objColCancelados = constructor.getSolicitudesCanceladas(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColCancelados = m_objColCancelados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColCancelados ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColCancelados(ByVal objNewValue As Variant)

    Set m_objColCancelados = objNewValue
    
End Property

Public Property Get ColPendienteEnvioTraspasoONS() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPendienteEnvioTraspasoONS Is Nothing Then
        Set ColPendienteEnvioTraspasoONS = m_objColPendienteEnvioTraspasoONS
        Exit Property
    End If
    Set m_objColPendienteEnvioTraspasoONS = constructor.getSolicitudesPendientesEnvioTraspasoONS(p_Tramitador:=m_ObjTramitadorParaTareas.UsuarioRed, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPendienteEnvioTraspasoONS = m_objColPendienteEnvioTraspasoONS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendienteEnvioTraspasoONS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColPendienteEnvioTraspasoONS(ByVal objNewValue As Variant)

    Set m_objColPendienteEnvioTraspasoONS = objNewValue
    
End Property

Public Property Get ColEnTramite() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColEnTramite Is Nothing Then
        Set ColEnTramite = m_objColEnTramite
        Exit Property
    End If
    Set m_objColEnTramite = constructor.getSolicitudesEnTramite(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEnTramite = m_objColEnTramite
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEnTramite ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColEnTramite(ByVal objNewValue As Variant)

    Set m_objColEnTramite = objNewValue
    
End Property

Public Property Get ColSolicitudes() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColSolicitudes Is Nothing Then
        Set ColSolicitudes = m_objColSolicitudes
        Exit Property
    End If
    Set m_objColSolicitudes = constructor.getSolicitudes(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColSolicitudes = m_objColSolicitudes
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColSolicitudes ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColSolicitudes(ByVal objNewValue As Variant)

    Set m_objColSolicitudes = objNewValue
    
End Property

Public Property Get Coltramitadores() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColtramitadores Is Nothing Then
        Set Coltramitadores = m_objColtramitadores
        Exit Property
    End If
    Set m_objColtramitadores = constructor.getUsuariosTramitadores(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Coltramitadores = m_objColtramitadores
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Coltramitadores ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set Coltramitadores(ByVal objNewValue As Variant)

    Set m_objColtramitadores = objNewValue
    
End Property



Public Property Get ColMotivos() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColMotivos Is Nothing Then
        Set ColMotivos = m_objColMotivos
        Exit Property
    End If
    Set m_objColMotivos = constructor.getMotivos(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColMotivos = m_objColMotivos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColMotivos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColMotivos(ByVal objNewValue As Variant)

    Set m_objColMotivos = objNewValue
    
End Property

Public Property Get colExplicacionTipoSeleccionTarea() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjcolExplicacionTipoSeleccionTarea Is Nothing Then
        Set colExplicacionTipoSeleccionTarea = m_ObjcolExplicacionTipoSeleccionTarea
        Exit Property
    End If
    Set m_ObjcolExplicacionTipoSeleccionTarea = New Scripting.Dictionary
    m_ObjcolExplicacionTipoSeleccionTarea.CompareMode = TextCompare
    With m_ObjcolExplicacionTipoSeleccionTarea
        
        .Add CStr(EnumEstado.PendienteEnvioExcel), _
            "Son las solicitudes que aún no se han entregado el Excel al solicitante"
        .Add CStr(EnumEstado.PendienteRecordatorioExcel1), _
            "Son las solicitudes que aún no se han entregado el Excel al solicitante y ya se ha alcanzado el tiempo del primer correo"
        .Add CStr(EnumEstado.PendienteRecordatorioExcel2), _
            "Son las solicitudes que aún no se han entregado el Excel al solicitante y ya se ha alcanzado el tiempo del segundo correo"
        .Add CStr(EnumEstado.PendienteRegistroMarga), _
            "Son las solicitudes que ya han entregado el Excel relleno y el tramitador todavía no ha indicado que se ha dado de alta en Marga"
        .Add CStr(EnumEstado.PendienteRecordatorioEnvioDPS1), _
            "Son las solicitudes que el tramitador ya ha dado el alta en Marga y el Solicitante aún no ha enviado el DPS. Ha llegado el momento del primer correo."
        .Add CStr(EnumEstado.PendienteRecordatorioEnvioDPS2), _
            "Son las solicitudes que el tramitador ya ha dado el alta en Marga y el Solicitante aún no ha enviado el DPS. Ha llegado el momento del segundo correo."
        .Add CStr(EnumEstado.PendienteEnvioDPS), "Son las solicitudes que ya han sido dadas de alta en Marga por el tramitados y aún el solicitante no ha enviado el DPS"
        .Add CStr(EnumEstado.PendienteRegistroONS), "Son las solicitudes que se ha enviado el DPS y aún no ha llegado el envío de ONS"
        .Add CStr(EnumEstado.PendienteEnvioTraspasoONS), "Son las solicitudes tipo Traspaso ya han entregado el Excel relleno y el tramitador todavía no enviado el formulario a la ONS"
        .Add CStr(EnumEstado.PendienteRegistroEnHPS), "Son las solicitudes que se ha hecho el envío ONS y aún no se ha trasladado a HPS"
    
    End With
    
    Set colExplicacionTipoSeleccionTarea = m_ObjcolExplicacionTipoSeleccionTarea
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colExplicacionTipoSeleccionTarea ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColEstadosTexto() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColEstadosTexto Is Nothing Then
        Set ColEstadosTexto = m_ObjColEstadosTexto
        Exit Property
    End If
    Set m_ObjColEstadosTexto = New Scripting.Dictionary
    m_ObjColEstadosTexto.CompareMode = TextCompare
    With m_ObjColEstadosTexto
        
        .Add CStr(EnumEstado.PendienteEnvioExcel), "PendienteEnvioExcel"
        .Add CStr(EnumEstado.PendienteRecordatorioExcel1), "PendienteRecordatorioExcel1"
        .Add CStr(EnumEstado.PendienteRecordatorioExcel2), "PendienteRecordatorioExcel2"
        .Add CStr(EnumEstado.PendienteExcelRelleno), "PendienteExcelRelleno"
        .Add CStr(EnumEstado.PendienteRegistroMarga), "PendienteRegistroMarga"
        .Add CStr(EnumEstado.PendienteRecordatorioEnvioDPS1), "PendienteRecordatorioEnvioDPS1"
        .Add CStr(EnumEstado.PendienteRecordatorioEnvioDPS2), "PendienteRecordatorioEnvioDPS2"
        .Add CStr(EnumEstado.PendienteEnvioDPS), "PendienteEnvioDPS"
        .Add CStr(EnumEstado.PendienteRegistroONS), "PendienteRegistroONS"
        .Add CStr(EnumEstado.PendienteRegistroEnHPS), "PendienteRegistroEnHPS"
        .Add CStr(EnumEstado.Tramitado), "Tramitado"
        .Add CStr(EnumEstado.Desestimado), "Desestimado"
        .Add CStr(EnumEstado.Cancelado), "Cancelado"
        .Add CStr(EnumEstado.PendienteEnvioTraspasoONS), "PendienteEnvioTraspasoONS"
        .Add CStr(EnumEstado.EnTramite), "EnTramite"
    End With
    
    Set ColEstadosTexto = m_ObjColEstadosTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColTextoEstados() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColTextoEstados Is Nothing Then
        Set ColTextoEstados = m_ObjColTextoEstados
        Exit Property
    End If
    Set m_ObjColTextoEstados = New Scripting.Dictionary
    m_ObjColTextoEstados.CompareMode = TextCompare
    With m_ObjColTextoEstados
        
        .Add "PendienteEnvioExcel", EnumEstado.PendienteEnvioExcel
        .Add "PendienteRecordatorioExcel1", EnumEstado.PendienteRecordatorioExcel1
        .Add "PendienteRecordatorioExcel2", EnumEstado.PendienteRecordatorioExcel2
        .Add "PendienteExcelRelleno", EnumEstado.PendienteExcelRelleno
        .Add "PendienteRegistroMarga", EnumEstado.PendienteRegistroMarga
        .Add "PendienteRecordatorioEnvioDPS1", EnumEstado.PendienteRecordatorioEnvioDPS1
        .Add "PendienteRecordatorioEnvioDPS2", EnumEstado.PendienteRecordatorioEnvioDPS2
        .Add "PendienteEnvioDPS", EnumEstado.PendienteEnvioDPS
        .Add "PendienteRegistroONS", EnumEstado.PendienteRegistroONS
        .Add "PendienteRegistroEnHPS", EnumEstado.PendienteRegistroEnHPS
        .Add "Tramitado", EnumEstado.Tramitado
        .Add "Desestimado", EnumEstado.Desestimado
        .Add "Cancelado", EnumEstado.Cancelado
        .Add "PendienteEnvioTraspasoONS", EnumEstado.PendienteEnvioTraspasoONS
        .Add "EnTramite", EnumEstado.EnTramite
    End With
    
    Set ColTextoEstados = m_ObjColTextoEstados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTextoEstados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColEstadosParaTitulo() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColEstadosParaTitulo Is Nothing Then
        Set ColEstadosParaTitulo = m_ObjColEstadosParaTitulo
        Exit Property
    End If
    Set m_ObjColEstadosParaTitulo = New Scripting.Dictionary
    m_ObjColEstadosParaTitulo.CompareMode = TextCompare
    With m_ObjColEstadosParaTitulo
        
        .Add CStr(EnumEstado.PendienteEnvioExcel), "Pendiente de Envío Excel"
        .Add CStr(EnumEstado.PendienteRecordatorioExcel1), "Pendiente de Recordatorio 1 Envio Excel"
        .Add CStr(EnumEstado.PendienteRecordatorioExcel2), "Pendiente de Recordatorio 2 Envio Excel"
        .Add CStr(EnumEstado.PendienteExcelRelleno), "Pendiente Excel Relleno"
        .Add CStr(EnumEstado.PendienteRegistroMarga), "Pendiente Alta Marga"
        .Add CStr(EnumEstado.PendienteRecordatorioEnvioDPS1), "Pendiente de Recordatorio 1 Envio DPS"
        .Add CStr(EnumEstado.PendienteRecordatorioEnvioDPS2), "Pendiente de Recordatorio 2 Envio DPS"
        .Add CStr(EnumEstado.PendienteEnvioDPS), "Pendiente Envío DPS"
        .Add CStr(EnumEstado.PendienteRegistroONS), "Pendiente Envío ONS"
        .Add CStr(EnumEstado.PendienteRegistroEnHPS), "Pendiente fin tramitación"
        .Add CStr(EnumEstado.Tramitado), "Tramitado"
        .Add CStr(EnumEstado.Desestimado), "Desestimado"
        .Add CStr(EnumEstado.Cancelado), "Cancelado"
        .Add CStr(EnumEstado.PendienteEnvioTraspasoONS), "Pendiente de Enviar Tramitación a la ONS (Traspaso)"
        .Add CStr(EnumEstado.EnTramite), "Pendiente de algún trámite"
    End With
    
    Set ColEstadosParaTitulo = m_ObjColEstadosParaTitulo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosParaTitulo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColTiposTexto() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColTiposTexto Is Nothing Then
        Set ColTiposTexto = m_ObjColTiposTexto
        Exit Property
    End If
    Set m_ObjColTiposTexto = New Scripting.Dictionary
    m_ObjColTiposTexto.CompareMode = TextCompare
    With m_ObjColTiposTexto
        .Add CStr(EnumTipo.Alta), "Alta"
        .Add CStr(EnumTipo.Traspaso), "Traspaso"
        .Add CStr(EnumTipo.Renovacion), "Renovación"
        .Add CStr(EnumTipo.ElevacionGrado), "Elevación Grado"
    End With
    
    Set ColTiposTexto = m_ObjColTiposTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTiposTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColTextoTipos() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColTextoTipos Is Nothing Then
        Set ColTextoTipos = m_ObjColTextoTipos
        Exit Property
    End If
    Set m_ObjColTextoTipos = New Scripting.Dictionary
    m_ObjColTextoTipos.CompareMode = TextCompare
    With m_ObjColTextoTipos
        .Add "Alta", EnumTipo.Alta
        .Add "Traspaso", EnumTipo.Traspaso
        .Add "Renovación", EnumTipo.Renovacion
        .Add "Elevación Grado", EnumTipo.ElevacionGrado
    End With
    
    Set ColTextoTipos = m_ObjColTextoTipos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTextoTipos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColNombrePlantillas() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColNombrePlantillas Is Nothing Then
        Set ColNombrePlantillas = m_ObjColNombrePlantillas
        Exit Property
    End If
    Set m_ObjColNombrePlantillas = New Scripting.Dictionary
    m_ObjColNombrePlantillas.CompareMode = TextCompare
    With m_ObjColNombrePlantillas
        .Add CStr(EnumTipoEnvioCorreo.EnvioExcel), "ENVIO_EXCEL.html"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioExcel1), "RECORDATORIO_EXCEL1.html"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioExcel2), "RECORDATORIO_EXCEL2.html"
        .Add CStr(EnumTipoEnvioCorreo.CancelacionPreMarga), "AUTOCANCELACION_PREMARGA.html"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioMarga1), "RECORDATORIO_MARGA1.html"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioMarga2), "RECORDATORIO_MARGA2.html"
        .Add CStr(EnumTipoEnvioCorreo.CancelacionMarga), "AUTOCANCELACION_MARGA.html"
    
    End With
    
    Set ColNombrePlantillas = m_ObjColNombrePlantillas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColNombrePlantillas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColTiposCorreos() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColTiposCorreos Is Nothing Then
        Set ColTiposCorreos = m_objColTiposCorreos
        Exit Property
    End If
    Set m_objColTiposCorreos = New Scripting.Dictionary
    m_objColTiposCorreos.CompareMode = TextCompare
    With m_objColTiposCorreos
        .Add CStr(EnumTipoEnvioCorreo.EnvioExcel), "SOLICITUD DATOS POR EXCEL"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioExcel1), "RECORDATORIO 1 SOLICITUD DATOS POR EXCEL"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioExcel2), "RECORDATORIO 2 SOLICITUD DATOS POR EXCEL"
        .Add CStr(EnumTipoEnvioCorreo.CancelacionPreMarga), "CANCELACIÓN PRE MARGA"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioMarga1), "RECORDATORIO 1 ENVÍO DPS"
        .Add CStr(EnumTipoEnvioCorreo.RecordatorioMarga2), "RECORDATORIO 2 ENVÍO DPS"
        .Add CStr(EnumTipoEnvioCorreo.CancelacionMarga), "CANCELACIÓN MARGA"
    End With
    
    Set ColTiposCorreos = m_objColTiposCorreos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTiposCorreos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get FaltaAlgunDatoConfiguracion() As Boolean
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_bFaltaAlgunDatoConfiguracion = DatosConfInsuficientes(Me.Configuracion, EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    FaltaAlgunDatoConfiguracion = m_bFaltaAlgunDatoConfiguracion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.FaltaAlgunDatoConfiguracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get CorreosAutomaticos() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sCorreosAutomaticos = Me.Configuracion.CorreosAutomaticos
    CorreosAutomaticos = m_sCorreosAutomaticos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CorreosAutomaticos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get EmailDirectorSeguridad() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sEmailDirectorSeguridad = Me.Configuracion.EmailDirectorSeguridad
    EmailDirectorSeguridad = m_sEmailDirectorSeguridad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EmailDirectorSeguridad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasParaRecordatorioExcel1() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sDiasParaRecordatorioExcel1 = Me.Configuracion.DiasParaRecordatorioExcel1
    DiasParaRecordatorioExcel1 = m_sDiasParaRecordatorioExcel1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasParaRecordatorioExcel1 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasParaRecordatorioExcel2() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sDiasParaRecordatorioExcel2 = Me.Configuracion.DiasParaRecordatorioExcel2
    DiasParaRecordatorioExcel2 = m_sDiasParaRecordatorioExcel2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasParaRecordatorioExcel2 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasCancelacionPreMARGA() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sDiasCancelacionPreMARGA = Me.Configuracion.DiasCancelacionPreMARGA
    DiasCancelacionPreMARGA = m_sDiasCancelacionPreMARGA
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasCancelacionPreMARGA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasParaRecordatorioRellenoMarga1() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sDiasParaRecordatorioRellenoMarga1 = Me.Configuracion.DiasParaRecordatorioRellenoMarga1
    DiasParaRecordatorioRellenoMarga1 = m_sDiasParaRecordatorioRellenoMarga1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasParaRecordatorioRellenoMarga1 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasParaRecordatorioRellenoMarga2() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sDiasParaRecordatorioRellenoMarga2 = Me.Configuracion.DiasParaRecordatorioRellenoMarga2
    DiasParaRecordatorioRellenoMarga2 = m_sDiasParaRecordatorioRellenoMarga2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasParaRecordatorioRellenoMarga2 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get BuzonSeguridad() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Configuracion Is Nothing Then
        Exit Property
    End If
    m_sBuzonSeguridad = Me.Configuracion.BuzonSeguridad
    BuzonSeguridad = m_sBuzonSeguridad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.BuzonSeguridad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Configuracion() As Configuracion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjConfiguracion Is Nothing Then
        Set Configuracion = m_ObjConfiguracion
        Exit Property
    End If
    Set m_ObjConfiguracion = constructor.getConfiguracion(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Configuracion = m_ObjConfiguracion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Configuracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set Configuracion(p_Value As Variant)
    
    Set m_ObjConfiguracion = p_Value
    
    
    
End Property

Public Property Get ColAdministradores() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColAdministradores Is Nothing Then
        Set ColAdministradores = m_objColAdministradores
        Exit Property
    End If
    Set m_objColAdministradores = getUsuariosAdministradores(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAdministradores = m_objColAdministradores
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColAdministradores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get URLAchivoIni() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLAchivoIni <> "" Then
        URLAchivoIni = m_sURLAchivoIni
        Exit Property
    End If
    m_sURLAchivoIni = Application.CurrentProject.Path & "\" & fso.GetBaseName(Application.CurrentProject.Name) & ".ini"
    If Not fso.FileExists(m_sURLAchivoIni) Then
        m_Error = "No es alcanzable m_sURLAchivoIni " & vbNewLine & m_sURLAchivoIni
        m_sURLAchivoIni = ""
        Err.Raise 1000
    End If
    URLAchivoIni = m_sURLAchivoIni
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLAchivoIni ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get VersionAplicacion() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sVersionAplicacion <> "" Then
        VersionAplicacion = m_sVersionAplicacion
        Exit Property
    End If
    m_sVersionAplicacion = LeerIni("VersionAplicacion", "#ERR#")
    
    VersionAplicacion = m_sVersionAplicacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VersionAplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLDirectorioDocumentacion() As String
    Dim m_URL As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioDocumentacion <> "" Then
        URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
        Exit Property
    End If
    m_URL = Me.URLDirectorioAplicacion
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirectorioDocumentacion = m_URL & "ANEXOS\"
    
    If Not fso.FolderExists(m_sURLDirectorioDocumentacion) Then
        m_Error = "No es alcanzable m_sURLDirectorioDocumentacion " & m_sURLDirectorioDocumentacion
        Err.Raise 1000
    End If
    
    URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioDocumentacion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioRecursos() As String
    Dim m_URL As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioRecursos <> "" Then
        URLDirectorioRecursos = m_sURLDirectorioRecursos
        Exit Property
    End If
    m_URL = Me.URLDirectorioAplicacion
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirectorioRecursos = m_URL & "RECURSOS\"
    
    If Not fso.FolderExists(m_sURLDirectorioRecursos) Then
        m_Error = "No es alcanzable m_sURLDirectorioRecursos " & m_sURLDirectorioRecursos
        Err.Raise 1000
    End If
    
    URLDirectorioRecursos = m_sURLDirectorioRecursos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioRecursos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLDirectorioDocumentacionAyuda() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioDocumentacionAyuda <> "" Then
        URLDirectorioDocumentacionAyuda = m_sURLDirectorioDocumentacionAyuda
        Exit Property
    End If
   
    m_sURLDirectorioDocumentacionAyuda = Me.URLDirectorioRecursos & "documentos de ayuda\"
    
    If Not fso.FolderExists(m_sURLDirectorioDocumentacionAyuda) Then
        m_Error = "No es alcanzable m_sURLDirectorioDocumentacionAyuda " & m_sURLDirectorioDocumentacionAyuda
        Err.Raise 1000
    End If
    
    URLDirectorioDocumentacionAyuda = m_sURLDirectorioDocumentacionAyuda
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioDocumentacionAyuda ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLDirectorioLocal() As String
    
    Dim m_URL As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioLocal <> "" Then
        URLDirectorioLocal = m_sURLDirectorioLocal
        Exit Property
    End If
    m_URL = Environ("APPDATA") & "\Aplicaciones DYSN\"
    If Not fso.FolderExists(m_URL) Then
        fso.CreateFolder m_URL
    End If
    m_sURLDirectorioLocal = m_URL & "SOLICITUDES HPS\"
    If Not fso.FolderExists(m_sURLDirectorioLocal) Then
        fso.CreateFolder m_sURLDirectorioLocal
    End If
    
    
    URLDirectorioLocal = m_sURLDirectorioLocal
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioLocal ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioAplicacion() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioAplicacion <> "" Then
        URLDirectorioAplicacion = m_sURLDirectorioAplicacion
        Exit Property
    End If
    
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_sURLDirectorioAplicacion = m_URLRutaAplicacionLocal
    Else
        m_sURLDirectorioAplicacion = m_URLRutaAplicacionRemota
    End If
    If Not fso.FolderExists(m_sURLDirectorioAplicacion) Then
        m_Error = "No es alcanzable m_sURLDirectorioAplicacion " & m_sURLDirectorioAplicacion
        Err.Raise 1000
    End If
    
    URLDirectorioAplicacion = m_sURLDirectorioAplicacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAplicacion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get UltimoCambio() As UltimoCambio
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUltimoCambio Is Nothing Then
        Set UltimoCambio = m_ObjUltimoCambio
        Exit Property
    End If
    Set m_ObjUltimoCambio = constructor.getUltimoCambio(p_IDUsuario:=m_ObjUsuarioConectado.id, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UltimoCambio = m_ObjUltimoCambio
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UltimoCambio ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set UltimoCambio(p_Valor As Variant)
    
    Set m_ObjUltimoCambio = p_Valor
    
End Property
Public Property Get TotalPendientesGlobal() As Long
    TotalPendientesGlobal = m_lTotalPendientesGlobal
End Property

Public Property Get TotalPendientesUsuario() As Long
    TotalPendientesUsuario = m_lTotalPendientesUsuario
End Property
' En Entorno.cls

Public Property Get ColPendienteRegistroEnHPS() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    ' Si ya está cargado en memoria, lo devolvemos directamente
    If Not m_objColPendienteRegistroEnHPS Is Nothing Then
        Set ColPendienteRegistroEnHPS = m_objColPendienteRegistroEnHPS
        Exit Property
    End If
    
    ' Si no, llamamos al constructor para buscarlo en la BD
    Set m_objColPendienteRegistroEnHPS = constructor.getSolicitudesPendientesRegistroEnHPS(p_Error:=m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColPendienteRegistroEnHPS = m_objColPendienteRegistroEnHPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPendienteRegistroEnHPS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColPendienteRegistroEnHPS(ByVal objNewValue As Variant)
    Set m_objColPendienteRegistroEnHPS = objNewValue
End Property
Public Sub CalcularTotalesPendientes(p_Usuario As Usuario, p_EsTecnico As EnumSiNo)
    ' Este procedimiento:
    ' 1. Resetea contadores a 0.
    ' 2. Limpia la memoria (Caché) para obligar a recargar datos frescos de la BD.
    ' 3. Ejecuta las consultas y suma los totales.
    
    Dim m_Col As Scripting.Dictionary
    Dim p_Error As String
    
    On Error GoTo errores
    
    ' 1. RESETEAR CONTADORES
    m_lTotalPendientesGlobal = 0
    m_lTotalPendientesUsuario = 0
    
    ' 2. LIMPIAR CACHÉ (Variables privadas a Nothing)
    ' Esto es fundamental para que no acumule datos viejos y traiga lo real de la BD.
    Set m_objColPendientesEnvioExcel = Nothing
    Set m_objColPendientesRecordatorioExcel1 = Nothing
    Set m_objColPendientesRecordatorioExcel2 = Nothing
    Set m_objColPendientesExcelRelleno = Nothing
    Set m_objColPendienteRegistroMarga = Nothing
    Set m_objColPendienteRecordatorioEnvioDPS1 = Nothing
    Set m_objColPendienteRecordatorioEnvioDPS2 = Nothing
    Set m_objColPendienteEnvioDPS = Nothing
    Set m_objColPendienteRegistroONS = Nothing
    Set m_objColPendienteRegistroEnHPS = Nothing
    Set m_objColPendienteEnvioTraspasoONS = Nothing
    
    ' 3. CÁLCULO DE TOTALES GLOBALES (Al acceder a las Propiedades Públicas, se dispara la carga)
    ' Usamos las propiedades Me.Propiedad, no las variables m_obj...
    
    If Not Me.ColPendientesEnvioExcel Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendientesEnvioExcel.Count
    If Not Me.ColPendientesRecordatorioExcel1 Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendientesRecordatorioExcel1.Count
    If Not Me.ColPendientesRecordatorioExcel2 Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendientesRecordatorioExcel2.Count
    If Not Me.ColPendienteRegistroMarga Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendienteRegistroMarga.Count
    If Not Me.ColPendienteRecordatorioEnvioDPS1 Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendienteRecordatorioEnvioDPS1.Count
    If Not Me.ColPendienteRecordatorioEnvioDPS2 Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendienteRecordatorioEnvioDPS2.Count
    If Not Me.ColPendienteEnvioDPS Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendienteEnvioDPS.Count
    If Not Me.ColPendienteRegistroONS Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendienteRegistroONS.Count
    If Not Me.ColPendienteRegistroEnHPS Is Nothing Then m_lTotalPendientesGlobal = m_lTotalPendientesGlobal + Me.ColPendienteRegistroEnHPS.Count
    
    ' Si es técnico, paramos aquí (solo ve globales)
    If p_EsTecnico = EnumSiNo.Sí Then Exit Sub

    ' 4. CÁLCULO DE TOTALES DEL USUARIO (Constructor queries directas)
    Dim m_NombreUsuario As String
    If Not p_Usuario Is Nothing Then
        m_NombreUsuario = p_Usuario.Nombre
    Else
        m_NombreUsuario = ""
    End If

    If m_NombreUsuario <> "" Then
        ' Llamamos al constructor pasando el nombre del tramitador
        Set m_Col = constructor.getSolicitudesPendientesEnvioExcel(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioExcel1(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioExcel2(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesAltaMarga(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioEnvioDPS1(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioEnvioDPS2(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesEnvioDPS(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesRegistroONS(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
        
        Set m_Col = constructor.getSolicitudesPendientesRegistroEnHPS(p_Tramitador:=m_NombreUsuario, p_Error:=p_Error)
        If Not m_Col Is Nothing Then m_lTotalPendientesUsuario = m_lTotalPendientesUsuario + m_Col.Count
    End If
    
    Exit Sub
    
errores:
    Debug.Print "Error en CalcularTotalesPendientes: " & Err.Description
End Sub

Public Property Get ColItems() As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColItems Is Nothing Then
        Set ColItems = m_objColItems
        Exit Property
    End If
    Set m_objColItems = New Scripting.Dictionary
    With m_objColItems
        .Add "ColAdministradores", "o"
        .Add "Coltramitadores", "o"
        .Add "URLAchivoIni", ""
        .Add "VersionAplicacion", ""
        .Add "URLDirectorioAplicacion", ""
        .Add "URLDirectorioDocumentacion", ""
        .Add "URLDirectorioRecursos", ""
        .Add "URLDirectorioLocal", ""
        .Add "URLDirectorioDocumentacionAyuda", ""
        .Add "Configuracion", "o"
        .Add "DiasParaRecordatorioExcel1", ""
        .Add "DiasParaRecordatorioExcel2", ""
        .Add "DiasCancelacionPreMARGA", ""
        .Add "DiasParaRecordatorioRellenoMarga1", ""
        .Add "DiasParaRecordatorioRellenoMarga2", ""
        .Add "DiasCancelacionMARGA", ""
        .Add "BuzonSeguridad", ""
        .Add "FaltaAlgunDatoConfiguracion", ""
        .Add "ColEstadosTexto", "o"
        .Add "ColTextoEstados", "o"
        .Add "ColEstadosParaTitulo", "o"
        .Add "colExplicacionTipoSeleccionTarea", "o"
        .Add "URLDirectorioPlantillas", ""
        .Add "URLDirectorioPlantillasHTML", ""
        .Add "URLDirectorioPlantillasExcel", ""
        .Add "URLPlantillaAlta", ""
        .Add "URLPlantillaTraspaso", ""
        .Add "ColTiposTexto", "o"
        .Add "ColTextoTipos", "o"
        .Add "URLDirectorioAplicaciones", ""
        .Add "AutocancelacionPreMARGA", ""
        .Add "AutocancelacionMARGA", ""
        .Add "CharsetISO", ""
        .Add "CharsetUTF", ""
        .Add "CorreosAutomaticos", ""
        .Add "URLPlantillasHTMLExcelSolicitud", ""
        .Add "URLHTMLRecordatorioExcel1", ""
        .Add "URLHTMLRecordatorioExcel2", ""
        .Add "URLHTMLRecordatorioMarga1", ""
        .Add "URLHTMLRecordatorioMarga2", ""
        .Add "URLHTMLAutocancelacionPREMARGA", ""
        .Add "URLHTMLAutocancelacionMARGA", ""
        
        .Add "ColPlantillas", "o"
        
        .Add "ConCorreoCopiaGestor", ""
        .Add "CSS", ""
        .Add "URLArchivoCSS", ""
        .Add "Entidades", "o"
        .Add "EntidadesFormularios", "o"
        .Add "ColNombrePlantillas", "o"
        .Add "ColTiposCorreos", "o"
    End With
    Set ColItems = m_objColItems
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColItems ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Function getPropiedad(p_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    
    On Error GoTo errores
    Me.Error = ""
    
    If p_NombrePropiedad = "ColAdministradores" Then
        Set getPropiedad = Me.ColAdministradores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "Coltramitadores" Then
        Set getPropiedad = Me.Coltramitadores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLAchivoIni" Then
        getPropiedad = Me.URLAchivoIni
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "VersionAplicacion" Then
        getPropiedad = Me.VersionAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioAplicacion" Then
        getPropiedad = Me.URLDirectorioAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLDirectorioDocumentacion" Then
        getPropiedad = Me.URLDirectorioDocumentacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioRecursos" Then
        getPropiedad = Me.URLDirectorioRecursos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioLocal" Then
        getPropiedad = Me.URLDirectorioLocal
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioDocumentacionAyuda" Then
        getPropiedad = Me.URLDirectorioDocumentacionAyuda
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColAdministradores" Then
        Set getPropiedad = Me.ColAdministradores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "Configuracion" Then
        Set getPropiedad = Me.Configuracion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "DiasParaRecordatorioExcel1" Then
        getPropiedad = Me.DiasParaRecordatorioExcel1
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "DiasParaRecordatorioExcel2" Then
        getPropiedad = Me.DiasParaRecordatorioExcel2
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "DiasCancelacionPreMARGA" Then
        getPropiedad = Me.DiasCancelacionPreMARGA
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "DiasParaRecordatorioRellenoMarga1" Then
        getPropiedad = Me.DiasParaRecordatorioRellenoMarga1
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "DiasParaRecordatorioRellenoMarga2" Then
        getPropiedad = Me.DiasParaRecordatorioRellenoMarga2
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "DiasCancelacionMARGA" Then
        getPropiedad = Me.DiasCancelacionMARGA
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "BuzonSeguridad" Then
        getPropiedad = Me.BuzonSeguridad
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CorreosAutomaticos" Then
        getPropiedad = Me.CorreosAutomaticos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "FaltaAlgunDatoConfiguracion" Then
        getPropiedad = Me.FaltaAlgunDatoConfiguracion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColEstadosTexto" Then
        Set getPropiedad = Me.ColEstadosTexto
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColTextoEstados" Then
        Set getPropiedad = Me.ColTextoEstados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColEstadosParaTitulo" Then
        Set getPropiedad = Me.ColEstadosParaTitulo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombrePropiedad = "colExplicacionTipoSeleccionTarea" Then
        Set getPropiedad = Me.colExplicacionTipoSeleccionTarea
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColTextoEstados" Then
        Set getPropiedad = Me.ColTextoEstados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioPlantillas" Then
        getPropiedad = Me.URLDirectorioPlantillas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioPlantillasHTML" Then
        getPropiedad = Me.URLDirectorioPlantillasHTML
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioPlantillasExcel" Then
        getPropiedad = Me.URLDirectorioPlantillasExcel
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLPlantillaAlta" Then
        getPropiedad = Me.URLPlantillaAlta
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLPlantillaTraspaso" Then
        getPropiedad = Me.URLPlantillaTraspaso
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColTiposTexto" Then
        Set getPropiedad = Me.ColTiposTexto
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColTextoTipos" Then
        Set getPropiedad = Me.ColTextoTipos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "Coltramitadores" Then
        Set getPropiedad = Me.Coltramitadores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLDirectorioAplicaciones" Then
        getPropiedad = Me.URLDirectorioAplicaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLPlantillasHTMLExcelSolicitud" Then
        getPropiedad = Me.URLPlantillasHTMLExcelSolicitud
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "AutocancelacionPreMARGA" Then
        getPropiedad = Me.AutocancelacionPreMARGA
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "AutocancelacionMARGA" Then
        getPropiedad = Me.AutocancelacionMARGA
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CharsetISO" Then
        getPropiedad = Me.CharsetISO
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CharsetUTF" Then
        getPropiedad = Me.CharsetUTF
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLPlantillasHTMLExcelSolicitud" Then
        getPropiedad = Me.URLPlantillasHTMLExcelSolicitud
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLHTMLRecordatorioExcel1" Then
        getPropiedad = Me.URLHTMLRecordatorioExcel1
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLHTMLRecordatorioExcel2" Then
        getPropiedad = Me.URLHTMLRecordatorioExcel2
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLHTMLRecordatorioMarga1" Then
        getPropiedad = Me.URLHTMLRecordatorioMarga1
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLHTMLRecordatorioMarga2" Then
        getPropiedad = Me.URLHTMLRecordatorioMarga2
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLHTMLAutocancelacionPREMARGA" Then
        getPropiedad = Me.URLHTMLAutocancelacionPREMARGA
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLHTMLAutocancelacionMARGA" Then
        getPropiedad = Me.URLHTMLAutocancelacionMARGA
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "ConCorreoCopiaGestor" Then
        getPropiedad = Me.ConCoreroCopiaGestor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CSS" Then
        getPropiedad = Me.CSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLArchivoCSS" Then
        getPropiedad = Me.URLArchivoCSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "Entidades" Then
        Set getPropiedad = Me.Entidades
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "EntidadesFormularios" Then
        Set getPropiedad = Me.EntidadesFormularios
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColNombrePlantillas" Then
        Set getPropiedad = Me.ColNombrePlantillas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColPlantillas" Then
        Set getPropiedad = Me.ColPlantillas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLArchivoCSS" Then
        getPropiedad = Me.URLArchivoCSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColTiposCorreos" Then
        Set getPropiedad = Me.ColTiposCorreos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    p_Error = "La propiedad " & p_NombrePropiedad & " no está codificada en Entorno.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Entorno.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

