VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDSolicitud As String
Public Gestor As String
Public DNI As String
Public Nombre As String
Public Apellido1 As String
Public Apellido2 As String
Public FNacimiento As String
Public LugarNacimiento As String
Public email As String
Public Telefono As String
Public IDExpediente As String
Public IDEmpresaUsuario As String
Public IDEmpresaTramitadora As String
Public URLAdjunto As String
Public emailResponsable As String
Public Estado As String
Public Observaciones As String
Public FechaCreacion As String
Public FechaUltimoCambio As String
Public UsuarioCreacion As String
Public UsuarioUltimoCambio As String
Public TIPO As String
Public CorreosAutomaticos As String

Public Motivo_HPS As String
Public IDJustificacion As String
Public IDUsuarioHPS As String
Public URLAdjuntoEnvioONS As String
Public Nemotecnico As String


Public HTMLParaCorreo As EnumSiNo

Private m_ObjGestorObj As Usuario
Private m_objUsuarioHPS As UsuarioHPS
Private m_sIDSolicitudCalculado As String
Private m_sURLAdjuntoCalculado As String
Private m_sURLDirectorioAnexo As String
Private m_ObjSolicitudFechas As SolicitudFechas
Private m_eTipoCalculado As EnumTipo


Private m_eEstadoCalculado As EnumTipo
Private m_eEstadoEnum As EnumEstado
Private m_sEstadoCalculadoTexto As String
Private m_sEstadoCalculadoTitulo As String
Private m_sEstadoTexto As String
Private m_sEstadoTitulo As String
Private m_sURLPlantillaExcel As String


Private m_ObjExpediente As Expediente
Private m_objEmpresaUsuario As Suministrador

Private m_ObjEmpresaTramitadora As Suministrador

Private m_sNombreCompleto As String
Private m_sNombreCompletoCapitalizado As String

Private m_ObjColParametrosParaHTML As Scripting.Dictionary
Private m_sColParametrosParaHTMLTexto As String
Private m_sDiasParaCancelacionPreMarga As String
Private m_sDiasParaCancelacionMarga As String

Private m_sDiasDesdeEnvioExcel As String
Private m_sDiasDesdeAltaEnMarga As String
Private m_ObjTramitador As Usuario
Private m_sHTMLEnvioExcel As String
Private m_sAsuntoCoreoEnvioExcel As String
Private m_sAsuntoRecordatorioExcel1 As String
Private m_sAsuntoRecordatorioExcel2 As String
Private m_sAsuntoCancelaconPREMARGA As String
Private m_sAsuntoRecordatorioMarga1 As String
Private m_sAsuntoRecordatorioMarga2 As String
Private m_sAsuntoCancelaconMARGA As String
Private m_ObjCorreos As Scripting.Dictionary
Private m_ObjUsuarioUltimoCambioObj As Usuario
Private m_ObjUsuarioCreacionObj As Usuario
Private m_sUltimaModificacionTexto As String
Private m_sURLAdjuntoEnvioONSCalculado As String



Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get UsuarioHPS() As UsuarioHPS
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioHPS Is Nothing Then
        Set UsuarioHPS = m_objUsuarioHPS
        Exit Property
    End If
    If Me.IDUsuarioHPS = "" And Me.DNI = "" Then
        Exit Property
    End If
    If Me.IDUsuarioHPS <> "" Then
        Set m_objUsuarioHPS = constructor.getUsuarioHPS(p_IDUsuario:=Me.IDUsuarioHPS, p_Error:=m_Error)
    Else
        Set m_objUsuarioHPS = constructor.getUsuarioHPS(p_DNI:=Me.DNI, p_Error:=m_Error)
    End If
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioHPS = m_objUsuarioHPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.UsuarioHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EmpresaUsuario() As Suministrador
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objEmpresaUsuario Is Nothing Then
        Set EmpresaUsuario = m_objEmpresaUsuario
        Exit Property
    End If
    If Me.IDEmpresaUsuario = "" Then
        Exit Property
    End If
    Set m_objEmpresaUsuario = constructor.getSuministrador(p_ID:=Me.IDEmpresaUsuario, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EmpresaUsuario = m_objEmpresaUsuario
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EmpresaUsuario ha devuelto el error: " & vbNewLine & Err.Description
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
    Set m_ObjExpediente = constructor.getExpediente(p_IDExpediente:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Expediente = m_ObjExpediente
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.Expediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get UsuarioUltimoCambioObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioUltimoCambioObj Is Nothing Then
        Set UsuarioUltimoCambioObj = m_ObjUsuarioUltimoCambioObj
        Exit Property
    End If
    If Me.UsuarioUltimoCambio = "" Then
        Exit Property
    End If
    Set m_ObjUsuarioUltimoCambioObj = constructor.getUsuario(p_UsuarioRed:=Me.UsuarioUltimoCambio, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioUltimoCambioObj = m_ObjUsuarioUltimoCambioObj
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.UsuarioUltimoCambioObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioCreacionObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioCreacionObj Is Nothing Then
        Set UsuarioCreacionObj = m_ObjUsuarioCreacionObj
        Exit Property
    End If
    If Me.UsuarioCreacion = "" Then
        Exit Property
    End If
    Set m_ObjUsuarioCreacionObj = constructor.getUsuario(p_UsuarioRed:=Me.UsuarioCreacion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioCreacionObj = m_ObjUsuarioCreacionObj
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.UsuarioCreacionObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get UltimaModificacionTexto() As String
    
    Dim m_NombreUsuario As String
    Dim m_Fecha As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    With Me
        If Not .UsuarioUltimoCambioObj Is Nothing Then
            m_NombreUsuario = .UsuarioUltimoCambioObj.Nombre
            If IsDate(.FechaUltimoCambio) Then
                m_Fecha = Format(.FechaUltimoCambio, "dd/mm/yyyy")
            Else
                m_Fecha = "Fecha desconocida"
            End If
        Else
            If Not .UsuarioCreacionObj Is Nothing Then
                m_NombreUsuario = .UsuarioCreacionObj.Nombre
            Else
                m_NombreUsuario = "Desconocido"
            End If
            If IsDate(.FechaCreacion) Then
                m_Fecha = Format(.FechaCreacion, "dd/mm/yyyy")
            Else
                m_Fecha = "Fecha desconocida"
            End If
        End If
       
    End With
    m_sUltimaModificacionTexto = "Última modificación por " & m_NombreUsuario & " el " & m_Fecha
    
    UltimaModificacionTexto = m_sUltimaModificacionTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.UltimaModificacionTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Correos() As Scripting.Dictionary
    
    Dim m_Parametro As Variant
    Dim m_Valor As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjCorreos Is Nothing Then
        Set Correos = m_ObjCorreos
        Exit Property
    End If
    If Me.IDSolicitud = "" Then
        Exit Property
    End If
    Set m_ObjCorreos = constructor.getCorreos(p_IDSolicitud:=Me.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Correos = m_ObjCorreos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.Correos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColParametrosParaHTMLTexto() As String
    
    Dim parser As ParametrosParser
    Set parser = New ParametrosParser
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.ColParametrosParaHTML Is Nothing Then
        ColParametrosParaHTMLTexto = ""
        Exit Property
    End If
    
    ' Usar ParametrosParser para generar la cadena
    ColParametrosParaHTMLTexto = parser.GenerarCadenaRecursos(Me.ColParametrosParaHTML)
    
    ' Verificar si hubo errores
    If parser.Error <> "" Then
        m_Error = parser.Error
        Err.Raise 1000
    End If
    
    ' Guardar en la variable privada para compatibilidad
    m_sColParametrosParaHTMLTexto = ColParametrosParaHTMLTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.ColParametrosParaHTMLTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    ColParametrosParaHTMLTexto = ""
End Property
Public Property Get ColParametrosParaHTML() As Scripting.Dictionary
    
    On Error GoTo errores
  
    Me.Error = ""
    m_Error = ""
    
    Set m_ObjColParametrosParaHTML = New Scripting.Dictionary
    m_ObjColParametrosParaHTML.CompareMode = TextCompare
    With m_ObjColParametrosParaHTML
        .Add "||BuzonSeguridad||", m_ObjEntorno.BuzonSeguridad
        
        If m_ObjEntorno.ConCoreroCopiaGestor = EnumSiNo.Sí Then
            If Not Me.GestorObj Is Nothing Then
                .Add "||mailto||", Me.GestorObj.CorreoUsuario & ";" & m_ObjEntorno.BuzonSeguridad
            Else
                .Add "||mailto||", m_ObjEntorno.BuzonSeguridad
            End If
        Else
            .Add "||mailto||", m_ObjEntorno.BuzonSeguridad
        End If
        
        .Add "||Gestor||", IIf(Not Me.GestorObj Is Nothing, Me.GestorObj.Nombre, "")

        ' --- NUEVO: AÑADIDO EMAIL RESPONSABLE ---
        .Add "||emailResponsable||", Nz(Me.emailResponsable, "")
        ' ----------------------------------------
        
        .Add "||DiasParaRecordatorioExcel1||", m_ObjEntorno.DiasParaRecordatorioExcel1
        .Add "||DiasParaRecordatorioExcel2||", m_ObjEntorno.DiasParaRecordatorioExcel2
        
        If Me.TipoCalculado = EnumTipo.Alta Or Me.TipoCalculado = EnumTipo.Renovacion Then
            .Add "||NombreBaseArchivoTipo||", fso.GetBaseName(m_ObjEntorno.URLPlantillaAlta)
        ElseIf Me.TipoCalculado = EnumTipo.Traspaso Then
            .Add "||NombreBaseArchivoTipo||", fso.GetBaseName(m_ObjEntorno.URLPlantillaTraspaso)
        End If
        
        .Add "||DiasParaCancelacionPreMARGA||", Me.DiasParaCancelacionPreMarga
        .Add "||DiasParaCancelacionMARGA||", Me.DiasParaCancelacionMarga
        .Add "||DiasDesdeEnvioExcel||", Me.DiasDesdeEnvioExcel
        .Add "||DiasDesdeAltaEnMarga||", Me.DiasDesdeAltaEnMarga
    End With
        
    Set ColParametrosParaHTML = m_ObjColParametrosParaHTML
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.ColParametrosParaHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get GestorObj() As Usuario
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjGestorObj Is Nothing Then
        Set GestorObj = m_ObjGestorObj
        Exit Property
    End If
    If Me.Gestor = "" Then
        Exit Property
    End If
    
    Set m_ObjGestorObj = constructor.getUsuario(p_Nombre:=Me.Gestor, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
        
    Set GestorObj = m_ObjGestorObj
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.GestorObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasParaCancelacionPreMarga() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sDiasParaCancelacionPreMarga <> "" Then
        DiasParaCancelacionPreMarga = m_sDiasParaCancelacionPreMarga
        Exit Property
    End If
    If Me.SolicitudFechas Is Nothing Then
        m_sDiasParaCancelacionPreMarga = m_ObjEntorno.DiasCancelacionPreMARGA
        Exit Property
    End If
    m_sDiasParaCancelacionPreMarga = Me.SolicitudFechas.DiasParaCancelacionPreMarga
    m_Error = Me.SolicitudFechas.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DiasParaCancelacionPreMarga = m_sDiasParaCancelacionPreMarga
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.DiasParaCancelacionPreMarga ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get DiasParaCancelacionMarga() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sDiasParaCancelacionMarga <> "" Then
        DiasParaCancelacionMarga = m_sDiasParaCancelacionMarga
        Exit Property
    End If
    If Me.SolicitudFechas Is Nothing Then
        m_sDiasParaCancelacionMarga = m_ObjEntorno.DiasCancelacionMARGA
        Exit Property
    End If
    m_sDiasParaCancelacionMarga = Me.SolicitudFechas.DiasParaCancelacionMarga
    m_Error = Me.SolicitudFechas.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DiasParaCancelacionMarga = m_sDiasParaCancelacionMarga
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.DiasParaCancelacionMarga ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasDesdeEnvioExcel() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sDiasDesdeEnvioExcel <> "" Then
        DiasDesdeEnvioExcel = m_sDiasDesdeEnvioExcel
        Exit Property
    End If
    If Me.SolicitudFechas Is Nothing Then
        m_sDiasDesdeEnvioExcel = m_ObjEntorno.DiasCancelacionMARGA
        Exit Property
    End If
    m_sDiasDesdeEnvioExcel = Me.SolicitudFechas.DiasDesdeEnvioExcel
    m_Error = Me.SolicitudFechas.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DiasDesdeEnvioExcel = m_sDiasDesdeEnvioExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.DiasDesdeEnvioExcel ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasDesdeAltaEnMarga() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sDiasDesdeAltaEnMarga <> "" Then
        DiasDesdeAltaEnMarga = m_sDiasDesdeAltaEnMarga
        Exit Property
    End If
    If Me.SolicitudFechas Is Nothing Then
        m_sDiasDesdeAltaEnMarga = m_ObjEntorno.DiasCancelacionMARGA
        Exit Property
    End If
    m_sDiasDesdeAltaEnMarga = Me.SolicitudFechas.DiasDesdeAltaEnMarga
    m_Error = Me.SolicitudFechas.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DiasDesdeAltaEnMarga = m_sDiasDesdeAltaEnMarga
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.DiasDesdeAltaEnMarga ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get AsuntoCoreoEnvioExcel() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sAsuntoCoreoEnvioExcel <> "" Then
        AsuntoCoreoEnvioExcel = m_sAsuntoCoreoEnvioExcel
        Exit Property
    End If
    If Me.TipoCalculado = Empty Then
        m_Error = "No se sabe el tipo de Solicitud"
        Err.Raise 1000
    End If
    
    m_sAsuntoCoreoEnvioExcel = "SOLICITUD HPS – " & Format(CDate(Date), "dd") & " – " & _
                        Format(CDate(Date), "mm") & " – " & Format(CDate(Date), "yyyy") & " – " & _
                        Me.NombreCompleto
    
    
    AsuntoCoreoEnvioExcel = m_sAsuntoCoreoEnvioExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.AsuntoCoreoEnvioExcel ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AsuntoRecordatorioExcel1() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sAsuntoRecordatorioExcel1 <> "" Then
        AsuntoRecordatorioExcel1 = m_sAsuntoRecordatorioExcel1
        Exit Property
    End If
    If Me.TipoCalculado = Empty Then
        m_Error = "No se sabe el tipo de Solicitud"
        Err.Raise 1000
    End If
    
    m_sAsuntoRecordatorioExcel1 = "1º RECORDATORIO SOLICITUD HPS – " & Format(CDate(Date), "dd") & " – " & _
                        Format(CDate(Date), "mm") & " – " & Format(CDate(Date), "yyyy") & " – " & _
                        Me.NombreCompleto
    
    
    AsuntoRecordatorioExcel1 = m_sAsuntoRecordatorioExcel1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.AsuntoRecordatorioExcel1 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AsuntoRecordatorioExcel2() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sAsuntoRecordatorioExcel2 <> "" Then
        AsuntoRecordatorioExcel2 = m_sAsuntoRecordatorioExcel2
        Exit Property
    End If
    If Me.TipoCalculado = Empty Then
        m_Error = "No se sabe el tipo de Solicitud"
        Err.Raise 1000
    End If
    
    m_sAsuntoRecordatorioExcel2 = "2º RECORDATORIO SOLICITUD HPS – " & Format(CDate(Date), "dd") & " – " & _
                        Format(CDate(Date), "mm") & " – " & Format(CDate(Date), "yyyy") & " – " & _
                        Me.NombreCompleto
    
    
    AsuntoRecordatorioExcel2 = m_sAsuntoRecordatorioExcel2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.AsuntoRecordatorioExcel2 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AsuntoCancelaconPREMARGA() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sAsuntoCancelaconPREMARGA <> "" Then
        AsuntoCancelaconPREMARGA = m_sAsuntoCancelaconPREMARGA
        Exit Property
    End If
    
    
    m_sAsuntoCancelaconPREMARGA = "AUTOCANCELACIÓN DE SOLICITUD HPS – " & Format(CDate(Date), "dd") & " – " & _
                        Format(CDate(Date), "mm") & " – " & Format(CDate(Date), "yyyy") & " – " & _
                        Me.NombreCompleto
    
    
    AsuntoCancelaconPREMARGA = m_sAsuntoCancelaconPREMARGA
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.AsuntoCancelaconPREMARGA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get AsuntoCancelaconMARGA() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sAsuntoCancelaconMARGA <> "" Then
        AsuntoCancelaconMARGA = m_sAsuntoCancelaconMARGA
        Exit Property
    End If
    
    
    m_sAsuntoCancelaconMARGA = "AUTOCANCELACIÓN DE SOLICITUD HPS – " & Format(CDate(Date), "dd") & " – " & _
                        Format(CDate(Date), "mm") & " – " & Format(CDate(Date), "yyyy") & " – " & _
                        Me.NombreCompleto
    
    
    AsuntoCancelaconMARGA = m_sAsuntoCancelaconMARGA
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.AsuntoCancelaconMARGA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get AsuntoRecordatorioMarga1() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sAsuntoRecordatorioMarga1 <> "" Then
        AsuntoRecordatorioMarga1 = m_sAsuntoRecordatorioMarga1
        Exit Property
    End If
    If Me.TipoCalculado = Empty Then
        m_Error = "No se sabe el tipo de Solicitud"
        Err.Raise 1000
    End If
    
    m_sAsuntoRecordatorioMarga1 = "1º RECORDATORIO: Cumplimentación DPS en plataforma MARGA " & _
                            "(Gobernanza de Ciberseguridad Nacional) – " & Format(CDate(Date), "dd") & " – " & _
                            Format(CDate(Date), "mm") & " – " & Format(CDate(Date), "yyyy") & " – " & _
                            Me.NombreCompleto
    
    
    AsuntoRecordatorioMarga1 = m_sAsuntoRecordatorioMarga1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.AsuntoRecordatorioMarga1 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get AsuntoRecordatorioMarga2() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sAsuntoRecordatorioMarga2 <> "" Then
        AsuntoRecordatorioMarga2 = m_sAsuntoRecordatorioMarga2
        Exit Property
    End If
    If Me.TipoCalculado = Empty Then
        m_Error = "No se sabe el tipo de Solicitud"
        Err.Raise 1000
    End If
    
    m_sAsuntoRecordatorioMarga2 = "2º RECORDATORIO: Cumplimentación DPS en plataforma MARGA " & _
                            "(Gobernanza de Ciberseguridad Nacional) – " & Format(CDate(Date), "dd") & " – " & _
                            Format(CDate(Date), "mm") & " – " & Format(CDate(Date), "yyyy") & " – " & _
                            Me.NombreCompleto
    
    
    AsuntoRecordatorioMarga2 = m_sAsuntoRecordatorioMarga2
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.AsuntoRecordatorioMarga2 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get HTMLEnvioExcel() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sHTMLEnvioExcel <> "" Then
        HTMLEnvioExcel = m_sHTMLEnvioExcel
        Exit Property
    End If
    If Me.TipoCalculado = Empty Then
        m_Error = "No se sabe el tipo de Solicitud"
        Err.Raise 1000
    End If
    m_sHTMLEnvioExcel = getTextoDeHTML(m_ObjEntorno.URLPlantillasHTMLExcelSolicitud, Me.ColParametrosParaHTML, , Me.HTMLParaCorreo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLEnvioExcel = m_sHTMLEnvioExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.HTMLEnvioExcel ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Tramitador() As Usuario
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjTramitador Is Nothing Then
        Set Tramitador = m_ObjTramitador
        Exit Property
    End If
    If Me.Gestor = "" Then
        Exit Property
    End If
    
    Set m_ObjTramitador = constructor.getUsuario(p_UsuarioRed:=Me.Gestor, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Tramitador = m_ObjTramitador
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.Tramitador ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get NombreCompleto() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
'    If m_sNombreCompleto <> "" Then
'        NombreCompleto = m_sNombreCompleto
'        Exit Property
'    End If
    
    
    If Me.Apellido2 <> "" Then
        m_sNombreCompleto = Me.Nombre & " " & Me.Apellido1 & " " & Me.Apellido2
    Else
        m_sNombreCompleto = Me.Nombre & " " & Me.Apellido1
    End If
    
    
    NombreCompleto = m_sNombreCompleto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.NombreCompleto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get NombreCompletoCapitalizado() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreCompletoCapitalizado <> "" Then
        NombreCompletoCapitalizado = m_sNombreCompletoCapitalizado
        Exit Property
    End If
    
    m_sNombreCompletoCapitalizado = FormatearNombreCompleto(Me.NombreCompleto)
    
    
    NombreCompletoCapitalizado = m_sNombreCompletoCapitalizado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.NombreCompletoCapitalizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get EmpresaTramitadora() As Suministrador
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEmpresaTramitadora Is Nothing Then
        Set EmpresaTramitadora = m_ObjEmpresaTramitadora
        Exit Property
    End If
    If Me.IDEmpresaTramitadora = "" Then
        Exit Property
    End If
    
    Set m_ObjEmpresaTramitadora = constructor.getSuministrador(p_ID:=Me.IDEmpresaTramitadora, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EmpresaTramitadora = m_ObjEmpresaTramitadora
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EmpresaTramitadora ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLPlantillaExcel() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLPlantillaExcel <> "" Then
        URLPlantillaExcel = m_sURLPlantillaExcel
        Exit Property
    End If
    If Me.TIPO = "" Then
        Exit Property
    End If
    If Me.TipoCalculado = EnumTipo.Alta Then
        m_sURLPlantillaExcel = m_ObjEntorno.URLPlantillaAlta
    ElseIf Me.TipoCalculado = EnumTipo.Renovacion Then
        m_sURLPlantillaExcel = m_ObjEntorno.URLPlantillaAlta
    ElseIf Me.TipoCalculado = EnumTipo.Traspaso Then
        m_sURLPlantillaExcel = m_ObjEntorno.URLPlantillaTraspaso
    End If
    
    URLPlantillaExcel = m_sURLPlantillaExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.URLPlantillaExcel ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TipoCalculado() As EnumTipo
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eTipoCalculado <> Empty Then
        TipoCalculado = m_eTipoCalculado
        Exit Property
    End If
    
    m_eTipoCalculado = m_ObjEntorno.ColTextoTipos(Me.TIPO)
   
    
    TipoCalculado = m_eTipoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.TipoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLAdjuntoCalculado() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.URLAdjunto = "" Then
        Exit Property
    End If
    If Me.IDSolicitud = "" Then
        Exit Property
    End If
    m_sURLAdjuntoCalculado = getURLFinal(Me.URLAdjunto, Format(Me.IDSolicitud, "000000"), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    URLAdjuntoCalculado = m_sURLAdjuntoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.URLAdjuntoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get URLDirectorioAnexo() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    If Me.IDSolicitud = "" Then
        Exit Property
    End If
    m_sURLDirectorioAnexo = m_ObjEntorno.URLDirectorioDocumentacion & "Solicitudes\" & Format(Me.IDSolicitud, "000000") & "\"
    
    URLDirectorioAnexo = m_sURLDirectorioAnexo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.URLDirectorioAnexo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EstadoEnum() As EnumEstado
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    m_eEstadoEnum = m_ObjEntorno.ColTextoEstados(Me.Estado)
   
    
    EstadoEnum = m_eEstadoEnum
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EstadoEnum ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get EstadoCalculado() As EnumEstado
    Dim i As Long
    Dim m_SolicitudFechas As SolicitudFechas
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_SolicitudFechas = Me.SolicitudFechas
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_SolicitudFechas Is Nothing Then
        m_eEstadoCalculado = EnumEstado.PendienteEnvioExcel
        EstadoCalculado = m_eEstadoCalculado
        Exit Property
    End If
recalcular:
    m_eEstadoCalculado = m_SolicitudFechas.EstadoCalculado
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_eEstadoCalculado = Empty Then
        i = i + 1
        If i < 3 Then
            GoTo recalcular
        End If
        Stop
    End If
    EstadoCalculado = m_eEstadoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EstadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get EstadoCalculadoTexto() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""

    Me.Estado = Me.EstadoCalculado
    
    m_sEstadoCalculadoTexto = m_ObjEntorno.ColEstadosTexto(CStr(Me.Estado))
    
    EstadoCalculadoTexto = m_sEstadoCalculadoTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EstadoCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get EstadoTexto() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    m_sEstadoTexto = m_ObjEntorno.ColEstadosTexto(CStr(Me.Estado))
    
    EstadoTexto = m_sEstadoTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EstadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get EstadoCalculadoTitulo() As String
    
    Dim m_EstadoCalculado As EnumEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_EstadoCalculado = Me.EstadoCalculado
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sEstadoCalculadoTitulo = m_ObjEntorno.ColEstadosParaTitulo(CStr(m_EstadoCalculado))
    
    EstadoCalculadoTitulo = m_sEstadoCalculadoTitulo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EstadoCalculadoTitulo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get EstadoTitulo() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    m_sEstadoTitulo = m_ObjEntorno.ColEstadosParaTitulo(CStr(Me.EstadoEnum))
    
    EstadoTitulo = m_sEstadoTitulo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.EstadoTitulo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get SolicitudFechas() As SolicitudFechas
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjSolicitudFechas Is Nothing Then
        Set SolicitudFechas = m_ObjSolicitudFechas
        Exit Property
    End If
    If Me.IDSolicitud = "" Then
        Exit Property
    End If
    Set m_ObjSolicitudFechas = constructor.getSolicitudFechas(p_IDSolicitud:=Me.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set SolicitudFechas = m_ObjSolicitudFechas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.SolicitudFechas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set SolicitudFechas(p_Valor As Variant)
    
    Set m_ObjSolicitudFechas = p_Valor
    
End Property
Public Property Get IDSolicitudCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDSolicitudCalculado = DameID("TbSolicitudes", "IDSolicitud", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDSolicitudCalculado = m_sIDSolicitudCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.IDSolicitudCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get URLAdjuntoEnvioONSCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLAdjuntoEnvioONSCalculado <> "" Then
        URLAdjuntoEnvioONSCalculado = m_sURLAdjuntoEnvioONSCalculado
        Exit Property
    End If
    If Me.URLAdjuntoEnvioONS = "" Then
        Exit Property
    End If
    m_sURLAdjuntoEnvioONSCalculado = getURLFinal(Me.URLAdjuntoEnvioONS, Format(Me.IDSolicitud, "000000"), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    URLAdjuntoEnvioONSCalculado = m_sURLAdjuntoEnvioONSCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.URLAdjuntoEnvioONSCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Function EstadoGrabar( _
                            Optional p_EstadoTexto As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim m_SQL As String
     Dim i As Long
    On Error GoTo errores
    If p_EstadoTexto = "" Then
recalcular:
        Me.Estado = Me.EstadoCalculadoTexto
        
    Else
        Me.Estado = p_EstadoTexto
    End If
    If Me.Estado = "" Then
       i = i + 1
       If i < 3 Then
            GoTo recalcular
       End If
        Exit Function
    End If
    m_SQL = "UPDATE TbSolicitudes SET Estado = '" & Me.Estado & "' " & _
            "WHERE IDSolicitud=" & Me.IDSolicitud & ";"
    getdb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.EstadoGrabar ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function AsuntoCalculado( _
                                    p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    On Error GoTo errores
    
  
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.EnvioExcel Then
       AsuntoCalculado = Me.AsuntoCoreoEnvioExcel
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel1 Then
       AsuntoCalculado = Me.AsuntoRecordatorioExcel1
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel2 Then
       AsuntoCalculado = Me.AsuntoRecordatorioExcel2
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionPreMarga Then
       AsuntoCalculado = Me.AsuntoCancelaconPREMARGA
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga1 Then
       AsuntoCalculado = Me.AsuntoRecordatorioMarga1
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga2 Then
       AsuntoCalculado = Me.AsuntoRecordatorioMarga2
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionMarga Then
       AsuntoCalculado = Me.AsuntoCancelaconMARGA
       Exit Function
    End If
    p_Error = "No se ha podido saber el tipo de correo para obtener el asunto"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.AsuntoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function DestinatariosConCopiaCalculados( _
                                                p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                                Optional ByRef p_Error As String _
                                                ) As String

    On Error GoTo errores
    
  
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.EnvioExcel Then
       DestinatariosConCopiaCalculados = m_ObjEntorno.BuzonSeguridad
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel1 Then
       DestinatariosConCopiaCalculados = Me.emailResponsable & ";" & m_ObjEntorno.BuzonSeguridad
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel2 Then
       DestinatariosConCopiaCalculados = Me.emailResponsable & ";" & m_ObjEntorno.BuzonSeguridad
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionPreMarga Then
       DestinatariosConCopiaCalculados = m_ObjEntorno.BuzonSeguridad
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga1 Then
       DestinatariosConCopiaCalculados = Me.emailResponsable & ";" & m_ObjEntorno.BuzonSeguridad
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga2 Then
       DestinatariosConCopiaCalculados = Me.emailResponsable & ";" & m_ObjEntorno.BuzonSeguridad
       Exit Function
    End If
    If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionMarga Then
       DestinatariosConCopiaCalculados = m_ObjEntorno.BuzonSeguridad
       Exit Function
    End If
    p_Error = "No se ha podido saber el tipo de correo para obtener los destinatarios"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.DestinatariosConCopiaCalculados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function AccionParaCorreoCalculada( _
                                                    p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String

    On Error GoTo errores
    
    AccionParaCorreoCalculada = TituloParaFormulario(p_EnumTipoEnvioCorreo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.AccionParaCorreoCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function NombrePlantillaParaCorreoCalculada( _
                                                    p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String

    On Error GoTo errores
    
    NombrePlantillaParaCorreoCalculada = m_ObjEntorno.ColNombrePlantillas(CStr(p_EnumTipoEnvioCorreo))
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.NombrePlantillaParaCorreoCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function TituloParaFormulario( _
                                        p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    On Error GoTo errores
    
    TituloParaFormulario = m_ObjEntorno.ColTiposCorreos(CStr(p_EnumTipoEnvioCorreo))
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.TituloParaFormulario ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function DestinatariosCalculados( _
                                        Optional ByRef p_Error As String _
                                        ) As String

    On Error GoTo errores
    
    DestinatariosCalculados = Me.email
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.DestinatariosCalculados ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDSolicitud"
        .Add "Gestor"
        .Add "DNI"
        .Add "Nombre"
        .Add "Apellido1"
        .Add "Apellido2"
        .Add "FNacimiento"
        .Add "LugarNacimiento"
        .Add "email"
        .Add "Telefono"
        .Add "IDExpediente"
        .Add "IDEmpresaUsuario"
        .Add "IDEmpresaTramitadora"
        .Add "URLAdjunto"
        .Add "FechaCreacion"
        .Add "emailResponsable"
        .Add "Estado"
        .Add "Observaciones"
        .Add "FechaUltimoCambio"
        .Add "UsuarioCreacion"
        .Add "UsuarioUltimoCambio"
        .Add "Tipo"
        .Add "CorreosAutomaticos"
        .Add "Motivo_HPS"
        .Add "IDJustificacion"
        .Add "IDUsuarioHPS"
        .Add "URLAdjuntoEnvioONS"
        .Add "Nemotecnico"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Solicitud.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDSolicitud" Then
        getPropiedad = Me.IDSolicitud
        Exit Function
    End If
    If m_NombrePropiedad = "Gestor" Then
        getPropiedad = Me.Gestor
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
    If m_NombrePropiedad = "Apellido1" Then
        getPropiedad = Me.Apellido1
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido2" Then
        getPropiedad = Me.Apellido2
        Exit Function
    End If
    If m_NombrePropiedad = "FNacimiento" Then
        getPropiedad = Me.FNacimiento
        Exit Function
    End If
    If m_NombrePropiedad = "LugarNacimiento" Then
        getPropiedad = Me.LugarNacimiento
        Exit Function
    End If
    If m_NombrePropiedad = "email" Then
        getPropiedad = Me.email
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        getPropiedad = Me.Telefono
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        getPropiedad = Me.IDExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        getPropiedad = Me.IDEmpresaUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaTramitadora" Then
        getPropiedad = Me.IDEmpresaTramitadora
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjunto" Then
        getPropiedad = Me.URLAdjunto
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        getPropiedad = Me.FechaCreacion
        Exit Function
    End If
    If m_NombrePropiedad = "emailResponsable" Then
        getPropiedad = Me.emailResponsable
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        getPropiedad = Me.FechaUltimoCambio
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioCreacion" Then
        getPropiedad = Me.UsuarioCreacion
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioUltimoCambio" Then
        getPropiedad = Me.UsuarioUltimoCambio
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        getPropiedad = Me.TIPO
        Exit Function
    End If
    If m_NombrePropiedad = "CorreosAutomaticos" Then
        getPropiedad = Me.CorreosAutomaticos
        Exit Function
    End If
    
    If m_NombrePropiedad = "Motivo_HPS" Then
        getPropiedad = Me.Motivo_HPS
        Exit Function
    End If
    If m_NombrePropiedad = "IDJustificacion" Then
        getPropiedad = Me.IDJustificacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioHPS" Then
        getPropiedad = Me.IDUsuarioHPS
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjuntoEnvioONS" Then
        getPropiedad = Me.URLAdjuntoEnvioONS
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        getPropiedad = Me.Nemotecnico
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Solicitud.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDSolicitud" Then
        Me.IDSolicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Gestor" Then
        Me.Gestor = m_ValorPropiedad
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
    If m_NombrePropiedad = "Apellido1" Then
        Me.Apellido1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido2" Then
        Me.Apellido2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FNacimiento" Then
        Me.FNacimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "LugarNacimiento" Then
        Me.LugarNacimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "email" Then
        Me.email = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        Me.Telefono = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        Me.IDExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        Me.IDEmpresaUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaTramitadora" Then
        Me.IDEmpresaTramitadora = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjunto" Then
        Me.URLAdjunto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        Me.FechaCreacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "emailResponsable" Then
        Me.emailResponsable = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        Me.FechaUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioCreacion" Then
        Me.UsuarioCreacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioUltimoCambio" Then
        Me.UsuarioUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        Me.TIPO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CorreosAutomaticos" Then
        Me.CorreosAutomaticos = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "Motivo_HPS" Then
        Me.Motivo_HPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDJustificacion" Then
        Me.IDJustificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioHPS" Then
        Me.IDUsuarioHPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjuntoEnvioONS" Then
        Me.URLAdjuntoEnvioONS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        Me.Nemotecnico = m_ValorPropiedad
        Exit Function
    End If
   
    p_Error = "Propiedad no reconocida en Solicitud.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Solicitud.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

