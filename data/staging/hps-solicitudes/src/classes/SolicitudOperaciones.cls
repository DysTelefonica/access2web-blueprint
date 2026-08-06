VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolicitudOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public solicitud As solicitud

Private Function FechaValidaNoFutura(ByVal p_Fecha As Variant, ByRef p_Error As String) As Date
    If Not IsDate(p_Fecha) Then
        p_Error = "La fecha indicada no es válida."
        Err.Raise 1000
    End If
    FechaValidaNoFutura = CDate(p_Fecha)
    If FechaValidaNoFutura > Date Then
        p_Error = "No se permite una fecha posterior a hoy."
        Err.Raise 1000
    End If
End Function

Private Sub ActualizarFechaHitoPorTipo(ByVal p_Tipo As EnumTipoEnvioCorreo, ByVal p_Fecha As Date, ByRef p_Solicitud As solicitud)
    With p_Solicitud.SolicitudFechas
        Select Case p_Tipo
            Case EnumTipoEnvioCorreo.EnvioExcel:         .FechaEnvioExcel = p_Fecha
            Case EnumTipoEnvioCorreo.RecordatorioExcel1:  .FechaCorreoRecordatorioExcel1 = p_Fecha
            Case EnumTipoEnvioCorreo.RecordatorioExcel2:  .FechaCorreoRecordatorioExcel2 = p_Fecha
            Case EnumTipoEnvioCorreo.CancelacionPreMarga: .FechaCancelado = p_Fecha
            Case EnumTipoEnvioCorreo.RecordatorioMarga1:  .FechaCorreoRecordatorioRellenoMarga1 = p_Fecha
            Case EnumTipoEnvioCorreo.RecordatorioMarga2:  .FechaCorreoRecordatorioRellenoMarga2 = p_Fecha
            Case EnumTipoEnvioCorreo.CancelacionMarga:    .FechaCancelado = p_Fecha
        End Select
    End With
End Sub


Public Function RegistrarDatos( _
                                Optional p_SolicitudAlInicio As solicitud, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_Estado As String
    
    On Error GoTo errores
    
    With Me.solicitud
        Avance "Comprobando datos "
        m_MotivoNoOK = MotivoNoOK(p_SolicitudAlInicio, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = m_MotivoNoOK
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        .Estado = .EstadoCalculadoTexto
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Avance "Registro de de la Solicitud"
        If p_SolicitudAlInicio Is Nothing Then
            m_SQL = "TbSolicitudes"
        Else
            m_SQL = "SELECT * " & _
                    "FROM TbSolicitudes " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        End If
        
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
       If p_SolicitudAlInicio Is Nothing Then
            .IDSolicitud = .IDSolicitudCalculado
            rcdDatos.AddNew
                rcdDatos.Fields("IDSolicitud") = .IDSolicitud
                If Not .GestorObj Is Nothing Then
                    rcdDatos.Fields("Gestor") = .GestorObj.Nombre
                End If
                If .DNI <> "" Then
                    rcdDatos.Fields("DNI") = .DNI
                End If
                If .Nombre <> "" Then
                    rcdDatos.Fields("Nombre") = .Nombre
                End If
                If .Apellido1 <> "" Then
                    rcdDatos.Fields("Apellido1") = .Apellido1
                End If
                If .Apellido2 <> "" Then
                    rcdDatos.Fields("Apellido2") = .Apellido2
                End If
                If .FNacimiento <> "" Then
                    rcdDatos.Fields("FNacimiento") = .FNacimiento
                End If
                If .LugarNacimiento <> "" Then
                    rcdDatos.Fields("LugarNacimiento") = .LugarNacimiento
                End If
                If .email <> "" Then
                    rcdDatos.Fields("email") = .email
                End If
                If .Telefono <> "" Then
                    rcdDatos.Fields("Telefono") = .Telefono
                End If
                If .IDExpediente <> "" Then
                    rcdDatos.Fields("IDExpediente") = .IDExpediente
                End If
                If .IDEmpresaUsuario <> "" Then
                    rcdDatos.Fields("IDEmpresaUsuario") = .IDEmpresaUsuario
                End If
                If .IDEmpresaTramitadora <> "" Then
                    rcdDatos.Fields("IDEmpresaTramitadora") = .IDEmpresaTramitadora
                End If
                If .CorreosAutomaticos <> "" Then
                    rcdDatos.Fields("CorreosAutomaticos") = .CorreosAutomaticos
                End If
                If .Motivo_HPS <> "" Then
                    rcdDatos.Fields("Motivo_HPS") = .Motivo_HPS
                End If
                If .IDJustificacion <> "" Then
                    rcdDatos.Fields("IDJustificacion") = .IDJustificacion
                End If
                If .emailResponsable <> "" Then
                    rcdDatos.Fields("emailResponsable") = .emailResponsable
                End If
                If .TIPO <> "" Then
                    rcdDatos.Fields("TIPO") = .TIPO
                End If
                If .Estado <> "" Then
                    rcdDatos.Fields("Estado") = .Estado
                End If
                If .Observaciones <> "" Then
                    rcdDatos.Fields("Observaciones") = .Observaciones
                End If
                If .IDUsuarioHPS <> "" Then
                    rcdDatos.Fields("IDUsuarioHPS") = .IDUsuarioHPS
                End If
                rcdDatos.Fields("FechaCreacion") = Ahora
                .UsuarioCreacion = m_ObjUsuarioConectado.UsuarioRed
                rcdDatos.Fields("UsuarioCreacion") = .UsuarioCreacion
               
                
                
               
            rcdDatos.Update
            
           
        Else
            If rcdDatos.EOF Then
                p_Error = "No se ha encontrado la factura registrada"
                Err.Raise 1000
            End If
            rcdDatos.Edit
               If Not .GestorObj Is Nothing Then
                    rcdDatos.Fields("Gestor") = .GestorObj.Nombre
                Else
                    rcdDatos.Fields("Gestor") = Null
                End If
                If .DNI <> "" Then
                    rcdDatos.Fields("DNI") = .DNI
                Else
                    rcdDatos.Fields("DNI") = Null
                End If
                If .Nombre <> "" Then
                    rcdDatos.Fields("Nombre") = .Nombre
                Else
                    rcdDatos.Fields("Nombre") = Null
                End If
                If .Apellido1 <> "" Then
                    rcdDatos.Fields("Apellido1") = .Apellido1
                Else
                    rcdDatos.Fields("Apellido1") = Null
                End If
                If .Apellido2 <> "" Then
                    rcdDatos.Fields("Apellido2") = .Apellido2
                Else
                    rcdDatos.Fields("Apellido2") = Null
                End If
                If .FNacimiento <> "" Then
                    rcdDatos.Fields("FNacimiento") = .FNacimiento
                Else
                    rcdDatos.Fields("FNacimiento") = Null
                End If
                If .LugarNacimiento <> "" Then
                    rcdDatos.Fields("LugarNacimiento") = .LugarNacimiento
                Else
                    rcdDatos.Fields("LugarNacimiento") = Null
                End If
                If .email <> "" Then
                    rcdDatos.Fields("email") = .email
                Else
                    rcdDatos.Fields("email") = Null
                End If
                If .Telefono <> "" Then
                    rcdDatos.Fields("Telefono") = .Telefono
                Else
                    rcdDatos.Fields("Telefono") = Null
                End If
                If .IDExpediente <> "" Then
                    rcdDatos.Fields("IDExpediente") = .IDExpediente
                Else
                    rcdDatos.Fields("IDExpediente") = Null
                End If
                If .IDEmpresaUsuario <> "" Then
                    rcdDatos.Fields("IDEmpresaUsuario") = .IDEmpresaUsuario
                Else
                    rcdDatos.Fields("IDEmpresaUsuario") = Null
                End If
                If .IDEmpresaTramitadora <> "" Then
                    rcdDatos.Fields("IDEmpresaTramitadora") = .IDEmpresaTramitadora
                Else
                    rcdDatos.Fields("IDEmpresaTramitadora") = Null
                End If
                If .CorreosAutomaticos <> "" Then
                    rcdDatos.Fields("CorreosAutomaticos") = .CorreosAutomaticos
                Else
                    rcdDatos.Fields("CorreosAutomaticos") = Null
                End If
                If .Motivo_HPS <> "" Then
                    rcdDatos.Fields("Motivo_HPS") = .Motivo_HPS
                Else
                    rcdDatos.Fields("Motivo_HPS") = Null
                End If
                If .IDJustificacion <> "" Then
                    rcdDatos.Fields("IDJustificacion") = .IDJustificacion
                Else
                    rcdDatos.Fields("IDJustificacion") = Null
                End If
                If .emailResponsable <> "" Then
                    rcdDatos.Fields("emailResponsable") = .emailResponsable
                Else
                    rcdDatos.Fields("emailResponsable") = Null
                End If
                If .TIPO <> "" Then
                    rcdDatos.Fields("TIPO") = .TIPO
                Else
                    rcdDatos.Fields("TIPO") = Null
                End If
                If .Estado <> "" Then
                    rcdDatos.Fields("Estado") = .Estado
                Else
                    rcdDatos.Fields("Estado") = Null
                End If
                If .Observaciones <> "" Then
                    rcdDatos.Fields("Observaciones") = .Observaciones
                Else
                    rcdDatos.Fields("Observaciones") = Null
                End If
                If .IDUsuarioHPS <> "" Then
                    rcdDatos.Fields("IDUsuarioHPS") = .IDUsuarioHPS
                Else
                    rcdDatos.Fields("IDUsuarioHPS") = Null
                End If
                .FechaUltimoCambio = Ahora
                .UsuarioUltimoCambio = m_ObjUsuarioConectado.UsuarioRed
                rcdDatos.Fields("UsuarioUltimoCambio") = .UsuarioUltimoCambio
                rcdDatos.Fields("FechaUltimoCambio") = .FechaUltimoCambio
            rcdDatos.Update
            
            
            
        End If
    
        rcdDatos.Close
        Set rcdDatos = Nothing
        
       
    End With
    RegistrarActualizacion Me.solicitud
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistrarDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    Dim m_URLFinal As String
    Dim m_URLCarpeta As String
    
    On Error GoTo errores
    
    ' 1. Gestión de Archivos Adjuntos (Código existente)
    With Me.solicitud
        m_URLFinal = .URLAdjuntoCalculado
        If fso.FileExists(m_URLFinal) Then
             If fso.FileExists(m_URLFinal) Then
                If FicheroAbierto(m_URLFinal) Then
                    p_Error = "Ha de cerrar el anexo antes de eliminar la solicitud."
                    Err.Raise 1000
                End If
            End If
            m_URLCarpeta = fso.GetParentFolderName(m_URLFinal)
        End If
        
        ' 2. NUEVO: Borrado de Correos (Enviados y Programados)
        ' Borramos todo lo que haya en la tabla de correos vinculado a esta solicitud
        m_SQL = "DELETE * FROM TbCorreosEnviados " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
        ' 3. Borrado de Tablas Principales (Código existente)
        m_SQL = "DELETE * FROM TbSolicitudesFechas " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
        m_SQL = "DELETE * FROM TbSolicitudes " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
        ' 4. Limpieza Física de Archivos (Código existente)
        If fso.FileExists(m_URLFinal) Then
            fso.DeleteFile m_URLFinal, True
        End If
        If fso.FolderExists(m_URLCarpeta) Then
            ' Borra la carpeta si está vacía (o forzado)
            If fso.GetFolder(m_URLCarpeta).Files.Count = 0 Then
                fso.DeleteFolder m_URLCarpeta, True
            End If
        End If
       
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.Eliminar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function MotivoNoOK( _
                                Optional p_SolicitudAlInicio As solicitud, _
                                Optional ByRef p_Error As String _
                                ) As String
        
    
    Dim blnComprobar As Boolean
    Dim m_Expediente As Expediente
    Dim m_Suministrador As Suministrador
    Dim m_Solicitud As solicitud
    Dim blnDNIValido As Boolean
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos para guardar"
        Err.Raise 1000
    End If
    'AL DAR DE ALTA ES OBLIGATORIO
    'tipo,Correoautomatico,Gestor,Nombre,Apellido1,email,emailResponsable,EmpresaTramitacion
    With Me.solicitud
        
        If .GestorObj Is Nothing Then
            MotivoNoOK = "No se conoce el Gestor"
            Exit Function
        End If
        
        If .Nombre = "" Then
            MotivoNoOK = "No se conoce el Nombre"
            Exit Function
        End If
        If .Apellido1 = "" Then
            MotivoNoOK = "Al menos ha de tener el primer apellido"
            Exit Function
        End If
        If Not EsCorreoValido(.email) Then
            MotivoNoOK = "email no tiene un formato correcto de email"
            Exit Function
        End If
        If .emailResponsable <> "" Then
            If Not EsCorreoValido(.emailResponsable) Then
                MotivoNoOK = "emailResponsable no tiene un formato correcto de email"
                Exit Function
            End If
        End If
        If .IDExpediente <> "" Then
            If .Expediente Is Nothing Then
                MotivoNoOK = "El expediente es obligatorio"
                Exit Function
            End If
            If .Expediente.CadenaContratistas = "" Then
                MotivoNoOK = "El expediente al menos ha de tener un contratista principal"
                Exit Function
            End If
        End If
        
        If .CorreosAutomaticos <> "Sí" And .CorreosAutomaticos <> "No" Then
            MotivoNoOK = "Ha de indicar si va a haber correos automáticos para este usuario"
            Exit Function
        End If
        If .IDEmpresaTramitadora <> "" Then
             Set m_Suministrador = constructor.getSuministrador(p_ID:=.IDEmpresaTramitadora, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_Suministrador Is Nothing Then
                MotivoNoOK = "No se conoce ha rellenado una empresa tramitadora registrada"
                Exit Function
            End If
        End If
        
       
        
        
        If .DNI <> "" Then
            blnDNIValido = EsDNIValido(.DNI)
            If blnDNIValido = False Then
                MotivoNoOK = "El DNI no es válido. Rellene los números para que haya 8 y ponga la letra"
                Exit Function
            End If
'            If p_SolicitudAlInicio Is Nothing Then
'                blnComprobar = True
'            Else
'                If .DNI <> p_SolicitudAlInicio.DNI Then
'                    blnComprobar = True
'                End If
'            End If
'            If blnComprobar = True Then
'                Set m_Solicitud = constructor.getSolicitud(p_DNI:=.DNI, p_Error:=p_Error)
'                If p_Error <> "" Then
'                    Err.Raise 1000
'                End If
'                If Not m_Solicitud Is Nothing Then
'                    MotivoNoOK = "El DNI ya se ha registrado para la solicitud Nº " & m_Solicitud.IDSolicitud
'                    Exit Function
'                End If
'            End If
        End If
        
        blnComprobar = False
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método SolicitudOperaciones.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Public Function RegistroEnvioExcel( _
                                    Optional p_ConCorreo As EnumSiNo = EnumSiNo.Sí, _
                                    Optional p_Fecha As String, _
                                    Optional p_DESTINATARIOS As String, _
                                    Optional p_DestinatariosConCopia As String, _
                                    Optional p_asunto As String, _
                                    Optional p_SoloEnvioCorreo As Boolean = False, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_Correo As Correo
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    On Error GoTo errores
    
    If Me.solicitud Is Nothing Then
        p_Error = "No se ha indicado la solicitud"
        Err.Raise 1000
    End If

    ' --- BLOQUE DE REGISTRO EN BD (Solo si NO es solo envío) ---
    If p_SoloEnvioCorreo = False Then
        If Not IsDate(p_Fecha) Then p_Fecha = Date
        If CDate(p_Fecha) > CDate(Date) Then
            p_Error = "La fecha no puede ser superior a la del día de hoy"
            Err.Raise 1000
        End If
        
        With Me.solicitud
            If .IDSolicitud = "" Then
                p_Error = "La solicitud no tiene ID"
                Err.Raise 1000
            End If
            
            If .SolicitudFechas Is Nothing Then
                Set .SolicitudFechas = New SolicitudFechas
                .SolicitudFechas.IDSolicitud = .IDSolicitud
                m_Alta = EnumSiNo.Sí
            Else
                m_Alta = EnumSiNo.no
            End If
        End With
        
        With Me.solicitud.SolicitudFechas
            .FechaEnvioExcel = p_Fecha
            .FechaPrevistaCorreoRecordatorioExcel1 = .FechaPrevistaCorreoRecordatorioExcel1Calculado
            If IsDate(.FechaPrevistaAutocancelacionPreMargaCalculado) Then
                .FechaPrevistaAutocancelacionPreMarga = .FechaPrevistaAutocancelacionPreMargaCalculado
            End If
            
            Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
            Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
            m_SolicitudfechasOp.Registrar m_Alta, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End With
        
        ' Automatización y Limpieza (Solo si registramos hito)
        If EsAutomatizacionActiva() Then
            Dim dias1 As Long, dias2 As Long, diasCancel As Long
            dias1 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioExcel1, 0))
            dias2 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioExcel2, 0))
            diasCancel = CLng(Nz(m_ObjEntorno.Configuracion.DiasCancelacionPreMARGA, 0))
            
            ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioExcel1, CDate(p_Fecha), dias1, p_Error
            If dias2 > 0 Then ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioExcel2, CDate(p_Fecha), (dias1 + dias2), p_Error
            If m_ObjEntorno.Configuracion.AutocancelacionPreMARGA = "Sí" Then ProgramarCorreoFuturo EnumTipoEnvioCorreo.CancelacionPreMarga, CDate(p_Fecha), diasCancel, p_Error
        End If
        
        AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "RecepcionExcel", p_Error
        RegistrarActualizacion Me.solicitud
    End If
    ' -----------------------------------------------------------

    ' --- ENVÍO DE CORREO ---
    If p_ConCorreo = EnumSiNo.Sí Then
        Set m_Correo = EnviarCorreoNoAutomatico(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.EnvioExcel, _
                                p_Solicitud:=Me.solicitud, p_DESTINATARIOS:=p_DESTINATARIOS, _
                                p_DestinatariosConCopia:=p_DestinatariosConCopia, _
                                p_asunto:=p_asunto, _
                                p_Error:=p_Error)
        If p_Error <> "" Then
            p_Error = "Registro OK, pero error al enviar correo: " & vbNewLine & p_Error
        Else
            RegistroEnvioExcel = m_Correo.IDCORREO
        End If
    End If
    
    ' --- LOG ---
    If p_ConCorreo = EnumSiNo.Sí Then
        DescParaLog = "Enviado a " & m_Correo.DESTINATARIOS
        If p_SoloEnvioCorreo Then DescParaLog = DescParaLog & " (REENVÍO SIN CAMBIO DE FECHA)"
    Else
        DescParaLog = "Solo registro de fecha (Sin correo)"
    End If
    
    RegistrarLogGeneral p_Accion:="Registro Envío Excel", p_Descripcion:=DescParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en SolicitudOperaciones.RegistroEnvioExcel: " & Err.Description
    End If
End Function


Public Function RegistroEnvioRecordatorioExcel1( _
                                                Optional p_ConCorreo As EnumSiNo = EnumSiNo.Sí, _
                                                Optional p_Fecha As String, _
                                                Optional p_DESTINATARIOS As String, _
                                                Optional p_DestinatariosConCopia As String, _
                                                Optional p_asunto As String, _
                                                Optional p_SoloEnvioCorreo As Boolean = False, _
                                                Optional ByRef p_Error As String _
                                                ) As String
   
    Dim m_Correo As Correo
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    On Error GoTo errores
    
    If Me.solicitud Is Nothing Then
        p_Error = "No se ha indicado la solicitud"
        Err.Raise 1000
    End If

    ' --- BLOQUE DE REGISTRO EN BD ---
    If p_SoloEnvioCorreo = False Then
        If Not IsDate(p_Fecha) Then p_Fecha = Date
        
        With Me.solicitud.SolicitudFechas
            .FechaCorreoRecordatorioExcel1 = p_Fecha
            .FechaPrevistaCorreoRecordatorioExcel2 = .FechaPrevistaCorreoRecordatorioExcel2Calculado
            
            Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
            Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
            
            ' Determinar si es alta o edición (si ya existía el objeto fechas)
            m_Alta = IIf(Me.solicitud.SolicitudFechas Is Nothing, EnumSiNo.Sí, EnumSiNo.no)
            m_SolicitudfechasOp.Registrar m_Alta, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End With
        RegistrarActualizacion Me.solicitud
    End If
    ' -------------------------------

    ' --- ENVÍO DE CORREO ---
    If p_ConCorreo = EnumSiNo.Sí Then
        Set m_Correo = EnviarCorreoNoAutomatico(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.RecordatorioExcel1, _
                                p_Solicitud:=Me.solicitud, p_DESTINATARIOS:=p_DESTINATARIOS, _
                                p_DestinatariosConCopia:=p_DestinatariosConCopia, _
                                p_asunto:=p_asunto, _
                                p_Error:=p_Error)
        If p_Error <> "" Then Err.Raise 1000
        RegistroEnvioRecordatorioExcel1 = m_Correo.IDCORREO
    End If
    
    ' --- LOG ---
    If p_ConCorreo = EnumSiNo.Sí Then
        DescParaLog = "Enviado a " & m_Correo.DESTINATARIOS
        If p_SoloEnvioCorreo Then DescParaLog = DescParaLog & " (REENVÍO)"
    Else
        DescParaLog = "Registro Manual"
    End If
    RegistrarLogGeneral p_Accion:="Recordatorio Excel 1", p_Descripcion:=DescParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    
    Exit Function
errores:
    If Err.Number <> 1000 Then p_Error = "Error en RegistroEnvioRecordatorioExcel1: " & Err.Description
End Function
Public Function RegistroEnvioRecordatorioExcel2( _
                                                Optional p_ConCorreo As EnumSiNo = EnumSiNo.Sí, _
                                                Optional p_Fecha As String, _
                                                Optional p_DESTINATARIOS As String, _
                                                Optional p_DestinatariosConCopia As String, _
                                                Optional p_asunto As String, _
                                                Optional p_SoloEnvioCorreo As Boolean = False, _
                                                Optional ByRef p_Error As String _
                                                ) As String
   
    Dim m_Correo As Correo
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    On Error GoTo errores
    
    If Me.solicitud Is Nothing Then
        p_Error = "No se ha indicado la solicitud"
        Err.Raise 1000
    End If

    ' --- BLOQUE DE REGISTRO EN BD ---
    If p_SoloEnvioCorreo = False Then
        If Not IsDate(p_Fecha) Then p_Fecha = Date
        If CDate(p_Fecha) > CDate(Date) Then
            p_Error = "La fecha no puede ser superior a la del día de hoy"
            Err.Raise 1000
        End If
        
        With Me.solicitud
            If .IDSolicitud = "" Then
                p_Error = "La solicitud no tiene ID"
                Err.Raise 1000
            End If
            
            If .SolicitudFechas Is Nothing Then
                Set .SolicitudFechas = New SolicitudFechas
                .SolicitudFechas.IDSolicitud = .IDSolicitud
                m_Alta = EnumSiNo.Sí
            Else
                m_Alta = EnumSiNo.no
            End If
        End With
        
        With Me.solicitud.SolicitudFechas
            .FechaCorreoRecordatorioExcel2 = p_Fecha
            .FechaPrevistaAutocancelacionPreMarga = .FechaPrevistaAutocancelacionPreMargaCalculado
            
            ' Validamos la fecha futura solo si estamos registrando
            If Not IsDate(.FechaPrevistaAutocancelacionPreMarga) Then
                ' Nota: No lanzamos error fatal, pero registramos aviso si es necesario
                ' p_Error = "No se ha podido determinar la fecha prevista de autocancelación"
            End If
            
            Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
            Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
            m_SolicitudfechasOp.Registrar m_Alta, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End With
        
        RegistrarActualizacion Me.solicitud
    End If
    ' -------------------------------

    ' --- ENVÍO DE CORREO ---
    If p_ConCorreo = EnumSiNo.Sí Then
        Set m_Correo = EnviarCorreoNoAutomatico(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.RecordatorioExcel2, _
                                p_Solicitud:=Me.solicitud, p_DESTINATARIOS:=p_DESTINATARIOS, _
                                p_DestinatariosConCopia:=p_DestinatariosConCopia, _
                                p_asunto:=p_asunto, _
                                p_Error:=p_Error)
        If p_Error <> "" Then
            p_Error = "Registro OK, pero error al enviar correo: " & vbNewLine & p_Error
        Else
            RegistroEnvioRecordatorioExcel2 = m_Correo.IDCORREO
        End If
    End If
    
    ' --- LOG ---
    If p_ConCorreo = EnumSiNo.Sí Then
        DescParaLog = "Enviado a " & m_Correo.DESTINATARIOS
        If p_SoloEnvioCorreo Then DescParaLog = DescParaLog & " (REENVÍO)"
    Else
        DescParaLog = "Registro Manual"
    End If
    
    RegistrarLogGeneral p_Accion:="Registro Recordatorio Envio Excel 2", p_Descripcion:=DescParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroEnvioRecordatorioExcel2 ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function RegistroEnvioAutocancelaconPREMARGA( _
                                                    Optional p_ConCorreo As EnumSiNo = EnumSiNo.Sí, _
                                                    Optional p_Fecha As String, _
                                                    Optional p_DESTINATARIOS As String, _
                                                    Optional p_DestinatariosConCopia As String, _
                                                    Optional p_asunto As String, _
                                                    Optional p_SoloEnvioCorreo As Boolean = False, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
   
    Dim m_Correo As Correo
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    On Error GoTo errores
    
    If Me.solicitud Is Nothing Then
        p_Error = "No se ha indicado la solicitud"
        Err.Raise 1000
    End If

    ' --- BLOQUE DE REGISTRO EN BD ---
    If p_SoloEnvioCorreo = False Then
        If Not IsDate(p_Fecha) Then p_Fecha = Date
        If CDate(p_Fecha) > CDate(Date) Then
            p_Error = "La fecha no puede ser superior a la del día de hoy"
            Err.Raise 1000
        End If
        
        With Me.solicitud
            If .IDSolicitud = "" Then
                p_Error = "La solicitud no tiene ID"
                Err.Raise 1000
            End If
            
            If .SolicitudFechas Is Nothing Then
                Set .SolicitudFechas = New SolicitudFechas
                .SolicitudFechas.IDSolicitud = .IDSolicitud
                m_Alta = EnumSiNo.Sí
            Else
                m_Alta = EnumSiNo.no
            End If
        End With
        
        With Me.solicitud.SolicitudFechas
            .FechaCancelado = p_Fecha ' Autocancelación implica cancelar la solicitud
            
            Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
            Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
            m_SolicitudfechasOp.Registrar m_Alta, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End With
        
        ' Al cancelar, borramos cualquier programación futura
        AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "TODO", p_Error
        RegistrarActualizacion Me.solicitud
    End If
    ' -------------------------------

    ' --- ENVÍO DE CORREO ---
    If p_ConCorreo = EnumSiNo.Sí Then
        Set m_Correo = EnviarCorreoNoAutomatico(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.CancelacionPreMarga, _
                                p_Solicitud:=Me.solicitud, p_DESTINATARIOS:=p_DESTINATARIOS, _
                                p_DestinatariosConCopia:=p_DestinatariosConCopia, _
                                p_asunto:=p_asunto, _
                                p_Error:=p_Error)
        If p_Error <> "" Then
            p_Error = "Registro OK, error envío correo: " & vbNewLine & p_Error
        Else
            RegistroEnvioAutocancelaconPREMARGA = m_Correo.IDCORREO
        End If
    End If
    
    ' --- LOG ---
    If p_ConCorreo = EnumSiNo.Sí Then
        DescParaLog = "Enviado a " & m_Correo.DESTINATARIOS
        If p_SoloEnvioCorreo Then DescParaLog = DescParaLog & " (REENVÍO)"
    Else
        DescParaLog = "Registro Manual"
    End If
    
    RegistrarLogGeneral p_Accion:="Registro Autocancelación Premarga", p_Descripcion:=DescParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroEnvioAutocancelaconPREMARGA ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function RegistroEnvioRecordatorioMarga1( _
                                                Optional p_ConCorreo As EnumSiNo = EnumSiNo.Sí, _
                                                Optional p_Fecha As String, _
                                                Optional p_DESTINATARIOS As String, _
                                                Optional p_DestinatariosConCopia As String, _
                                                Optional p_asunto As String, _
                                                Optional p_SoloEnvioCorreo As Boolean = False, _
                                                Optional ByRef p_Error As String _
                                                ) As String
   
    Dim m_Correo As Correo
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    On Error GoTo errores
    
    If Me.solicitud Is Nothing Then
        p_Error = "No se ha indicado la solicitud"
        Err.Raise 1000
    End If

    ' --- BLOQUE DE REGISTRO EN BD ---
    If p_SoloEnvioCorreo = False Then
        If Not IsDate(p_Fecha) Then p_Fecha = Date
        If CDate(p_Fecha) > CDate(Date) Then
            p_Error = "La fecha no puede ser superior a la del día de hoy"
            Err.Raise 1000
        End If
        
        With Me.solicitud
            If .IDSolicitud = "" Then
                p_Error = "La solicitud no tiene ID"
                Err.Raise 1000
            End If
            
            If .SolicitudFechas Is Nothing Then
                Set .SolicitudFechas = New SolicitudFechas
                .SolicitudFechas.IDSolicitud = .IDSolicitud
                m_Alta = EnumSiNo.Sí
            Else
                m_Alta = EnumSiNo.no
            End If
        End With
        
        With Me.solicitud.SolicitudFechas
            .FechaCorreoRecordatorioRellenoMarga1 = p_Fecha
            .FechaPrevistaCorreoRecordatorioRellenoMarga2 = .FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado
            If IsDate(.FechaPrevistaAutocancelacionMargaCalculado) Then
                .FechaPrevistaAutocancelacionMarga = .FechaPrevistaAutocancelacionMargaCalculado
            End If
            
            Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
            Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
            m_SolicitudfechasOp.Registrar m_Alta, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End With
        
        RegistrarActualizacion Me.solicitud
    End If
    ' -------------------------------

    ' --- ENVÍO DE CORREO ---
    If p_ConCorreo = EnumSiNo.Sí Then
        Set m_Correo = EnviarCorreoNoAutomatico(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.RecordatorioMarga1, _
                                p_Solicitud:=Me.solicitud, p_DESTINATARIOS:=p_DESTINATARIOS, _
                                p_DestinatariosConCopia:=p_DestinatariosConCopia, _
                                p_asunto:=p_asunto, _
                                p_Error:=p_Error)
        If p_Error <> "" Then
            p_Error = "Registro OK, error envío correo: " & vbNewLine & p_Error
        Else
            RegistroEnvioRecordatorioMarga1 = m_Correo.IDCORREO
        End If
    End If
    
    ' --- LOG ---
    If p_ConCorreo = EnumSiNo.Sí Then
        DescParaLog = "Enviado a " & m_Correo.DESTINATARIOS
        If p_SoloEnvioCorreo Then DescParaLog = DescParaLog & " (REENVÍO)"
    Else
        DescParaLog = "Registro Manual"
    End If
    
    RegistrarLogGeneral p_Accion:="Registro Recordatorio Marga 1", p_Descripcion:=DescParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroEnvioRecordatorioMarga1 ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function RegistroEnvioRecordatorioMarga2( _
                                                Optional p_ConCorreo As EnumSiNo = EnumSiNo.Sí, _
                                                Optional p_Fecha As String, _
                                                Optional p_DESTINATARIOS As String, _
                                                Optional p_DestinatariosConCopia As String, _
                                                Optional p_asunto As String, _
                                                Optional p_SoloEnvioCorreo As Boolean = False, _
                                                Optional ByRef p_Error As String _
                                                ) As String
   
    Dim m_Correo As Correo
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    On Error GoTo errores
    
    If Me.solicitud Is Nothing Then
        p_Error = "No se ha indicado la solicitud"
        Err.Raise 1000
    End If

    ' --- BLOQUE DE REGISTRO EN BD ---
    If p_SoloEnvioCorreo = False Then
        If Not IsDate(p_Fecha) Then p_Fecha = Date
        If CDate(p_Fecha) > CDate(Date) Then
            p_Error = "La fecha no puede ser superior a la del día de hoy"
            Err.Raise 1000
        End If
        
        With Me.solicitud
            If .IDSolicitud = "" Then
                p_Error = "La solicitud no tiene ID"
                Err.Raise 1000
            End If
            
            If .SolicitudFechas Is Nothing Then
                Set .SolicitudFechas = New SolicitudFechas
                .SolicitudFechas.IDSolicitud = .IDSolicitud
                m_Alta = EnumSiNo.Sí
            Else
                m_Alta = EnumSiNo.no
            End If
        End With
        
        With Me.solicitud.SolicitudFechas
            .FechaCorreoRecordatorioRellenoMarga2 = p_Fecha
            .FechaPrevistaAutocancelacionMarga = .FechaPrevistaAutocancelacionMargaCalculado
            
            Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
            Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
            m_SolicitudfechasOp.Registrar m_Alta, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End With
        
        RegistrarActualizacion Me.solicitud
    End If
    ' -------------------------------

    ' --- ENVÍO DE CORREO ---
    If p_ConCorreo = EnumSiNo.Sí Then
        Set m_Correo = EnviarCorreoNoAutomatico(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.RecordatorioMarga2, _
                                p_Solicitud:=Me.solicitud, p_DESTINATARIOS:=p_DESTINATARIOS, _
                                p_DestinatariosConCopia:=p_DestinatariosConCopia, _
                                p_asunto:=p_asunto, _
                                p_Error:=p_Error)
        If p_Error <> "" Then
            p_Error = "Registro OK, error envío correo: " & vbNewLine & p_Error
        Else
            RegistroEnvioRecordatorioMarga2 = m_Correo.IDCORREO
        End If
    End If
    
    ' --- LOG ---
    If p_ConCorreo = EnumSiNo.Sí Then
        DescParaLog = "Enviado a " & m_Correo.DESTINATARIOS
        If p_SoloEnvioCorreo Then DescParaLog = DescParaLog & " (REENVÍO)"
    Else
        DescParaLog = "Registro Manual"
    End If
    
    RegistrarLogGeneral p_Accion:="Registro Recordatorio Marga 2", p_Descripcion:=DescParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroEnvioRecordatorioMarga2 ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function RegistroEnvioAutocancelaconMARGA( _
                                                Optional p_ConCorreo As EnumSiNo = EnumSiNo.Sí, _
                                                Optional p_Fecha As String, _
                                                Optional p_DESTINATARIOS As String, _
                                                Optional p_DestinatariosConCopia As String, _
                                                Optional p_asunto As String, _
                                                Optional p_SoloEnvioCorreo As Boolean = False, _
                                                Optional ByRef p_Error As String _
                                                ) As String
   
    Dim m_Correo As Correo
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    On Error GoTo errores
    
    If Me.solicitud Is Nothing Then
        p_Error = "No se ha indicado la solicitud"
        Err.Raise 1000
    End If

    ' --- BLOQUE DE REGISTRO EN BD ---
    If p_SoloEnvioCorreo = False Then
        If Not IsDate(p_Fecha) Then p_Fecha = Date
        If CDate(p_Fecha) > CDate(Date) Then
            p_Error = "La fecha no puede ser superior a la del día de hoy"
            Err.Raise 1000
        End If
        
        With Me.solicitud
            If .IDSolicitud = "" Then
                p_Error = "La solicitud no tiene ID"
                Err.Raise 1000
            End If
            
            If .SolicitudFechas Is Nothing Then
                Set .SolicitudFechas = New SolicitudFechas
                .SolicitudFechas.IDSolicitud = .IDSolicitud
                m_Alta = EnumSiNo.Sí
            Else
                m_Alta = EnumSiNo.no
            End If
        End With
        
        With Me.solicitud.SolicitudFechas
            .FechaCancelado = p_Fecha ' La cancelación en Marga también rellena FechaCancelado
            
            Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
            Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
            m_SolicitudfechasOp.Registrar m_Alta, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End With
        
        ' Limpiar automatización futura
        AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "TODO", p_Error
        RegistrarActualizacion Me.solicitud
    End If
    ' -------------------------------

    ' --- ENVÍO DE CORREO ---
    If p_ConCorreo = EnumSiNo.Sí Then
        Set m_Correo = EnviarCorreoNoAutomatico(p_EnumTipoEnvioCorreo:=EnumTipoEnvioCorreo.CancelacionMarga, _
                                p_Solicitud:=Me.solicitud, p_DESTINATARIOS:=p_DESTINATARIOS, _
                                p_DestinatariosConCopia:=p_DestinatariosConCopia, _
                                p_asunto:=p_asunto, _
                                p_Error:=p_Error)
        If p_Error <> "" Then
            p_Error = "Registro OK, error envío correo: " & vbNewLine & p_Error
        Else
            RegistroEnvioAutocancelaconMARGA = m_Correo.IDCORREO
        End If
    End If
    
    ' --- LOG ---
    If p_ConCorreo = EnumSiNo.Sí Then
        DescParaLog = "Enviado a " & m_Correo.DESTINATARIOS
        If p_SoloEnvioCorreo Then DescParaLog = DescParaLog & " (REENVÍO)"
    Else
        DescParaLog = "Registro Manual"
    End If
    
    RegistrarLogGeneral p_Accion:="Registro Autocancelación Marga", p_Descripcion:=DescParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroEnvioAutocancelaconMARGA ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function MotivoNoOKRegistrarUsuarioHPSEnHPS( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    
    On Error GoTo errores
    If m_EnOficina <> EnumSiNo.Sí Then
        p_Error = "Esta operación sólo se puede hacer en la oficina"
        Err.Raise 1000
    End If
    With Me.solicitud
        
        If .DNI = "" Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El DNI es obligatorio"
            Exit Function
        End If
        
        If .Nombre = "" Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El Nombre es obligatorio"
            Exit Function
        End If
        If .Apellido1 = "" Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El Apellido1 es obligatorio"
            Exit Function
        End If
        If .Telefono = "" Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El Telefono es obligatorio"
            Exit Function
        End If
        If .EmpresaUsuario Is Nothing Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "La empresa de usuario es obligatoria"
            Exit Function
        End If
        If .EmpresaTramitadora Is Nothing Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "La empresa de Tramitación del HPS es obligatoria"
            Exit Function
        End If
        If InStr(1, .email, "@") = 0 Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "Falta el Correo electrónico de la solicitud o no tiene la forma adecuada"
            Exit Function
        End If
        If .Motivo_HPS = "" Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El Motivo_HPS es obligatorio"
            Exit Function
        End If
        If .Expediente Is Nothing Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El expediente es obligatorio"
            Exit Function
        End If
        If .Expediente.CadenaContratistas = "" Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El expediente ha de tener al menos un contratista"
            Exit Function
        End If
        If .LugarNacimiento = "" Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "El lugar de nacimiento es obligatoria"
            Exit Function
        End If
        If .Expediente Is Nothing Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "Se ha de indicar el expediente"
            Exit Function
        End If
        If Not IsDate(.FNacimiento) Then
            MotivoNoOKRegistrarUsuarioHPSEnHPS = "La FNacimiento es obligatoria"
            Exit Function
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOKRegistrarUsuarioHPSEnHPS ha devuelto el error: " & Err.Description
     End If
                                
End Function

Public Function RegistrarUsuarioHPSEnHPS( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    Dim m_MotivoNoOKRegistrarUsuarioHPSEnHPS As String
    Dim m_Resultado As String
    On Error GoTo errores
    
   
    m_MotivoNoOKRegistrarUsuarioHPSEnHPS = MotivoNoOKRegistrarUsuarioHPSEnHPS(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOKRegistrarUsuarioHPSEnHPS
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbUsuarios " & _
            "WHERE DNI='" & Me.solicitud.DNI & "';"
            
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        rcdDatos.AddNew
            m_ID = DameID("TbUsuarios", "ID", getdbHPS(), p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Me.solicitud.IDUsuarioHPS = m_ID
            
            rcdDatos.Fields("ID") = m_ID
            rcdDatos.Fields("DNI") = Me.solicitud.DNI
            rcdDatos.Fields("Nombre") = Me.solicitud.Nombre
            rcdDatos.Fields("Apellido_1") = Me.solicitud.Apellido1
            If Me.solicitud.Apellido2 <> "" Then
                rcdDatos.Fields("Apellido_2") = Me.solicitud.Apellido2
            End If
            rcdDatos.Fields("Telefono") = Me.solicitud.Telefono
            rcdDatos.Fields("Correo_e") = Me.solicitud.email
            rcdDatos.Fields("F_Nacimiento") = Me.solicitud.FNacimiento
            rcdDatos.Fields("LugarNacimiento") = Me.solicitud.LugarNacimiento
            rcdDatos.Fields("Motivo_HPS") = Me.solicitud.Motivo_HPS
            rcdDatos.Fields("IDEmpresaUsuario") = Me.solicitud.IDEmpresaUsuario
            rcdDatos.Fields("IDEmpresaHPS") = Me.solicitud.IDEmpresaTramitadora
            rcdDatos.Fields("IDExpediente") = Me.solicitud.IDExpediente
            
            
        rcdDatos.Update
        
        
        m_Resultado = "0"
    Else
        rcdDatos.Edit
            rcdDatos.Fields("Nombre") = Me.solicitud.Nombre
            rcdDatos.Fields("Apellido_1") = Me.solicitud.Apellido1
            If Me.solicitud.Apellido2 <> "" Then
                rcdDatos.Fields("Apellido_2") = Me.solicitud.Apellido2
            End If
            rcdDatos.Fields("Telefono") = Me.solicitud.Telefono
            rcdDatos.Fields("Correo_e") = Me.solicitud.email
            rcdDatos.Fields("F_Nacimiento") = Me.solicitud.FNacimiento
            rcdDatos.Fields("LugarNacimiento") = Me.solicitud.LugarNacimiento
            rcdDatos.Fields("Motivo_HPS") = Me.solicitud.Motivo_HPS
            rcdDatos.Fields("IDEmpresaUsuario") = Me.solicitud.IDEmpresaUsuario
            rcdDatos.Fields("IDEmpresaHPS") = Me.solicitud.IDEmpresaTramitadora
            rcdDatos.Fields("IDExpediente") = Me.solicitud.IDExpediente
            
        rcdDatos.Update
        m_Resultado = "1"
    End If
    RegistrarEnHPS Date, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    RegistrarUsuarioHPSEnHPS = m_Resultado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistrarUsuarioHPSEnHPS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Private Function RegistrarEnHPS( _
                                p_Fecha As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    
    On Error GoTo errores
    
    Me.solicitud.SolicitudFechas.FechaRegistroEnHPS = Format(p_Fecha, "mm/dd/yyyy")
    m_SQL = "UPDATE TbSolicitudesFechas SET FechaRegistroEnHPS =#" & Format(Me.solicitud.SolicitudFechas.FechaRegistroEnHPS, "mm/dd/yyyy") & "# " & _
            "WHERE IDSolicitud=" & Me.solicitud.IDSolicitud & ";"
    getdb().Execute m_SQL
    Me.solicitud.EstadoGrabar
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistrarEnHPS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function CambiarTipo( _
                                p_Tipo As String, _
                                Optional ByRef p_Error As String _
                                ) As String
        
    Dim m_SQL As String
    Dim m_TipoCalculado As EnumTipo
    Dim m_TipoAlInicio As String
    On Error GoTo errores
    If p_Tipo = "" Then
        p_Error = "No se ha indicado el tipo"
        Err.Raise 1000
    End If
    With Me.solicitud
        m_TipoAlInicio = .TIPO
        If m_TipoAlInicio = p_Tipo Then
            p_Error = "Ya tiene ese tipo"
            Err.Raise 1000
        End If
        m_TipoCalculado = .TipoCalculado
        If m_TipoCalculado = EnumTipo.Traspaso Then
            p_Error = "El tipo actual es traspaso y no admite cambio de tipo"
            Err.Raise 1000
        End If
        .TIPO = p_Tipo
        m_SQL = "UPDATE TbSolicitudes SET " & _
                "TIPO ='" & p_Tipo & "' " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
    End With
    
    
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Cambio del tipo " & m_TipoAlInicio & " al " & p_Tipo, p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.CambiarTipo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function EliminarFase( _
                            p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SolFechaOp As New SolicitudFechasOperaciones
    Dim m_SQL As String
    Dim m_AccionParaLog As String
    
    On Error GoTo errores
    If Me.solicitud.SolicitudFechas Is Nothing Then
        p_Error = "No hay registro de fechas"
        Err.Raise 1000
    End If
    Set m_SolFechaOp.SolicitudFechas = Me.solicitud.SolicitudFechas
    With Me.solicitud.SolicitudFechas
        If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.EnvioExcel Then
            If IsDate(.FechaRecepcionExcel) Then
                p_Error = "Ya se ha registrado el envío del Excel"
                Err.Raise 1000
            End If
            .FechaCorreoRecordatorioExcel1 = ""
            .FechaCorreoRecordatorioExcel2 = ""
            .FechaPrevistaCorreoRecordatorioExcel1 = ""
            .FechaPrevistaCorreoRecordatorioExcel2 = ""
            .FechaPrevistaAutocancelacionPreMarga = ""
           m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                "FechaEnvioExcel = Null, " & _
                "FechaCorreoRecordatorioExcel1 = Null, " & _
                "FechaCorreoRecordatorioExcel2 = Null, " & _
                "FechaPrevistaCorreoRecordatorioExcel1 = Null, " & _
                "FechaPrevistaCorreoRecordatorioExcel2 = Null, " & _
                "FechaPrevistaAutocancelacionPreMarga = Null " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Envio Excel"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel1 Then
            .FechaCorreoRecordatorioExcel1 = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaCorreoRecordatorioExcel1 = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Recordatorio Excel 1"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioExcel2 Then
            .FechaCorreoRecordatorioExcel2 = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaCorreoRecordatorioExcel2 = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Recordatorio Excel 2"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionPreMarga Then
            .FechaPrevistaAutocancelacionPreMarga = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaPrevistaAutocancelacionPreMarga = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Cancelación PreMarga"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.ExcelAdjuntado Then
            EliminarExcelDevuelto p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.AltaEnMarga Then
            EliminarMarga p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Function
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga1 Then
            .FechaCorreoRecordatorioRellenoMarga1 = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaCorreoRecordatorioRellenoMarga1 = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Recordatorio 1 Envío DPS"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.RecordatorioMarga2 Then
            .FechaCorreoRecordatorioRellenoMarga2 = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaCorreoRecordatorioRellenoMarga2 = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Recordatorio 2 Envío DPS"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.CancelacionMarga Then
            .FechaPrevistaAutocancelacionMarga = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaPrevistaAutocancelacionMarga = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Cancelación Marga"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.EnvioDPS Then
            If IsDate(.FechaEnvioONS) Then
                p_Error = "Ya se ha registrado el envío ONS"
                Err.Raise 1000
            End If
            .FechaEnvioONS = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaEnvioDPS = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Envío DPS"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.EnvioONS Then
            
            .FechaEnvioONS = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaEnvioONS = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Envío ONS"
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.Desestimado Then
            .FechaDesestimado = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaDesestimado = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Registro Desestimado"
        
        ElseIf p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.cancelacion Then
            .FechaCancelado = ""
            m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaCancelado = Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
            m_AccionParaLog = "Eliminado Cancelación"
        Else
    
            p_Error = "Acción no codificada"
            Err.Raise 1000
        End If
        getdb().Execute m_SQL
        
        
        
    End With
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    
    If m_AccionParaLog <> "" Then
        p_Error = ""
        RegistrarLogGeneral p_Accion:=m_AccionParaLog, p_Solicitud:=Me.solicitud, p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.Registrar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function RegistroExcelDevuelto( _
                                        p_URLLocal As String, _
                                        Optional p_Fecha As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_SQL As String
    Dim m_URLFinal As String
    Dim m_URLFinalActual As String
    Dim m_CarpetaAnexo As String
    Dim m_NombreArchivo As String
    
    On Error GoTo errores
    'aquí ha de estar ya el expediente, y las empresas
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me.solicitud
        If Not fso.FileExists(p_URLLocal) Then
            p_Error = "El adjunto no es alcanzable"
            Err.Raise 1000
        End If
        If .IDSolicitud = "" Then
            p_Error = "Primero hemos de registrar la solicitud"
            Err.Raise 1000
        End If
        .SolicitudFechas.FechaRecepcionExcel = p_Fecha
        If Not IsDate(.SolicitudFechas.FechaEnvioExcel) Then
            p_Error = "Primero hay que registrar el envío del Excel"
            Err.Raise 1000
        End If
        If CDate(.SolicitudFechas.FechaRecepcionExcel) < CDate(.SolicitudFechas.FechaEnvioExcel) Then
            p_Error = "La fecha no puede ser anterior a la del envío del Excel"
            Err.Raise 1000
        End If
        m_URLFinal = getURLFinal(p_URLLocal, .IDSolicitud, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If fso.FileExists(m_URLFinal) Then
            If FicheroAbierto(m_URLFinal) Then
                p_Error = "Parece que el adjunto existe en el destino y está abierto o de sólo lectura"
                Err.Raise 1000
            End If
        End If
        m_URLFinalActual = .URLAdjuntoCalculado
        If fso.FileExists(m_URLFinalActual) Then
            If FicheroAbierto(m_URLFinalActual) Then
                p_Error = "Parece que el adjunto existe en el destino y está abierto o de sólo lectura"
                Err.Raise 1000
            End If
        End If
        m_CarpetaAnexo = .URLDirectorioAnexo
        If Not fso.FolderExists(m_CarpetaAnexo) Then
            fso.CreateFolder m_CarpetaAnexo
        End If
        m_NombreArchivo = fso.GetFileName(p_URLLocal)
        fso.CopyFile p_URLLocal, m_CarpetaAnexo & m_NombreArchivo
        .URLAdjunto = m_NombreArchivo
        m_SQL = "UPDATE TbSolicitudes SET " & _
                    "URLAdjunto='" & .URLAdjunto & "' " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
        
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaRecepcionExcel=#" & Format(.SolicitudFechas.FechaRecepcionExcel, "mm/dd/yyyy") & "# " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Registro Excel del Solicitante", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
     ' Limpiar la basura futura que ya no sirve
    AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "RecepcionExcel", p_Error
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudFechasOperaciones.RegistroExcelDevuelto ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EliminarExcelDevuelto( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_SQL As String
    Dim m_URLFinalActual As String
    Dim m_CarpetaAnexo As String
   
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    With Me.solicitud
        m_URLFinalActual = .URLAdjuntoCalculado
        If fso.FileExists(m_URLFinalActual) Then
            If FicheroAbierto(m_URLFinalActual) Then
                p_Error = "Parece que el adjuntestá abierto o de sólo lectura"
                Err.Raise 1000
            End If
        End If
        m_CarpetaAnexo = fso.GetParentFolderName(m_URLFinalActual)
        
    End With
    
    With Me.solicitud.SolicitudFechas
        If Me.solicitud.TipoCalculado = EnumTipo.Alta Then
            If IsDate(.FechaTramitacionAltaMarga) Then
                p_Error = "Tiene registrada el alta de la ficha en Marga"
                Err.Raise 1000
            End If
        End If
        If Me.solicitud.TipoCalculado = EnumTipo.Renovacion Or Me.solicitud.TipoCalculado = EnumTipo.ElevacionGrado Then
            If IsDate(.FechaEnvioONS) Then
                p_Error = "Tiene registrada la fecha de envío del formulario de traspaso a la ONS"
                Err.Raise 1000
            End If
        End If
    
        .FechaRecepcionExcel = ""
        m_SQL = "UPDATE TbSolicitudes SET " & _
                    "URLAdjunto=Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaRecepcionExcel=Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    If fso.FileExists(m_URLFinalActual) Then
        fso.DeleteFile m_URLFinalActual, True
    End If
    If fso.FolderExists(m_CarpetaAnexo) Then
        If fso.GetFolder(m_CarpetaAnexo).Files.Count = 0 Then
            fso.DeleteFolder m_CarpetaAnexo, True
        End If
    End If
    Me.solicitud.URLAdjunto = ""
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Eliminado Excel del Solicitante", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    '
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.EliminarExcelDevuelto ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function RegistroEnvioTraspasoONS( _
                                            p_URLLocal As String, _
                                            Optional p_Fecha As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_SQL As String
    Dim m_URLFinal As String
    Dim m_URLFinalActual As String
    Dim m_CarpetaAnexo As String
    Dim m_NombreArchivo As String
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me.solicitud
        If Not fso.FileExists(p_URLLocal) Then
            p_Error = "El adjunto no es alcanzable"
            Err.Raise 1000
        End If
        If .IDSolicitud = "" Then
            p_Error = "Primero hemos de registrar la solicitud"
            Err.Raise 1000
        End If
        .SolicitudFechas.FechaEnvioONS = p_Fecha
        If Not IsDate(.SolicitudFechas.FechaRecepcionExcel) Then
            p_Error = "Primero hay que registrar la recepción del excel del usuario"
            Err.Raise 1000
        End If
        If CDate(.SolicitudFechas.FechaRecepcionExcel) > CDate(.SolicitudFechas.FechaEnvioONS) Then
            p_Error = "La fecha no puede ser anterior a la de recepción del excel del usuario"
            Err.Raise 1000
        End If
        m_URLFinal = getURLFinal(p_URLLocal, .IDSolicitud, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If fso.FileExists(m_URLFinal) Then
            If FicheroAbierto(m_URLFinal) Then
                p_Error = "Parece que el adjunto existe en el destino y está abierto o de sólo lectura"
                Err.Raise 1000
            End If
        End If
        m_URLFinalActual = .URLAdjuntoEnvioONSCalculado
        If fso.FileExists(m_URLFinalActual) Then
            If FicheroAbierto(m_URLFinalActual) Then
                p_Error = "Parece que el adjunto existe en el destino y está abierto o de sólo lectura"
                Err.Raise 1000
            End If
        End If
        m_CarpetaAnexo = .URLDirectorioAnexo
        If Not fso.FolderExists(m_CarpetaAnexo) Then
            fso.CreateFolder m_CarpetaAnexo
        End If
        m_NombreArchivo = fso.GetFileName(p_URLLocal)
        fso.CopyFile p_URLLocal, m_CarpetaAnexo & m_NombreArchivo
        .URLAdjuntoEnvioONS = m_NombreArchivo
        m_SQL = "UPDATE TbSolicitudes SET " & _
                    "URLAdjuntoEnvioONS='" & .URLAdjuntoEnvioONS & "' " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
        
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaEnvioONS=#" & Format(.SolicitudFechas.FechaEnvioONS, "mm/dd/yyyy") & "# " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Registro Envío formulario Traspaso a la ONS", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudFechasOperaciones.RegistroEnvioTraspasoONS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EliminarEnvioTraspasoONS( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_SQL As String
    Dim m_URLFinalActual As String
    Dim m_CarpetaAnexo As String
   
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    With Me.solicitud
        m_URLFinalActual = .URLAdjuntoEnvioONSCalculado
        If fso.FileExists(m_URLFinalActual) Then
            If FicheroAbierto(m_URLFinalActual) Then
                p_Error = "Parece que el adjuntestá abierto o de sólo lectura"
                Err.Raise 1000
            End If
        End If
        m_CarpetaAnexo = fso.GetParentFolderName(m_URLFinalActual)
        
    End With
    
    With Me.solicitud.SolicitudFechas
        
        
        .FechaEnvioONS = ""
        m_SQL = "UPDATE TbSolicitudes SET " & _
                    "URLAdjuntoEnvioONS=Null " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                    "FechaEnvioONS=Null, " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    If fso.FileExists(m_URLFinalActual) Then
        fso.DeleteFile m_URLFinalActual, True
    End If
    If fso.FolderExists(m_CarpetaAnexo) Then
        If fso.GetFolder(m_CarpetaAnexo).Files.Count = 0 Then
            fso.DeleteFolder m_CarpetaAnexo, True
        End If
    End If
    Me.solicitud.URLAdjuntoEnvioONS = ""
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Eliminado Envío formulario Traspaso a la ONS", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    '
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.EliminarEnvioTraspasoONS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function



Public Function RegistroAltaMarga( _
                                Optional p_Fecha As String, _
                                Optional ByRef p_Error As String _
                                ) As String
        
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me.solicitud
        If .SolicitudFechas Is Nothing Then
            .SolicitudFechas = New SolicitudFechas
            .SolicitudFechas.IDSolicitud = .IDSolicitud
            m_Alta = EnumSiNo.Sí
        Else
            m_Alta = EnumSiNo.no
        End If
    End With
    With Me.solicitud.SolicitudFechas
        .FechaTramitacionAltaMarga = p_Fecha
        If Not IsDate(.FechaRecepcionExcel) Then
            p_Error = "Primero hay que registrar el envío del Excel"
            Err.Raise 1000
        End If
        If CDate(.FechaRecepcionExcel) > CDate(.FechaTramitacionAltaMarga) Then
            p_Error = "La fecha no puede ser anterior a la de alta de recepción del Excel"
            Err.Raise 1000
        End If
        
        .FechaPrevistaCorreoRecordatorioRellenoMarga1 = .FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado
        .FechaPrevistaCorreoRecordatorioRellenoMarga2 = .FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado
        .FechaPrevistaAutocancelacionMarga = .FechaPrevistaAutocancelacionMargaCalculado
        Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
        Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
         m_SolicitudfechasOp.Registrar m_Alta, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Registro Alta en Marga", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
     ' --- INICIO BLOQUE AUTOMATIZACIÓN MARGA ---
    If EsAutomatizacionActiva() Then
        Dim diasM1 As Long, diasM2 As Long, diasCancelM As Long
        
        diasM1 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioRellenoMarga1, 0))
        diasM2 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioRellenoMarga2, 0))
        diasCancelM = CLng(Nz(m_ObjEntorno.Configuracion.DiasCancelacionMARGA, 0))
        
        ' 1. Programar Recordatorio Marga 1
        ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioMarga1, CDate(p_Fecha), diasM1, p_Error
        
        ' 2. Programar Recordatorio Marga 2 (Suma de tiempos igual que antes)
        If diasM2 > 0 Then
            ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioMarga2, CDate(p_Fecha), (diasM1 + diasM2), p_Error
        End If
        
        ' 3. Programar Autocancelación Marga
        If m_ObjEntorno.Configuracion.AutocancelacionMARGA = "Sí" Then
            ProgramarCorreoFuturo EnumTipoEnvioCorreo.CancelacionMarga, CDate(p_Fecha), diasCancelM, p_Error
        End If
        
        ' LIMPIEZA DE FASE ANTERIOR:
        ' Por seguridad, al entrar en Marga, nos aseguramos que no quede nada de la fase Excel
        AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "AltaMarga", p_Error
    End If
    ' --- FIN BLOQUE AUTOMATIZACIÓN ---
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistroAltaMarga ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EliminarMarga( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Me.solicitud.SolicitudFechas Is Nothing Then
        p_Error = "No hay datos Marga Registrados"
        Err.Raise 1000
    End If
    With Me.solicitud.SolicitudFechas
        If IsDate(.FechaEnvioDPS) Then
            p_Error = "Ya se ha registrado el envío del DPS"
            Err.Raise 1000
        End If
        .FechaTramitacionAltaMarga = ""
        .FechaCorreoRecordatorioRellenoMarga1 = ""
        .FechaPrevistaCorreoRecordatorioRellenoMarga1 = ""
        .FechaCorreoRecordatorioRellenoMarga2 = ""
        .FechaPrevistaCorreoRecordatorioRellenoMarga2 = ""
        .FechaPrevistaAutocancelacionMarga = ""
            
            
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                "FechaTramitacionAltaMarga=Null, " & _
                "FechaCorreoRecordatorioRellenoMarga1=Null, " & _
                "FechaPrevistaCorreoRecordatorioRellenoMarga1=Null, " & _
                "FechaCorreoRecordatorioRellenoMarga2=Null, " & _
                "FechaPrevistaCorreoRecordatorioRellenoMarga2=Null, " & _
                "FechaPrevistaAutocancelacionMarga=Null " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Eliminado Registro Alta Marga", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    '
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.EliminarMarga ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function RegistroAltaDPS( _
                                Optional p_Fecha As String, _
                                Optional ByRef p_Error As String _
                                ) As String
        
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me.solicitud
        If .SolicitudFechas Is Nothing Then
            .SolicitudFechas = New SolicitudFechas
            .SolicitudFechas.IDSolicitud = .IDSolicitud
            m_Alta = EnumSiNo.Sí
        Else
            m_Alta = EnumSiNo.no
        End If
    End With
    With Me.solicitud.SolicitudFechas
        .FechaEnvioDPS = p_Fecha
        If Not IsDate(.FechaTramitacionAltaMarga) Then
            p_Error = "Primero hay que registrar el alta en Marga"
            Err.Raise 1000
        End If
        If CDate(.FechaTramitacionAltaMarga) > CDate(.FechaEnvioDPS) Then
            p_Error = "La fecha no puede ser anterior a la de alta de tramitación Marga"
            Err.Raise 1000
        End If
        Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
        Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
         m_SolicitudfechasOp.Registrar m_Alta, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Alta Envío DPS", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ' --- LIMPIEZA ---
    ' El usuario ha cumplido con Marga. Borramos recordatorios y cancelaciones de Marga pendientes.
    AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "RecepcionDPS", p_Error
    ' ----------------
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistroAltaDPS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EliminarDPS( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Me.solicitud.SolicitudFechas Is Nothing Then
        p_Error = "No hay datos Marga Registrados"
        Err.Raise 1000
    End If
    
    With Me.solicitud.SolicitudFechas
        .FechaEnvioDPS = ""
        
            
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                "FechaEnvioDPS=Null " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Eliminar Registro DPS", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    '
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.EliminarDPS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function RegistroAltaONS( _
                                Optional p_Fecha As String, _
                                Optional ByRef p_Error As String _
                                ) As String
        
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me.solicitud
        If .SolicitudFechas Is Nothing Then
            .SolicitudFechas = New SolicitudFechas
            .SolicitudFechas.IDSolicitud = .IDSolicitud
            m_Alta = EnumSiNo.Sí
        Else
            m_Alta = EnumSiNo.no
        End If
    End With
    With Me.solicitud.SolicitudFechas
        .FechaEnvioONS = p_Fecha
        If Not IsDate(.FechaEnvioDPS) Then
            p_Error = "Primero hay que registrar el envío del DPS"
            Err.Raise 1000
        End If
        If CDate(.FechaEnvioDPS) > CDate(.FechaEnvioONS) Then
            p_Error = "La fecha no puede ser anterior a la del envío del DPS"
            Err.Raise 1000
        End If
        Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
        Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
         m_SolicitudfechasOp.Registrar m_Alta, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Alta Registro ONS", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ' --- LIMPIEZA EXTRA ---
     ' Si registran ONS directamente, asumir que DPS está hecho y limpiar Marga
     AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "RecepcionDPS", p_Error
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistroAltaONS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EliminarONS( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Me.solicitud.SolicitudFechas Is Nothing Then
        p_Error = "No hay datos DPS Registrados"
        Err.Raise 1000
    End If
    
    With Me.solicitud.SolicitudFechas
        .FechaEnvioONS = ""
        
            
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                "FechaEnvioONS=Null " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Eliminar Registro ONS", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    '
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.EliminarONS ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function RegistroDesestimado( _
                                Optional p_Fecha As String, _
                                Optional ByRef p_Error As String _
                                ) As String
        
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me.solicitud
        If .SolicitudFechas Is Nothing Then
            .SolicitudFechas = New SolicitudFechas
            .SolicitudFechas.IDSolicitud = .IDSolicitud
            m_Alta = EnumSiNo.Sí
        Else
            m_Alta = EnumSiNo.no
        End If
    End With
    With Me.solicitud.SolicitudFechas
        .FechaDesestimado = p_Fecha
        
        Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
        Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
         m_SolicitudfechasOp.Registrar m_Alta, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Alta Desestimado", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "TODO", p_Error
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistroDesestimado ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EliminarDesestimado( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Me.solicitud.SolicitudFechas Is Nothing Then
        p_Error = "No hay datos Registrados"
        Err.Raise 1000
    End If
    
    With Me.solicitud.SolicitudFechas
        .FechaDesestimado = ""
        
            
        m_SQL = "UPDATE TbSolicitudesFechas SET " & _
                "FechaDesestimado=Null " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        getdb().Execute m_SQL
        
    End With
    Me.solicitud.EstadoGrabar
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Eliminar Desestimado", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    '
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.EliminarDesestimado ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function RegistroCancelado( _
                                    Optional p_Fecha As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
        
    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    Dim m_Alta As EnumSiNo
    
    
    On Error GoTo errores
    If Me.solicitud Is Nothing Then
        p_Error = "No hay datos"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me.solicitud
        If .SolicitudFechas Is Nothing Then
            .SolicitudFechas = New SolicitudFechas
            .SolicitudFechas.IDSolicitud = .IDSolicitud
            m_Alta = EnumSiNo.Sí
        Else
            m_Alta = EnumSiNo.no
        End If
    End With
    With Me.solicitud.SolicitudFechas
        .FechaCancelado = p_Fecha
        
        Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
        Set m_SolicitudfechasOp.SolicitudFechas = Me.solicitud.SolicitudFechas
         m_SolicitudfechasOp.Registrar m_Alta, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    RegistrarActualizacion Me.solicitud
    RegistrarLogGeneral p_Accion:="Alta Desestimado", p_Solicitud:=Me.solicitud, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    AutomatizacionRepositorio.AnularProgramadosPorFase Me.solicitud.IDSolicitud, "TODO", p_Error
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudOperaciones.RegistroCancelado ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
' Método genérico para inyectar un correo en la cola del futuro
Private Sub ProgramarCorreoFuturo( _
                                p_Tipo As EnumTipoEnvioCorreo, _
                                p_FechaBase As Date, _
                                p_DiasDesfase As Long, _
                                Optional ByRef p_Error As String)
                                
    Dim m_Correo As Correo
    Dim m_CorreoOp As CorreoOperaciones
    Dim m_FechaProgramada As Date
    
    On Error GoTo errores
    
    ' Validar días
    If p_DiasDesfase <= 0 Then Exit Sub ' Si la config es 0 o vacía, no programamos
    
    ' 1. Calcular fecha futura
    m_FechaProgramada = DateAdd("d", p_DiasDesfase, p_FechaBase)
    
    ' 2. Configurar el objeto Correo
    Set m_Correo = New Correo
    With m_Correo
        .IDSolicitud = Me.solicitud.IDSolicitud
        .tipoCorreo = p_Tipo
        .Programado = "Sí"
        .FechaOrdenEnvio = m_FechaProgramada
        .DesencadenadoPor = "Sistema Automático (Programado)"
        
        ' Usamos la lógica existente en Solicitud para calcular asuntos y plantillas
        .Asunto = Me.solicitud.AsuntoCalculado(p_Tipo)
        .DESTINATARIOS = Me.solicitud.DestinatariosCalculados
        .DestinatariosConCopia = Me.solicitud.DestinatariosConCopiaCalculados(p_Tipo)
        .nombrePlantilla = Me.solicitud.NombrePlantillaParaCorreoCalculada(p_Tipo)
        .VersionPlantilla = m_ObjEntorno.Configuracion.VersionPlantillasHTML
        .cadenaRecursos = Me.solicitud.ColParametrosParaHTMLTexto
        .Accion = Me.solicitud.AccionParaCorreoCalculada(p_Tipo)
    End With
    
    ' 3. Guardar en la cola
    Set m_CorreoOp = New CorreoOperaciones
    Set m_CorreoOp.Correo = m_Correo
    m_CorreoOp.Registrar , p_Error
    
    If p_Error <> "" Then Err.Raise 1000, , "Error programando correo tipo " & p_Tipo
    
    ' 4. LOG
    RegistrarLogGeneral "Programación Automática", "Programado correo " & m_Correo.Accion & " para el " & m_FechaProgramada, Me.solicitud, p_Error

    Exit Sub
errores:
    ' No interrumpimos el flujo principal, solo devolvemos el error
    p_Error = "Error warning en ProgramarCorreoFuturo: " & Err.Description
End Sub
' ============================================================================================
' REGENERACIÓN DE AVISOS (REACTIVACIÓN)
' ============================================================================================

Public Sub RegenerarAvisos(Optional ByRef p_Error As String)
    On Error GoTo errores
    
    ' 1. Validaciones previas
    If Me.solicitud.CorreosAutomaticos <> "Sí" Then Exit Sub
    If m_ObjEntorno.Configuracion.CorreosAutomaticos <> "Sí" Then Exit Sub
    
    Dim dias1 As Long, dias2 As Long, diasCancel As Long
    Dim fechaBase As Date
    
    ' --------------------------------------------------------------------------------
    ' FASE 1: ENVÍO EXCEL -> ESPERANDO RECEPCIÓN
    ' --------------------------------------------------------------------------------
    If Me.solicitud.SolicitudFechas Is Nothing Then Exit Sub
    
    With Me.solicitud.SolicitudFechas
        ' Solo si se envió el Excel Y NO se ha recibido respuesta Y NO está cancelado/desestimado
        If IsDate(.FechaEnvioExcel) And Not IsDate(.FechaRecepcionExcel) And _
           Not IsDate(.FechaCancelado) And Not IsDate(.FechaDesestimado) Then
            
            fechaBase = CDate(.FechaEnvioExcel)
            dias1 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioExcel1, 0))
            dias2 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioExcel2, 0))
            diasCancel = CLng(Nz(m_ObjEntorno.Configuracion.DiasCancelacionPreMARGA, 0))
            
            ' Recordatorio 1
            If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioExcel1, p_Error) Then
                ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioExcel1, fechaBase, dias1, p_Error
            End If
            
            ' Recordatorio 2
            If dias2 > 0 Then
                If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioExcel2, p_Error) Then
                    ' Nota: Asumo dias2 desde el inicio. Si es acumulativo, sumar dias1.
                    ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioExcel2, fechaBase, dias2, p_Error
                End If
            End If
            
            ' Cancelación Pre-Marga
            If m_ObjEntorno.Configuracion.AutocancelacionPreMARGA = "Sí" Then
                If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.CancelacionPreMarga, p_Error) Then
                    ProgramarCorreoFuturo EnumTipoEnvioCorreo.CancelacionPreMarga, fechaBase, diasCancel, p_Error
                End If
            End If
            
        End If
        
        ' --------------------------------------------------------------------------------
        ' FASE 2: ALTA MARGA -> ESPERANDO DPS
        ' --------------------------------------------------------------------------------
        ' Solo si se dio alta en Marga Y NO se ha recibido DPS/ONS Y NO está cancelado
        If IsDate(.FechaTramitacionAltaMarga) And Not IsDate(.FechaEnvioDPS) And Not IsDate(.FechaEnvioONS) And _
           Not IsDate(.FechaCancelado) And Not IsDate(.FechaDesestimado) Then
           
            fechaBase = CDate(.FechaTramitacionAltaMarga)
            dias1 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioRellenoMarga1, 0))
            dias2 = CLng(Nz(m_ObjEntorno.Configuracion.DiasParaRecordatorioRellenoMarga2, 0))
            diasCancel = CLng(Nz(m_ObjEntorno.Configuracion.DiasCancelacionMARGA, 0))
            
            ' Recordatorio Marga 1
            If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioMarga1, p_Error) Then
                ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioMarga1, fechaBase, dias1, p_Error
            End If
            
            ' Recordatorio Marga 2
            If dias2 > 0 Then
                If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioMarga2, p_Error) Then
                    ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioMarga2, fechaBase, dias2, p_Error
                End If
            End If
            
            ' Cancelación Marga
            If m_ObjEntorno.Configuracion.AutocancelacionMARGA = "Sí" Then
                If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.CancelacionMarga, p_Error) Then
                    ProgramarCorreoFuturo EnumTipoEnvioCorreo.CancelacionMarga, fechaBase, diasCancel, p_Error
                End If
            End If
            
        End If
        
    End With
    
    Exit Sub
errores:
    p_Error = "Error en RegenerarAvisos: " & Err.Description
End Sub
' ============================================================================================
' MÉTODOS PRIVADOS DE AYUDA (AÑADIR ESTO EN SOLICITUDOPERACIONES.CLS)
' ============================================================================================

' Verifica si se deben programar correos basándose en Configuración Global y Preferencia de Solicitud
Private Function EsAutomatizacionActiva() As Boolean
    On Error Resume Next
    
    ' 1. Nivel Global
    If m_ObjEntorno.CorreosAutomaticos <> "Sí" Then
        EsAutomatizacionActiva = False
        Exit Function
    End If
    
    ' 2. Nivel Solicitud
    If Me.solicitud.CorreosAutomaticos <> "Sí" Then
        EsAutomatizacionActiva = False
        Exit Function
    End If
    
    EsAutomatizacionActiva = True
End Function


Public Sub ReprogramarAvisosFuturos(Optional ByRef p_Error As String)
    Dim opFechas As SolicitudFechasOperaciones
    Dim hayCambiosEnFechas As Boolean
    Dim fechaTeorica As Date
    Dim fechaStr As String
    Dim hoy As Date
    
    ' Variables para la transacción
    Dim wks As DAO.Workspace
    Dim enTransaccion As Boolean
    
    On Error GoTo errores
    
    ' 1. Validaciones de Seguridad (Antes de iniciar transacción)
    If m_ObjEntorno.Configuracion.CorreosAutomaticos <> "Sí" Then
        p_Error = "La configuración global de correos automáticos está desactivada."
        Err.Raise 1000
    End If
    
    If Me.solicitud.CorreosAutomaticos <> "Sí" Then
        p_Error = "Esta solicitud tiene desactivado el envío de correos automáticos."
        Err.Raise 1000
    End If
    
    If Me.solicitud.SolicitudFechas Is Nothing Then Exit Sub
    
    ' 2. INICIO DE LA TRANSACCIÓN
    Set wks = DBEngine.Workspaces(0)
    wks.BeginTrans
    enTransaccion = True
    
    hoy = Date
    hayCambiosEnFechas = False
    
    With Me.solicitud.SolicitudFechas
        
        ' --------------------------------------------------------------------------------
        ' FASE 1: PENDIENTE DE RECEPCIÓN DE EXCEL
        ' --------------------------------------------------------------------------------
        If IsDate(.FechaEnvioExcel) And Not IsDate(.FechaRecepcionExcel) And _
           Not IsDate(.FechaCancelado) And Not IsDate(.FechaDesestimado) Then
            
            ' A. Recordatorio Excel 1
            fechaStr = .FechaPrevistaCorreoRecordatorioExcel1
            If Not IsDate(fechaStr) Then
                fechaStr = .FechaPrevistaCorreoRecordatorioExcel1Calculado
                If IsDate(fechaStr) Then
                    .FechaPrevistaCorreoRecordatorioExcel1 = fechaStr
                    hayCambiosEnFechas = True
                End If
            End If
            
            If IsDate(fechaStr) Then
                fechaTeorica = CDate(fechaStr)
                If fechaTeorica >= hoy Then
                    If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioExcel1, p_Error) Then
                        ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioExcel1, fechaTeorica, 0, p_Error
                    End If
                End If
            End If
            
            ' B. Recordatorio Excel 2
            fechaStr = .FechaPrevistaCorreoRecordatorioExcel2
            If Not IsDate(fechaStr) Then
                fechaStr = .FechaPrevistaCorreoRecordatorioExcel2Calculado
                If IsDate(fechaStr) Then
                    .FechaPrevistaCorreoRecordatorioExcel2 = fechaStr
                    hayCambiosEnFechas = True
                End If
            End If
            
            If IsDate(fechaStr) Then
                fechaTeorica = CDate(fechaStr)
                If fechaTeorica >= hoy Then
                    If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioExcel2, p_Error) Then
                        ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioExcel2, fechaTeorica, 0, p_Error
                    End If
                End If
            End If
            
            ' C. Autocancelación Pre-Marga
            If m_ObjEntorno.Configuracion.AutocancelacionPreMARGA = "Sí" Then
                fechaStr = .FechaPrevistaAutocancelacionPreMarga
                If Not IsDate(fechaStr) Then
                    fechaStr = .FechaPrevistaAutocancelacionPreMargaCalculado
                    If IsDate(fechaStr) Then
                        .FechaPrevistaAutocancelacionPreMarga = fechaStr
                        hayCambiosEnFechas = True
                    End If
                End If
                
                If IsDate(fechaStr) Then
                    fechaTeorica = CDate(fechaStr)
                    If fechaTeorica >= hoy Then
                        If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.CancelacionPreMarga, p_Error) Then
                            ProgramarCorreoFuturo EnumTipoEnvioCorreo.CancelacionPreMarga, fechaTeorica, 0, p_Error
                        End If
                    End If
                End If
            End If
        End If
        
        ' --------------------------------------------------------------------------------
        ' FASE 2: PENDIENTE DE DPS (ALTA MARGA REALIZADA)
        ' --------------------------------------------------------------------------------
        If IsDate(.FechaTramitacionAltaMarga) And Not IsDate(.FechaEnvioDPS) And Not IsDate(.FechaEnvioONS) And _
           Not IsDate(.FechaCancelado) And Not IsDate(.FechaDesestimado) Then
           
            ' A. Recordatorio Marga 1
            fechaStr = .FechaPrevistaCorreoRecordatorioRellenoMarga1
            If Not IsDate(fechaStr) Then
                fechaStr = .FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado
                If IsDate(fechaStr) Then
                    .FechaPrevistaCorreoRecordatorioRellenoMarga1 = fechaStr
                    hayCambiosEnFechas = True
                End If
            End If
            
            If IsDate(fechaStr) Then
                fechaTeorica = CDate(fechaStr)
                If fechaTeorica >= hoy Then
                    If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioMarga1, p_Error) Then
                        ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioMarga1, fechaTeorica, 0, p_Error
                    End If
                End If
            End If
            
            ' B. Recordatorio Marga 2
            fechaStr = .FechaPrevistaCorreoRecordatorioRellenoMarga2
            If Not IsDate(fechaStr) Then
                fechaStr = .FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado
                If IsDate(fechaStr) Then
                    .FechaPrevistaCorreoRecordatorioRellenoMarga2 = fechaStr
                    hayCambiosEnFechas = True
                End If
            End If
            
            If IsDate(fechaStr) Then
                fechaTeorica = CDate(fechaStr)
                If fechaTeorica >= hoy Then
                    If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.RecordatorioMarga2, p_Error) Then
                        ProgramarCorreoFuturo EnumTipoEnvioCorreo.RecordatorioMarga2, fechaTeorica, 0, p_Error
                    End If
                End If
            End If
            
            ' C. Autocancelación Marga
            If m_ObjEntorno.Configuracion.AutocancelacionMARGA = "Sí" Then
                fechaStr = .FechaPrevistaAutocancelacionMarga
                If Not IsDate(fechaStr) Then
                    fechaStr = .FechaPrevistaAutocancelacionMargaCalculado
                    If IsDate(fechaStr) Then
                        .FechaPrevistaAutocancelacionMarga = fechaStr
                        hayCambiosEnFechas = True
                    End If
                End If
                
                If IsDate(fechaStr) Then
                    fechaTeorica = CDate(fechaStr)
                    If fechaTeorica >= hoy Then
                        If Not CorreoRepositorio.ExisteCorreoPendiente(Me.solicitud.IDSolicitud, EnumTipoEnvioCorreo.CancelacionMarga, p_Error) Then
                            ProgramarCorreoFuturo EnumTipoEnvioCorreo.CancelacionMarga, fechaTeorica, 0, p_Error
                        End If
                    End If
                End If
            End If
        End If
        
    End With
    
    ' 3. PERSISTENCIA FINAL EN TbSolicitudesFechas (DENTRO DE LA TRANSACCIÓN)
    If hayCambiosEnFechas Then
        Set opFechas = New SolicitudFechasOperaciones
        Set opFechas.SolicitudFechas = Me.solicitud.SolicitudFechas
        
        ' Importante: opFechas.Registrar usa getdb().Execute o Recordsets.
        ' Como getdb() usa el mismo Workspace(0), esta operación se une a la transacción.
        opFechas.Registrar EnumSiNo.no, p_Error
        
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    ' 4. CONFIRMACIÓN (COMMIT)
    wks.CommitTrans
    enTransaccion = False
    
    Exit Sub

errores:
    ' 5. GESTIÓN DE ERRORES (ROLLBACK)
    If enTransaccion Then
        On Error Resume Next ' Evitar error si el Rollback falla por algún motivo raro
        wks.Rollback
        On Error GoTo 0
    End If
    
    p_Error = "Error en ReprogramarAvisosFuturos (Rollback ejecutado): " & Err.Description
    Err.Raise 1000
End Sub


' Borra TODOS los correos programados pendientes para esta solicitud
' Se usa cuando se desactiva el check de Correos Automáticos
Public Function AnularTodosLosProgramados(Optional ByRef p_Error As String) As Boolean
    Dim m_SQL As String
    
    On Error GoTo errores
    
    ' Borramos de la tabla de correos todo lo que sea de esta solicitud,
    ' esté marcado como programado y aún no tenga fecha de envío (no haya salido)
    m_SQL = "DELETE * FROM TbCorreosEnviados " & _
            "WHERE IDSolicitud = " & Me.solicitud.IDSolicitud & " " & _
            "AND FechaEnvio Is Null " & _
            "AND Programado = 'Sí';"
            
    getdb().Execute m_SQL
    
    AnularTodosLosProgramados = True
    Exit Function

errores:
    p_Error = "Error al anular correos programados: " & Err.Description
    AnularTodosLosProgramados = False
End Function
