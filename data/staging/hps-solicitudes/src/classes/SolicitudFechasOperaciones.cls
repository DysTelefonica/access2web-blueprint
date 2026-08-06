VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolicitudFechasOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public SolicitudFechas As SolicitudFechas
' Verifica si se deben programar correos basándose en Configuración Global y Preferencia de Solicitud
Private Function EsAutomatizacionActiva() As Boolean
    On Error Resume Next
    
    ' 1. Nivel Global
    If m_ObjEntorno.Configuracion.CorreosAutomaticos <> "Sí" Then
        EsAutomatizacionActiva = False
        Exit Function
    End If
    
    ' 2. Nivel Solicitud (ACCESO CORREGIDO A TRAVÉS DE SolicitudFechas)
    If Me.SolicitudFechas.solicitud.CorreosAutomaticos <> "Sí" Then
        EsAutomatizacionActiva = False
        Exit Function
    End If
    
    EsAutomatizacionActiva = True
End Function
Private Function CamposCambian( _
                                Optional p_SolicitudFechasAlInicio As SolicitudFechas, _
                                Optional ByRef p_Error As String _
                                ) As String
     
    
    Dim m_Campos As String
    
    On Error GoTo errores
    If Me.SolicitudFechas Is Nothing Then
        Exit Function
    End If
    If p_SolicitudFechasAlInicio Is Nothing Then
        Exit Function
    End If
    With p_SolicitudFechasAlInicio
        If .FechaAutocancelacion <> Me.SolicitudFechas.FechaAutocancelacion Then
            m_Campos = "FechaAutocancelacion"
        End If
        If .FechaCorreoRecordatorioExcel1 <> Me.SolicitudFechas.FechaCorreoRecordatorioExcel1 Then
            If m_Campos = "" Then
                m_Campos = "FechaCorreoRecordatorioExcel1"
            Else
                m_Campos = m_Campos & ";" & "FechaCorreoRecordatorioExcel1"
            End If
        End If
        If .FechaCorreoRecordatorioExcel2 <> Me.SolicitudFechas.FechaCorreoRecordatorioExcel2 Then
            If m_Campos = "" Then
                m_Campos = "FechaCorreoRecordatorioExcel2"
            Else
                m_Campos = m_Campos & ";" & "FechaCorreoRecordatorioExcel2"
            End If
        End If
        If .FechaCorreoRecordatorioRellenoMarga1 <> Me.SolicitudFechas.FechaCorreoRecordatorioRellenoMarga1 Then
            If m_Campos = "" Then
                m_Campos = "FechaCorreoRecordatorioRellenoMarga1"
            Else
                m_Campos = m_Campos & ";" & "FechaCorreoRecordatorioRellenoMarga1"
            End If
        End If
        If .FechaCorreoRecordatorioRellenoMarga2 <> Me.SolicitudFechas.FechaCorreoRecordatorioRellenoMarga2 Then
            If m_Campos = "" Then
                m_Campos = "FechaCorreoRecordatorioRellenoMarga2"
            Else
                m_Campos = m_Campos & ";" & "FechaCorreoRecordatorioRellenoMarga2"
            End If
        End If
        If .FechaDesestimado <> Me.SolicitudFechas.FechaDesestimado Then
            If m_Campos = "" Then
                m_Campos = "FechaDesestimado"
            Else
                m_Campos = m_Campos & ";" & "FechaDesestimado"
            End If
        End If
        If .FechaEnvioExcel <> Me.SolicitudFechas.FechaEnvioExcel Then
            If m_Campos = "" Then
                m_Campos = "FechaEnvioExcel"
            Else
                m_Campos = m_Campos & ";" & "FechaEnvioExcel"
            End If
        End If
        If .FechaEnvioDPS <> Me.SolicitudFechas.FechaEnvioDPS Then
            If m_Campos = "" Then
                m_Campos = "FechaEnvioDPS"
            Else
                m_Campos = m_Campos & ";" & "FechaEnvioDPS"
            End If
        End If
        If .FechaEnvioONS <> Me.SolicitudFechas.FechaEnvioONS Then
            If m_Campos = "" Then
                m_Campos = "FechaEnvioONS"
            Else
                m_Campos = m_Campos & ";" & "FechaEnvioONS"
            End If
        End If
        If .FechaRegistroEnHPS <> Me.SolicitudFechas.FechaRegistroEnHPS Then
            If m_Campos = "" Then
                m_Campos = "FechaRegistroEnHPS"
            Else
                m_Campos = m_Campos & ";" & "FechaRegistroEnHPS"
            End If
        End If
        If .FechaRecepcionExcel <> Me.SolicitudFechas.FechaRecepcionExcel Then
            If m_Campos = "" Then
                m_Campos = "FechaRecepcionExcel"
            Else
                m_Campos = m_Campos & ";" & "FechaRecepcionExcel"
            End If
        End If
        
        If .FechaTramitacionAltaMarga <> Me.SolicitudFechas.FechaTramitacionAltaMarga Then
            If m_Campos = "" Then
                m_Campos = "FechaTramitacionAltaMarga"
            Else
                m_Campos = m_Campos & ";" & "FechaTramitacionAltaMarga"
            End If
        End If
        
        If .FechaPrevistaCorreoRecordatorioExcel1 <> Me.SolicitudFechas.FechaPrevistaCorreoRecordatorioExcel1 Then
            If m_Campos = "" Then
                m_Campos = "FechaPrevistaCorreoRecordatorioExcel1"
            Else
                m_Campos = m_Campos & ";" & "FechaPrevistaCorreoRecordatorioExcel1"
            End If
        End If
        If .FechaPrevistaCorreoRecordatorioExcel2 <> Me.SolicitudFechas.FechaPrevistaCorreoRecordatorioExcel2 Then
            If m_Campos = "" Then
                m_Campos = "FechaPrevistaCorreoRecordatorioExcel2"
            Else
                m_Campos = m_Campos & ";" & "FechaPrevistaCorreoRecordatorioExcel2"
            End If
        End If
        If .FechaPrevistaAutocancelacionPreMarga <> Me.SolicitudFechas.FechaPrevistaAutocancelacionPreMarga Then
            If m_Campos = "" Then
                m_Campos = "FechaPrevistaAutocancelacionPreMarga"
            Else
                m_Campos = m_Campos & ";" & "FechaPrevistaAutocancelacionPreMarga"
            End If
        End If
        If .FechaPrevistaCorreoRecordatorioRellenoMarga1 <> Me.SolicitudFechas.FechaPrevistaCorreoRecordatorioRellenoMarga1 Then
            If m_Campos = "" Then
                m_Campos = "FechaPrevistaCorreoRecordatorioRellenoMarga1"
            Else
                m_Campos = m_Campos & ";" & "FechaPrevistaCorreoRecordatorioRellenoMarga1"
            End If
        End If
        If .FechaPrevistaCorreoRecordatorioRellenoMarga2 <> Me.SolicitudFechas.FechaPrevistaCorreoRecordatorioRellenoMarga2 Then
            If m_Campos = "" Then
                m_Campos = "FechaPrevistaCorreoRecordatorioRellenoMarga2"
            Else
                m_Campos = m_Campos & ";" & "FechaPrevistaCorreoRecordatorioRellenoMarga2"
            End If
        End If
        If .FechaPrevistaAutocancelacionMarga <> Me.SolicitudFechas.FechaPrevistaAutocancelacionMarga Then
            If m_Campos = "" Then
                m_Campos = "FechaPrevistaAutocancelacionMarga"
            Else
                m_Campos = m_Campos & ";" & "FechaPrevistaAutocancelacionMarga"
            End If
        End If
        If .FechaCancelado <> Me.SolicitudFechas.FechaCancelado Then
            If m_Campos = "" Then
                m_Campos = "FechaCancelado"
            Else
                m_Campos = m_Campos & ";" & "FechaCancelado"
            End If
        End If
        
        
    End With
    CamposCambian = m_Campos
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudFechasOperaciones.CamposCambian ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function Registrar( _
                            p_Alta As EnumSiNo, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_Estado As String
    
    On Error GoTo errores
    '
    With Me.SolicitudFechas
        Avance "Comprobando datos "
'        m_MotivoNoOK = MotivoNoOK(p_Error)
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
'        p_Error = m_MotivoNoOK
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
       
        Avance "Registro de de la SolicitudFechas"
        If p_Alta = EnumSiNo.Sí Then
            m_SQL = "TbSolicitudesFechas"
        Else
            m_SQL = "SELECT * " & _
                    "FROM TbSolicitudesFechas " & _
                    "WHERE IDSolicitud=" & .IDSolicitud & ";"
        End If
        
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
       If p_Alta = EnumSiNo.Sí Then
            
            rcdDatos.AddNew
                rcdDatos.Fields("IDSolicitud") = .IDSolicitud
                If .FechaEnvioExcel <> "" Then
                    rcdDatos.Fields("FechaEnvioExcel") = .FechaEnvioExcel
                End If
                If .FechaRecepcionExcel <> "" Then
                    rcdDatos.Fields("FechaRecepcionExcel") = .FechaRecepcionExcel
                End If
                If .FechaTramitacionAltaMarga <> "" Then
                    rcdDatos.Fields("FechaTramitacionAltaMarga") = .FechaTramitacionAltaMarga
                End If
                If .FechaEnvioDPS <> "" Then
                    rcdDatos.Fields("FechaEnvioDPS") = .FechaEnvioDPS
                End If
                If .FechaEnvioONS <> "" Then
                    rcdDatos.Fields("FechaEnvioONS") = .FechaEnvioONS
                End If
                If .FechaRegistroEnHPS <> "" Then
                    rcdDatos.Fields("FechaRegistroEnHPS") = .FechaRegistroEnHPS
                End If
                If .FechaCorreoRecordatorioExcel1 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioExcel1") = .FechaCorreoRecordatorioExcel1
                End If
                If .FechaCorreoRecordatorioExcel2 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioExcel2") = .FechaCorreoRecordatorioExcel2
                End If
                If .FechaAutocancelacion <> "" Then
                    rcdDatos.Fields("FechaAutocancelacion") = .FechaAutocancelacion
                End If
                If .FechaCorreoRecordatorioRellenoMarga1 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioRellenoMarga1") = .FechaCorreoRecordatorioRellenoMarga1
                End If
                If .FechaCorreoRecordatorioRellenoMarga2 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioRellenoMarga2") = .FechaCorreoRecordatorioRellenoMarga2
                End If
                If .FechaDesestimado <> "" Then
                    rcdDatos.Fields("FechaDesestimado") = .FechaDesestimado
                End If
                If .FechaPrevistaCorreoRecordatorioExcel1 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioExcel1") = .FechaPrevistaCorreoRecordatorioExcel1
                End If
                If .FechaPrevistaCorreoRecordatorioExcel2 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioExcel2") = .FechaPrevistaCorreoRecordatorioExcel2
                End If
                If .FechaPrevistaCorreoRecordatorioRellenoMarga1 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioRellenoMarga1") = .FechaPrevistaCorreoRecordatorioRellenoMarga1
                End If
                If .FechaPrevistaCorreoRecordatorioRellenoMarga2 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioRellenoMarga2") = .FechaPrevistaCorreoRecordatorioRellenoMarga2
                End If
                If .FechaPrevistaAutocancelacionPreMarga <> "" Then
                    rcdDatos.Fields("FechaPrevistaAutocancelacionPreMarga") = .FechaPrevistaAutocancelacionPreMarga
                End If
                
                If .FechaPrevistaAutocancelacionMarga <> "" Then
                    rcdDatos.Fields("FechaPrevistaAutocancelacionMarga") = .FechaPrevistaAutocancelacionMarga
                End If
                If .FechaCancelado <> "" Then
                    rcdDatos.Fields("FechaCancelado") = .FechaCancelado
                End If
                
                
            rcdDatos.Update
'            Log.TipoOperacion = "Alta Solicitud Fechas de DNI:" & .Solicitud.DNI
'            Log.IDSolicitud = .IDSolicitud
        Else
            If rcdDatos.EOF Then
                p_Error = "No se ha encontrado la Solicitud"
                Err.Raise 1000
            End If
            rcdDatos.Edit
                If .FechaEnvioExcel <> "" Then
                    rcdDatos.Fields("FechaEnvioExcel") = .FechaEnvioExcel
                Else
                    rcdDatos.Fields("FechaEnvioExcel") = Null
                End If
                If .FechaRecepcionExcel <> "" Then
                    rcdDatos.Fields("FechaRecepcionExcel") = .FechaRecepcionExcel
                Else
                    rcdDatos.Fields("FechaRecepcionExcel") = Null
                End If
                If .FechaTramitacionAltaMarga <> "" Then
                    rcdDatos.Fields("FechaTramitacionAltaMarga") = .FechaTramitacionAltaMarga
                Else
                    rcdDatos.Fields("FechaTramitacionAltaMarga") = Null
                End If
                If .FechaEnvioDPS <> "" Then
                    rcdDatos.Fields("FechaEnvioDPS") = .FechaEnvioDPS
                Else
                    rcdDatos.Fields("FechaEnvioDPS") = Null
                End If
                If .FechaEnvioONS <> "" Then
                    rcdDatos.Fields("FechaEnvioONS") = .FechaEnvioONS
                Else
                    rcdDatos.Fields("FechaEnvioONS") = Null
                End If
                If .FechaRegistroEnHPS <> "" Then
                    rcdDatos.Fields("FechaRegistroEnHPS") = .FechaRegistroEnHPS
                Else
                    rcdDatos.Fields("FechaRegistroEnHPS") = Null
                End If
                If .FechaCorreoRecordatorioExcel1 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioExcel1") = .FechaCorreoRecordatorioExcel1
                Else
                    rcdDatos.Fields("FechaCorreoRecordatorioExcel1") = Null
                End If
                If .FechaCorreoRecordatorioExcel2 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioExcel2") = .FechaCorreoRecordatorioExcel2
                Else
                    rcdDatos.Fields("FechaCorreoRecordatorioExcel2") = Null
                End If
                If .FechaAutocancelacion <> "" Then
                    rcdDatos.Fields("FechaAutocancelacion") = .FechaAutocancelacion
                Else
                    rcdDatos.Fields("FechaAutocancelacion") = Null
                End If
                If .FechaCorreoRecordatorioRellenoMarga1 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioRellenoMarga1") = .FechaCorreoRecordatorioRellenoMarga1
                Else
                    rcdDatos.Fields("FechaCorreoRecordatorioRellenoMarga1") = Null
                End If
                If .FechaCorreoRecordatorioRellenoMarga2 <> "" Then
                    rcdDatos.Fields("FechaCorreoRecordatorioRellenoMarga2") = .FechaCorreoRecordatorioRellenoMarga2
                Else
                    rcdDatos.Fields("FechaCorreoRecordatorioRellenoMarga2") = Null
                End If
                If .FechaDesestimado <> "" Then
                    rcdDatos.Fields("FechaDesestimado") = .FechaDesestimado
                Else
                    rcdDatos.Fields("FechaDesestimado") = Null
                End If
                If .FechaPrevistaCorreoRecordatorioExcel1 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioExcel1") = .FechaPrevistaCorreoRecordatorioExcel1
                Else
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioExcel1") = Null
                End If
                If .FechaPrevistaCorreoRecordatorioExcel2 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioExcel2") = .FechaPrevistaCorreoRecordatorioExcel2
                Else
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioExcel2") = Null
                End If
                If .FechaPrevistaCorreoRecordatorioRellenoMarga1 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioRellenoMarga1") = .FechaPrevistaCorreoRecordatorioRellenoMarga1
                
                Else
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioRellenoMarga1") = Null
                End If
                If .FechaPrevistaCorreoRecordatorioRellenoMarga2 <> "" Then
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioRellenoMarga2") = .FechaPrevistaCorreoRecordatorioRellenoMarga2
                Else
                    rcdDatos.Fields("FechaPrevistaCorreoRecordatorioRellenoMarga2") = Null
                End If
                If .FechaPrevistaAutocancelacionPreMarga <> "" Then
                    rcdDatos.Fields("FechaPrevistaAutocancelacionPreMarga") = .FechaPrevistaAutocancelacionPreMarga
                Else
                    rcdDatos.Fields("FechaPrevistaAutocancelacionPreMarga") = Null
                End If
                
                If .FechaPrevistaAutocancelacionMarga <> "" Then
                    rcdDatos.Fields("FechaPrevistaAutocancelacionMarga") = .FechaPrevistaAutocancelacionMarga
                Else
                    rcdDatos.Fields("FechaPrevistaAutocancelacionMarga") = Null
                End If
                If .FechaCancelado <> "" Then
                    rcdDatos.Fields("FechaCancelado") = .FechaCancelado
                Else
                    rcdDatos.Fields("FechaCancelado") = Null
                End If
               
                
            rcdDatos.Update
'            Log.TipoOperacion = "Edición Solicitud Fechas de DNI:" & .Solicitud.DNI
'            Log.IDSolicitud = .IDSolicitud
'            Log.DESCRIPCION = "Campos que cambian" & vbNewLine & Replace(CamposCambian(p_SolicitudFechasAlInicio), ";", vbNewLine)
        End If
    
        rcdDatos.Close
        Set rcdDatos = Nothing
        RegistrarActualizacion .solicitud
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudFechasOperaciones.Registrar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    
    With Me.SolicitudFechas
       
        
        m_SQL = "DELETE * " & _
                "FROM TbSolicitudesFechas " & _
                "WHERE IDSolicitud='" & .IDSolicitud & "';"
        getdb().Execute m_SQL
        
        RegistrarActualizacion .solicitud
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudFechasOperaciones.Eliminar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function MotivoNoOK( _
                            Optional ByRef p_Error As String _
                            ) As String
    
        
    
    Dim blnComprobar As Boolean
    Dim m_Expediente As Expediente
    Dim m_SolicitudFechas As SolicitudFechas
    Dim m_URLFinal As String
    On Error GoTo errores
        
    With Me.SolicitudFechas
        If .FechaRecepcionExcel <> "" And .FechaEnvioExcel = "" Then
            MotivoNoOK = "No se puede registrar FechaRecepcionExcel sin haber registrado FechaEnvioExcel"
            Exit Function
        End If
        If .FechaTramitacionAltaMarga <> "" And .FechaRecepcionExcel = "" Then
            MotivoNoOK = "No se conoce el Gestor"
            Exit Function
        End If
    
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método SolicitudFechasOperaciones.MotivoNoOK ha devuelto el error: " & Err.Description
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
    Dim m_SolicitudRef As solicitud ' Referencia auxiliar para escribir menos código
    
    On Error GoTo errores
    
    ' Referencia a la solicitud padre
    Set m_SolicitudRef = Me.SolicitudFechas.solicitud
    
    ' Validar días
    If p_DiasDesfase <= 0 Then Exit Sub
    
    ' 1. Calcular fecha futura
    m_FechaProgramada = DateAdd("d", p_DiasDesfase, p_FechaBase)
    
    ' 2. Configurar el objeto Correo
    Set m_Correo = New Correo
    With m_Correo
        .IDSolicitud = m_SolicitudRef.IDSolicitud
        .tipoCorreo = p_Tipo
        .Programado = "Sí"
        .FechaOrdenEnvio = m_FechaProgramada
        .DesencadenadoPor = "Sistema Automático (Programado)"
        
        ' Usamos la lógica existente en la Solicitud padre
        .Asunto = m_SolicitudRef.AsuntoCalculado(p_Tipo)
        .DESTINATARIOS = m_SolicitudRef.DestinatariosCalculados
        .DestinatariosConCopia = m_SolicitudRef.DestinatariosConCopiaCalculados(p_Tipo)
        .nombrePlantilla = m_SolicitudRef.NombrePlantillaParaCorreoCalculada(p_Tipo)
        .VersionPlantilla = m_ObjEntorno.Configuracion.VersionPlantillasHTML
        .cadenaRecursos = m_SolicitudRef.ColParametrosParaHTMLTexto
        .Accion = m_SolicitudRef.AccionParaCorreoCalculada(p_Tipo)
    End With
    
    ' 3. Guardar en la cola
    Set m_CorreoOp = New CorreoOperaciones
    Set m_CorreoOp.Correo = m_Correo
    m_CorreoOp.Registrar , p_Error
    
    If p_Error <> "" Then Err.Raise 1000, , "Error programando correo tipo " & p_Tipo
    
    ' 4. LOG
    RegistrarLogGeneral "Programación Automática", "Programado correo " & m_Correo.Accion & " para el " & m_FechaProgramada, m_SolicitudRef, p_Error

    Exit Sub
errores:
    p_Error = "Error warning en ProgramarCorreoFuturo: " & Err.Description
End Sub
