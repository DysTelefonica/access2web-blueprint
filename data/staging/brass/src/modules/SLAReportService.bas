Attribute VB_Name = "SLAReportService"
Option Compare Database
Option Explicit

'===========================================================
' SLAReportService
' Servicio compartido para generación de informes SLA
' Extraído de Form_FormInformeSLA para reutilización y
' futura salida HTML
'===========================================================

'-----------------------------------------------------------
' TRSS Calculation Mode Configuration
' Boolean switch to toggle between corridos (calendar) and
' laborables (business days) for TRSS elapsed time calculation.
' FALSE = días corridos (default, current business decision)
' TRUE  = días laborables (not yet implemented — switch point)
' PUBLIC so SLAHTMLService can reference it for rendering logic.
'-----------------------------------------------------------
Public Const TRSS_USA_DIAS_LABORABLES As Boolean = False

'-----------------------------------------------------------
' GetTRSSObjetivo
' Retorna el objetivo de días para TRSS según criticidad.
' Actualmente retorna días corridos (configurable).
' Switch point: cambiar TRSS_USA_DIAS_LABORABLES a True
' para migrar a días laborables sin reescribir lógica.
'
' Parámetros:
'   p_Criticidad - Valor numérico de criticidad (1, 3, 5)
'
' Retorna:
'   Long con el objetivo en días según tabla:
'   Criticidad 1 (Crítica) -> 1 día
'   Criticidad 3 (Alta)    -> 5 días
'   Criticidad 5 (Media)   -> 15 días
'   Otro                   -> 0 (desconocido)
'-----------------------------------------------------------
Public Function GetTRSSObjetivo(ByVal p_Criticidad As Long) As Long
    Select Case p_Criticidad
        Case 1: GetTRSSObjetivo = 1
        Case 3: GetTRSSObjetivo = 5
        Case 5: GetTRSSObjetivo = 15
        Case Else: GetTRSSObjetivo = 0
    End Select
End Function

'-----------------------------------------------------------
' GetTRSSDiasTranscurridos
' Calcula los días transcurridos entre recepción y restablecimiento
' usando el modo configurado (corridos o laborables).
'
' Parámetros:
'   p_FechaRecepcion        - Fecha de recepción de notificación
'   p_FechaRestablecimiento - Fecha de restablecimiento del servicio
'
' Retorna:
'   Long con días transcurridos, o -1 si fechas inválidas.
'   Modo corridos: simple DateDiff("d", start, end)
'   Modo laborables: estructurado para futura implementación
'-----------------------------------------------------------
Private Function GetTRSSDiasTranscurridos( _
    ByVal p_FechaRecepcion As Date, _
    ByVal p_FechaRestablecimiento As Date) As Long
    
    If TRSS_USA_DIAS_LABORABLES Then
        ' TODO: Implementar cálculo de días laborables
        ' Usar DateDiff con evaluación de fines de semana/festivos
        ' Estructura preparada para no reescribir la lógica principal
        GetTRSSDiasTranscurridos = DateDiff("d", p_FechaRecepcion, p_FechaRestablecimiento)
    Else
        ' Modo corridos (default actual)
        GetTRSSDiasTranscurridos = DateDiff("d", p_FechaRecepcion, p_FechaRestablecimiento)
    End If
End Function

'-----------------------------------------------------------
' ConstruirSQLEventosFranqueados
' Construye la SQL para filtrar eventos franqueados con SLA
'
' NOTA: Todos los campos SLA necesarios para el cálculo de cumplimiento
'       deben estar en el SELECT para evitar runtime error 3265
'       (Item not found in collection).
'
' Parámetros:
'   p_FechaInicio - Fecha inicio del filtro (formato dd/mm/aaaa)
'   p_FechaFin    - Fecha fin del filtro (formato dd/mm/aaaa)
'   p_Error       - Parámetro de error (opcional, por referencia)
'
' Retorna:
'   String con la consulta SQL completa
'-----------------------------------------------------------
Public Function ConstruirSQLEventosFranqueados( _
    ByVal p_FechaInicio As String, _
    ByVal p_FechaFin As String, _
    Optional ByRef p_Error As String) As String
    
    Dim arrSQL() As String
    Dim m_FechaInicio As Date
    Dim m_FechaFin As Date
    
    p_Error = ""
    
    On Error GoTo errores
    
    ' Validación de fechas
    If Not IsDate(p_FechaInicio) Then
        p_Error = "La fecha de inicio no es válida: " & p_FechaInicio
        ConstruirSQLEventosFranqueados = ""
        Exit Function
    End If
    
    If Not IsDate(p_FechaFin) Then
        p_Error = "La fecha de fin no es válida: " & p_FechaFin
        ConstruirSQLEventosFranqueados = ""
        Exit Function
    End If
    
    m_FechaInicio = CDate(p_FechaInicio)
    m_FechaFin = CDate(p_FechaFin)
    
    ' Validación de rango: fecha inicio no puede ser mayor que fecha fin
    If m_FechaInicio > m_FechaFin Then
        p_Error = "La fecha de inicio no puede ser mayor que la fecha fin."
        ConstruirSQLEventosFranqueados = ""
        Exit Function
    End If
    
    ReDim arrSQL(0 To 22)
    arrSQL(0) = "SELECT"
    arrSQL(1) = "E.IDEVENTO,"
    arrSQL(2) = "E.FECHAALTAEVENTO,"
    arrSQL(3) = "E.BUI,"
    arrSQL(4) = "E.SUBSISTEMA,"
    arrSQL(5) = "E.CRITICIDAD,"
    arrSQL(6) = "E.FechaRecepcionNotificacion,"
    arrSQL(7) = "E.FechaInicioContactoCliente,"
    arrSQL(8) = "E.IncidenciaAveriaOReparacion,"
    arrSQL(9) = "E.TipoReparacion,"
    arrSQL(10) = "E.Urgente,"
    arrSQL(11) = "E.EventoConServicioAfectado,"
    arrSQL(12) = "E.FechaRestablecimientoServicio,"
    arrSQL(13) = "E.FechaInicioTiempoAdquisicion,"
    arrSQL(14) = "E.FechaFinTiempoAdquisicion,"
    arrSQL(15) = "E.TipoRepInsitu,"
    arrSQL(16) = "E.TipoRepNoSMT,"
    arrSQL(17) = "E.TipoRepValvulas"
    arrSQL(18) = "FROM TbEventos AS E"
    arrSQL(19) = "WHERE E.Franqueado = True"
    arrSQL(20) = "AND E.FECHAALTAEVENTO >= #" & Format(m_FechaInicio, "yyyy-mm-dd") & "#"
    arrSQL(21) = "AND E.FECHAALTAEVENTO <= #" & Format(m_FechaFin, "yyyy-mm-dd") & "#"
    ' Mostrar todos los eventos franqueados en rango de fecha,
    ' aunque aún no tengan todos los campos SLA cargados.
    arrSQL(22) = "ORDER BY E.FECHAALTAEVENTO DESC, E.IDEVENTO DESC;"
    
    ConstruirSQLEventosFranqueados = Join(arrSQL, " ")
    Exit Function
    
errores:
    p_Error = "Error al construir SQL de eventos franqueados: " & Err.Number & " - " & Err.Description
    ConstruirSQLEventosFranqueados = ""
End Function

'-----------------------------------------------------------
' ObtenerEventosFranqueados
' Ejecuta la consulta y devuelve un recordset con los eventos
' Util para quien necesite el recordset directamente (ej: HTML)
'
' Parámetros:
'   p_FechaInicio - Fecha inicio del filtro
'   p_FechaFin    - Fecha fin del filtro
'   p_Recordset   - Output: recordset con los eventos (Nothing si error)
'   p_Error       - Parámetro de error (opcional, por referencia)
'
' Retorna:
'   True si OK, False si error (ver p_Error)
'-----------------------------------------------------------
Public Function ObtenerEventosFranqueados( _
    ByVal p_FechaInicio As String, _
    ByVal p_FechaFin As String, _
    ByRef p_Recordset As DAO.Recordset, _
    Optional ByRef p_Error As String) As Boolean
    
    Dim strSQL As String
    
    p_Error = ""
    Set p_Recordset = Nothing
    
    On Error GoTo errores
    
    strSQL = ConstruirSQLEventosFranqueados(p_FechaInicio, p_FechaFin, p_Error)
    If p_Error <> "" Or strSQL = "" Then
        ObtenerEventosFranqueados = False
        Exit Function
    End If
    
    Set p_Recordset = CurrentDb().OpenRecordset(strSQL)
    ObtenerEventosFranqueados = True
    Exit Function
    
errores:
    p_Error = "Error al obtener eventos franqueados: " & Err.Number & " - " & Err.Description
    Set p_Recordset = Nothing
    ObtenerEventosFranqueados = False
End Function

'===========================================================
' CalcularCumplimientoSLA
' Calcula los KPIs de cumplimiento SLA para un evento
' Extraído de SLAHTMLService.CalcularCumplimiento para
' reutilización en servicios HTML y Excel
'
' Parámetros:
'   p_Fields       - Colección de campos del registro actual
'   p_CumpleTRES   - Output: True si cumple TRES
'   p_CumpleTRCM   - Output: True si cumple TRCM
'   p_CumpleTRSS   - Output: True si cumple TRSS
'   p_CumpleGlobal - Output: True si cumple todos
'===========================================================
Public Sub CalcularCumplimientoSLA( _
    ByRef p_Fields As DAO.Fields, _
    ByRef p_CumpleTRES As Boolean, _
    ByRef p_CumpleTRCM As Boolean, _
    ByRef p_CumpleTRSS As Boolean, _
    ByRef p_CumpleGlobal As Boolean)
    
    Dim m_Incidencia As Variant
    Dim m_ServicioAfectado As Variant
    
    p_CumpleTRES = False
    p_CumpleTRCM = True
    p_CumpleTRSS = True
    p_CumpleGlobal = False
    
    On Error Resume Next
    
    ' TRES: FechaRecepcionNotificacion <= FechaInicioContactoCliente
    If IsDate(p_Fields("FechaRecepcionNotificacion")) And IsDate(p_Fields("FechaInicioContactoCliente")) Then
        p_CumpleTRES = (CDate(p_Fields("FechaInicioContactoCliente")) >= CDate(p_Fields("FechaRecepcionNotificacion")))
    End If
    
    ' TRCM: Si IncidenciaAveriaOReparacion = True, requiere fechas de adquisición y tipo de reparación
    ' Nota: TRCM tiene 3 estados (Cumple/NoCumple/Inconsistente). El output Boolean p_CumpleTRCM
    ' es True solo para "Cumple", False para "NoCumple" o "Inconsistente" (ambos tratados como no-cumple
    ' en el sentido Boolean para backwards compatibility del resumen de donuts).
    ' La distinción Inconsistente se captura en ObtenerDetalleCumplimientoEvento para el detalle.
    p_CumpleTRCM = True
    m_Incidencia = p_Fields("IncidenciaAveriaOReparacion")
    If Not IsNull(m_Incidencia) Then
        If m_Incidencia = True Then
            Dim blnTRCMFechaOK As Boolean
            Dim blnTRCMTipoOK As Boolean
            blnTRCMFechaOK = IsDate(p_Fields("FechaInicioTiempoAdquisicion")) And _
                            IsDate(p_Fields("FechaFinTiempoAdquisicion"))
            blnTRCMTipoOK = (Nz(p_Fields("TipoReparacion"), "") <> "")
            
            If blnTRCMFechaOK And blnTRCMTipoOK Then
                p_CumpleTRCM = (CDate(p_Fields("FechaFinTiempoAdquisicion")) >= CDate(p_Fields("FechaInicioTiempoAdquisicion")))
            Else
                ' Datos incompletos = Inconsistente (no "no cumple con objetivo 0")
                ' Boolean output = False para mantener backward compat del resumen
                p_CumpleTRCM = False
            End If
        End If
    End If
    
    ' TRSS: Si EventoConServicioAfectado = True, requiere FechaRestablecimiento
    ' Comparar con objetivo según criticidad (1=1d, 3=5d, 5=15d) usando modo configurado
    p_CumpleTRSS = True
    m_ServicioAfectado = p_Fields("EventoConServicioAfectado")
    If Not IsNull(m_ServicioAfectado) Then
        If m_ServicioAfectado = True Then
            If IsDate(p_Fields("FechaRestablecimientoServicio")) And IsDate(p_Fields("FechaRecepcionNotificacion")) Then
                Dim datRecepSvc As Date, datRestabSvc As Date
                Dim lngCritSvc As Long, lngObjetivoSvc As Long, lngDiasSvc As Long
                datRecepSvc = CDate(p_Fields("FechaRecepcionNotificacion"))
                datRestabSvc = CDate(p_Fields("FechaRestablecimientoServicio"))
                
                ' Solo validar cronología si FechaRestablecimiento >= FechaRecepcion
                If datRestabSvc >= datRecepSvc Then
                    lngCritSvc = Nz(p_Fields("CRITICIDAD"), 0)
                    lngObjetivoSvc = GetTRSSObjetivo(lngCritSvc)
                    lngDiasSvc = GetTRSSDiasTranscurridos(datRecepSvc, datRestabSvc)
                    
                    If lngObjetivoSvc > 0 Then
                        p_CumpleTRSS = (lngDiasSvc <= lngObjetivoSvc)
                    Else
                        ' Criticidad sin objetivo: no podemos evaluar
                        ' Boolean output = True (inconcluso, no realmente No Cumple)
                        p_CumpleTRSS = True
                    End If
                Else
                    p_CumpleTRSS = False
                End If
            Else
                p_CumpleTRSS = False
            End If
        End If
    End If
    
    p_CumpleGlobal = (p_CumpleTRES And p_CumpleTRCM And p_CumpleTRSS)
    
    On Error GoTo 0
End Sub

'-----------------------------------------------------------
' ObtenerDetalleCumplimientoEvento
' Busca el registro con IDEVENTO en el recordset ya abierto
' y retorna un Dictionary con todos los campos detallados para F3.
' Usa la MISMA lógica de cálculo que CalcularCumplimientoSLA.
'
' Parámetros:
'   p_IDEvento  - IDEVENTO a buscar (String)
'   p_rs        - Recordset ya abierto y posicionado (ByRef)
'
' Retorna:
'   Scripting.Dictionary con todos los campos para el panel F3
'   Keys: IDEVENTO, BUI, Criticidad,
'         TRES_FechaRecepcion, TRES_FechaContacto, TRES_Horas, TRES_ObjetivoHoras, TRES_Cumple,
'         TRCM_Incidencia, TRCM_TipoReparacion, TRCM_Urgente,
'         TRCM_FechaInicioAdq, TRCM_FechaFinAdq, TRCM_Dias, TRCM_ObjetivoDias, TRCM_Cumple,
'         TRSS_ServicioAfectado, TRSS_FechaRestablecimiento, TRSS_Cumple,
'         SLA4_Numerador, SLA4_Denominador, SLA4_Porcentaje,
'         Global_CumpleGlobal
'-----------------------------------------------------------
Public Function ObtenerDetalleCumplimientoEvento( _
    ByVal p_IDEvento As String, _
    ByRef p_rs As DAO.Recordset) As Object
    
    Dim dict As Object
    Dim varBookmark As Variant
    Dim blnEncontrado As Boolean
    Dim blnCumpleTRES As Boolean, blnCumpleTRCM As Boolean, blnCumpleTRSS As Boolean, blnCumpleGlobal As Boolean
    Dim lngErr As Long
    Dim strErr As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    blnEncontrado = False
    
    On Error GoTo errores
    
    ' Guardar posición actual del recordset via Bookmark (más confiable que AbsolutePosition)
    varBookmark = Null
    If Not p_rs Is Nothing Then
        On Error Resume Next
        varBookmark = p_rs.Bookmark
        If Err.Number <> 0 Then varBookmark = Null
        On Error GoTo errores
    End If
    
    ' Buscar el registro con el IDEVENTO matching
    If Not p_rs Is Nothing And Not p_rs.EOF Then
        p_rs.MoveFirst
        Do While Not p_rs.EOF
            If Nz(p_rs.Fields("IDEVENTO"), "") = p_IDEvento Then
                blnEncontrado = True
                Exit Do
            End If
            p_rs.MoveNext
        Loop
        
        If blnEncontrado Then
            '---- Campos básicos ----
            dict("IDEVENTO") = Nz(p_rs.Fields("IDEVENTO"), "")
            dict("BUI") = Nz(p_rs.Fields("BUI"), "")
            dict("Criticidad") = Nz(p_rs.Fields("CRITICIDAD"), "")
            
            '---- TRES ----
            If IsDate(p_rs.Fields("FechaRecepcionNotificacion")) Then
                dict("TRES_FechaRecepcion") = Format(p_rs.Fields("FechaRecepcionNotificacion"), "dd/mm/yyyy hh:mm")
            Else
                dict("TRES_FechaRecepcion") = ""
            End If
            
            If IsDate(p_rs.Fields("FechaInicioContactoCliente")) Then
                dict("TRES_FechaContacto") = Format(p_rs.Fields("FechaInicioContactoCliente"), "dd/mm/yyyy hh:mm")
            Else
                dict("TRES_FechaContacto") = ""
            End If
            
            ' Calcular horas TRES y objetivo según criticidad
            If IsDate(p_rs.Fields("FechaRecepcionNotificacion")) And IsDate(p_rs.Fields("FechaInicioContactoCliente")) Then
                Dim datRecep As Date, datCont As Date
                datRecep = CDate(p_rs.Fields("FechaRecepcionNotificacion"))
                datCont = CDate(p_rs.Fields("FechaInicioContactoCliente"))
                dict("TRES_Horas") = Round((datCont - datRecep) * 24, 2)
            Else
                dict("TRES_Horas") = ""
            End If
            
            ' Objetivo según criticidad
            Dim lngCrit As Long
            lngCrit = Nz(p_rs.Fields("CRITICIDAD"), 0)
            Select Case lngCrit
                Case 1: dict("TRES_ObjetivoHoras") = 1
                Case 3: dict("TRES_ObjetivoHoras") = 3
                Case 5: dict("TRES_ObjetivoHoras") = 8
                Case Else: dict("TRES_ObjetivoHoras") = 0
            End Select
            
            ' Calcular TRES Cumple (misma lógica que CalcularCumplimientoSLA)
            blnCumpleTRES = False
            If IsDate(p_rs.Fields("FechaRecepcionNotificacion")) And IsDate(p_rs.Fields("FechaInicioContactoCliente")) Then
                blnCumpleTRES = (CDate(p_rs.Fields("FechaInicioContactoCliente")) >= CDate(p_rs.Fields("FechaRecepcionNotificacion")))
            End If
            dict("TRES_Cumple") = blnCumpleTRES
            
            '---- TRCM ----
            dict("TRCM_Incidencia") = IIf(Nz(p_rs.Fields("IncidenciaAveriaOReparacion"), False) = True, "Sí", "No")
            dict("TRCM_TipoReparacion") = Nz(p_rs.Fields("TipoReparacion"), "")
            dict("TRCM_Urgente") = IIf(Nz(p_rs.Fields("Urgente"), False) = True, "Sí", "No")
            
            If IsDate(p_rs.Fields("FechaInicioTiempoAdquisicion")) Then
                dict("TRCM_FechaInicioAdq") = Format(p_rs.Fields("FechaInicioTiempoAdquisicion"), "dd/mm/yyyy")
            Else
                dict("TRCM_FechaInicioAdq") = ""
            End If
            
            If IsDate(p_rs.Fields("FechaFinTiempoAdquisicion")) Then
                dict("TRCM_FechaFinAdq") = Format(p_rs.Fields("FechaFinTiempoAdquisicion"), "dd/mm/yyyy")
            Else
                dict("TRCM_FechaFinAdq") = ""
            End If
            
            ' Calcular días TRCM y objetivo
            ' TRCM tiene 3 estados: Cumple (True), NoCumple (False), Inconsistente (dato incompleto)
            Dim strTipoRep As String, blnUrgente As Boolean
            Dim lngTRCMEstadoRaw As Long ' 1=Cumple, 2=Inconsistente, 3=NoCumple, 0=N/A
            strTipoRep = Nz(p_rs.Fields("TipoReparacion"), "")
            blnUrgente = Nz(p_rs.Fields("Urgente"), False)
            lngTRCMEstadoRaw = 0
            blnCumpleTRCM = True
            
            If Nz(p_rs.Fields("IncidenciaAveriaOReparacion"), False) = True Then
                If IsDate(p_rs.Fields("FechaInicioTiempoAdquisicion")) And IsDate(p_rs.Fields("FechaFinTiempoAdquisicion")) Then
                    If strTipoRep <> "" Then
                        Dim datInicioAdq As Date, datFinAdq As Date
                        datInicioAdq = CDate(p_rs.Fields("FechaInicioTiempoAdquisicion"))
                        datFinAdq = CDate(p_rs.Fields("FechaFinTiempoAdquisicion"))
                        dict("TRCM_Dias") = Round((datFinAdq - datInicioAdq), 2)
                        
                        ' Objetivo según tipo y urgencia
                        If strTipoRep = "Contratista" Then
                            dict("TRCM_ObjetivoDias") = IIf(blnUrgente, 7, 30)
                        ElseIf strTipoRep = "Fabricante" Then
                            dict("TRCM_ObjetivoDias") = IIf(blnUrgente, 60, 150)
                        Else
                            ' TipoReparacion no reconocido: inconsistency, not a valid target
                            lngTRCMEstadoRaw = 2 ' Inconsistente
                            dict("TRCM_ObjetivoDias") = "N/A"
                        End If
                        
                        blnCumpleTRCM = (datFinAdq >= datInicioAdq)
                        If lngTRCMEstadoRaw = 0 Then
                            lngTRCMEstadoRaw = IIf(blnCumpleTRCM, 1, 3)
                        End If
                    Else
                        ' TipoReparacion faltante = inconsistencia, no "no cumple con objetivo 0"
                        dict("TRCM_Dias") = ""
                        dict("TRCM_ObjetivoDias") = "N/A"
                        lngTRCMEstadoRaw = 2 ' Inconsistente
                        blnCumpleTRCM = False
                    End If
                Else
                    ' Faltan fechas de adquisición = inconsistencia
                    dict("TRCM_Dias") = ""
                    dict("TRCM_ObjetivoDias") = "N/A"
                    lngTRCMEstadoRaw = 2 ' Inconsistente
                    blnCumpleTRCM = False
                End If
            Else
                ' Incidencia = No: TRCM no aplica
                dict("TRCM_Dias") = "N/A"
                dict("TRCM_ObjetivoDias") = "N/A"
                lngTRCMEstadoRaw = 0
            End If
            
            ' Mapear estado raw a texto
            Select Case lngTRCMEstadoRaw
                Case 0: dict("TRCM_Estado") = "N/A"
                Case 1: dict("TRCM_Estado") = "Cumple"
                Case 2: dict("TRCM_Estado") = "Inconsistente"
                Case 3: dict("TRCM_Estado") = "No cumple"
            End Select
            
            ' Para backwards: Cumple = (Estado="Cumple")
            dict("TRCM_Cumple") = (lngTRCMEstadoRaw = 1)
            
            '---- TRSS ----
            dict("TRSS_ServicioAfectado") = IIf(Nz(p_rs.Fields("EventoConServicioAfectado"), False) = True, "Sí", "No")
            
            If IsDate(p_rs.Fields("FechaRestablecimientoServicio")) Then
                dict("TRSS_FechaRestablecimiento") = Format(p_rs.Fields("FechaRestablecimientoServicio"), "dd/mm/yyyy")
            Else
                dict("TRSS_FechaRestablecimiento") = ""
            End If
            
            ' Calcular estado TRSS con objetivo por criticidad y modo configurado
            Dim lngTRSSCrit As Long
            Dim lngTRSSObjetivo As Long
            Dim lngTRSSDias As Long
            Dim lngTRSSEstadoRaw As Long ' 1=Cumple, 2=Inconsistente, 3=NoCumple, 0=N/A
            
            lngTRSSEstadoRaw = 0
            lngTRSSDias = 0
            lngTRSSObjetivo = 0
            
            If Nz(p_rs.Fields("EventoConServicioAfectado"), False) = True Then
                If IsDate(p_rs.Fields("FechaRestablecimientoServicio")) And IsDate(p_rs.Fields("FechaRecepcionNotificacion")) Then
                    ' Datos completos: calcular días y comparar con objetivo
                    Dim datRecepTRSS As Date, datRestabTRSS As Date
                    datRecepTRSS = CDate(p_rs.Fields("FechaRecepcionNotificacion"))
                    datRestabTRSS = CDate(p_rs.Fields("FechaRestablecimientoServicio"))
                    
                    ' Obtener objetivo según criticidad
                    lngTRSSCrit = Nz(p_rs.Fields("CRITICIDAD"), 0)
                    lngTRSSObjetivo = GetTRSSObjetivo(lngTRSSCrit)
                    lngTRSSDias = GetTRSSDiasTranscurridos(datRecepTRSS, datRestabTRSS)
                    
                    ' Determinar estado: consistencia cronológica + comparación vs objetivo
                    blnCumpleTRSS = (datRestabTRSS >= datRecepTRSS)
                    If blnCumpleTRSS Then
                        If lngTRSSObjetivo > 0 Then
                            blnCumpleTRSS = (lngTRSSDias <= lngTRSSObjetivo)
                            If blnCumpleTRSS Then
                                lngTRSSEstadoRaw = 1 ' Cumple
                            Else
                                lngTRSSEstadoRaw = 3 ' No cumple
                            End If
                        Else
                            ' Criticidad sin objetivo definido
                            lngTRSSEstadoRaw = 2 ' Inconsistente (objetivo desconocido)
                        End If
                    Else
                        lngTRSSEstadoRaw = 2 ' Inconsistente (fechas inválidas)
                    End If
                Else
                    ' Servicio afectado pero sin fechas: Inconsistente
                    lngTRSSEstadoRaw = 2
                    blnCumpleTRSS = False
                End If
            Else
                ' No aplica TRSS
                lngTRSSEstadoRaw = 0
                blnCumpleTRSS = True
            End If
            
            ' Guardar campos TRSS expandidos
            If lngTRSSEstadoRaw = 0 Then
                dict("TRSS_Estado") = "N/A"
                dict("TRSS_ObjetivoDias") = "N/A"
                dict("TRSS_DiasTranscurridos") = "N/A"
            ElseIf lngTRSSEstadoRaw = 2 Then
                dict("TRSS_Estado") = "Inconsistente"
                dict("TRSS_ObjetivoDias") = "N/A"
                dict("TRSS_DiasTranscurridos") = IIf(lngTRSSDias > 0, CStr(lngTRSSDias), "N/A")
            Else
                dict("TRSS_Estado") = IIf(lngTRSSEstadoRaw = 1, "Cumple", "No cumple")
                dict("TRSS_ObjetivoDias") = CStr(lngTRSSObjetivo)
                dict("TRSS_DiasTranscurridos") = CStr(lngTRSSDias)
            End If
            dict("TRSS_Cumple") = (lngTRSSEstadoRaw = 1)
            
            '---- SLA-4 ----
            Dim blnTipoRepInsitu As Boolean, blnTipoRepNoSMT As Boolean, blnTipoRepValvulas As Boolean
            blnTipoRepInsitu = Nz(p_rs.Fields("TipoRepInsitu"), False)
            blnTipoRepNoSMT = Nz(p_rs.Fields("TipoRepNoSMT"), False)
            blnTipoRepValvulas = Nz(p_rs.Fields("TipoRepValvulas"), False)
            
            If blnTipoRepInsitu = True And (blnTipoRepNoSMT = True Or blnTipoRepValvulas = True) Then
                dict("SLA4_Numerador") = 1
            Else
                dict("SLA4_Numerador") = 0
            End If
            
            If blnTipoRepNoSMT = True Or blnTipoRepValvulas = True Then
                dict("SLA4_Denominador") = 1
            Else
                dict("SLA4_Denominador") = 0
            End If
            
            If dict("SLA4_Denominador") > 0 Then
                dict("SLA4_Porcentaje") = Round((dict("SLA4_Numerador") / dict("SLA4_Denominador")) * 100, 1)
            Else
                dict("SLA4_Porcentaje") = 0
            End If
            
            '---- Global ----
            blnCumpleGlobal = (blnCumpleTRES And blnCumpleTRCM And blnCumpleTRSS)
            dict("Global_CumpleGlobal") = blnCumpleGlobal
        End If
    End If
    
    ' Restaurar posición original del recordset via Bookmark
    If Not p_rs Is Nothing And Not p_rs.EOF Then
        If Not IsNull(varBookmark) Then
            On Error Resume Next
            p_rs.Bookmark = varBookmark
            If Err.Number <> 0 Then
                ' Bookmark inválido: intentar volver al inicio
                On Error GoTo 0
                On Error Resume Next
                p_rs.MoveFirst
            End If
            On Error GoTo 0
        End If
    End If
    
    Set ObtenerDetalleCumplimientoEvento = dict
    Exit Function
    
errores:
    lngErr = Err.Number
    strErr = Err.Description

    ' En caso de error, retornar diccionario con detalle (no vacío — se preservan los campos básicos)
    On Error Resume Next
    If Not p_rs Is Nothing And Not p_rs.EOF Then
        If Not IsNull(varBookmark) Then
            p_rs.Bookmark = varBookmark
        End If
    End If
    On Error GoTo 0
    Set dict = CreateObject("Scripting.Dictionary")
    dict("ERROR") = lngErr & " - " & strErr
    Set ObtenerDetalleCumplimientoEvento = dict
End Function


