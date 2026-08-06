Attribute VB_Name = "modEstadoRiesgosActivo"
' =============================================================================
' modEstadoRiesgosActivo.bas
'
' Singleton C2 del epic forms-thin-refactor-2026-06-23.
' Project: gestion_riesgos
' Block: 1B (forms-thin-phase0-testeable-2026-06-25)
' Work item: WI-4
'
' DECISION DE CACHE (CRITICA): Esta funcion NO crea segunda dictionary en
' Variables Globales.bas. Deriva el estado on-demand desde m_DicRiesgos
' (single source of truth). La razon: si hubiera una segunda dictionary,
' toda invalidacion de m_DicRiesgos (AceptacionRegistrar, RetiroRegistrar,
' etc.) tendria que recordar invalidar TAMBIEN la segunda, y el sistema
' tendria DOS fuentes de verdad sobre el estado de un riesgo. Single source
' of truth preserva la calidad del dato (calidad del dato > performance).
'
' Si en el futuro hay evidencia de que este computo es caro, agregar cache
' en un PR dedicado con su propio set de invalidacion testeado.
'
' Contexto:
'   Reemplaza la logica inter-form que calculaba el estado activo del
'   riesgo. El form sigue mostrando UI; el helper expone el estado via JSON.
'
' Reglas:
'   - Resultado como JSON contrato {ok, value, payload, error, logs} via
'     TestCore_BuildOk / TestCore_BuildFail.
'   - Convencion Telefonica D&S p_Error ByRef: vacio en exito, poblado en
'     error real. El helper retorna JSON SIEMPRE (aun en error).
'   - CONSUME CACHE via GetCachedRiesgo. NO bypass a Constructor.getRiesgo.
'   - El parametro db es inyeccion explicita (atoms -> sandbox via
'     Test_Fixtures.GetTestDb). El helper NO tiene UI ni MsgBox.
' =============================================================================
Option Compare Database
Option Explicit

' =============================================================================
' EstadoRiesgosActivo_ObtenerEstado
'
' Devuelve JSON {ok, value, payload, error, logs} con el estado activo del
' riesgo y su fecha. Deriva el estado del riesgo cacheado.
'
' Reglas:
'   - p_IDRiesgo vacio/whitespace -> TestCore_BuildFail
'   - Riesgo no encontrado en cache ni en DB -> TestCore_BuildFail
'   - Caso normal: GetCachedRiesgo -> leer Estado y FechaMaterializado del
'     riesgo, serializar como JSON, devolver TestCore_BuildOk.
' =============================================================================
Public Function EstadoRiesgosActivo_ObtenerEstado( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String _
) As String

    Dim m_Logs(0 To 5) As String
    Dim m_LogIdx As Long
    m_LogIdx = 0

    On Error GoTo errores

    p_Error = ""
    EstadoRiesgosActivo_ObtenerEstado = ""

    ' --- Validacion del ID ---
    If Len(Trim$(p_IDRiesgo)) = 0 Then
        p_Error = "EstadoRiesgosActivo_ObtenerEstado: p_IDRiesgo esta vacio"
        m_Logs(m_LogIdx) = "1. ID rejected (empty)"
        m_LogIdx = m_LogIdx + 1
        EstadoRiesgosActivo_ObtenerEstado = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    m_Logs(m_LogIdx) = "1. ID recibido: " & p_IDRiesgo
    m_LogIdx = m_LogIdx + 1

    ' --- Cache hit (consume m_DicRiesgos via GetCachedRiesgo) ---
    ' Forward db explicitly: per skill §5.4, no bypass to getdb().
    Dim m_Riesgo As Object
    Set m_Riesgo = GetCachedRiesgo(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error, db:=db)
    If p_Error <> "" Then
        EstadoRiesgosActivo_ObtenerEstado = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If
    If m_Riesgo Is Nothing Then
        p_Error = "EstadoRiesgosActivo_ObtenerEstado: GetCachedRiesgo no devolvio instancia para ID " & p_IDRiesgo
        EstadoRiesgosActivo_ObtenerEstado = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    m_Logs(m_LogIdx) = "2. Cache hit OK"
    m_LogIdx = m_LogIdx + 1

    ' --- Derivar estado y fecha del riesgo ---
    ' fechaEstado: FechaMaterializado si existe, sino FechaRetirado, sino null.
    Dim m_Estado As String
    m_Estado = ""
    On Error Resume Next
    m_Estado = CStr(m_Riesgo.Estado)
    On Error GoTo errores

    Dim m_FechaEstado As String
    m_FechaEstado = ""
    On Error Resume Next
    If Not IsNull(m_Riesgo.FechaMaterializado) And Len(CStr(m_Riesgo.FechaMaterializado)) > 0 Then
        m_FechaEstado = Format(m_Riesgo.FechaMaterializado, "yyyy-mm-dd")
    ElseIf Not IsNull(m_Riesgo.FechaRetirado) And Len(CStr(m_Riesgo.FechaRetirado)) > 0 Then
        m_FechaEstado = Format(m_Riesgo.FechaRetirado, "yyyy-mm-dd")
    Else
        m_FechaEstado = ""
    End If
    On Error GoTo errores

    ' --- Construir JSON payload: {idRiesgo, estado, fechaEstado} ---
    Dim m_Payload As String
    m_Payload = "{""idRiesgo"":""" & CStr(m_Riesgo.IDRiesgo) & """, " & _
                """estado"":""" & m_Estado & """, " & _
                """fechaEstado"":""" & m_FechaEstado & """}"

    m_Logs(m_LogIdx) = "3. Estado=" & m_Estado & ", fechaEstado=" & m_FechaEstado
    m_LogIdx = m_LogIdx + 1

    EstadoRiesgosActivo_ObtenerEstado = TestCore_BuildOk(m_Payload, m_Logs)
    Exit Function

errores:
    If p_Error = "" Then
        p_Error = "EstadoRiesgosActivo_ObtenerEstado: " & Err.Number & " - " & Err.Description
    End If
    On Error GoTo 0
    EstadoRiesgosActivo_ObtenerEstado = TestCore_BuildFail(p_Error, m_Logs)
End Function
