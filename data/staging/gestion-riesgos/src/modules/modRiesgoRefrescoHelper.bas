Attribute VB_Name = "modRiesgoRefrescoHelper"
' =============================================================================
' modRiesgoRefrescoHelper.bas
'
' Helper para refresco del arbol de riesgos (categoria C1 del epic
' forms-thin-refactor-2026-06-23).
' Project: gestion_riesgos
' Block: 1B (forms-thin-phase0-testeable-2026-06-25)
' Work item: WI-3
'
' Exports con prefijo "Refresco_" (alineado con "DetalleRefresco_" del
' Block 1A gemelo en modRiesgoDetalleRefrescoHelper.bas).
'
' Contexto:
'   Reemplaza la logica inter-form de Form_FormRiesgosGestion.
'   RefrescarArbolRiesgosScope y Form_FormRiesgosGestion.
'   RefrescarNodoRiesgoActual por un helper testeable. El form sigue
'   haciendo la parte UI (treeview); el helper expone data (lista de IDs
'   / riesgo) que el form consume para reconstruir el arbol.
'
' Reglas:
'   - Resultado como JSON contrato {ok, value, payload, error, logs} via
'     TestCore_BuildOk / TestCore_BuildFail (TestCore_EscapeJsonString para
'     strings seguras).
'   - Convencion Telefonica D&S p_Error ByRef: vacio en exito, poblado en
'     error real. El helper retorna JSON SIEMPRE (aun en error).
'   - CONSUME CACHE: el helper del refresco de nodo usa GetCachedRiesgo
'     (NO GetCachedRiesgoFresh). Razon: el callsite canonico es
'     modRiesgoDetalleRefrescoHelper.bas:78, que ya invalida entre los dos
'     GetCachedRiesgo para forzar instancia distinta. Llamar Fresh aca
'     seria invalidacion doble innecesaria.
'   - Helper de scope del arbol: queries TbRiesgos via SQL directo con
'     el filtro del caller. NO cache de resultados de filtro (cada llamada
'     es deterministica y barata con el sandbox).
'   - El parametro db es inyeccion explicita (atoms pasan sandbox via
'     Test_Fixtures.GetTestDb); helper no tiene UI ni MsgBox.
' =============================================================================
Option Compare Database
Option Explicit

' --- Constantes ---
Private Const SCOPE_JSON_DELIM As String = ","

' ============================================================
' Refresco_RefrescarArbolRiesgosScope
'
' Devuelve JSON {ok, value, payload, error, logs} con un array de IDs
' de riesgos que cumplen el filtro SQL del caller (p_Filtro).
'
' Reglas:
'   - p_Filtro vacio/whitespace -> TestCore_BuildFail
'   - p_Filtro valido -> ejecuta SQL en TbRiesgos con WHERE p_Filtro,
'     serializa IDs como JSON array, devuelve TestCore_BuildOk con value
'     siendo el array serializado.
'   - NO bypass: usa db invectado (atoms -> sandbox).
'   - NO usa GetCachedRiesgo (el filtro es dinamico; el cache es por IDRiesgo).
' ============================================================
Public Function Refresco_RefrescarArbolRiesgosScope( _
    ByVal p_Filtro As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String _
) As String

    Dim m_Logs(0 To 5) As String
    Dim m_LogIdx As Long
    m_LogIdx = 0

    On Error GoTo errores

    p_Error = ""
    Refresco_RefrescarArbolRiesgosScope = ""

    ' --- Validacion explicita del filtro ---
    If Len(Trim$(p_Filtro)) = 0 Then
        p_Error = "Refresco_RefrescarArbolRiesgosScope: p_Filtro esta vacio"
        m_Logs(m_LogIdx) = "2. Filter rejected (empty)"
        m_LogIdx = m_LogIdx + 1
        Refresco_RefrescarArbolRiesgosScope = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    m_Logs(m_LogIdx) = "1. Filtro recibido: " & p_Filtro
    m_LogIdx = m_LogIdx + 1

    ' --- Resolver db: si Nothing, usar CurrentDb (modo dual-test/prod) ---
    Dim m_Db As DAO.Database
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- Construir SQL: SELECT IDRiesgo FROM TbRiesgos WHERE <filtro> ---
    Dim m_SQL As String
    m_SQL = "SELECT IDRiesgo FROM TbRiesgos WHERE " & p_Filtro

    m_Logs(m_LogIdx) = "2. SQL: " & m_SQL
    m_LogIdx = m_LogIdx + 1

    ' --- Ejecutar query y serializar IDs a JSON array ---
    Dim m_RS As DAO.Recordset
    Set m_RS = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)

    Dim m_IDs As String
    m_IDs = ""
    Dim m_Count As Long
    m_Count = 0
    Do While Not m_RS.EOF
        If Len(m_IDs) > 0 Then m_IDs = m_IDs & SCOPE_JSON_DELIM
        m_IDs = m_IDs & CStr(m_RS.Fields("IDRiesgo").value)
        m_Count = m_Count + 1
        m_RS.MoveNext
    Loop
    On Error Resume Next
    m_RS.Close
    Set m_RS = Nothing
    On Error GoTo errores

    Dim m_ArrayJson As String
    m_ArrayJson = "[" & m_IDs & "]"

    m_Logs(m_LogIdx) = "3. IDs encontrados: " & m_Count
    m_LogIdx = m_LogIdx + 1

    Refresco_RefrescarArbolRiesgosScope = TestCore_BuildOk(m_ArrayJson, m_Logs)
    Exit Function

errores:
    If p_Error = "" Then
        p_Error = "Refresco_RefrescarArbolRiesgosScope: " & Err.Number & " - " & Err.Description
    End If
    On Error Resume Next
    If Not m_RS Is Nothing Then
        m_RS.Close
        Set m_RS = Nothing
    End If
    On Error GoTo 0
    Refresco_RefrescarArbolRiesgosScope = TestCore_BuildFail(p_Error, m_Logs)
End Function

' ============================================================
' Refresco_RefrescarNodoRiesgoActual
'
' Refresca el riesgo activo (cache hit) y lo expone via ByRef.
' Decisión: usa GetCachedRiesgo (NO GetCachedRiesgoFresh) — ver WI-3
' Pregunta Clave del spec y diseño §2.3 D3.
'
' Reglas:
'   - p_IDRiesgo vacio/whitespace -> TestCore_BuildFail
'   - p_IDRiesgo valido -> GetCachedRiesgo, expone instancia via ByRef,
'     JSON TestCore_BuildOk con id canonico.
'   - Si GetCachedRiesgo retorna Nothing -> TestCore_BuildFail con error.
' ============================================================
Public Function Refresco_RefrescarNodoRiesgoActual( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_ObjRiesgoActivo As Object, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String _
) As String

    Dim m_Logs(0 To 5) As String
    Dim m_LogIdx As Long
    m_LogIdx = 0

    On Error GoTo errores

    p_Error = ""
    Refresco_RefrescarNodoRiesgoActual = ""
    Set p_ObjRiesgoActivo = Nothing

    ' --- Validacion del ID ---
    If Len(Trim$(p_IDRiesgo)) = 0 Then
        p_Error = "Refresco_RefrescarNodoRiesgoActual: p_IDRiesgo esta vacio"
        m_Logs(m_LogIdx) = "2. ID rejected (empty)"
        m_LogIdx = m_LogIdx + 1
        Refresco_RefrescarNodoRiesgoActual = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    m_Logs(m_LogIdx) = "1. ID recibido: " & p_IDRiesgo
    m_LogIdx = m_LogIdx + 1

    ' --- Cache hit (consume m_DicRiesgos via GetCachedRiesgo) ---
    ' Forward db explicitly: GetCachedRiesgo's db parameter must propagate to
    ' Constructor.getRiesgo on cache miss. Without this, db is decorative.
    Set p_ObjRiesgoActivo = GetCachedRiesgo(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error, db:=db)
    If p_Error <> "" Then
        Refresco_RefrescarNodoRiesgoActual = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    If p_ObjRiesgoActivo Is Nothing Then
        p_Error = "Refresco_RefrescarNodoRiesgoActual: GetCachedRiesgo no devolvio instancia para ID " & p_IDRiesgo
        Refresco_RefrescarNodoRiesgoActual = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    m_Logs(m_LogIdx) = "2. Cache hit OK (ID canonico=" & CStr(p_ObjRiesgoActivo.IDRiesgo) & ")"
    m_LogIdx = m_LogIdx + 1

    ' --- JSON con id canonico (NO el input crudo, que pudo tener whitespace) ---
    Dim m_Result As Scripting.Dictionary
    Set m_Result = New Scripting.Dictionary
    m_Result("idRiesgo") = CStr(p_ObjRiesgoActivo.IDRiesgo)
    m_Result("codigoUnico") = CStr(p_ObjRiesgoActivo.CodigoUnico)
    m_Result("estado") = CStr(p_ObjRiesgoActivo.Estado)

    Dim m_Payload As String
    m_Payload = JsonConverter.ConvertToJson(m_Result, 0)

    Refresco_RefrescarNodoRiesgoActual = TestCore_BuildOk(m_Payload, m_Logs)
    Exit Function

errores:
    If p_Error = "" Then
        p_Error = "Refresco_RefrescarNodoRiesgoActual: " & Err.Number & " - " & Err.Description
    End If
    On Error GoTo 0
    Refresco_RefrescarNodoRiesgoActual = TestCore_BuildFail(p_Error, m_Logs)
End Function