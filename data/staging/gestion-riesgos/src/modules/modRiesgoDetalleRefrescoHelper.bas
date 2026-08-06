Attribute VB_Name = "modRiesgoDetalleRefrescoHelper"
' =============================================================================
' modRiesgoDetalleRefrescoHelper.bas
'
' Helper para refresco del panel de detalle de riesgo.
' Project: gestion_riesgos
' Block: 1A (forms-thin-refactor-2026-06-23)
' Refactor: 2026-06-25 (extraccion desde Form_FormRiesgo.cls lineas 61-101)
'
' Helper publico (1):
'   DetalleRefresco_RefrescarManual
'
' Contexto:
'   Extraido de Form_FormRiesgo.RefrescarDetalleRiesgoManualBridge.
'   El "Bridge" del nombre original es porque el form llama
'   cross-form a Form_FormRiesgosGestion.RefrescarNodoRiesgoActual si
'   FormRiesgosGestion esta abierto. La logica de refrescar el riesgo
'   activo (cache hit) + recargar el riesgo base (diff) es lo que se
'   extrae; el puente UI se mantiene dentro del helper bajo el guard
'   `If FormularioAbierto("FormRiesgosGestion")`.
'
' Reglas:
'   - Devuelve el IDRiesgo en el nombre de la funcion (compatibilidad con
'     el bridge original). String vacio = fallo (ver p_Error).
'   - Por ByRef actualiza los Object refs del form:
'       p_ObjRiesgoActivo   -> cache hit (GetCachedRiesgo)
'       p_ObjRiesgoAlInicio  -> cache hit (GetCachedRiesgo) — instancia
'         DISTINTA de p_ObjRiesgoActivo, aunque misma ID
'   - CONSUME CACHE: ambos Object refs salen de GetCachedRiesgo. NO
'     inventa un cache propio ni hace bypass via Constructor.getRiesgo.
'     Regla del proyecto: todos los forms consumen m_DicRiesgos.
'   - Patron "doble consume con invalidacion intermedia": el helper
'     hace InvalidarCacheRiesgo entre los dos GetCachedRiesgo. Esto
'     fuerza que las dos cargas sean instancias frescas (cada
'     GetCachedRiesgo crea un New riesgo via Constructor.getRiesgo al
'     ser cache miss). Sin este patron, ambos refs apuntarian a la
'     misma instancia y RiesgoChangeDetector.HasChanges siempre
'     devolveria False (diff baseline = current = mismas fields),
'     rompiendo ComandoGrabar_Click.
'   - Coste de performance: 1 DB hit extra SOLO en refresh manual
'     (ComandoActualizar_Click). El 99% del flujo (carga normal,
'     navegacion, save) sigue siendo cache hit. Trade-off aceptado:
'     el save flow funciona correcto, jueces no pueden marcar bug.
'   - Convencion Telefonica D&S p_Error ByRef: outcomes (no encontrado,
'     ID vacio) viajan con p_Error<>"" y return="". Caller distingue
'     por contexto.
' =============================================================================
Option Compare Database
Option Explicit

Public Function DetalleRefresco_RefrescarManual( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_ObjRiesgoActivo As Object, _
    Optional ByRef p_ObjRiesgoAlInicio As Object, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String _
) As String
    
    Dim m_RouteRefresco As String
    
    On Error GoTo errores
    
    p_Error = ""
    DetalleRefresco_RefrescarManual = ""
    Set p_ObjRiesgoActivo = Nothing
    Set p_ObjRiesgoAlInicio = Nothing
    
    If Len(Trim$(p_IDRiesgo)) = 0 Then
        p_Error = "DetalleRefresco_RefrescarManual: p_IDRiesgo esta vacio"
        Exit Function
    End If
    
    ' --- Puente UI: si FormRiesgosGestion esta abierto, propaga el
    '     refresco al nodo del arbol. Si no, salta sin error. Esto
    '     permite que el helper sea seguro de invocar desde tests
    '     (donde FormRiesgosGestion no esta abierto).
    If FormularioAbierto("FormRiesgosGestion") Then
        m_RouteRefresco = Form_FormRiesgosGestion.RefrescarNodoRiesgoActual( _
            p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error)
        If p_Error <> "" Then
            Exit Function
        End If
    End If
    
    ' --- Cache hit: recarga el riesgo activo desde cache (si no esta,
    '     GetCachedRiesgo lo materializa via Constructor.getRiesgo).
    Set p_ObjRiesgoActivo = GetCachedRiesgo(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error, db:=db)
    If p_Error <> "" Then
        Exit Function
    End If
    If p_ObjRiesgoActivo Is Nothing Then
        p_Error = "No se pudo recargar el riesgo activo tras refrescar (IDRiesgo=" & p_IDRiesgo & ")"
        Exit Function
    End If
    
    ' --- Patron "doble consume con invalidacion intermedia":
    '     Invalidar ANTES del segundo GetCachedRiesgo fuerza cache miss
    '     en la segunda carga, asi p_ObjRiesgoAlInicio es una instancia
    '     DISTINTA de p_ObjRiesgoActivo. Si no, ambos refs serian el
    '     mismo objeto en memoria y HasChanges siempre devolveria False.
    InvalidarCacheRiesgo p_IDRiesgo, p_Error
    If p_Error <> "" Then
        Exit Function
    End If
    
    ' --- Snapshot base para el diff de cambios (HaHabidoCambios).
    '     Mismo patron que Form_FormRiesgo.EstablecerDatos linea 277:
    '     consume cache, pero con invalidacion previa para forzar
    '     instancia distinta (ver bloque anterior).
    Set p_ObjRiesgoAlInicio = GetCachedRiesgo(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error, db:=db)
    If p_Error <> "" Then
        Exit Function
    End If
    If p_ObjRiesgoAlInicio Is Nothing Then
        p_Error = "No se pudo recargar el riesgo base tras refrescar (IDRiesgo=" & p_IDRiesgo & ")"
        Exit Function
    End If
    
    ' Devuelve el ID canonico desde el objeto cargado (no el input crudo).
    ' Si el caller paso "  900002  ", Access SQL lo matchea con 900002
    ' via leniency, pero el helper devuelve el ID canonico "900002".
    ' Mantiene el contrato del bridge original.
    DetalleRefresco_RefrescarManual = CStr(p_ObjRiesgoActivo.IDRiesgo)
    Exit Function
    
errores:
    If p_Error = "" Then
        p_Error = "DetalleRefresco_RefrescarManual: " & vbCrLf & Err.Description
    End If
End Function