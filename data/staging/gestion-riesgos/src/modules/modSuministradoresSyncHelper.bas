Attribute VB_Name = "modSuministradoresSyncHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modSuministradoresSyncHelper.bas
'
' Helper module para sincronizacion de suministradores entre jerarquia y edicion.
' Wrapper testable sobre SuministradoresHelper.SincronizarSuministradoresEnEdicion.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor: 2026-06-19
'
' Helper publico (1):
'   SincronizarSuministradoresEnEdicion (Sub) - sincroniza ColSuministradoresFaltan/Sobran
'
' Tablas:
'   Destino: TbProyectosEdicionesSuministradores (PK ID, IDEdicion, IDSuministrador, IDAnexo)
'   Fuente: TbExpedientesSuministradores (jerarquia, lazy via Edicion.ColSuministradoresFaltan)
'   Maestro: TbSuministradores
' =============================================================================

'------------------------------------------------------------------------------
' SincronizarSuministradoresEnEdicion
'
' Sincroniza los subcontratistas de la jerarquia del expediente con la tabla
' TbProyectosEdicionesSuministradores de una edicion.
'
' Paso 1: si p_ObjEdicion es Nothing, resuelve via Constructor.getEdicion.
' Paso 2: si la edicion resuelta es Nothing -> p_Error + Exit Sub.
' Paso 3: acceder a ColSuministradoresSobran y ColSuministradoresFaltan (lazy).
' Paso 4: para cada ID en Sobran -> EliminarSuministradorEvidenciaEnEdicion.
' Paso 5: para cada ID en Faltan -> AltaSuministradorEvidenciaEnEdicion.
' Paso 6: vaciar colecciones.
' Paso 7: propagar error via Err.Raise 1000.
'
' Parametros:
'   p_IDEdicion    (ByVal String)  - ID de la edicion a sincronizar
'   p_ObjEdicion   (ByVal Edicion) - objeto Edicion ya resuelto (Nothing = resolver)
'   p_Error        (Optional ByRef) - cadena de error, vazia si OK
'------------------------------------------------------------------------------
Public Sub SincronizarSuministradoresParaEdicion( _
    ByVal p_IDEdicion As String, _
    ByVal p_ObjEdicion As Edicion, _
    Optional ByRef p_Error As String)

    Dim m_Edicion As Edicion

    On Error GoTo errores
    p_Error = ""

    ' Paso 1: resolver edicion si no fue inyectada
    If p_ObjEdicion Is Nothing Then
        Set m_Edicion = Constructor.getEdicion(p_IDEdicion:=p_IDEdicion, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set m_Edicion = p_ObjEdicion
    End If

    ' Paso 2: validar que la edicion exista
    If m_Edicion Is Nothing Then
        p_Error = "No se ha podido determinar la edicion"
        Err.Raise 1000
    End If

    ' Pasos 3-6: delegar al helper existente
    ' SuministradoresHelper resuelve ColSuministradoresFaltan/Sobran via lazy loading
    ' y ejecuta Alta/Eliminar para cada ID de las colecciones.
    SuministradoresHelper.SincronizarSuministradoresEnEdicion _
        p_IDEdicion:=m_Edicion.IDEdicion, _
        p_Edicion:=m_Edicion, _
        p_Error:=p_Error

    If p_Error <> "" Then
        Err.Raise 1000
    End If

    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "SincronizarSuministradoresParaEdicion: " & Err.description
    End If
End Sub

' =============================================================================
' Helper: ListarEvidenciasPendientesEnInforme
'
' Devuelve un resumen textual de los subcontratistas con evidencia pendiente
' para una edición dada. Diseñado para alimentar el panel "informe de calidad"
' (ver Bloque 3 / REQ-CAL-12). El helper NO bloquea publicación — solo
' reporta; el caller (form o flow superior, p.ej. modPublicacionCalidadExecutionHelper)
' decide qué hacer con la señal "PUBLICACIÓN BLOQUEADA".
'
' Forma del retorno (string multilínea):
'   - "Sin pendientes" si todos los subcontratistas tienen IDAnexo no nulo.
'   - "PUBLICACIÓN BLOQUEADA: <detalle>" si hay al menos un subcontratista
'     con evidencia pendiente. El detalle lista uno por línea:
'         "  - <Nombre> (CIF <CIF>): <N> riesgo(s) sin evidencia"
'     Si CIF está vacío para un subcontratista, se sustituye por el placeholder
'     "(sin CIF)" — el caller lo muestra tal cual al técnico de Calidad.
'   - "" si p_Error está poblado (validación o DAO falló).
'
' Decisión de fuente (2026-06-22, sdd-apply Bloque 3 / PR-4):
'   * Tablas: TbProyectosEdicionesSuministradores (PK ID, IDEdicion,
'     IDSuministrador, IDAnexo Variant — null = sin evidencia)
'     JOIN TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio).
'   * Heurística de "evidencia pendiente" (documentada en el helper):
'       IDAnexo IS NULL OR Nz(IDAnexo, 0) = 0
'     IDAnexo es Variant en el schema (puede ser Long con valor 0 o Null);
'     la heurística acepta ambas formas como "pendiente". Si Calidad decide
'     en el futuro cambiar el contrato (p.ej. IDAnexo = -1 explícito), este
'     helper se ajusta sin tocar la firma.
'   * NIT/CIF: el schema de TbSuministradores tiene la columna "CIF"
'     (Código de Identificación Fiscal — convención española equivalente al
'     NIT colombiano). El placeholder "(sin CIF)" cubre el caso Null.
'   * Conteo de riesgos sin evidencia: COUNT(*) sobre las filas del subcontratista
'     en TbProyectosEdicionesSuministradores donde IDAnexo es pendiente.
'     El término "riesgos" aquí es un poco laxo (la tabla es subcontratista ×
'     edición, no subcontratista × riesgo); si Calidad necesita conteo por
'     riesgo real, se ajusta en una iteración posterior.
'
' Parámetros:
'   p_IDEdicion      - ID de la edición (String por convención del proyecto).
'   db               - DAO.Database opcional (inyectado por el átomo TDD).
'   p_PromptResult   - Long, ByRef. Reservado para futura migración del MsgBox
'                      que el form pueda mostrar antes de listar. No se usa en
'                      este helper (la consulta es no interactiva).
'   p_Error (out)    - String, ByRef. Cadena de error si falla validación o
'                      DAO. Cadena vacía si OK.
'
' Retorna:
'   "Sin pendientes" si no hay pendientes.
'   "PUBLICACIÓN BLOQUEADA: ..." si hay pendientes.
'   "" si p_Error está poblado.
' =============================================================================
Public Function ListarEvidenciasPendientesEnInforme( _
    ByRef p_IDEdicion As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long, _
    Optional ByRef p_Error As String) As String

    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_SQL As String
    Dim m_EdicionObj As Edicion
    Dim m_LngEdicion As Long

    Dim m_Nombre As String
    Dim m_CIF As String
    Dim m_Cantidad As Long
    Dim m_Lineas As String
    Dim m_FirstRow As Boolean

    Const PLACEHOLDER_SIN_CIF As String = "(sin CIF)"
    Const PLACEHOLDER_SIN_NOMBRE As String = "(sin nombre)"
    Const MENSAJE_SIN_PENDIENTES As String = "Sin pendientes"
    Const PREFIJO_BLOQUEO As String = "PUBLICACIÓN BLOQUEADA"

    On Error GoTo errores
    p_Error = ""
    ListarEvidenciasPendientesEnInforme = ""

    ' --- 1. Validación: p_IDEdicion no vacío ---
    If Len(Trim$(Nz(p_IDEdicion, ""))) = 0 Then
        p_Error = "ListarEvidenciasPendientesEnInforme: p_IDEdicion está vacío"
        Exit Function
    End If
    If Not IsNumeric(p_IDEdicion) Then
        p_Error = "ListarEvidenciasPendientesEnInforme: p_IDEdicion no es numérico"
        Exit Function
    End If
    m_LngEdicion = CLng(p_IDEdicion)

    ' --- 2. Resolver db (inyectado por el átomo, o CurrentDb en producción) ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 3. Resolver edición vía Constructor.getEdicion (valida FK) ---
    '     Si la edición no existe, Constructor devuelve Nothing y rellena p_Error.
    Set m_EdicionObj = Constructor.getEdicion(p_IDEdicion:=CStr(m_LngEdicion), p_Error:=p_Error)
    If p_Error <> "" Then
        Set m_Db = Nothing
        Exit Function
    End If
    If m_EdicionObj Is Nothing Then
        Set m_Db = Nothing
        p_Error = "ListarEvidenciasPendientesEnInforme: no se encontró la edición"
        Exit Function
    End If

    ' --- 4. Query: subcontratistas con evidencia pendiente ---
    '     SQL documentado (ver header del helper):
    '       INNER JOIN porque sólo nos interesan subcontratistas que tienen al
    '       menos una fila en TbProyectosEdicionesSuministradores (huérfanos no
    '       aparecen en el informe). Filtro (IDAnexo IS NULL OR Nz(IDAnexo, 0) = 0)
    '       detecta evidencia pendiente.
    m_SQL = "SELECT s.Nombre, s.CIF, COUNT(*) AS Pendientes " & _
            "FROM TbProyectosEdicionesSuministradores pes " & _
            "INNER JOIN TbSuministradores s " & _
            "  ON s.IDSuministrador = pes.IDSuministrador " & _
            "WHERE pes.IDEdicion = " & m_LngEdicion & " " & _
            "  AND (pes.IDAnexo IS NULL OR Nz(pes.IDAnexo, 0) = 0) " & _
            "GROUP BY s.Nombre, s.CIF " & _
            "ORDER BY s.Nombre"

    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)

    m_Lineas = ""
    m_FirstRow = True
    Do While Not m_Rs.EOF
        m_Nombre = Trim$(CStr(Nz(m_Rs.fields("Nombre").value, "")))
        m_CIF = Trim$(CStr(Nz(m_Rs.fields("CIF").value, "")))
        m_Cantidad = CLng(Nz(m_Rs.fields("Pendientes").value, 0))

        ' Apply HTMLSafe to user-controlled fields (subcontratista name, CIF)
        ' para que el caller pueda pintar el string sin inyección de tags.
        m_Nombre = HTMLSafe(m_Nombre)
        m_CIF = HTMLSafe(m_CIF)
        If Len(m_Nombre) = 0 Then m_Nombre = PLACEHOLDER_SIN_NOMBRE

        ' CRLF separator entre filas (NO prepender CRLF en la primera fila).
        If Not m_FirstRow Then m_Lineas = m_Lineas & vbCrLf

        ' Formato condicional: cuando CIF está vacío, el placeholder "(sin CIF)"
        ' ya implica la ausencia del identificador — NO prepender "CIF " (eso
        ' produciría "CIF (sin CIF)", que es ruido visual en el report).
        If Len(m_CIF) = 0 Then
            m_Lineas = m_Lineas & "  - " & m_Nombre & " " & PLACEHOLDER_SIN_CIF & _
                        ": " & CStr(m_Cantidad) & " riesgo(s) sin evidencia"
        Else
            m_Lineas = m_Lineas & "  - " & m_Nombre & " (CIF " & m_CIF & _
                        "): " & CStr(m_Cantidad) & " riesgo(s) sin evidencia"
        End If
        m_FirstRow = False

        m_Rs.MoveNext
    Loop
    m_Rs.Close
    Set m_Rs = Nothing

    ' --- 5. Build return string ---
    If Len(m_Lineas) = 0 Then
        ListarEvidenciasPendientesEnInforme = MENSAJE_SIN_PENDIENTES
    Else
        ListarEvidenciasPendientesEnInforme = PREFIJO_BLOQUEO & _
            ": hay subcontratista(s) con evidencia pendiente en la edición " & _
            CStr(m_LngEdicion) & ":" & vbCrLf & m_Lineas
    End If

    Set m_Db = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ListarEvidenciasPendientesEnInforme: " & Err.Number & " - " & Err.description
    End If
    ListarEvidenciasPendientesEnInforme = ""
End Function

