Attribute VB_Name = "modEdicionFechasPresenter"
Option Compare Database
Option Explicit

' =============================================================================
' modEdicionFechasPresenter.bas
'
' Helper module "presenter" para resolver la etiqueta visible de los campos
' de fecha del formulario de detalle de edición y del histórico de ediciones.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-08
'
' Helper público (2):
'   ObtenerEtiquetaFecha     - devuelve la etiqueta legible (en español) que
'                              el form y los informes deben mostrar junto al
'                              valor del campo de fecha, considerando el
'                              estado de la edición.
'   ObtenerFechaBaseProyecto - devuelve MIN(FechaEdicion) de
'                              TbProyectosEdiciones filtrado por IDProyecto,
'                              usado como constraint (b) del issue #90
'                              (Punto 04 parte 2). Cache-first: el form
'                              resuelve este valor UNA SOLA VEZ en Form_Load
'                              y lo guarda en m_FechaBaseProyecto.
'
' Decisión 2026-06-22 (sdd-apply Bloque 4 / PR-7):
'   * Helper semi-puro: la etiqueta es función del (p_NombreCampo,
'     p_IDEdicion) más el estado persistido en TbProyectosEdiciones. Se
'     incluye `Optional ByRef db As DAO.Database = Nothing` por convención
'     del proyecto (skill access-vba-e2e-methodology §hard rule 2):
'     el helper lee desde DAO, así que la inyección es REQUERIDA (no
'     opcional) — quitar el parámetro cambiaría el output.
'   * p_IDEdicion es String por convención del proyecto (los ID* son
'     String, NO Long — esto difiere del design §2.2 que usaba Long,
'     corregido per project rules verificadas por el orchestrator).
'   * El helper NO modifica el form; solo devuelve el string de la
'     etiqueta. El caller decide cómo pintarlo (label control, caption,
'     Debug.Print de defensa, etc.).
'   * Si p_Error está poblado, el retorno es "" (regla del proyecto:
'     JSON body lleva datos válidos, error viaja por p_Error ByRef).
'   * DAO opcional: si db=Nothing usa CurrentDb; los átomos TDD inyectan
'     el sandbox vía Test_Fixtures.GetTestDb().
'   * Composición basada en p_NombreCampo + estado de la edición.
'
' Regla de etiquetas (acordada con Calidad 2026-07-08 — Fase A scope):
'   - FechaPreparadaParaPublicar     ? "Fecha propuesta para publicar"
'   - FechaPublicacionEdicionAnterior ? "Fecha publicación edición anterior"
'   - FechaCierre (campo del form):
'       * si la edición tiene FechaPublicacion persistida:
'           ? "Fecha Edición"
'       * si NO tiene FechaPublicacion (primera edición, activa):
'           ? "Fecha publicación anterior"
'   - ColumnaTablaHistorico (columna "Fecha" del histórico de ediciones):
'       * si la edición tiene FechaPublicacion persistida:
'           ? "Fecha Edición" (header) + celda con prefijo "Pub. "
'       * si NO tiene FechaPublicacion:
'           ? "Fecha de inicio" (header) + celda con prefijo "Creac. "
'
' Notas de modelado (validado por usuario 2026-07-08):
'   - Para Calidad, "Fecha Edición" significa "fecha de publicación" (publicar
'     la edición = editarla). La fecha de creación de la edición activa
'     (FechaEdicion system) coincide con la fecha de publicación de la edición
'     anterior (o con la fecha de inicio del proyecto si es la primera edición),
'     por lo que el value que se muestra en el form es el mismo en ambos casos
'     (FechaEdicion system == FechaPublicacionEdicionAnterior).
'   - El refactor es PURO UI: no se tocan nombres de campos VBA (Public
'     Property FechaPublicacion, FechaEdicion), ni tablas/columnas de BD,
'     ni nombres de variables internas. Solo cambian las etiquetas
'     visibles y la selección del source value en el caller.
'
' Notas de modelado — issue #90 (Punto 04 parte 2):
'   - FechaBaseProyecto (= MIN(FechaEdicion) por IDProyecto) es la
'     "primera edición" / "fecha de inicio del proyecto" (validado por
'     usuario 2026-07-09). Se resuelve UNA SOLA VEZ en Form_Load vía
'     ObtenerFechaBaseProyecto y se cachea en m_FechaBaseProyecto del
'     form. NO se usa una constante global del sistema.
'   - Esta función complementa ObtenerEtiquetaFecha (también del
'     Bloque 4 / REQ-CAL-08) — juntas dan el "paquete de fechas" del
'     form: etiqueta visible + constraint cronológica contra el
'     proyecto.
' =============================================================================

' -----------------------------------------------------------------------------
' ObtenerEtiquetaFecha
'
' Resuelve la etiqueta legible para el campo de fecha cuyo nombre canónico
' se pasa como argumento. La etiqueta depende del estado de la edición
' (si tiene FechaPublicacion persistida o no).
'
' Parámetros:
'   p_IDEdicion (String)               - ID de la edición (convención String).
'   p_NombreCampo (String)            - nombre canónico del campo
'                                       (case-insensitive):
'                                         "FECHAPREPARADAPARAPUBLICAR"
'                                         "FECHAPUBLICACIONEDICIONANTERIOR"
'                                         "FECHACIERRE"
'                                         "COLUMNATABLAHISTORICO"
'   db (DAO.Database, opcional)       - inyectado por átomos TDD; usa
'                                       CurrentDb si Nothing.
'   p_Error (out String, opcional)    - mensaje legible si retorna "".
'                                       Vacío si retorna una etiqueta.
'
' Retorna:
'   La etiqueta en español para el campo/estado. "" si p_Error está poblado.
'
' Reglas (en orden):
'   1. p_IDEdicion no vacío            ? si vacío, p_Error poblado, return "".
'   2. p_NombreCampo no vacío          ? si vacío, p_Error poblado, return "".
'   3. Resolver db (inyectado o CurrentDb).
'   4. Leer FechaPublicacion y FechaEdicion de TbProyectosEdiciones filtrado
'      por IDEdicion.
'   5. Switch sobre p_NombreCampo (case-insensitive):
'        - "FECHAPREPARADAPARAPUBLICAR"      ? etiqueta fija
'        - "FECHAPUBLICACIONEDICIONANTERIOR" ? etiqueta fija
'        - "FECHACIERRE"                     ? depende de IsDate(FechaPublicacion)
'        - "COLUMNATABLAHISTORICO"           ? depende de IsDate(FechaPublicacion)
'        - otro                              ? p_Error poblado, return "".
' -----------------------------------------------------------------------------
Public Function ObtenerEtiquetaFecha( _
    ByVal p_IDEdicion As String, _
    ByVal p_NombreCampo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String

    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_SQL As String
    Dim m_FechaPublicacion As Variant
    Dim m_TienePublicacion As Boolean
    Dim m_Nombre As String

    On Error GoTo errores
    p_Error = ""
    ObtenerEtiquetaFecha = ""

    ' --- 1. Validación: p_IDEdicion no vacío ---
    If Len(Trim$(Nz(p_IDEdicion, ""))) = 0 Then
        p_Error = "ObtenerEtiquetaFecha: p_IDEdicion está vacío"
        Exit Function
    End If

    ' --- 2. Validación: p_NombreCampo no vacío ---
    m_Nombre = UCase$(Trim$(Nz(p_NombreCampo, "")))
    If Len(m_Nombre) = 0 Then
        p_Error = "ObtenerEtiquetaFecha: p_NombreCampo está vacío"
        Exit Function
    End If

    ' --- 3. Resolver db ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 4. Leer FechaPublicacion desde TbProyectosEdiciones ---
    m_SQL = "SELECT FechaPublicacion " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion=" & CLng(p_IDEdicion)
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If m_Rs.EOF Then
        m_Rs.Close
        Set m_Rs = Nothing
        Set m_Db = Nothing
        p_Error = "ObtenerEtiquetaFecha: no se encontró la edición con IDEdicion=" & p_IDEdicion
        Exit Function
    End If

    m_FechaPublicacion = m_Rs.fields("FechaPublicacion").value
    m_Rs.Close
    Set m_Rs = Nothing

    ' VBA no hace short-circuit de Or/And (vba-access §1.6.1); separar checks.
    m_TienePublicacion = False
    If Not IsNull(m_FechaPublicacion) Then
        If Len(Trim$(CStr(m_FechaPublicacion))) > 0 Then
            If IsDate(m_FechaPublicacion) Then
                m_TienePublicacion = True
            End If
        End If
    End If

    Set m_Db = Nothing

    ' --- 5. Switch sobre p_NombreCampo ---
    Select Case m_Nombre
        Case "FECHAPREPARADAPARAPUBLICAR"
            ' Etiqueta fija: no depende del estado de la edición. Es la
            ' fecha que Calidad/Técnico propusieron para publicar — siempre
            ' se llama "propuesta" porque puede o no concretarse.
            ObtenerEtiquetaFecha = "Fecha propuesta para publicar"

        Case "FECHAPUBLICACIONEDICIONANTERIOR"
            ' Etiqueta fija: el source value es la fecha de publicación de
            ' la edición anterior (no la edición actual).
            ObtenerEtiquetaFecha = "Fecha publicación edición anterior"

        Case "FECHACIERRE"
            ' Punto 08 / Fase A — label visible per Calidad 2026-07-08:
            '   con FechaPublicacion -> "Fecha Edición" (publicación de ESTA edición)
            '   sin FechaPublicacion  -> "Fecha publicación anterior" (la última
            '     publicación conocida, que == FechaEdicion system == valor
            '     de FechaPublicacionEdicionAnterior para ediciones subsiguientes
            '     o == project start date para la primera edición)
            If m_TienePublicacion Then
                ObtenerEtiquetaFecha = "Fecha Edición"
            Else
                ObtenerEtiquetaFecha = "Fecha publicación anterior"
            End If

        Case "COLUMNATABLAHISTORICO"
            ' Punto 18 / B3 — header del histórico de ediciones.
            ' Spec de Fase A 2026-07-08: con publicación -> "Fecha Edición";
            ' sin publicación -> "Fecha de inicio". El prefijo de celda
            ' ("Pub. " / "Creac. ") se conserva para identificar el tipo de
            ' valor que contiene la celda.
            If m_TienePublicacion Then
                ObtenerEtiquetaFecha = "Fecha Edición"
            Else
                ObtenerEtiquetaFecha = "Fecha de inicio"
            End If

        Case Else
            p_Error = "ObtenerEtiquetaFecha: p_NombreCampo='" & p_NombreCampo & _
                      "' no reconocido (esperado: FECHAPREPARADAPARAPUBLICAR, " & _
                      "FECHAPUBLICACIONEDICIONANTERIOR, FECHACIERRE, COLUMNATABLAHISTORICO)"
    End Select
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ObtenerEtiquetaFecha: " & Err.Number & " - " & Err.description
    End If
    ObtenerEtiquetaFecha = ""
End Function

' -----------------------------------------------------------------------------
' ObtenerFechaBaseProyecto
'
' Resuelve la fecha base del proyecto (= MIN(FechaEdicion) de
' TbProyectosEdiciones filtrado por IDProyecto). Es la constraint (b)
' del issue #90 — la fecha de inicio de una acción no puede ser
' anterior a esta fecha.
'
' NOTA — Decisión 2026-07-09 (per aclaración del usuario):
'   NO se usa una constante global del sistema (ej. "fecha de
'   nacimiento del GR"). La fecha base se resuelve SIEMPRE por query
'   SQL a TbProyectosEdiciones, filtrando por el IDProyecto del
'   riesgo/acción en edición.
'
' Parámetros:
'   p_IDProyecto (String)             - ID del proyecto (convención String).
'   db (DAO.Database, opcional)       - inyectado por átomos TDD; usa
'                                       CurrentDb si Nothing.
'   p_Error (out String, opcional)    - mensaje legible si retorna Null.
'                                       Vacío si retorna una fecha.
'
' Retorna:
'   Variant con la fecha MIN(FechaEdicion) como Date si hay ediciones;
'   Null si no hay ediciones para ese proyecto (o si p_Error está
'   poblado).
'
' Reglas:
'   1. p_IDProyecto no vacío            → si vacío, p_Error poblado, return Null.
'   2. Resolver db (inyectado o CurrentDb).
'   3. Query SQL única: SELECT MIN(FechaEdicion) FROM TbProyectosEdiciones
'      WHERE IDProyecto = p_IDProyecto.
'   4. Cache-first: el caller (form layer) debe invocar este helper UNA
'      sola vez en Form_Load y guardar el resultado en m_FechaBaseProyecto.
'      El helper NO cachea internamente — es responsabilidad del caller.
' -----------------------------------------------------------------------------
Public Function ObtenerFechaBaseProyecto( _
    ByVal p_IDProyecto As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As Variant

    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_SQL As String
    Dim m_FechaMin As Variant

    On Error GoTo errores
    p_Error = ""
    ObtenerFechaBaseProyecto = Null

    ' --- 1. Validación: p_IDProyecto no vacío ---
    If Len(Trim$(Nz(p_IDProyecto, ""))) = 0 Then
        p_Error = "ObtenerFechaBaseProyecto: p_IDProyecto está vacío"
        Exit Function
    End If

    ' --- 2. Resolver db ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 3. Query SQL única: MIN(FechaEdicion) por IDProyecto ---
    m_SQL = "SELECT MIN(FechaEdicion) AS FechaMin " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDProyecto=" & CLng(p_IDProyecto)
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)

    If m_Rs.EOF Then
        ' No hay ediciones para este proyecto. La fecha base es
        ' conceptualmente "indefinida" — el helper NO la inventa.
        m_Rs.Close
        Set m_Rs = Nothing
        Set m_Db = Nothing
        ObtenerFechaBaseProyecto = Null
        Exit Function
    End If

    m_FechaMin = m_Rs.fields("FechaMin").value
    m_Rs.Close
    Set m_Rs = Nothing
    Set m_Db = Nothing

    ' Si MIN(FechaEdicion) es Null (todas las ediciones tienen FechaEdicion
    ' nula), retornamos Null. La constraint (b) del issue #90 NO se puede
    ' aplicar sin este dato — el form layer trata Null como "skip".
    If IsNull(m_FechaMin) Then
        ObtenerFechaBaseProyecto = Null
    Else
        ObtenerFechaBaseProyecto = CDate(m_FechaMin)
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ObtenerFechaBaseProyecto: " & Err.Number & " - " & Err.description
    End If
    ObtenerFechaBaseProyecto = Null
End Function

